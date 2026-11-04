import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import { createServer as createViteServer } from 'vite';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const WEBHOOK_URL =
  process.env.GAS_WEBHOOK_URL ||
  'https://script.google.com/macros/s/AKfycbxH_ahjT5hIPL9QwYAM5m1XKSy25nAkq44A7uxB6BpeojePRXVjGdln-qFHtRqgXErK/exec';

interface CacheEntry<T = unknown> {
  data: T;
  timestamp: number;
  latencyMs: number;
}

const memoryCache = new Map<string, CacheEntry>();
const inFlightRequests = new Map<string, Promise<{ data: any; latencyMs: number }>>();

// Cache TTL: 15 seconds fresh, 5 minutes stale-while-revalidate
const FRESH_TTL_MS = 15_000;
const MAX_STALE_TTL_MS = 300_000;

async function fetchFromGasGet(action: string, params: Record<string, string> = {}): Promise<{ data: any; latencyMs: number }> {
  const query = new URLSearchParams({ action, ...params }).toString();
  const cacheKey = `GET:${query}`;

  // Deduplicate concurrent identical requests
  const existingFlight = inFlightRequests.get(cacheKey);
  if (existingFlight) {
    return existingFlight;
  }

  const flightPromise = (async () => {
    const start = Date.now();
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 25_000);
    try {
      const url = `${WEBHOOK_URL}?${query}`;
      const res = await fetch(url, {
        method: 'GET',
        redirect: 'follow',
        signal: controller.signal,
        headers: {
          Accept: 'application/json',
        },
      });
      const text = await res.text();
      let parsed: any;
      try {
        parsed = JSON.parse(text);
      } catch {
        throw new Error(`Invalid JSON from GAS (${res.status}): ${text.slice(0, 180)}`);
      }
      const latencyMs = Date.now() - start;
      if (parsed && parsed.status !== 'error') {
        memoryCache.set(cacheKey, {
          data: parsed,
          timestamp: Date.now(),
          latencyMs,
        });
      }
      return { data: parsed, latencyMs };
    } finally {
      clearTimeout(timeout);
      inFlightRequests.delete(cacheKey);
    }
  })();

  inFlightRequests.set(cacheKey, flightPromise);
  return flightPromise;
}

async function getWithSwr(action: string, params: Record<string, string> = {}, forceRefresh = false) {
  const query = new URLSearchParams({ action, ...params }).toString();
  const cacheKey = `GET:${query}`;
  const cached = memoryCache.get(cacheKey);
  const now = Date.now();

  if (!forceRefresh && cached) {
    const age = now - cached.timestamp;
    if (age <= FRESH_TTL_MS) {
      return {
        data: cached.data,
        cacheStatus: 'HIT',
        ageMs: age,
        upstreamLatencyMs: cached.latencyMs,
      };
    }
    if (age <= MAX_STALE_TTL_MS) {
      // Trigger background revalidation without blocking client response
      fetchFromGasGet(action, params).catch((err) => {
        console.error(`[SWR Background Error] ${action}:`, err.message);
      });
      return {
        data: cached.data,
        cacheStatus: 'STALE',
        ageMs: age,
        upstreamLatencyMs: cached.latencyMs,
      };
    }
  }

  try {
    const { data, latencyMs } = await fetchFromGasGet(action, params);
    return {
      data,
      cacheStatus: 'MISS',
      ageMs: 0,
      upstreamLatencyMs: latencyMs,
    };
  } catch (err: any) {
    // Fallback to expired cache if upstream fails
    if (cached) {
      return {
        data: cached.data,
        cacheStatus: 'FALLBACK',
        ageMs: now - cached.timestamp,
        upstreamLatencyMs: cached.latencyMs,
      };
    }
    throw err;
  }
}

function invalidateCache(prefixes?: string[]) {
  if (!prefixes || prefixes.length === 0) {
    memoryCache.clear();
    return;
  }
  for (const key of memoryCache.keys()) {
    if (prefixes.some((p) => key.includes(p))) {
      memoryCache.delete(key);
    }
  }
}

async function startServer() {
  const app = express();
  const PORT = Number(process.env.PORT) || 3000;

  app.use(express.json({ limit: '10mb' }));
  app.use(express.text({ limit: '10mb' }));

  // 1. Combined Ultra-Fast Sync Endpoint (/api/sync)
  app.get('/api/sync', async (req, res) => {
    const start = Date.now();
    const userId = String(req.query.userId || '');
    const force = req.query.force === 'true';

    try {
      const [usersResult, reportsResult, quotesResult, messagesResult] = await Promise.all([
        getWithSwr('get_users', {}, force),
        getWithSwr('get_reports', {}, force),
        getWithSwr('get_quotes', {}, force),
        getWithSwr('get_messages', userId ? { userId } : {}, force),
      ]);

      const serverTimeMs = Date.now() - start;
      const maxUpstreamMs = Math.max(
        usersResult.upstreamLatencyMs,
        reportsResult.upstreamLatencyMs,
        quotesResult.upstreamLatencyMs,
        messagesResult.upstreamLatencyMs
      );

      res.json({
        status: 'success',
        users: usersResult.data?.users || [],
        reports: reportsResult.data?.reports || [],
        quotes: quotesResult.data?.quotes || [],
        messages: messagesResult.data?.messages || [],
        diagnostics: {
          serverResponseMs: serverTimeMs,
          upstreamLatencyMs: maxUpstreamMs,
          cacheStates: {
            users: usersResult.cacheStatus,
            reports: reportsResult.cacheStatus,
            quotes: quotesResult.cacheStatus,
            messages: messagesResult.cacheStatus,
          },
          syncedAt: new Date().toISOString(),
          webhookUrl: WEBHOOK_URL,
        },
      });
    } catch (error: any) {
      console.error('[/api/sync Error]:', error);
      res.status(502).json({
        status: 'error',
        message: error?.message || 'Ошибка синхронизации с Google Sheets',
      });
    }
  });

  // 2. Individual GET proxy (/api/gas?action=...)
  app.get('/api/gas', async (req, res) => {
    const action = String(req.query.action || 'get_all');
    const force = req.query.force === 'true';
    const params: Record<string, string> = {};
    for (const [k, v] of Object.entries(req.query)) {
      if (k !== 'action' && k !== 'force' && typeof v === 'string') {
        params[k] = v;
      }
    }
    try {
      const result = await getWithSwr(action, params, force);
      res.json(result.data);
    } catch (error: any) {
      res.status(502).json({ status: 'error', message: error?.message || 'GAS GET failed' });
    }
  });

  // 3. POST Proxy (/api/gas) with automatic cache invalidation
  app.post('/api/gas', async (req, res) => {
    const start = Date.now();
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 25_000);

    try {
      const bodyObj = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
      const action = bodyObj?.action || '';

      const upstreamRes = await fetch(WEBHOOK_URL, {
        method: 'POST',
        redirect: 'follow',
        signal: controller.signal,
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(bodyObj),
      });

      const text = await upstreamRes.text();
      let parsed: any;
      try {
        parsed = JSON.parse(text);
      } catch {
        parsed = { status: 'error', message: `Ошибка ответа сервера: ${text.slice(0, 120)}` };
      }

      // Immediately invalidate relevant caches on mutation so next read is 100% fresh
      if (action.includes('report')) {
        invalidateCache(['action=get_reports', 'action=get_users']);
      } else if (action.includes('user')) {
        invalidateCache(['action=get_users']);
      } else if (action.includes('quote')) {
        invalidateCache(['action=get_quotes']);
      } else if (action.includes('message')) {
        invalidateCache(['action=get_messages']);
      } else if (action !== 'login') {
        invalidateCache();
      }

      res.json({
        ...parsed,
        _latencyMs: Date.now() - start,
      });
    } catch (error: any) {
      console.error('[/api/gas POST Error]:', error);
      res.status(502).json({
        status: 'error',
        message: error?.message || 'Ошибка соединения с сервером Google Apps Script',
      });
    } finally {
      clearTimeout(timeout);
    }
  });

  // 4. Server Health & Audit Check Endpoint
  app.get('/api/health', async (_req, res) => {
    const start = Date.now();
    try {
      const { data, latencyMs } = await fetchFromGasGet('get_users');
      res.json({
        status: 'ok',
        gasReachable: data?.status === 'success',
        gasLatencyMs: latencyMs,
        totalCheckMs: Date.now() - start,
        usersCount: Array.isArray(data?.users) ? data.users.length : 0,
        cacheKeysCount: memoryCache.size,
        webhookUrl: WEBHOOK_URL,
      });
    } catch (err: any) {
      res.status(500).json({
        status: 'error',
        gasReachable: false,
        message: err?.message || 'Unreachable',
      });
    }
  });

  // Warm up cache on server start in background
  Promise.all([
    fetchFromGasGet('get_users'),
    fetchFromGasGet('get_reports'),
    fetchFromGasGet('get_quotes'),
  ]).catch(() => {
    // Ignore warmup errors
  });

  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(__dirname, 'dist');
    app.use(express.static(distPath));
    app.get('*', (_req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://0.0.0.0:${PORT}`);
  });
}

startServer();
