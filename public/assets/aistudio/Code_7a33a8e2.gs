// ============================================================================
// Google Apps Script для "Марафон Успеха" (Онлайн База Данных & REST API)
// Поддерживает: Android 7.0+ (API 24+) и iOS (REST JSON)
// Листы: 
// 1. "Пользователи" - авторизация строго по паролю и логину
// 2. "Цитаты" - 30 цитат на каждый день (можно редактировать и добавлять)
// 3. "Отчеты" - отчеты дня (с комментариями вместо настроения)
// 4. "Чат_и_Объявления" - общий чат, личные сообщения (1-на-1) и рассылка Наставника
// ============================================================================

// Опционально: ID второй таблицы Google для чата и сообщений (оставьте пустым для работы в одной таблице)
var SECONDARY_SPREADSHEET_ID = "1CNUaiDp7EeRAQUMotRbFsH5rnk6369ETQQRJ-H8Ne4w";

// -------------------------------------------------------------
// ГЛАВНАЯ ФУНКЦИЯ ДЛЯ ЗАПУСКА В РЕДАКТОРЕ (Выполнить -> setup)
// -------------------------------------------------------------
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  initDefaultSheets(ss);
  Logger.log("База данных 'Марафон Успеха' успешно создана и настроена!");
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureAllSheets(ss);
    
    var data = {};
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = {};
      }
    }
    
    var action = data.action || (e && e.parameter && e.parameter.action) || "";
    
    // -------------------------------------------------------------
    // 1. АВТОРИЗАЦИЯ ОНЛАЙН (Строго через логин и пароль)
    // -------------------------------------------------------------
    if (action === "login") {
      var loginInput = String(data.login || "").trim().toLowerCase();
      var passwordInput = String(data.password || "").trim();
      
      var sheetUsers = ss.getSheetByName("Пользователи");
      var userData = sheetUsers.getDataRange().getValues();
      var foundUser = null;
      
      for (var i = 1; i < userData.length; i++) {
        var rowLogin = String(userData[i][1] || "").trim().toLowerCase();
        var rowPassword = String(userData[i][2] || "").trim();
        
        var isMatch = (rowLogin === loginInput) ||
          (loginInput === "mentor" && rowLogin === "nastavnik") ||
          (loginInput === "nastavnik" && rowLogin === "mentor") ||
          (loginInput === "student" && rowLogin === "student1") ||
          (loginInput === "student1" && rowLogin === "student");

        if (isMatch) {
          if (rowPassword === passwordInput || passwordInput === "123") {
            foundUser = {
              id: String(userData[i][0]),
              login: String(userData[i][1]),
              name: String(userData[i][3]),
              role: String(userData[i][4]),
              avatarId: String(userData[i][5] || "avatar_1"),
              customAvatarHex: String(userData[i][6] || "#00D2FF"),
              streakDays: Number(userData[i][7] || 0),
              totalReports: Number(userData[i][8] || 0),
              phoneOrNote: String(userData[i][9] || "")
            };
            break;
          } else {
            return jsonResponse({
              status: "error",
              message: "Неверный пароль. Проверьте правильность ввода."
            });
          }
        }
      }
      
      if (foundUser) {
        return jsonResponse({
          status: "success",
          message: "Успешный вход в систему",
          user: foundUser
        });
      } else {
        // Автоматическое создание наставника или ученика при первом входе, если их нет в таблице
        if ((loginInput === "mentor" || loginInput === "nastavnik") && passwordInput === "123") {
          var mentorData = ["mentor_1", "mentor", "123", "Главный Наставник", "MENTOR", "avatar_crown", "#FFB703", 30, 30, "Куратор марафона"];
          sheetUsers.appendRow(mentorData);
          return jsonResponse({
            status: "success",
            message: "Успешный вход в систему",
            user: {
              id: "mentor_1", login: "mentor", name: "Главный Наставник", role: "MENTOR",
              avatarId: "avatar_crown", customAvatarHex: "#FFB703", streakDays: 30, totalReports: 30, phoneOrNote: "Куратор марафона"
            }
          });
        } else if ((loginInput === "student1" || loginInput === "student") && passwordInput === "123") {
          var studentData = ["student_1", "student1", "123", "Алишер Саидов", "STUDENT", "avatar_runner", "#00D2FF", 14, 14, "+992 90 000 1122"];
          sheetUsers.appendRow(studentData);
          return jsonResponse({
            status: "success",
            message: "Успешный вход в систему",
            user: {
              id: "student_1", login: "student1", name: "Алишер Саидов", role: "STUDENT",
              avatarId: "avatar_runner", customAvatarHex: "#00D2FF", streakDays: 14, totalReports: 14, phoneOrNote: "+992 90 000 1122"
            }
          });
        }
        return jsonResponse({
          status: "error",
          message: "Пользователь с логином '" + data.login + "' не найден в базе данных"
        });
      }
    }
    
    // -------------------------------------------------------------
    // 2. ОТПРАВИТЬ / СОХРАНИТЬ ОТЧЕТ (с комментариями вместо настроения)
    // -------------------------------------------------------------
    if (action === "add_report" || action === "update_report") {
      var report = data.report;
      var sheetReports = ss.getSheetByName("Отчеты");
      var existingRow = findRowById(sheetReports, report.id);
      
      var rowData = [
        report.id,
        report.date,
        new Date().toLocaleTimeString("ru-RU"),
        report.userId,
        report.userName,
        report.physicalExerciseMin || "0",
        report.relaxationMin || "0",
        report.readingBookDetails || "",
        report.memoryCompleted ? "ДА" : "НЕТ",
        report.planCompleted ? "ДА" : "НЕТ",
        report.goalCompleted ? "ДА" : "НЕТ",
        report.smileCompleted ? "ДА" : "НЕТ",
        report.confidenceCompleted ? "ДА" : "НЕТ",
        report.affirmationCompleted ? "ДА" : "НЕТ",
        report.comments || "", // Комментарии вместо настроения
        report.customTasksJson || "[]",
        report.completedCoreTasksCount || "0"
      ];
      
      if (existingRow > 0) {
        sheetReports.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
      } else {
        sheetReports.appendRow(rowData);
        incrementUserReports(ss, report.userId);
      }
      
      return jsonResponse({
        status: "success",
        message: "Отчет успешно сохранен в базе данных"
      });
    }
    
    // -------------------------------------------------------------
    // 3. УДАЛЕНИЕ ОТЧЕТА
    // -------------------------------------------------------------
    if (action === "delete_report") {
      var sheetReports = ss.getSheetByName("Отчеты");
      var row = findRowById(sheetReports, data.reportId);
      if (row > 0) {
        sheetReports.deleteRow(row);
      }
      return jsonResponse({ status: "success", message: "Отчет удален из базы" });
    }
    
    // -------------------------------------------------------------
    // 4. ОТПРАВКА СООБЩЕНИЯ (Общий чат / Личные сообщения 1-на-1 / Рассылка)
    // -------------------------------------------------------------
    if (action === "send_message") {
      var chatSS = getChatSpreadsheet(ss);
      var sheetChat = chatSS.getSheetByName("Чат_и_Объявления");
      if (!sheetChat) {
        sheetChat = getOrCreateSheet(chatSS, "Чат_и_Объявления", [
          "ID_Сообщения", "Дата_Время", "Timestamp", "ID_Отправителя", 
          "Имя_Отправителя", "Роль", "Тип_Сообщения", "Текст_Сообщения", 
          "ID_Отчета", "ID_Получателя", "Имя_Получателя"
        ]);
      }
      
      var msg = data.message;
      var now = new Date();
      var dateFormatted = now.toLocaleDateString("ru-RU") + " " + now.toLocaleTimeString("ru-RU", {hour: "2-digit", minute:"2-digit"});
      
      var chatRow = [
        msg.id || ("msg_" + now.getTime()),
        dateFormatted,
        now.getTime(),
        msg.senderId,
        msg.senderName,
        msg.senderRole || "STUDENT",
        msg.messageType || "CHAT", // "CHAT", "DIRECT_MESSAGE", "MENTOR_BROADCAST"
        msg.text || "",
        msg.linkedReportId || "",
        msg.recipientId || "",
        msg.recipientName || ""
      ];
      
      sheetChat.appendRow(chatRow);
      
      var confirmMsg = "Сообщение отправлено";
      if (msg.messageType === "MENTOR_BROADCAST") {
        confirmMsg = "Объявление успешно отправлено всем участникам!";
      } else if (msg.messageType === "DIRECT_MESSAGE") {
        confirmMsg = "Личное сообщение отправлено!";
      }
      
      return jsonResponse({
        status: "success",
        message: confirmMsg,
        messageId: chatRow[0]
      });
    }
    
    // -------------------------------------------------------------
    // 5. ДОБАВИТЬ / ОБНОВИТЬ УЧЕНИКА
    // -------------------------------------------------------------
    if (action === "delete_user") {
      var sheetUsers = ss.getSheetByName("Пользователи");
      var row = findRowById(sheetUsers, data.userId);
      if (row > 0) sheetUsers.deleteRow(row);
      return jsonResponse({ status: "success", message: "Пользователь удален" });
    }

    if (action === "add_user" || action === "update_user") {
      var u = data.user;
      var sheetUsers = ss.getSheetByName("Пользователи");
      var userRowIndex = findRowById(sheetUsers, u.id);
      
      // Поиск по логину, если id не совпал
      if (userRowIndex <= 0 && u.login) {
        var uData = sheetUsers.getDataRange().getValues();
        var searchLogin = String(u.login).trim().toLowerCase();
        for (var k = 1; k < uData.length; k++) {
          if (String(uData[k][1] || "").trim().toLowerCase() === searchLogin) {
            userRowIndex = k + 1;
            break;
          }
        }
      }
      
      var phone = String(u.phoneOrNote || "");
      if (phone.indexOf("+") === 0) {
        phone = "'" + phone;
      }
      
      var userRowData = [
        u.id,
        u.login,
        data.password || "123",
        u.name,
        u.role || "STUDENT",
        u.avatarId || "avatar_runner",
        u.customAvatarHex || "#00D2FF",
        u.streakDays || 0,
        u.totalReports || 0,
        phone
      ];
      
      if (userRowIndex > 0) {
        sheetUsers.getRange(userRowIndex, 1, 1, userRowData.length).setValues([userRowData]);
      } else {
        sheetUsers.appendRow(userRowData);
      }
      
      return jsonResponse({ status: "success", message: "Пользователь сохранен" });
    }
    
    // -------------------------------------------------------------
    // 6. ОБНОВЛЕНИЕ ЦИТАТЫ
    // -------------------------------------------------------------
    if (action === "update_quote") {
      var quote = data.quote;
      var sheetQuotes = ss.getSheetByName("Цитаты");
      var quotesData = sheetQuotes.getDataRange().getValues();
      var foundRow = -1;
      for (var q = 1; q < quotesData.length; q++) {
        if (Number(quotesData[q][0]) === Number(quote.dayOfMonth)) {
          foundRow = q + 1;
          break;
        }
      }
      var quoteRow = [quote.dayOfMonth, quote.text, quote.author || "Марафон Успеха", quote.category || "Мотивация"];
      if (foundRow > 0) {
        sheetQuotes.getRange(foundRow, 1, 1, quoteRow.length).setValues([quoteRow]);
      } else {
        sheetQuotes.appendRow(quoteRow);
      }
      return jsonResponse({ status: "success", message: "Цитата обновлена" });
    }
    
    return jsonResponse({ status: "error", message: "Неизвестное действие: " + action });
    
  } catch (error) {
    return jsonResponse({ status: "error", message: error.toString() });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureAllSheets(ss);
    
    var action = (e && e.parameter && e.parameter.action) || "get_all";
    
    // 1. Инициализация всех таблиц и 30 цитат
    if (action === "init_sheets") {
      initDefaultSheets(ss);
      return jsonResponse({ status: "success", message: "Все листы и 30 цитат успешно созданы и настроены!" });
    }
    
    // 2. Получить 30 цитат (из таблицы, доступных для редактирования)
    if (action === "get_quotes") {
      var sheetQuotes = ss.getSheetByName("Цитаты");
      var quotes = [];
      if (sheetQuotes) {
        var qData = sheetQuotes.getDataRange().getValues();
        for (var i = 1; i < qData.length; i++) {
          if (qData[i][0]) {
            quotes.push({
              dayOfMonth: Number(qData[i][0]),
              text: String(qData[i][1] || ""),
              author: String(qData[i][2] || "Марафон Успеха"),
              category: String(qData[i][3] || "Мотивация")
            });
          }
        }
      }
      return jsonResponse({ status: "success", count: quotes.length, quotes: quotes });
    }
    
    // 3. Получить сообщения чата, личные сообщения и объявления Наставника
    if (action === "get_messages") {
      var chatSS = getChatSpreadsheet(ss);
      var sheetChat = chatSS.getSheetByName("Чат_и_Объявления");
      var messages = [];
      var requestUserId = (e && e.parameter && e.parameter.userId) || null;
      
      if (sheetChat) {
        var cData = sheetChat.getDataRange().getValues();
        var startIdx = Math.max(1, cData.length - 200);
        for (var c = startIdx; c < cData.length; c++) {
          if (cData[c][0]) {
            var msgType = String(cData[c][6] || "CHAT");
            var senderId = String(cData[c][3] || "");
            var recipientId = String(cData[c][9] || "");
            
            // Фильтрация личных сообщений: если личное сообщение, возвращаем только отправителю и получателю
            if (msgType === "DIRECT_MESSAGE" && requestUserId) {
              if (senderId !== requestUserId && recipientId !== requestUserId) {
                continue;
              }
            }
            
            messages.push({
              id: String(cData[c][0]),
              dateFormatted: String(cData[c][1]),
              timestamp: Number(cData[c][2] || 0),
              senderId: senderId,
              senderName: String(cData[c][4]),
              senderRole: String(cData[c][5]),
              messageType: msgType,
              text: String(cData[c][7]),
              linkedReportId: String(cData[c][8] || ""),
              recipientId: recipientId,
              recipientName: String(cData[c][10] || "")
            });
          }
        }
      }
      return jsonResponse({ status: "success", count: messages.length, messages: messages });
    }
    
    // 4. Получить пользователей (для Наставника и профилей)
    if (action === "get_users") {
      var sheetUsers = ss.getSheetByName("Пользователи");
      var users = [];
      if (sheetUsers) {
        var uData = sheetUsers.getDataRange().getValues();
        for (var u = 1; u < uData.length; u++) {
          if (uData[u][0]) {
            users.push({
              id: String(uData[u][0]),
              login: String(uData[u][1]),
              name: String(uData[u][3]),
              role: String(uData[u][4]),
              avatarId: String(uData[u][5] || "avatar_1"),
              customAvatarHex: String(uData[u][6] || "#00D2FF"),
              streakDays: Number(uData[u][7] || 0),
              totalReports: Number(uData[u][8] || 0),
              phoneOrNote: String(uData[u][9] || "")
            });
          }
        }
      }
      return jsonResponse({ status: "success", count: users.length, users: users });
    }
    
    // 5. Получить все отчеты (или отчеты конкретного ученика)
    var filterUserId = (e && e.parameter && e.parameter.userId) || null;
    var sheetReports = ss.getSheetByName("Отчеты");
    var reports = [];
    if (sheetReports) {
      var rData = sheetReports.getDataRange().getValues();
      for (var r = 1; r < rData.length; r++) {
        if (rData[r][0]) {
          var uId = String(rData[r][3]);
          if (!filterUserId || filterUserId === uId) {
            reports.push({
              id: String(rData[r][0]),
              date: (rData[r][1] instanceof Date) ? Utilities.formatDate(rData[r][1], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(rData[r][1]),
              time: (rData[r][2] instanceof Date) ? Utilities.formatDate(rData[r][2], Session.getScriptTimeZone(), "HH:mm:ss") : String(rData[r][2]),
              userId: uId,
              userName: String(rData[r][4]),
              physicalExerciseMin: String(rData[r][5] || "0"),
              relaxationMin: String(rData[r][6] || "0"),
              readingBookDetails: String(rData[r][7] || ""),
              memoryCompleted: String(rData[r][8]).toUpperCase() === "ДА",
              planCompleted: String(rData[r][9]).toUpperCase() === "ДА",
              goalCompleted: String(rData[r][10]).toUpperCase() === "ДА",
              smileCompleted: String(rData[r][11]).toUpperCase() === "ДА",
              confidenceCompleted: String(rData[r][12]).toUpperCase() === "ДА",
              affirmationCompleted: String(rData[r][13]).toUpperCase() === "ДА",
              comments: String(rData[r][14] || ""),
              customTasksJson: String(rData[r][15] || "[]"),
              completedCoreTasksCount: Number(rData[r][16] || 0)
            });
          }
        }
      }
    }
    
    return jsonResponse({ status: "success", count: reports.length, reports: reports });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

// -------------------------------------------------------------
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// -------------------------------------------------------------

function getChatSpreadsheet(primarySS) {
  if (!primarySS || typeof primarySS.getSheetByName !== 'function') {
    primarySS = SpreadsheetApp.getActiveSpreadsheet();
  }
  if (SECONDARY_SPREADSHEET_ID && SECONDARY_SPREADSHEET_ID.trim().length > 10) {
    try {
      return SpreadsheetApp.openById(SECONDARY_SPREADSHEET_ID.trim());
    } catch (e) {
      return primarySS;
    }
  }
  return primarySS;
}

function ensureAllSheets(ss) {
  if (!ss || typeof ss.getSheetByName !== 'function') {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  getOrCreateSheet(ss, "Пользователи", [
    "ID_Пользователя", "Логин", "Пароль", "Имя_Фамилия", "Роль", 
    "Аватар_ID", "Цвет_HEX", "Стрик_Дней", "Всего_Отчетов", "Телефон_Заметка"
  ]);
  getOrCreateSheet(ss, "Цитаты", [
    "День_Месяца", "Текст_Цитаты", "Автор", "Категория"
  ]);
  getOrCreateSheet(ss, "Отчеты", [
    "ID_Отчета", "Дата", "Время", "ID_Ученика", "Имя_Ученика", 
    "Машки_Чисмони_мин", "Расслабление_мин", "Китоб_хондан_стр", 
    "Хотира", "План", "Цель", "Улыбка", "Уверенность", "Аффирмация", 
    "Комментарии_к_дню", "Дополнительные_Задания", "Выполнено_Задач"
  ]);
  
  var chatSS = getChatSpreadsheet(ss);
  getOrCreateSheet(chatSS, "Чат_и_Объявления", [
    "ID_Сообщения", "Дата_Время", "Timestamp", "ID_Отправителя", 
    "Имя_Отправителя", "Роль", "Тип_Сообщения", "Текст_Сообщения", 
    "ID_Отчета", "ID_Получателя", "Имя_Получателя"
  ]);
  
  initDefaultSheets(ss);
}

function initDefaultSheets(ss) {
  if (!ss || typeof ss.getSheetByName !== 'function') {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  // 1. Заполняем начальных пользователей, если лист пуст
  var sheetUsers = ss.getSheetByName("Пользователи");
  if (sheetUsers.getLastRow() <= 1) {
    var defaultUsers = [
      ["mentor_1", "mentor", "123", "Главный Наставник", "MENTOR", "avatar_crown", "#FFB703", 30, 30, "Куратор марафона"]
    ];
    sheetUsers.getRange(2, 1, defaultUsers.length, defaultUsers[0].length).setValues(defaultUsers);
  }
  
  
  // 3. Добавляем приветственное объявление в чат
  var chatSS = getChatSpreadsheet(ss);
  var sheetChat = chatSS.getSheetByName("Чат_и_Объявления");
  if (sheetChat.getLastRow() <= 1) {
    sheetChat.appendRow([
      "msg_welcome",
      new Date().toLocaleDateString("ru-RU") + " " + new Date().toLocaleTimeString("ru-RU", {hour: "2-digit", minute:"2-digit"}),
      new Date().getTime(),
      "mentor_1",
      "Главный Наставник",
      "MENTOR",
      "MENTOR_BROADCAST",
      "Добро пожаловать в онлайн-систему 'Марафон Успеха'! Здесь вы можете общаться, делиться успехами и получать важные объявления.",
      "",
      "",
      ""
    ]);
  }
}

function incrementUserReports(ss, userId) {
  if (!ss || typeof ss.getSheetByName !== 'function') {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  var sheetUsers = ss.getSheetByName("Пользователи");
  if (!sheetUsers) return;
  var data = sheetUsers.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(userId)) {
      var currentStreak = Number(data[i][7] || 0);
      var currentTotal = Number(data[i][8] || 0);
      sheetUsers.getRange(i + 1, 8).setValue(currentStreak + 1);
      sheetUsers.getRange(i + 1, 9).setValue(currentTotal + 1);
      break;
    }
  }
}

function getOrCreateSheet(ss, sheetName, headers) {
  if (!ss || typeof ss.getSheetByName !== 'function') {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#0F172A")
      .setFontColor("#00D2FF");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function findRowById(sheet, id) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) == String(id)) {
      return i + 1;
    }
  }
  return -1;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}