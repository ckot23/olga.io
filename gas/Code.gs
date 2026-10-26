/**
 * Гостевой дом «Ольга» — приём заявок с сайта.
 *
 * doPost():  сохраняет заявку (Имя, Телефон, Источник) в лист «Заявки»
 *            Google Таблицы и отправляет уведомление модератору в Telegram.
 * doGet():   проверка доступности веб-приложения (обычный JSON и JSONP).
 *
 * СЕКРЕТЫ хранятся только в Script Properties (Настройки проекта →
 * Свойства скрипта). Никогда не вставляйте токен бота в index.html:
 *   TELEGRAM_BOT_TOKEN — токен от @BotFather
 *   TELEGRAM_CHAT_ID   — id чата модератора (можно несколько через запятую)
 *   SPREADSHEET_ID     — НЕОБЯЗАТЕЛЬНО, если скрипт не привязан к таблице
 *
 * Инструкция по установке: README.md в корне проекта.
 */

const SHEET_NAME = 'Заявки';
const TIMEZONE = 'Europe/Moscow';
const HEADERS = ['Дата и время (МСК)', 'Имя', 'Телефон', 'Источник', 'Telegram'];

/* ==================== HTTP-точки веб-приложения ==================== */

/**
 * Проверка канала. Сайт дёргает этот адрес через JSONP (?callback=...),
 * потому что Apps Script не отдаёт CORS-заголовки для чтения ответа.
 */
function doGet(e) {
  const info = {
    ok: true,
    service: 'olga-booking',
    sheet: SHEET_NAME,
    time: new Date().toISOString()
  };
  const params = (e && e.parameter) || {};
  const cb = params.callback;
  if (cb && /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(cb)) {
    return ContentService
      .createTextOutput(cb + '(' + JSON.stringify(info) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(info);
}

/** Приём заявки с сайта. Тело — JSON-строка с Content-Type: text/plain. */
function doPost(e) {
  try {
    const d = parseRequest_(e);

    const name = String(d.name || '').trim().slice(0, 80);
    const phone = String(d.phone || '').trim().slice(0, 40);
    const source = String(d.source || 'Сайт').trim().slice(0, 120) || 'Сайт';

    if (name.length < 2) return json_({ ok: false, error: 'Укажите имя' });
    if (phone.replace(/\D/g, '').length < 11) return json_({ ok: false, error: 'Некорректный телефон' });

    // Сначала убеждаемся, что лист существует, затем шлём уведомление и пишем строку:
    // статус отправки попадёт в ту же строку. Если Telegram недоступен —
    // заявка ВСЁ РАВНО сохранится, а в колонке будет причина.
    const sheet = getSheet_();
    let tgStatus = 'Отправлено';
    try {
      const sent = sendTelegram_(name, phone, source);
      if (!sent) tgStatus = 'Не отправлено';
    } catch (err) {
      tgStatus = 'Ошибка: ' + String(err).slice(0, 120);
      console.error('Telegram: ' + err);
    }

    sheet.appendRow([
      Utilities.formatDate(new Date(), TIMEZONE, 'dd.MM.yyyy HH:mm:ss'),
      safeCell_(name),
      safeCell_(phone),
      safeCell_(source),
      tgStatus
    ]);

    return json_({ ok: true, telegram: tgStatus });
  } catch (err) {
    console.error('doPost: ' + err);
    return json_({ ok: false, error: String(err) });
  }
}

/* ==================== Google Таблица ==================== */

/** Возвращает лист «Заявки», создаёт его и шапку при первом обращении. */
function getSheet_() {
  const props = PropertiesService.getScriptProperties();
  const ssId = props.getProperty('SPREADSHEET_ID');
  const ss = ssId ? SpreadsheetApp.openById(ssId) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Таблица не найдена: привяжите скрипт к таблице или задайте SPREADSHEET_ID в свойствах скрипта.');
  }

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  } else if (sheet.getRange(1, 1).getValue() !== HEADERS[0]) {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  // Чтобы телефон вида +7 (905) ... не превратился в формулу.
  sheet.getRange('C:C').setNumberFormat('@');
  return sheet;
}

/** Защита от интерпретации пользовательского ввода как формулы. */
function safeCell_(v) {
  const s = String(v == null ? '' : v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

/* ==================== Telegram ==================== */

/**
 * Шлёт уведомление всем модераторам из TELEGRAM_CHAT_ID.
 * Возвращает false, если секреты не заданы или API ответил ошибкой.
 */
function sendTelegram_(name, phone, source) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('TELEGRAM_BOT_TOKEN');
  const chatId = props.getProperty('TELEGRAM_CHAT_ID');
  if (!token || !chatId) {
    console.warn('TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID не заданы — заявка сохранится без уведомления.');
    return false;
  }

  const text =
    '🏡 Новая заявка — Гостевой дом «Ольга»\n\n' +
    '👤 Имя: ' + name + '\n' +
    '📞 Телефон: ' + phone + '\n' +
    '📥 Источник: ' + source + '\n' +
    '🕒 ' + Utilities.formatDate(new Date(), TIMEZONE, 'dd.MM.yyyy HH:mm') + ' (МСК)';

  const ids = String(chatId).split(',').map(function (s) { return s.trim(); }).filter(Boolean);
  let allOk = ids.length > 0;

  ids.forEach(function (id) {
    const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ chat_id: id, text: text, disable_web_page_preview: true }),
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) {
      allOk = false;
      console.error('Telegram ' + id + ': HTTP ' + res.getResponseCode() + ' ' + res.getContentText());
    }
  });

  return allOk;
}

/* ==================== Служебное ==================== */

function parseRequest_(e) {
  const raw = e && e.postData && e.postData.contents;
  if (raw) {
    try { return JSON.parse(raw) || {}; } catch (err) { /* ниже — form-encoded */ }
  }
  return (e && e.parameter) || {};
}

function getProp_(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ==================== Функции для ручного запуска в редакторе ==================== */

/** Шаг 1: создаёт лист с шапкой и отправляет тестовое сообщение в Telegram. */
function setup() {
  const sheet = getSheet_();
  const ok = sendTelegram_('Проверка связи', '+7 (905) 495-96-12', 'Настройка Apps Script');
  Logger.log('Лист: ' + sheet.getName() + ', строк: ' + sheet.getLastRow() + ', Telegram: ' + (ok ? 'OK' : 'НЕ настроен/ошибка'));
}

/** Помогает узнать chat_id: сначала напишите боту любое сообщение, потом запустите. */
function logChatIds() {
  const token = getProp_('TELEGRAM_BOT_TOKEN');
  if (!token) throw new Error('Сначала задайте TELEGRAM_BOT_TOKEN в свойствах скрипта.');
  const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/getUpdates', { muteHttpExceptions: true });
  const data = JSON.parse(res.getContentText());
  const updates = data.result || [];
  if (!updates.length) {
    Logger.log('Обновлений нет: напишите боту любое сообщение в Telegram и запустите функцию ещё раз.');
    return;
  }
  updates.forEach(function (u) {
    const chat = (u.message && u.message.chat) || {};
    Logger.log('chat_id=' + chat.id + ' (' + (chat.first_name || '') + ' ' + (chat.username || '') + ')');
  });
}

/** Прогоняет полный цикл без сайта: строка в таблице + сообщение в Telegram. */
function testRequest() {
  const out = doPost({
    parameter: {},
    postData: {
      contents: JSON.stringify({
        name: 'Тестовая заявка',
        phone: '+7 (905) 495-96-12',
        source: 'Тест из Apps Script'
      })
    }
  });
  Logger.log(out.getContent());
}
