/**
 * Гостевой дом «Ольга» — приём заявок с сайта.
 *
 * doPost():  сохраняет заявку (Имя, Телефон, Источник, ID) в лист «Заявки»
 *            Google Таблицы и отправляет уведомление модератору в Telegram.
 * doGet():   диагностика канала, проверка «строка сохранилась?» и резервная
 *            отправка заявки — всё через JSONP (?callback=...), потому что
 *            Apps Script не отдаёт CORS-заголовки.
 *
 * Адрес веб-приложения (…/exec) умеет:
 *   ?action=diag   — состояние канала: доступна ли таблица, настроен ли Telegram
 *   ?action=check&id=ID — сохранилась ли заявка с таким ID (реальное подтверждение)
 *   ?action=book&name=…&phone=…&source=…&id=ID — резервная отправка через GET
 *   без параметров — то же, что diag
 *
 * СЕКРЕТЫ хранятся только в Script Properties (Настройки проекта →
 * Свойства скрипта). Никогда не вставляйте токен бота в index.html:
 *   TELEGRAM_BOT_TOKEN — токен от @BotFather
 *   TELEGRAM_CHAT_ID   — id чата модератора (можно несколько через запятую)
 *   SPREADSHEET_ID     — НЕОБЯЗАТЕЛЬНО, если скрипт не привязан к таблице
 *
 * ВАЖНО: после любой правки кода нужно Развернуть → Управление развёртываниями
 * → изменить → Новая версия → Развернуть. Иначе /exec работает по-старому.
 *
 * Инструкция по установке: README.md в корне проекта.
 */

const SHEET_NAME = 'Заявки';
const TIMEZONE = 'Europe/Moscow';
const HEADERS = ['Дата и время (МСК)', 'Имя', 'Телефон', 'Источник', 'Telegram', 'ID заявки'];
const COL_TG = 5;              // E — статус Telegram
const COL_ID = 6;              // F — ID заявки (нужен сайту для подтверждения)

/* ==================== HTTP-точки веб-приложения ==================== */

/** Приём заявки с сайта. Тело — JSON-строка с Content-Type: text/plain. */
function doPost(e) {
  let data = {};
  try {
    data = parseRequest_(e);
  } catch (err) {
    console.error('doPost/parse: ' + err);
  }
  return respond_(handleBooking_(data));
}

/**
 * GET-точка. Используется сайтом через JSONP: <script src="…/exec?callback=cb">.
 * Здесь же диагностика, проверка сохранения и резервная отправка.
 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = String(p.action || 'ping').toLowerCase();
  let out;
  try {
    if (action === 'check') {
      out = checkRequest_(p.id);
    } else if (action === 'book') {
      out = handleBooking_({
        name: p.name,
        phone: p.phone,
        source: p.source,
        clientId: p.id || p.clientId
      });
    } else {
      out = diagInfo_(); // ping / diag / любой другой GET
    }
  } catch (err) {
    out = { ok: false, saved: false, code: err.code || 'ERROR', error: errMessage_(err) };
  }
  return respond_(out, p.callback);
}

/* ==================== Заявка ==================== */

/**
 * Общая логика для doPost и doGet?action=book.
 * Сначала пишет строку в таблицу, потом шлёт Telegram и обновляет статус в E:
 * так заявка не теряется, даже если Telegram недоступен.
 * Если clientId уже есть в таблице — дубль не создаётся.
 */
function handleBooking_(d) {
  d = d || {};

  const name = String(d.name || '').trim().slice(0, 80);
  const phone = String(d.phone || '').trim().slice(0, 40);
  const source = String(d.source || 'Сайт').trim().slice(0, 120) || 'Сайт';
  const clientId = String(d.clientId || d.id || '')
    .replace(/[^A-Za-z0-9._-]/g, '')
    .slice(0, 64);

  if (name.length < 2) {
    return { ok: false, saved: false, code: 'NAME', error: 'Укажите имя' };
  }
  if (phone.replace(/\D/g, '').length < 11) {
    return { ok: false, saved: false, code: 'PHONE', error: 'Некорректный телефон' };
  }

  let sheet;
  try {
    sheet = getSheet_();
  } catch (err) {
    console.error('getSheet_: ' + err);
    return {
      ok: false,
      saved: false,
      code: err.code || 'NO_SHEET',
      error: errMessage_(err),
      message: 'Таблица недоступна — заявка НЕ сохранена.'
    };
  }

  // Пишем строку под блокировкой, чтобы параллельные заявки не задублировались.
  const lock = LockService.getScriptLock();
  let row = 0;
  try {
    lock.waitLock(15000);
  } catch (err) {
    return { ok: false, saved: false, code: 'LOCK', error: 'Скрипт занят, попробуйте ещё раз.' };
  }

  try {
    if (clientId) {
      const existing = findRowById_(sheet, clientId);
      if (existing) {
        return { ok: true, saved: true, duplicate: true, id: clientId, telegram: existing.telegram };
      }
    }
    sheet.appendRow([
      now_(),
      safeCell_(name),
      safeCell_(phone),
      safeCell_(source),
      'Отправляется…',
      safeCell_(clientId)
    ]);
    row = sheet.getLastRow();
  } catch (err) {
    console.error('appendRow: ' + err);
    return {
      ok: false,
      saved: false,
      code: 'WRITE',
      error: errMessage_(err),
      message: 'Не удалось записать строку в таблицу.'
    };
  } finally {
    try { lock.releaseLock(); } catch (err) { /* уже отпущена */ }
  }

  let tgStatus = 'Отправлено';
  try {
    if (!sendTelegram_(name, phone, source)) tgStatus = 'Не отправлено';
  } catch (err) {
    tgStatus = 'Ошибка: ' + errMessage_(err).slice(0, 100);
    console.error('Telegram: ' + err);
  }

  try {
    sheet.getRange(row, COL_TG).setValue(tgStatus);
  } catch (err) {
    console.error('status cell: ' + err);
  }

  return { ok: true, saved: true, id: clientId || null, row: row, telegram: tgStatus };
}

/** Сохранилась ли заявка с таким ID. Ответ читается сайтом через JSONP. */
function checkRequest_(id) {
  const cleanId = String(id || '').replace(/[^A-Za-z0-9._-]/g, '').slice(0, 64);
  if (!cleanId) return { ok: false, saved: false, code: 'NO_ID', error: 'Не указан ID заявки' };
  try {
    const found = findRowById_(getSheet_(), cleanId);
    if (!found) return { ok: true, saved: false, id: cleanId };
    return { ok: true, saved: true, id: cleanId, row: found.row, telegram: found.telegram };
  } catch (err) {
    console.error('check: ' + err);
    return {
      ok: false,
      saved: false,
      code: err.code || 'CHECK',
      error: errMessage_(err),
      message: 'Таблица недоступна.'
    };
  }
}

/* ==================== Диагностика ==================== */

/**
 * Состояние канала. Сайт показывает его в индикаторе у формы, владелец может
 * открыть в браузере: <…/exec?action=diag>
 */
function diagInfo_() {
  const info = {
    ok: true,
    saved: true,
    service: 'olga-booking',
    sheet: SHEET_NAME,
    time: new Date().toISOString(),
    sheetOk: false,
    telegramConfigured: false,
    spreadsheet: null,
    code: null,
    message: null
  };

  const props = PropertiesService.getScriptProperties();
  info.telegramConfigured = !!(
    props.getProperty('TELEGRAM_BOT_TOKEN') && props.getProperty('TELEGRAM_CHAT_ID')
  );

  try {
    const ss = getSpreadsheet_();
    info.spreadsheet = ss.getName();
    const sheet = getSheet_();
    info.sheetOk = true;
    info.rows = Math.max(0, sheet.getLastRow() - 1);
  } catch (err) {
    info.ok = false;
    info.saved = false;
    info.code = err.code || 'ERROR';
    info.message = errMessage_(err);
    console.error('diag: ' + err);
  }

  if (info.sheetOk && !info.telegramConfigured) {
    info.message = 'Таблица работает, Telegram не настроен (строки будут без уведомлений).';
  }

  return info;
}

/* ==================== Google Таблица ==================== */

/** Открывает таблицу: привязанную к скрипту или заданную в SPREADSHEET_ID. */
function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const ssId = String(props.getProperty('SPREADSHEET_ID') || '').trim();
  let ss = null;

  try {
    ss = ssId ? SpreadsheetApp.openById(ssId) : SpreadsheetApp.getActiveSpreadsheet();
  } catch (err) {
    throw appError_(
      'NO_SPREADSHEET',
      'Не удалось открыть таблицу по SPREADSHEET_ID — проверьте id и доступ. ' + errMessage_(err)
    );
  }

  if (!ss) {
    throw appError_(
      'NO_SPREADSHEET',
      'Скрипт не привязан к таблице. Создайте скрипт через «Расширения → Apps Script» ' +
      'в самой таблице или задайте SPREADSHEET_ID в свойствах скрипта.'
    );
  }
  return ss;
}

/** Возвращает лист «Заявки», создаёт его и шапку при первом обращении. */
function getSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  } else if (String(sheet.getRange(1, 1).getValue()).trim() !== HEADERS[0]) {
    // Первая строка — данные, а не шапка (например, лист создали вручную).
    sheet.insertRowBefore(1);
  }

  // Приводим шапку к актуальному виду (в т.ч. добавляем колонку «ID заявки»).
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  // Чтобы телефон вида +7 (905) ... не превратился в формулу.
  sheet.getRange('C:C').setNumberFormat('@');
  return sheet;
}

/** Ищет строку по ID заявки (колонка F). Возвращает {row, telegram} или null. */
function findRowById_(sheet, id) {
  if (!id) return null;
  const last = sheet.getLastRow();
  if (last < 2) return null;

  // Смотрим последние 1000 строк — этого достаточно и не тормозит на больших листах.
  const start = Math.max(2, last - 1000 + 1);
  const values = sheet.getRange(start, COL_ID, last - start + 1, 1).getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0]).trim() === id) {
      const row = start + i;
      let tg = '';
      try { tg = String(sheet.getRange(row, COL_TG).getValue() || ''); } catch (err) { /* не критично */ }
      return { row: row, telegram: tg };
    }
  }
  return null;
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

function now_() {
  return Utilities.formatDate(new Date(), TIMEZONE, 'dd.MM.yyyy HH:mm:ss');
}

function errMessage_(err) {
  return String((err && (err.message || err)) || 'Неизвестная ошибка').slice(0, 300);
}

/** Ошибка с машинным кодом — сайт показывает её причину. */
function appError_(code, message) {
  const err = new Error(message);
  err.code = code;
  return err;
}

function respond_(obj, cb) {
  const callback = cb ? String(cb) : '';
  if (callback && /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(obj) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
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

/** Показывает, куда именно пишутся заявки и что с Telegram. */
function diagnose() {
  const info = diagInfo_();
  Logger.log(JSON.stringify(info, null, 2));
  if (!info.sheetOk) {
    Logger.log('ПРИЧИНА: ' + info.code + ' — ' + info.message);
  } else {
    const ss = getSpreadsheet_();
    Logger.log('Таблица: «' + ss.getName() + '», лист «' + SHEET_NAME + '», заявок: ' + info.rows);
    Logger.log('Ссылка: ' + ss.getUrl());
  }
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
  const out = handleBooking_({
    name: 'Тестовая заявка',
    phone: '+7 (905) 495-96-12',
    source: 'Тест из Apps Script',
    clientId: 'test-' + Date.now()
  });
  Logger.log(JSON.stringify(out));
}
