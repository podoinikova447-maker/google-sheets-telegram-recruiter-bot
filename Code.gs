// ============================================================
//  СКРИПТ АВТОМАТИЗАЦИИ НАПОМИНАНИЙ РЕКРУТЕРАМ
//  Адаптирован под реальную структуру таблицы "Отклики на вакансии"
// ============================================================

// ===================== НАСТРОЙКИ =====================
const SHEET_NAME = 'Вакансии';
const TELEGRAM_BOT_TOKEN = '';
const TELEGRAM_CHAT_ID = '';
const TIMEZONE = Session.getScriptTimeZone() || 'Europe/Moscow';

const REMINDER_TEXT = '🔔 Пора напомнить рекрутеру';
const COLOR_GREEN = '#b6d7a8';
const COLOR_YELLOW = '#ffe599';
const COLOR_NONE = '#ffffff';

// Индексы колонок (0-based для массивов)
const COL = {
  DATE_OTKLIKA: 0,       // A — Дата отклика
  RESURS: 1,             // B — Ресурс
  LINK: 2,               // C — Ссылка на вакансию
  POSITION: 3,           // D — Позиция
  COMPANY: 4,            // E — Компания
  CONTACT: 5,            // F — Контакт
  CV: 6,                 // G — CV
  CL: 7,                 // H — CL (ссылка)
  STATUS: 8,             // I — Статус
  REASON: 9,             // J — Причина отказа
  COMMENT: 10,           // K — Комментарий
  DATE_NAPOMIN: 11,      // L — Дата напоминания
  NAPOMINANIE: 12,       // M — Напоминание
  DOGOVOR: 13,           // N — Тип договора
  SKILLS: 14             // O — Ключевые навыки
};
const TOTAL_COLS = 15;

// ===================== ПОЛЬЗОВАТЕЛЬСКОЕ МЕНЮ =====================
/**
 * Создает кастомное меню в интерфейсе Google Таблицы при открытии файла
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔔 Напоминания')
    .addItem('🚀 Запустить обработку (Авто)', 'runDailyReminder')
    .addSeparator()
    .addItem('🧪 Тест 1: Проверка связи с Telegram', 'testTelegramConnection')
    .addItem('🧪 Тест 2: Отправка реальной строки из таблицы', 'testSendRealReminder')
    .addSeparator()
    .addItem('⚙️ Создать ежедневный триггер (9:00)', 'createDailyTrigger')
    .addToUi();
}

// ===================== РУЧНОЕ ТЕСТИРОВАНИЕ =====================

/**
 * ТЕСТ 1: Отправляет тестовое сообщение для проверки Токена и Chat ID
 */
function testTelegramConnection() {
  const testData = {
    resurs: 'ТЕСТОВЫЙ РЕСУРС',
    dateOtklika: new Date(),
    company: 'Тестовая Компания ООО',
    position: 'Тестовая Позиция',
    dateNapomin: new Date(),
    linkText: 'Текст ссылки',
    linkUrl: 'https://hh.ru/vacancy/12345'
  };

  try {
    sendTelegramMessage(testData);
    SpreadsheetApp.getUi().alert('✅ ТЕСТ 1 Успешен!', 'Сообщение успешно доставлено в Telegram. Токен и ID чата верны.', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Ошибка отправки', 'Проверьте Токен бота и ID чата.\n\nТекст ошибки: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * ТЕСТ 2: Берет первую строку со статусом "Ждем ответа" и отправляет её в TG
 * Проверяет корректность чтения данных, дат и гиперссылок из таблицы.
 */
function testSendRealReminder() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ Ошибка', 'Лист "' + SHEET_NAME + '" не найден.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('❌ Ошибка', 'В таблице нет данных.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, TOTAL_COLS).getValues();
  const richTextLinks = sheet.getRange(2, COL.LINK + 1, lastRow - 1, 1).getRichTextValues();

  // Ищем первую строку со статусом "Ждем ответа"
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][COL.STATUS]).trim() === 'Ждем ответа') {
      const row = values[i];
      const richObj = richTextLinks[i][0];
      const linkUrl = (richObj && typeof richObj.getLinkUrl === 'function') ? richObj.getLinkUrl() : null;

      const testData = {
        resurs: row[COL.RESURS],
        dateOtklika: row[COL.DATE_OTKLIKA],
        company: row[COL.COMPANY],
        position: row[COL.POSITION],
        dateNapomin: new Date(), // Подставляем сегодняшнюю дату для теста
        linkText: row[COL.LINK],
        linkUrl: linkUrl
      };

      try {
        sendTelegramMessage(testData);
        SpreadsheetApp.getUi().alert('✅ ТЕСТ 2 Успешен!', 'Отправлено реальное сообщение для компании:\n\n"' + row[COL.COMPANY] + '"\n\nПроверьте форматирование и ссылки в Telegram.', SpreadsheetApp.getUi().ButtonSet.OK);
      } catch (e) {
        SpreadsheetApp.getUi().alert('❌ Ошибка отправки', 'Ошибка при формировании/отправке реального сообщения:\n\n' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
      }
      return; // Прерываем после первой найденной строки
    }
  }
  
  SpreadsheetApp.getUi().alert('⚠️ Внимание', 'В таблице не найдено ни одной строки со статусом "Ждем ответа".', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ===================== ОСНОВНАЯ ФУНКЦИЯ =====================
function runDailyReminder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('Лист "' + SHEET_NAME + '" не найден. Проверьте SHEET_NAME.');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const numRows = lastRow - 1;

  // --- Пакетное чтение ---
  const dataRange = sheet.getRange(2, 1, numRows, TOTAL_COLS);
  const values = dataRange.getValues();
  const richTextLinks = sheet.getRange(2, COL.LINK + 1, numRows, 1).getRichTextValues();

  const colL_updates = [];  
  const colM_updates = [];  
  const bgUpdates = [];     
  const tgQueue = [];       

  const today = zeroTime(new Date());

  for (let i = 0; i < numRows; i++) {
    const row = values[i];

    const dateOtklikaRaw = row[COL.DATE_OTKLIKA];
    const resurs = row[COL.RESURS];
    const position = row[COL.POSITION];
    const company = row[COL.COMPANY];
    const status = String(row[COL.STATUS]).trim();
    const dogovor = String(row[COL.DOGOVOR]).trim();

    const richObj = richTextLinks[i][0];
    const linkUrl = (richObj && typeof richObj.getLinkUrl === 'function') ? richObj.getLinkUrl() : null;
    const linkText = row[COL.LINK];

    const dateOtklika = parseDate(dateOtklikaRaw);
    let dateNapomin = parseDate(row[COL.DATE_NAPOMIN]);

    if (!dateOtklika) {
      colL_updates.push([row[COL.DATE_NAPOMIN]]); 
      colM_updates.push([row[COL.NAPOMINANIE]]);
      bgUpdates.push(Array(TOTAL_COLS).fill(COLOR_NONE));
      continue;
    }

    if (!dateNapomin) {
      dateNapomin = addWorkingDays(dateOtklika, 5);
    }

    const isToday = dateNapomin.getTime() === today.getTime();
    const isWaiting = (status === 'Ждем ответа');
    const isPast = dateNapomin < today;

    let napominText = '';
    let bgColor = COLOR_NONE;
    let sendToTg = false;

    if (isToday && isWaiting) {
      napominText = REMINDER_TEXT;
      bgColor = dogovor.includes('ТК РФ') ? COLOR_GREEN : COLOR_YELLOW;
      sendToTg = true;
    } else if (isWaiting && !isPast && !isToday) {
      napominText = '';
      bgColor = COLOR_NONE;
    } else {
      napominText = '';
      bgColor = COLOR_NONE;
    }

    colL_updates.push([dateNapomin]);
    colM_updates.push([napominText]);
    bgUpdates.push(Array(TOTAL_COLS).fill(bgColor));

    if (sendToTg) {
      tgQueue.push({
        resurs: resurs,
        dateOtklika: dateOtklikaRaw,
        company: company,
        position: position,
        dateNapomin: dateNapomin,
        linkText: linkText,
        linkUrl: linkUrl
      });
    }
  }

  if (numRows > 0) {
    sheet.getRange(2, COL.DATE_NAPOMIN + 1, numRows, 1).setValues(colL_updates);
    sheet.getRange(2, COL.NAPOMINANIE + 1, numRows, 1).setValues(colM_updates);
    sheet.getRange(2, 1, numRows, TOTAL_COLS).setBackgrounds(bgUpdates);
  }

  for (let i = 0; i < tgQueue.length; i++) {
    try {
      sendTelegramMessage(tgQueue[i]);
      if (i < tgQueue.length - 1) Utilities.sleep(1000); 
    } catch (e) {
      console.error('Ошибка TG для "' + tgQueue[i].company + '": ' + e.message);
    }
  }

  console.log('Готово! Обработано строк: ' + numRows + ', отправлено в TG: ' + tgQueue.length);
}

// ===================== TELEGRAM =====================
function sendTelegramMessage(data) {
  const url = 'https://api.telegram.org/bot' + TELEGRAM_BOT_TOKEN + '/sendMessage';

  const dateOtklikaStr = formatDateSafe(data.dateOtklika);
  const dateNapStr = formatDateSafe(data.dateNapomin);

  const linkHtml = data.linkUrl
    ? '<a href="' + escapeHtml(data.linkUrl) + '">' + escapeHtml(data.linkText || 'Открыть вакансию') + '</a>'
    : escapeHtml(data.linkText || 'Ссылка отсутствует');

  const text =
    '<b>🔔 Напоминание по отклику</b>\n\n' +
    '<b>Ресурс:</b> ' + escapeHtml(data.resurs) + '\n' +
    '<b>Дата отклика:</b> ' + dateOtklikaStr + '\n' +
    '<b>Компания:</b> ' + escapeHtml(data.company) + '\n' +
    '<b>Позиция:</b> ' + escapeHtml(data.position) + '\n' +
    '<b>Дата напоминания:</b> ' + dateNapStr + '\n' +
    '<b>Вакансия:</b> ' + linkHtml;

  const payload = {
    chat_id: TELEGRAM_CHAT_ID,
    text: text,
    parse_mode: 'HTML',
    disable_web_page_preview: true
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const result = JSON.parse(response.getContentText());
  if (!result.ok) {
    throw new Error('Telegram API: ' + result.description);
  }
}

// ===================== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ =====================
function parseDate(value) {
  if (!value) return null;
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return null;
    return zeroTime(value);
  }

  const str = String(value).trim();
  if (!str) return null;

  const dotMatch = str.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (dotMatch) {
    const d = new Date(parseInt(dotMatch[3]), parseInt(dotMatch[2]) - 1, parseInt(dotMatch[1]));
    return isNaN(d.getTime()) ? null : zeroTime(d);
  }

  const slashMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slashMatch) {
    const d = new Date(parseInt(slashMatch[3]), parseInt(slashMatch[1]) - 1, parseInt(slashMatch[2]));
    return isNaN(d.getTime()) ? null : zeroTime(d);
  }

  const fallback = new Date(str);
  return isNaN(fallback.getTime()) ? null : zeroTime(fallback);
}

function zeroTime(date) {
  if (!date) return null;
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  return d;
}

function addWorkingDays(startDate, days) {
  const result = new Date(startDate.getTime());
  let added = 0;
  while (added < days) {
    result.setDate(result.getDate() + 1);
    const dow = result.getDay();
    if (dow !== 0 && dow !== 6) added++;
  }
  return result;
}

function formatDateSafe(value) {
  if (!value) return '—';
  const d = value instanceof Date ? value : parseDate(value);
  if (!d) return String(value);
  return Utilities.formatDate(d, TIMEZONE, 'dd.MM.yyyy');
}

function escapeHtml(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// ===================== ТРИГГЕР =====================
function createDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'runDailyReminder') {
      ScriptApp.deleteTrigger(t);
    }
  }

  ScriptApp.newTrigger('runDailyReminder')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();

  SpreadsheetApp.getUi().alert('✅ Триггер создан!', 'Скрипт будет запускаться ежедневно в 09:00.', SpreadsheetApp.getUi().ButtonSet.OK);
}