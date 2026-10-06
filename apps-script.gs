/**
 * Google Apps Script для приёма ответов формы RSVP свадебного сайта.
 * Отдельный проект Apps Script (script.google.com), пишет в таблицу по SHEET_ID.
 * Разворачивается как веб-приложение (см. RSVP-SETUP.md).
 *
 * ВАЖНО: после любого изменения этого кода нужно заново развернуть:
 * Развернуть → Управление развертываниями → (карандаш) → Версия: Новая → Развернуть.
 */

// Куда слать уведомление о каждом ответе.
// Оставь пустым ("") — придёт на почту владельца таблицы.
var NOTIFY_EMAIL = "";

// Таблица «Свадьба Португалия · RSVP», куда пишутся ответы.
var SHEET_ID = "1wc4RYuhHH5uCmWnxSNfXkvGd482bCEAmujIhj2b2HEI";

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Дата', 'Имя', 'Статус', 'Дни', 'Напитки', 'Откуда', 'Сообщение', 'Страница']);
      sheet.setFrozenRows(1);
    }

    var data = {};
    try {
      data = JSON.parse(e.postData.contents);
    } catch (err) {
      data = {};
    }

    var list = function (v) { return Array.isArray(v) ? v.join(', ') : (v || ''); };
    var days = list(data.days);
    var drinks = list(data.drinks);
    var when = new Date();

    sheet.appendRow([
      when,                // Дата
      data.name || '',     // Имя
      data.attend || '',   // Статус (Точно буду / Пока не уверен(а) / Не смогу приехать)
      days,                // Дни (16 / 17 / 18 апреля)
      drinks,              // Напитки
      data.stay || '',     // Откуда (Лиссабон и рядом / Порту / Издалека)
      data.msg || '',      // Сообщение
      data.page || ''      // Страница
    ]);

    // Уведомление на почту (не роняем запись, если письмо не ушло)
    try {
      var to = NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
      if (to) {
        var status = data.attend || '—';
        var subject = '🌊 Свадьба в Португалии · ответ: ' + (data.name || 'Гость') + ' — ' + status;
        var body =
          'Новый ответ с сайта-приглашения:\n\n' +
          'Имя: ' + (data.name || '—') + '\n' +
          'Статус: ' + status + '\n' +
          'Дни: ' + (days || '—') + '\n' +
          'Напитки: ' + (drinks || '—') + '\n' +
          'Откуда: ' + (data.stay || '—') + '\n' +
          'Сообщение: ' + (data.msg || '—') + '\n\n' +
          'Время: ' + when.toLocaleString('ru-RU') + '\n' +
          'Страница: ' + (data.page || '—') + '\n';
        MailApp.sendEmail(to, subject, body);
      }
    } catch (mailErr) {
      // игнорируем ошибки почты
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Позволяет открыть URL /exec в браузере для быстрой проверки, что скрипт жив.
function doGet() {
  return ContentService.createTextOutput('RSVP endpoint is running.');
}
