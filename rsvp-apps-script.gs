/**
 * RSVP-приёмник для свадебных приглашений.
 *
 * Что делает:
 *   1) принимает ответы гостей с формы приглашения и пишет их строкой в Google Таблицу;
 *   2) по паролю отдаёт весь список ответов для страницы guests.html.
 *
 * Как подключить — см. файл RSVP-НАСТРОЙКА.md рядом с приглашениями.
 */

// ─────────── НАСТРОЙКИ (поменяйте на свои) ───────────

// Пароль для входа на страницу со списком гостей.
// Придумайте свой и не используйте простые вроде 1234.
var ADMIN_PASSWORD = 'СМЕНИТЕ_ЭТОТ_ПАРОЛЬ';

// Название листа внутри таблицы, куда писать ответы.
var SHEET_NAME = 'Ответы';

// ─────────────────────────────────────────────────────

var HEADERS = ['Дата ответа', 'Имя', 'Придёт', 'Гостей', 'Меню', 'Комментарий', 'Приглашение'];

/**
 * Единственная точка входа. Форма и админ-страница шлют сюда POST
 * с телом в виде JSON-строки (Content-Type: text/plain — так браузер
 * не делает предварительный CORS-запрос, который Apps Script не умеет обрабатывать).
 */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    if (data.action === 'submit') return json(handleSubmit(data));
    if (data.action === 'list') return json(handleList(data));

    return json({ ok: false, error: 'Неизвестное действие' });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/** Запись нового ответа гостя. */
function handleSubmit(data) {
  var name = String(data.name || '').trim();
  if (!name) return { ok: false, error: 'Не указано имя' };

  var sheet = getSheet();
  sheet.appendRow([
    new Date(),
    name,
    data.coming === 'yes' ? 'Да' : 'Нет',
    data.coming === 'yes' ? (Number(data.guests) || 1) : 0,
    String(data.menu || '').trim(),
    String(data.comment || '').trim(),
    String(data.style || '').trim()
  ]);

  return { ok: true };
}

/** Отдача всего списка ответов — только по правильному паролю. */
function handleList(data) {
  if (String(data.password || '') !== ADMIN_PASSWORD) {
    return { ok: false, error: 'Неверный пароль' };
  }

  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, rows: [], totals: emptyTotals() };

  var values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  var rows = [];
  var comingPeople = 0, comingAnswers = 0, declined = 0;

  for (var i = 0; i < values.length; i++) {
    var v = values[i];
    var isComing = v[2] === 'Да';
    var guests = Number(v[3]) || 0;

    if (isComing) { comingAnswers++; comingPeople += guests; } else { declined++; }

    rows.push({
      date: v[0] instanceof Date ? v[0].toISOString() : String(v[0]),
      name: String(v[1]),
      coming: isComing,
      guests: guests,
      menu: String(v[4]),
      comment: String(v[5]),
      style: String(v[6])
    });
  }

  return {
    ok: true,
    rows: rows,
    totals: {
      comingPeople: comingPeople,
      comingAnswers: comingAnswers,
      declined: declined,
      total: rows.length
    }
  };
}

function emptyTotals() {
  return { comingPeople: 0, comingAnswers: 0, declined: 0, total: 0 };
}

/** Лист с ответами; создаётся автоматически вместе с шапкой. */
function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
