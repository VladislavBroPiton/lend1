/**
 * Приём ответов от гостей с пригласительного и запись в Google Таблицу.
 *
 * Куда вставлять: Расширения → Apps Script в вашей Google Таблице.
 * Полная инструкция — в файле ИНСТРУКЦИЯ-google-таблица.md
 */

// Название листа, куда пишутся ответы. Лист создастся сам, если его нет.
var SHEET_NAME = 'Ответы';

function doPost(e) {
  // Блокировка, чтобы два гостя, ответившие одновременно,
  // не записались в одну и ту же строку
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var data = JSON.parse(e.postData.contents);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);

    // если листа ещё нет — создаём его вместе с шапкой
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Когда ответил',
        'Имя и фамилия',
        'Придёт',
        'Ночёвка',
        'Напитки',
        'Приглашение'
      ]);
      var head = sheet.getRange(1, 1, 1, 6);
      head.setFontWeight('bold');
      head.setBackground('#EFEBE3');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 150);
      sheet.setColumnWidth(2, 200);
      sheet.setColumnWidth(5, 260);
      sheet.setColumnWidth(6, 180);
    }

    sheet.appendRow([
      new Date(),
      data.name || '',
      data.attend || '',
      data.stay || '',
      data.drinks || '',
      data.page || ''
    ]);

    return json({ result: 'ok' });

  } catch (err) {
    return json({ result: 'error', message: String(err) });

  } finally {
    lock.releaseLock();
  }
}

// Открытие адреса скрипта в браузере — просто проверка, что он живой
function doGet() {
  return json({ result: 'ok', message: 'Скрипт работает. Ответы принимаются.' });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
