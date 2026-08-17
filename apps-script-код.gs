/**
 * Приём ответов от гостей с пригласительного и запись в Google Таблицу
 * + автоматический подсчёт итогов.
 *
 * Куда вставлять: Расширения → Apps Script в вашей Google Таблице.
 * Полная инструкция — в файле ИНСТРУКЦИЯ-google-таблица.md
 */

// Название листа, куда пишутся ответы. Лист создастся сам, если его нет.
var SHEET_NAME = 'Ответы';

// Варианты напитков — по ним считается, сколько чего заказывать.
// Список должен совпадать с галочками в анкете на странице приглашения.
var DRINKS = [
  'Шампанское',
  'Белое вино',
  'Красное вино',
  'Виски',
  'Водка',
  'Джин',
  'Ром',
  'Не пью алкоголь'
];


/* ═══════════════════════════════════════════════════════════
   ПРИЁМ ОТВЕТА ОТ ГОСТЯ
   ═══════════════════════════════════════════════════════════ */
function doPost(e) {
  // Блокировка, чтобы два гостя, ответившие одновременно,
  // не записались в одну и ту же строку
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getSheet();

    sheet.appendRow([
      new Date(),
      data.name || '',
      data.attend || '',
      data.stay || '',
      data.drinks || '',
      data.page || ''
    ]);

    ensureSummary(sheet);

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


/* ═══════════════════════════════════════════════════════════
   НАСТРОИТЬ ИТОГИ

   Запустите эту функцию один раз вручную:
   вверху редактора выберите в списке «настроитьИтоги»
   и нажмите «Выполнить».
   Итоги появятся справа от таблицы и дальше будут
   пересчитываться сами при каждом новом ответе.
   ═══════════════════════════════════════════════════════════ */
function настроитьИтоги() {
  var sheet = getSheet();
  ensureSummary(sheet, true);
  SpreadsheetApp.getActiveSpreadsheet().toast('Итоги настроены', 'Готово', 5);
}


/* ═══════════════════════════════════════════════════════════
   СЛУЖЕБНОЕ
   ═══════════════════════════════════════════════════════════ */

// Возвращает лист с ответами, создавая его при необходимости
function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

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

  return sheet;
}

/**
 * Создаёт блок итогов в колонках H и I.
 *
 * Итоги стоят СПРАВА от таблицы, а не снизу — потому что новые ответы
 * добавляются вниз и постоянно сдвигали бы нижний блок.
 *
 * Внутри — живые формулы, поэтому цифры обновляются сами,
 * даже если вы вручную удалите или добавите строку.
 *
 * force = true — переписать блок заново (для ручного запуска).
 */
function ensureSummary(sheet, force) {
  // если блок уже стоит и не просили пересоздать — ничего не делаем
  if (!force && sheet.getRange('H1').getValue() !== '') return;

  // на случай, если от прошлых запусков остались объединённые ячейки
  try { sheet.getRange(1, 8, 30, 2).breakApart(); } catch (ignore) {}

  var rows = [
    ['ИТОГИ', ''],
    ['Всего ответили', '=COUNTA(B2:B)'],
    ['Придут', '=COUNTIF(C2:C,"Придёт")'],
    ['Не смогут', '=COUNTIF(C2:C,"Не сможет")'],
    ['Остаются ночевать', '=COUNTIF(D2:D,"Останется ночевать")'],
    ['Уедут вечером', '=COUNTIF(D2:D,"Уедет вечером")'],
    ['', ''],
    ['НАПИТКИ', '']
  ];

  // строки по каждому напитку
  for (var i = 0; i < DRINKS.length; i++) {
    rows.push([DRINKS[i], '=COUNTIF(E2:E,"*' + DRINKS[i] + '*")']);
  }

  // записываем: подписи как текст, вторая колонка как формулы
  for (var r = 0; r < rows.length; r++) {
    var label = rows[r][0];
    var formula = rows[r][1];

    sheet.getRange(r + 1, 8).setValue(label);          // колонка H
    if (formula) {
      sheet.getRange(r + 1, 9).setFormula(formula);    // колонка I
    } else {
      sheet.getRange(r + 1, 9).clearContent();
    }
  }

  // оформление
  sheet.setColumnWidth(7, 30);   // пустая колонка-отступ G
  sheet.setColumnWidth(8, 190);  // H
  sheet.setColumnWidth(9, 70);   // I

  // заголовки блоков «ИТОГИ» и «НАПИТКИ»
  sheet.getRange('H1:I1')
    .setFontWeight('bold')
    .setBackground('#EFEBE3');
  sheet.getRange('H8:I8')
    .setFontWeight('bold')
    .setBackground('#EFEBE3');

  // цифры покрупнее и по центру
  sheet.getRange(2, 9, rows.length - 1, 1)
    .setHorizontalAlignment('center')
    .setFontWeight('bold');

  // рамка вокруг всего блока
  sheet.getRange(1, 8, rows.length, 2)
    .setBorder(true, true, true, true, true, true, '#D9D2C5', SpreadsheetApp.BorderStyle.SOLID);
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
