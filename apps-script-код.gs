/**
 * Приём ответов от гостей с пригласительного, запись в Google Таблицу,
 * подсчёт итогов, оформление и диаграммы.
 *
 * Куда вставлять: Расширения → Apps Script в вашей Google Таблице.
 * Полная инструкция — в файле ИНСТРУКЦИЯ-google-таблица.md
 */

// Название листа, куда пишутся ответы. Лист создастся сам, если его нет.
var SHEET_NAME = 'Ответы';

// Варианты напитков — по ним считается, сколько чего закупать.
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

// Палитра — та же, что на самом приглашении
var C_INK    = '#2B2B2B';
var C_HEAD   = '#EFEBE3';
var C_LINE   = '#D9D2C5';
var C_GREEN  = '#8FA8A6';
var C_ROSE   = '#C98B8B';
var C_SAND   = '#E5C79A';
var C_STONE  = '#B9B0A2';


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

    // Пишем в первую свободную строку ПОД данными.
    // Через appendRow нельзя: блок итогов в колонках H и I
    // «удлиняет» лист, и ответы улетали бы ниже него.
    var row = getLastDataRow(sheet) + 1;
    sheet.getRange(row, 1, 1, 6).setValues([[
      new Date(),
      data.name || '',
      data.attend || '',
      data.stay || '',
      data.drinks || '',
      data.page || ''
    ]]);

    refresh(sheet, false);

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
   РУЧНОЕ ОБНОВЛЕНИЕ

   Меню «Свадьба» вверху таблицы, либо запуск функции
   «настроитьИтоги» из редактора скрипта.
   ═══════════════════════════════════════════════════════════ */
function настроитьИтоги() {
  refresh(getSheet(), true);
  SpreadsheetApp.getActiveSpreadsheet().toast('Итоги, оформление и диаграммы обновлены', 'Готово', 5);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Свадьба')
    .addItem('Обновить итоги и диаграммы', 'настроитьИтоги')
    .addToUi();
}

// force = true — убрать пустые строки, пересоздать оформление и диаграммы
function refresh(sheet, force) {
  if (force) compactRows(sheet);
  writeSummary(sheet);
  styleData(sheet);
  ensureCharts(sheet, force);
}

/**
 * Последняя строка с ответом — считается по колонке A (дата),
 * а не по всему листу, чтобы блок итогов справа не влиял.
 */
function getLastDataRow(sheet) {
  var values = sheet.getRange('A1:A').getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0]).trim() !== '') return i + 1;
  }
  return 1;   // только шапка
}

/**
 * Убирает пустые строки между ответами.
 * Нужно, если записи уже разъехались из-за прежней версии скрипта.
 */
function compactRows(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return;

  var data = sheet.getRange(2, 1, last - 1, 6).getValues();
  var kept = [];

  for (var i = 0; i < data.length; i++) {
    var hasDate = String(data[i][0]).trim() !== '';
    var hasName = String(data[i][1]).trim() !== '';
    if (hasDate || hasName) kept.push(data[i]);
  }

  sheet.getRange(2, 1, last - 1, 6).clearContent();
  if (kept.length > 0) {
    sheet.getRange(2, 1, kept.length, 6).setValues(kept);
  }
}


/* ═══════════════════════════════════════════════════════════
   ЛИСТ И ШАПКА
   ═══════════════════════════════════════════════════════════ */
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
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(5, 260);
    sheet.setColumnWidth(6, 180);
  }

  return sheet;
}


/* ═══════════════════════════════════════════════════════════
   ИТОГИ (готовыми числами — не зависят от языка таблицы)
   ═══════════════════════════════════════════════════════════ */
function writeSummary(sheet) {
  var last = getLastDataRow(sheet);

  var total = 0, come = 0, cant = 0, stay = 0, leave = 0;
  var drinkCount = {};
  for (var d = 0; d < DRINKS.length; d++) drinkCount[DRINKS[d]] = 0;

  if (last > 1) {
    // читаем колонки B..E (имя, придёт, ночёвка, напитки)
    var rows = sheet.getRange(2, 2, last - 1, 4).getValues();

    for (var i = 0; i < rows.length; i++) {
      var name = String(rows[i][0]).trim();
      if (!name) continue;               // пустые строки не считаем

      total++;

      var attend = String(rows[i][1]).trim();
      var isComing = (attend === 'Придёт');
      if (isComing) come++;
      else if (attend === 'Не сможет') cant++;

      var night = String(rows[i][2]).trim();
      if (night === 'Останется ночевать') stay++;
      else if (night === 'Уедет вечером') leave++;

      // напитки считаем только у тех, кто придёт
      if (isComing) {
        var drinks = String(rows[i][3]);
        for (var k = 0; k < DRINKS.length; k++) {
          if (drinks.indexOf(DRINKS[k]) !== -1) drinkCount[DRINKS[k]]++;
        }
      }
    }
  }

  var block = [
    ['ИТОГИ', ''],
    ['Всего ответили', total],
    ['Придут', come],
    ['Не смогут', cant],
    ['Остаются ночевать', stay],
    ['Уедут вечером', leave],
    ['', ''],
    ['НАПИТКИ (среди тех, кто придёт)', '']
  ];
  for (var n = 0; n < DRINKS.length; n++) {
    block.push([DRINKS[n], drinkCount[DRINKS[n]]]);
  }

  try { sheet.getRange(1, 8, 40, 2).breakApart(); } catch (ignore) {}

  sheet.getRange(1, 8, 40, 2).clearContent();
  sheet.getRange(1, 8, block.length, 2).setValues(block);

  sheet.setColumnWidth(7, 30);    // отступ
  sheet.setColumnWidth(8, 250);   // подписи
  sheet.setColumnWidth(9, 70);    // числа

  sheet.getRange('H1:I1').setFontWeight('bold').setBackground(C_HEAD).setFontColor(C_INK);
  sheet.getRange('H8:I8').setFontWeight('bold').setBackground(C_HEAD).setFontColor(C_INK);

  sheet.getRange(2, 9, block.length - 1, 1)
    .setHorizontalAlignment('center')
    .setFontWeight('bold')
    .setFontSize(12);

  sheet.getRange(1, 8, block.length, 2)
    .setBorder(true, true, true, true, true, true, C_LINE, SpreadsheetApp.BorderStyle.SOLID);
}


/* ═══════════════════════════════════════════════════════════
   ОФОРМЛЕНИЕ ТАБЛИЦЫ С ОТВЕТАМИ
   ═══════════════════════════════════════════════════════════ */
function styleData(sheet) {
  var last = Math.max(getLastDataRow(sheet), 2);

  // шапка
  sheet.getRange(1, 1, 1, 6)
    .setFontWeight('bold')
    .setBackground(C_HEAD)
    .setFontColor(C_INK)
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 34);

  // дата и время читаемым видом
  sheet.getRange(2, 1, last - 1, 1).setNumberFormat('dd.MM.yyyy  HH:mm');

  // перенос текста в колонке с напитками
  sheet.getRange(2, 5, last - 1, 1).setWrap(true);

  // тонкая сетка по данным
  sheet.getRange(1, 1, last, 6)
    .setBorder(true, true, true, true, true, true, C_LINE, SpreadsheetApp.BorderStyle.SOLID);

  // цветовые метки ответов
  var rules = [];

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Придёт')
    .setBackground('#E6EFE9').setFontColor('#38603F').setBold(true)
    .setRanges([sheet.getRange('C2:C1000')]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Не сможет')
    .setBackground('#FBECEC').setFontColor('#8E3B3B')
    .setRanges([sheet.getRange('C2:C1000')]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Останется ночевать')
    .setBackground('#EAF0F6').setFontColor('#2F4C6E')
    .setRanges([sheet.getRange('D2:D1000')]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Уедет вечером')
    .setBackground('#F6F2E8').setFontColor('#6E5A24')
    .setRanges([sheet.getRange('D2:D1000')]).build());

  sheet.setConditionalFormatRules(rules);
}


/* ═══════════════════════════════════════════════════════════
   ДИАГРАММЫ

   Строятся по блоку итогов (колонки H и I).
   Значения в этом блоке обновляются при каждом ответе,
   поэтому диаграммы перерисовываются сами — пересоздавать
   их каждый раз не нужно.
   ═══════════════════════════════════════════════════════════ */
function ensureCharts(sheet, force) {
  var existing = sheet.getCharts();

  if (existing.length > 0) {
    if (!force) return;                 // уже стоят — не трогаем
    for (var i = 0; i < existing.length; i++) sheet.removeChart(existing[i]);
  }

  // подписи значений внутри секторов круговых диаграмм
  var sliceText = { fontSize: 15, bold: true, color: C_INK };

  // 1. Кольцевая: придут / не смогут
  sheet.insertChart(
    sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange('H3:I4'))
      .setNumHeaders(0)
      .setPosition(1, 11, 0, 0)
      .setOption('title', 'Кто придёт')
      .setOption('titleTextStyle', { fontSize: 14, bold: true, color: C_INK })
      .setOption('pieHole', 0.45)
      .setOption('colors', [C_GREEN, C_ROSE])
      .setOption('pieSliceText', 'value')          // число прямо в секторе
      .setOption('pieSliceTextStyle', sliceText)
      .setOption('width', 400)
      .setOption('height', 260)
      .setOption('backgroundColor', '#FFFFFF')
      .setOption('legend', { position: 'right', textStyle: { fontSize: 12 } })
      .build()
  );

  // 2. Кольцевая: ночёвка
  sheet.insertChart(
    sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange('H5:I6'))
      .setNumHeaders(0)
      .setPosition(15, 11, 0, 0)
      .setOption('title', 'Ночёвка на локации')
      .setOption('titleTextStyle', { fontSize: 14, bold: true, color: C_INK })
      .setOption('pieHole', 0.45)
      .setOption('colors', [C_SAND, C_STONE])
      .setOption('pieSliceText', 'value')
      .setOption('pieSliceTextStyle', sliceText)
      .setOption('width', 400)
      .setOption('height', 260)
      .setOption('backgroundColor', '#FFFFFF')
      .setOption('legend', { position: 'right', textStyle: { fontSize: 12 } })
      .build()
  );

  // 3. Горизонтальные полосы: напитки
  //    Ось делаем целочисленной: считаем максимум, чтобы вместо
  //    0,25 / 0,50 / 0,75 были ровные деления 0 / 1 / 2 …
  var vals = sheet.getRange(9, 9, DRINKS.length, 1).getValues();
  var maxDrink = 0;
  for (var v = 0; v < vals.length; v++) {
    var n = Number(vals[v][0]) || 0;
    if (n > maxDrink) maxDrink = n;
  }
  var ticks = Math.min(Math.max(maxDrink, 1) + 1, 11);

  sheet.insertChart(
    sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(sheet.getRange('H9:I16'))
      .setNumHeaders(0)
      .setPosition(29, 11, 0, 0)
      .setOption('title', 'Напитки — сколько закупать')
      .setOption('titleTextStyle', { fontSize: 14, bold: true, color: C_INK })
      .setOption('colors', [C_GREEN])
      .setOption('width', 520)
      .setOption('height', 340)
      .setOption('backgroundColor', '#FFFFFF')
      .setOption('legend', { position: 'none' })
      // число в конце каждой полосы
      .setOption('series', { 0: { dataLabel: 'value', color: C_GREEN } })
      .setOption('annotations', {
        alwaysOutside: true,
        textStyle: { fontSize: 12, bold: true, color: C_INK }
      })
      .setOption('hAxis', {
        format: '0',
        viewWindow: { min: 0 },
        gridlines: { color: '#EDE9E1', count: ticks },
        textStyle: { fontSize: 11 }
      })
      .setOption('vAxis', { textStyle: { fontSize: 12 } })
      .build()
  );
}


function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
