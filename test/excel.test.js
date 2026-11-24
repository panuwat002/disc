/* =========================================================================
   ชุดทดสอบการนำเข้า–ส่งออกไฟล์ .xls และ .xlsx
   รันด้วย: node test/excel.test.js  หรือ  npm test
   ========================================================================= */

'use strict';

var assert = require('assert');
var fs = require('fs');
var os = require('os');
var path = require('path');
var vm = require('vm');
var childProcess = require('child_process');
var XLSX = require('xlsx');

var ROOT = path.resolve(__dirname, '..');
var excel = require(path.join(ROOT, 'excel.js'));

var passed = 0;
var failed = 0;

function test(name, fn) {
  try {
    fn();
    passed++;
    console.log('  [PASS] ' + name);
  } catch (err) {
    failed++;
    console.error('  [FAIL] ' + name);
    console.error('         ' + (err && err.stack ? err.stack.split('\n')[0] : err));
  }
}

/* ---------- สร้างโฟลเดอร์ชั่วคราวสำหรับไฟล์ทดสอบ ---------- */

var tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'disc-excel-test-'));

function cleanup() {
  try {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  } catch (e) { /* ignore */ }
}

console.log('\n=== ทดสอบฝั่งเครื่องผู้ดูแล (excel.js) ===');

var exportedXlsx = path.join(tmpDir, 'questions.xlsx');
var exportedXls = path.join(tmpDir, 'questions.xls');

test('1. ส่งออกชุดคำถามปัจจุบันเป็น .xlsx (OOXML ZIP header)', function () {
  var res = excel.exportQuestions(exportedXlsx, {});
  assert.strictEqual(res.count, 37);
  assert.strictEqual(res.pairs, 8);
  assert.ok(fs.existsSync(exportedXlsx));
  var buf = fs.readFileSync(exportedXlsx);
  assert.strictEqual(buf[0], 0x50); // 'P'
  assert.strictEqual(buf[1], 0x4B); // 'K'
});

test('2. ส่งออกชุดคำถามปัจจุบันเป็น .xls (BIFF8 OLE2 header)', function () {
  var res = excel.exportQuestions(exportedXls, {});
  assert.strictEqual(res.count, 37);
  assert.strictEqual(res.pairs, 8);
  assert.ok(fs.existsSync(exportedXls));
  var buf = fs.readFileSync(exportedXls);
  assert.strictEqual(buf[0], 0xD0);
  assert.strictEqual(buf[1], 0xCF);
  assert.strictEqual(buf[2], 0x11);
  assert.strictEqual(buf[3], 0xE0);
});

test('3. อ่านกลับจาก .xlsx (round-trip) ภาษาไทยและข้อมูลครบ 100% ไม่คลาดเคลื่อน', function () {
  var res = excel.importQuestions(exportedXlsx, { 'dry-run': true });
  assert.strictEqual(res.dryRun, true);
  assert.strictEqual(res.count, 37);
  assert.strictEqual(res.pairs, 8);
  assert.strictEqual(res.diff.unchanged, 37);
  assert.strictEqual(res.diff.modified, 0);
  assert.strictEqual(res.diff.added, 0);
  assert.strictEqual(res.diff.removed, 0);
});

test('4. อ่านกลับจาก .xls (round-trip) ภาษาไทยและข้อมูลครบ 100% ไม่คลาดเคลื่อน', function () {
  var res = excel.importQuestions(exportedXls, { 'dry-run': true });
  assert.strictEqual(res.dryRun, true);
  assert.strictEqual(res.count, 37);
  assert.strictEqual(res.pairs, 8);
  assert.strictEqual(res.diff.unchanged, 37);
  assert.strictEqual(res.diff.modified, 0);
  assert.strictEqual(res.diff.added, 0);
  assert.strictEqual(res.diff.removed, 0);
});

test('5. จับคู่หัวคอลัมน์ภาษาอังกฤษ (q, D, I, S, C, tag) และข้ามแถวว่างได้ถูกต้อง', function () {
  var wb = XLSX.utils.book_new();
  var rows = [
    ['Question', 'D', 'I', 'S', 'C', 'tag'],
    ['คำถามทดสอบ 1', 'ลุยเลย', 'คุยสนุก', 'ใจเย็น', 'ละเอียด', 'คู่1'],
    ['', '', '', '', '', ''],
    ['คำถามทดสอบ 2', 'สั่งการ', 'สร้างสีสัน', 'รับฟัง', 'ตรวจทาน', 'คู่1']
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'Sheet1');
  var f = path.join(tmpDir, 'english-cols.xls');
  XLSX.writeFile(wb, f, { bookType: 'biff8' });

  var res = excel.importQuestions(f, { 'dry-run': true });
  assert.strictEqual(res.count, 2);
  assert.strictEqual(res.pairs, 1);
});

test('6. ปฏิเสธไฟล์ที่ข้อมูลไม่ครบหรือตัวเลือกซ้ำ โดยไม่เขียนทับไฟล์เดิม', function () {
  var wb = XLSX.utils.book_new();
  var rows = [
    ['คำถาม', 'D (กระทิง)', 'I (นกอินทรี)', 'S (หนู)', 'C (หมี)'],
    ['ข้อขาดตัวเลือก', 'ก', 'ข', '', 'ง'],
    ['ข้อตัวเลือกซ้ำ', 'เหมือนกัน', 'เหมือนกัน', 'ค', 'ง']
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'คำถาม');
  var f = path.join(tmpDir, 'invalid.xlsx');
  XLSX.writeFile(wb, f, { bookType: 'xlsx' });

  var threw = false;
  try {
    excel.importQuestions(f, { 'dry-run': true });
  } catch (err) {
    threw = true;
    assert.ok(Array.isArray(err.validationErrors));
    assert.strictEqual(err.validationErrors.length, 2);
  }
  assert.ok(threw, 'ควร throw error เมื่อข้อมูลไม่ครบ');
});

test('7. สร้างไฟล์ต้นแบบคำถามและผลประเมินได้ทั้ง .xlsx และ .xls', function () {
  var tq1 = path.join(tmpDir, 'tq.xlsx');
  var tq2 = path.join(tmpDir, 'tq.xls');
  var tr1 = path.join(tmpDir, 'tr.xlsx');
  var tr2 = path.join(tmpDir, 'tr.xls');

  excel.createQuestionsTemplate(tq1, {});
  excel.createQuestionsTemplate(tq2, {});
  excel.createResultsTemplate(tr1, {});
  excel.createResultsTemplate(tr2, {});

  [tq1, tq2, tr1, tr2].forEach(function (p) {
    assert.ok(fs.existsSync(p), 'ไม่พบไฟล์ ' + p);
    var wb = XLSX.readFile(p);
    assert.ok(wb.SheetNames.length >= 2);
  });
});

test('8. แปลงไฟล์ .xls <-> .xlsx รักษาทุกชีตและภาษาไทยครบถ้วน', function () {
  var converted = path.join(tmpDir, 'converted.xlsx');
  var cv = excel.convertFile(exportedXls, converted, {});
  assert.deepStrictEqual(cv.sheets, ['คำถาม', 'คู่มือ']);
  var res = excel.importQuestions(converted, { 'dry-run': true });
  assert.strictEqual(res.count, 37);
  assert.strictEqual(res.diff.unchanged, 37);
});

/* ---------- ทดสอบฝั่ง Google Apps Script (apps-script/Code.gs) ---------- */

console.log('\n=== ทดสอบฝั่ง Google Sheets (apps-script/Code.gs) ===');

function createAppsScriptSandbox() {
  var sheetsMap = {};
  var docProps = {};
  var menuCalls = [];

  function makeSheet(name) {
    var grid = [];
    var sheetObj = {
      _grid: grid,
      getName: function () { return name; },
      getLastRow: function () { return grid.length; },
      getLastColumn: function () {
        var max = 0;
        grid.forEach(function (r) { if (r.length > max) max = r.length; });
        return max;
      },
      getMaxRows: function () { return Math.max(100, grid.length + 10); },
      setFrozenRows: function () {},
      autoResizeColumns: function () {},
      clear: function () { grid.length = 0; },
      appendRow: function (row) { grid.push(row.slice()); },
      getRange: function (row, col, numRows, numCols) {
        numRows = numRows || 1;
        numCols = numCols || 1;
        return {
          getValues: function () {
            var out = [];
            for (var r = 0; r < numRows; r++) {
              var src = grid[row - 1 + r] || [];
              var line = [];
              for (var c = 0; c < numCols; c++) {
                line.push(src[col - 1 + c] === undefined ? '' : src[col - 1 + c]);
              }
              out.push(line);
            }
            return out;
          },
          setValues: function (vals) {
            for (var r = 0; r < vals.length; r++) {
              var rr = row - 1 + r;
              while (grid.length <= rr) grid.push([]);
              for (var c = 0; c < vals[r].length; c++) {
                grid[rr][col - 1 + c] = vals[r][c];
              }
            }
            return this;
          },
          setFontWeight: function () { return this; },
          setVerticalAlignment: function () { return this; }
        };
      }
    };
    sheetsMap[name] = sheetObj;
    return sheetObj;
  }

  function formatDateMock(d, tz, pattern) {
    var dt = new Date(d.getTime());
    if (tz === 'Asia/Bangkok') {
      dt = new Date(d.getTime() + 7 * 3600 * 1000);
    }
    var y = dt.getUTCFullYear();
    var m = ('0' + (dt.getUTCMonth() + 1)).slice(-2);
    var day = ('0' + dt.getUTCDate()).slice(-2);
    var hh = ('0' + dt.getUTCHours()).slice(-2);
    var mm = ('0' + dt.getUTCMinutes()).slice(-2);
    var ss = ('0' + dt.getUTCSeconds()).slice(-2);
    return pattern
      .replace('yyyy', String(y))
      .replace('MM', m)
      .replace('dd', day)
      .replace('HH', hh)
      .replace('mm', mm)
      .replace('ss', ss);
  }

  var uiMenuBuilder = {
    addItem: function (label, fn) { menuCalls.push({ label: label, fn: fn }); return this; },
    addSeparator: function () { menuCalls.push({ sep: true }); return this; },
    addToUi: function () { return this; }
  };

  var sandbox = {
    console: console,
    Date: Date,
    Math: Math,
    Number: Number,
    String: String,
    Array: Array,
    Object: Object,
    JSON: JSON,
    RegExp: RegExp,
    parseInt: parseInt,
    isNaN: isNaN,
    LockService: {
      getScriptLock: function () {
        return { waitLock: function () {}, releaseLock: function () {} };
      }
    },
    PropertiesService: {
      getDocumentProperties: function () {
        return {
          getProperty: function (k) { return docProps[k] || null; },
          setProperty: function (k, v) { docProps[k] = v; }
        };
      }
    },
    Utilities: {
      formatDate: formatDateMock
    },
    SpreadsheetApp: {
      getActiveSpreadsheet: function () {
        return {
          getName: function () { return 'ผลประเมิน DISC 2569'; },
          getSheetByName: function (n) { return sheetsMap[n] || null; },
          insertSheet: function (n) { return makeSheet(n); },
          getSheets: function () { return Object.keys(sheetsMap).map(function (k) { return sheetsMap[k]; }); },
          setActiveSheet: function () {}
        };
      },
      getUi: function () {
        return {
          createMenu: function () { return uiMenuBuilder; },
          showModalDialog: function () {},
          alert: function () {}
        };
      }
    },
    HtmlService: {
      createHtmlOutput: function (s) {
        return {
          _html: s,
          setWidth: function () { return this; },
          setHeight: function () { return this; }
        };
      }
    },
    ContentService: {
      MimeType: { JSON: 'application/json' },
      createTextOutput: function (t) {
        return { _text: t, setMimeType: function () { return this; } };
      }
    },
    Logger: { log: function () {} },
    _menuCalls: menuCalls,
    _sheetsMap: sheetsMap
  };

  var code = fs.readFileSync(path.join(ROOT, 'apps-script', 'Code.gs'), 'utf8');
  vm.runInNewContext(code, sandbox, { filename: 'Code.gs' });
  return sandbox;
}

var gs = createAppsScriptSandbox();

test('9. Code.gs โหลดได้ไม่มี syntax error และ onOpen() สร้างเมนูครบทั้ง 4 รายการ', function () {
  gs.onOpen();
  var items = gs._menuCalls.filter(function (x) { return !x.sep; });
  assert.strictEqual(items.length, 4);
  assert.strictEqual(items[0].fn, 'openExcelImportDialog');
  assert.strictEqual(items[1].fn, 'openExcelExportDialog');
  assert.strictEqual(items[2].fn, 'openExcelTemplateDialog');
  assert.strictEqual(items[3].fn, 'createSummarySheet');
});

test('10. HTML Dialog ของทั้ง 3 แท็บมี JavaScript ที่ syntax ถูกต้องสมบูรณ์', function () {
  ['import', 'export', 'template'].forEach(function (tab) {
    var html = gs.buildExcelDialogHtml_(tab);
    assert.ok(html.indexOf('XLSX.write') !== -1);
    assert.ok(html.indexOf('biff8') !== -1);
    var scripts = html.match(/<script>([\s\S]*?)<\/script>/g);
    assert.ok(scripts && scripts.length >= 2);
    scripts.forEach(function (block) {
      var body = block.replace(/^<script>/, '').replace(/<\/script>$/, '');
      new Function(body); // จะ throw ถ้ามี syntax error
    });
  });
});

test('11. normalizeDateString_ แปลงวันที่ได้ทั้ง ISO, dd/MM/yyyy, พ.ศ. และ Excel serial', function () {
  assert.strictEqual(gs.normalizeDateString_('2026-11-24 09:30:00'), '2026-11-24 09:30:00');
  assert.strictEqual(gs.normalizeDateString_('24/11/2569 09:30'), '2026-11-24 09:30:00');
  assert.strictEqual(gs.normalizeDateString_('2569-11-24'), '2026-11-24 00:00:00');
  assert.strictEqual(gs.normalizeDateString_(46350), '2026-11-24 00:00:00');
});

test('12. นำเข้าข้อมูลพนักงานจากไฟล์ .xls คำนวณสไตล์หลัก/รอง และข้ามรายการซ้ำได้', function () {
  var empFile = path.join(tmpDir, 'emp-import.xls');
  var wb = XLSX.utils.book_new();
  var aoa = [
    ['วันเวลาที่ทำ', 'ชื่อ-สกุล', 'แผนก', 'D', 'I', 'S', 'C'],
    ['2026-10-01 09:00:00', 'สมชาย ก', 'ฝ่ายขาย', 16, 10, 6, 5],
    ['2026-11-20 14:15:00', 'สมหญิง ข', 'ฝ่ายบัญชี', 4, 5, 10, 18],
    ['', '', '', '', '', '', '']
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), 'คำตอบ');
  XLSX.writeFile(wb, empFile, { bookType: 'biff8' });

  var readWb = XLSX.readFile(empFile);
  var rows = XLSX.utils.sheet_to_json(readWb.Sheets['คำตอบ'], { defval: '', raw: true });

  var res1 = gs.importExcelRows({
    targetSheet: 'คำตอบ',
    rows: rows,
    skipDuplicates: true,
    recalcScores: true
  });
  assert.strictEqual(res1.ok, true);
  assert.strictEqual(res1.importedCount, 2);
  assert.strictEqual(res1.skippedEmpty, 1);

  // ตรวจว่าคำนวณสไตล์หลัก/รองถูกต้อง
  var sh = gs._sheetsMap['คำตอบ'];
  var headers = sh._grid[0];
  var pCol = headers.indexOf('รหัสสไตล์หลัก');
  var sCol = headers.indexOf('รหัสสไตล์รอง');
  assert.strictEqual(sh._grid[1][pCol], 'D');
  assert.strictEqual(sh._grid[1][sCol], 'I');
  assert.strictEqual(sh._grid[2][pCol], 'C');
  assert.strictEqual(sh._grid[2][sCol], 'S');

  // นำเข้าซ้ำรอบสอง — ต้องข้ามทั้ง 2 แถว
  var res2 = gs.importExcelRows({
    targetSheet: 'คำตอบ',
    rows: rows,
    skipDuplicates: true,
    recalcScores: true
  });
  assert.strictEqual(res2.ok, true);
  assert.strictEqual(res2.importedCount, 0);
  assert.strictEqual(res2.skippedDuplicates, 2);
});

test('13. นำเข้าข้อมูลผู้สมัครงาน เติมคอลัมน์วิเคราะห์สัมภาษณ์ให้อัตโนมัติ', function () {
  var res = gs.importExcelRows({
    targetSheet: 'ผู้สมัคร',
    rows: [
      { 'วันเวลาที่ทำ': '2026-11-15 11:00:00', 'ชื่อ-สกุล': 'นายสมัคร ทดสอบ', 'ตำแหน่งที่สมัคร': 'พนักงานขาย', 'D': 5, 'I': 6, 'S': 14, 'C': 12 }
    ],
    skipDuplicates: true,
    recalcScores: true
  });
  assert.strictEqual(res.ok, true);
  assert.strictEqual(res.importedCount, 1);

  var sh = gs._sheetsMap['ผู้สมัคร'];
  var headers = sh._grid[0];
  var wantedIdx = headers.indexOf('ตำแหน่งนี้มองหา');
  var probesIdx = headers.indexOf('ประเด็นที่ควรถามเพิ่มในสัมภาษณ์');
  var watchIdx = headers.indexOf('จุดที่ควรสังเกต');
  assert.ok(String(sh._grid[1][wantedIdx]).indexOf('นกอินทรี') !== -1);
  assert.ok(String(sh._grid[1][probesIdx]).length > 20);
  assert.ok(String(sh._grid[1][watchIdx]).length > 20);
});

test('14. ส่งออกข้อมูลพร้อมกรองช่วงวันที่ และสร้างไฟล์ .xls / .xlsx ได้จริง', function () {
  var expAll = gs.getExportData({ sheets: ['คำตอบ', 'ผู้สมัคร'] });
  assert.strictEqual(expAll.ok, true);
  assert.strictEqual(expAll.sheets[0].exportedCount, 2);
  assert.strictEqual(expAll.sheets[1].exportedCount, 1);

  // กรองเฉพาะข้อมูลเก่าก่อนเดือน พ.ย. 2026
  var expOld = gs.getExportData({
    sheets: ['คำตอบ'],
    dateFrom: '2026-01-01',
    dateTo: '2026-10-31'
  });
  assert.strictEqual(expOld.ok, true);
  assert.strictEqual(expOld.sheets[0].exportedCount, 1);
  assert.strictEqual(expOld.sheets[0].rows[1][1], 'สมชาย ก');

  // เขียนออกเป็น .xls จริงแล้วอ่านกลับตรวจ
  var outWb = XLSX.utils.book_new();
  expAll.sheets.forEach(function (s) {
    XLSX.utils.book_append_sheet(outWb, XLSX.utils.aoa_to_sheet(s.rows), s.name);
  });
  var outXls = path.join(tmpDir, 'exported-results.xls');
  XLSX.writeFile(outWb, outXls, { bookType: 'biff8' });
  var checkWb = XLSX.readFile(outXls);
  assert.deepStrictEqual(checkWb.SheetNames, ['คำตอบ', 'ผู้สมัคร']);
});

test('15. รัน node build.js แล้วยังทำงานปกติและไม่ทำลายโค้ด Excel ใน Code.gs', function () {
  var out = childProcess.execFileSync(process.execPath, [path.join(ROOT, 'build.js')], {
    cwd: ROOT,
    encoding: 'utf8'
  });
  assert.ok(out.indexOf('สร้างไฟล์เรียบร้อย') !== -1);
  var gsAfter = fs.readFileSync(path.join(ROOT, 'apps-script', 'Code.gs'), 'utf8');
  assert.ok(gsAfter.indexOf('function importExcelRows(') !== -1);
  assert.ok(gsAfter.indexOf('function getExportData(') !== -1);
});

cleanup();

console.log('\nสรุปผลทดสอบ: ผ่าน ' + passed + ' / ' + (passed + failed) + ' รายการ\n');
if (failed > 0) process.exit(1);
