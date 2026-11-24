/* =========================================================================
   เครื่องมือนำเข้า–ส่งออกไฟล์ Excel (.xls / .xlsx) สำหรับแบบประเมิน DISC

   คำสั่งหลัก:
     node excel.js export-questions [ไฟล์ปลายทาง.xlsx|.xls] [--format xlsx|xls]
     node excel.js import-questions <ไฟล์ต้นทาง.xls|.xlsx> [--sheet ชื่อชีต] [--dry-run] [--no-build] [--no-backup]
     node excel.js template-questions [ไฟล์ปลายทาง.xlsx|.xls] [--format xlsx|xls]
     node excel.js template-results [ไฟล์ปลายทาง.xlsx|.xls] [--format xlsx|xls]
     node excel.js convert <ไฟล์ต้นทาง.xls|.xlsx> [ไฟล์ปลายทาง.xlsx|.xls] [--format xlsx|xls]

   รองรับทั้งไฟล์ .xls (Excel 97–2003 BIFF8) และ .xlsx (Excel 2007+)
   ========================================================================= */

'use strict';

var fs = require('fs');
var path = require('path');
var childProcess = require('child_process');
var XLSX = require('xlsx');

var ROOT = __dirname;
var SRC = path.join(ROOT, 'source', 'questions.js');
var BACKUP_DIR = path.join(ROOT, 'source', 'backups');
var EXPORT_DIR = path.join(ROOT, 'exports');

var STYLES = ['D', 'I', 'S', 'C'];

var QUESTION_HEADERS = [
  'ข้อที่',
  'คำถาม',
  'D (กระทิง)',
  'I (นกอินทรี)',
  'S (หนู)',
  'C (หมี)',
  'กลุ่มความสอดคล้อง (tag)'
];

/* ---------- ตัวช่วยทั่วไป ---------- */

function pad2(n) {
  return (n < 10 ? '0' : '') + n;
}

function dateStamp(withTime) {
  var d = new Date();
  var s = '' + d.getFullYear() + pad2(d.getMonth() + 1) + pad2(d.getDate());
  if (withTime) {
    s += '-' + pad2(d.getHours()) + pad2(d.getMinutes()) + pad2(d.getSeconds());
  }
  return s;
}

function ensureDir(dir) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

function resolveBookType(filePath, formatFlag) {
  var fmt = formatFlag ? String(formatFlag).toLowerCase().replace(/^\./, '') : '';
  var ext = filePath ? path.extname(filePath).toLowerCase().replace(/^\./, '') : '';
  var chosen = fmt || ext || 'xlsx';
  if (chosen === 'xls' || chosen === 'biff8') {
    return { ext: '.xls', bookType: 'biff8', label: '.xls (Excel 97–2003)' };
  }
  if (chosen === 'xlsx') {
    return { ext: '.xlsx', bookType: 'xlsx', label: '.xlsx (Excel)' };
  }
  throw new Error('ไม่รองรับรูปแบบไฟล์ "' + chosen + '" — กรุณาใช้ .xlsx หรือ .xls');
}

function parseArgs(argv) {
  var positionals = [];
  var flags = {};
  for (var i = 0; i < argv.length; i++) {
    var arg = argv[i];
    if (arg === '--dry-run' || arg === '--no-build' || arg === '--no-backup') {
      flags[arg.slice(2)] = true;
    } else if (arg === '--format' || arg === '-f') {
      flags.format = argv[++i];
    } else if (arg === '--sheet' || arg === '-s') {
      flags.sheet = argv[++i];
    } else if (arg.indexOf('--format=') === 0) {
      flags.format = arg.slice('--format='.length);
    } else if (arg.indexOf('--sheet=') === 0) {
      flags.sheet = arg.slice('--sheet='.length);
    } else if (arg === '--help' || arg === '-h') {
      flags.help = true;
    } else if (arg.indexOf('-') === 0) {
      throw new Error('ไม่รู้จักตัวเลือก: ' + arg);
    } else {
      positionals.push(arg);
    }
  }
  return { positionals: positionals, flags: flags };
}

function loadCurrentSource() {
  if (!fs.existsSync(SRC)) return [];
  delete require.cache[require.resolve(SRC)];
  var data = require(SRC);
  return Array.isArray(data) ? data : [];
}

function writeWorkbook(wb, outPath, bookType) {
  ensureDir(path.dirname(outPath));
  XLSX.writeFile(wb, outPath, { bookType: bookType });
}

/* ---------- จับคู่ชื่อคอลัมน์แบบยืดหยุ่น ---------- */

function normalizeHeader(h) {
  return String(h == null ? '' : h)
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function mapQuestionColumns(headerRow) {
  var map = { q: -1, D: -1, I: -1, S: -1, C: -1, tag: -1 };
  headerRow.forEach(function (raw, idx) {
    var h = normalizeHeader(raw);
    if (!h) return;
    if (map.q === -1 && (h === 'q' || h === 'question' || h === 'คำถาม' || h === 'ข้อความคำถาม' || h.indexOf('คำถาม') === 0)) {
      map.q = idx;
    } else if (map.D === -1 && (h === 'd' || /^d[\s(:\-]/.test(h) || h.indexOf('กระทิง') !== -1 || h === 'dominance')) {
      map.D = idx;
    } else if (map.I === -1 && (h === 'i' || /^i[\s(:\-]/.test(h) || h.indexOf('นกอินทรี') !== -1 || h === 'influence')) {
      map.I = idx;
    } else if (map.S === -1 && (h === 's' || /^s[\s(:\-]/.test(h) || h.indexOf('หนู') !== -1 || h === 'steadiness')) {
      map.S = idx;
    } else if (map.C === -1 && (h === 'c' || /^c[\s(:\-]/.test(h) || h.indexOf('หมี') !== -1 || h === 'conscientiousness')) {
      map.C = idx;
    } else if (map.tag === -1 && (h === 'tag' || h === 'pair' || h.indexOf('สอดคล้อง') !== -1 || h.indexOf('tag') !== -1)) {
      map.tag = idx;
    }
  });
  return map;
}

function pickQuestionSheet(wb, requestedName) {
  if (requestedName) {
    if (wb.SheetNames.indexOf(requestedName) === -1) {
      throw new Error('ไม่พบชีตชื่อ "' + requestedName + '" ในไฟล์ (มีชีต: ' + wb.SheetNames.join(', ') + ')');
    }
    return requestedName;
  }
  var preferred = ['คำถาม', 'questions', 'Questions', 'DISC'];
  for (var i = 0; i < preferred.length; i++) {
    if (wb.SheetNames.indexOf(preferred[i]) !== -1) return preferred[i];
  }
  for (var j = 0; j < wb.SheetNames.length; j++) {
    var name = wb.SheetNames[j];
    var rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
    if (rows.length) {
      var m = mapQuestionColumns(rows[0]);
      if (m.q !== -1 && m.D !== -1) return name;
    }
  }
  return wb.SheetNames[0];
}

/* ---------- ตรวจสอบความถูกต้องของชุดคำถาม ---------- */

function validateQuestions(items) {
  var errors = [];
  var warnings = [];
  var seenQ = {};
  var tagged = {};

  if (!items.length) {
    errors.push('ไม่พบรายการคำถามเลยแม้แต่ข้อเดียว');
    return { errors: errors, warnings: warnings, pairs: 0 };
  }

  items.forEach(function (e, i) {
    var prefix = e._row ? ('แถว ' + e._row + ' (ข้อ ' + (i + 1) + ')') : ('ข้อ ' + (i + 1));

    if (!e.q || !String(e.q).trim()) {
      errors.push(prefix + ': ไม่มีข้อความคำถาม');
      return;
    }

    var texts = [];
    STYLES.forEach(function (s) {
      if (!e[s] || !String(e[s]).trim()) {
        errors.push(prefix + ': ขาดตัวเลือก ' + s);
      } else {
        texts.push(String(e[s]).trim());
      }
    });

    texts.forEach(function (t, n) {
      if (texts.indexOf(t) !== n) {
        errors.push(prefix + ': มีตัวเลือกซ้ำกัน — "' + t + '"');
      }
    });

    var q = String(e.q).trim();
    if (seenQ[q]) {
      warnings.push(prefix + ': คำถามซ้ำกับข้อ ' + seenQ[q]);
    } else {
      seenQ[q] = i + 1;
    }

    if (e.tag && String(e.tag).trim()) {
      var tag = String(e.tag).trim();
      (tagged[tag] = tagged[tag] || []).push(i);
    }
  });

  var pairCount = 0;
  Object.keys(tagged).forEach(function (tag) {
    var n = tagged[tag].length;
    if (n === 2) {
      pairCount++;
    } else {
      warnings.push('tag "' + tag + '" ใช้ใน ' + n + ' ข้อ (ต้องเป็น 2 ข้อพอดีถึงจะวัดความสอดคล้องได้) — ข้ามคู่นี้');
    }
  });

  return { errors: errors, warnings: warnings, pairs: pairCount };
}

/* ---------- สร้างไฟล์ source/questions.js จากรายการคำถาม ---------- */

function formatQuestionsSource(items) {
  var header = [
    '/* =========================================================================',
    '   ไฟล์ต้นฉบับคำถาม — แก้ที่นี่ที่เดียว (หรือนำเข้าจาก Excel ด้วย node excel.js)',
    '',
    '   ไฟล์นี้ "ห้ามขึ้น GitHub" เพราะมีเฉลยอยู่ (ถูก .gitignore ไว้แล้ว)',
    '',
    '   แก้เสร็จแล้วสั่ง:  node build.js',
    '   สคริปต์จะไปเขียน index.html และ apps-script/Code.gs ให้ตรงกันเอง',
    '',
    '   รูปแบบแต่ละข้อ',
    '     q   ข้อความคำถาม',
    '     D   ตัวเลือกแบบผู้นำ (กระทิง)',
    '     I   ตัวเลือกแบบผู้มีอิทธิพล (นกอินทรี)',
    '     S   ตัวเลือกแบบผู้รักความสงบ (หนู)',
    '     C   ตัวเลือกแบบนักวิเคราะห์ (หมี)',
    '     tag ใส่ชื่อเดียวกันใน 2 ข้อที่ถามเรื่องเดียวกัน ระบบจะเอาไปวัดความสอดคล้อง',
    '         (ไม่ใส่ก็ได้ ข้อนั้นจะไม่ถูกนำไปวัด)',
    '   ========================================================================= */',
    '',
    'var SOURCE = ['
  ].join('\n');

  var body = items.map(function (item, idx) {
    var lines = [
      '  /* ข้อ ' + (idx + 1) + ' */',
      '  { q: ' + JSON.stringify(item.q) + ',',
      '    D: ' + JSON.stringify(item.D) + ',',
      '    I: ' + JSON.stringify(item.I) + ',',
      '    S: ' + JSON.stringify(item.S) + ',',
      '    C: ' + JSON.stringify(item.C) + (item.tag ? ',' : '')
    ];
    if (item.tag) {
      lines.push('    tag: ' + JSON.stringify(item.tag));
    }
    lines.push('  }');
    return lines.join('\n');
  }).join(',\n\n');

  return header + '\n' + body + '\n];\n\nif (typeof module !== "undefined") module.exports = SOURCE;\n';
}

/* ---------- สร้างชีตคำถามและชีตคู่มือ ---------- */

function buildQuestionsWorkbook(items) {
  var rows = [QUESTION_HEADERS];
  items.forEach(function (e, i) {
    rows.push([
      i + 1,
      String(e.q || '').trim(),
      String(e.D || '').trim(),
      String(e.I || '').trim(),
      String(e.S || '').trim(),
      String(e.C || '').trim(),
      e.tag ? String(e.tag).trim() : ''
    ]);
  });

  var wsQ = XLSX.utils.aoa_to_sheet(rows);
  wsQ['!cols'] = [
    { wch: 8 },
    { wch: 48 },
    { wch: 44 },
    { wch: 44 },
    { wch: 44 },
    { wch: 44 },
    { wch: 24 }
  ];

  var guideRows = [
    ['คู่มือการแก้ไขและนำเข้าชุดคำถาม DISC ด้วย Excel (.xls / .xlsx)'],
    [''],
    ['หัวข้อ', 'คำอธิบาย'],
    ['คอลัมน์บังคับ', 'คำถาม, D (กระทิง), I (นกอินทรี), S (หนู), C (หมี) — ห้ามเว้นว่าง และตัวเลือกทั้ง 4 ในข้อเดียวกันห้ามซ้ำกัน'],
    ['กลุ่มความสอดคล้อง (tag)', 'ใส่ชื่อเดียวกันใน 2 ข้อที่ถามเรื่องเดียวกัน เพื่อวัดว่าผู้ตอบตอบตรงกันหรือไม่ (เว้นว่างได้)'],
    ['คอลัมน์ ข้อที่', 'มีไว้ให้อ่านง่าย ระบบจะเรียงลำดับตามแถวในชีต "คำถาม" โดยอัตโนมัติตอนนำเข้า'],
    ['ลำดับตัวเลือก', 'พิมพ์ให้ตรงช่อง D / I / S / C ได้เลย ระบบจะสลับลำดับให้อัตโนมัติตอนสร้างหน้าเว็บ'],
    ['รองรับนามสกุลไฟล์', 'บันทึกเป็นได้ทั้ง .xlsx (Excel รุ่นใหม่) และ .xls (Excel 97–2003)'],
    ['วิธีนำเข้ากลับเข้าระบบ', 'รันคำสั่ง: node excel.js import-questions <ชื่อไฟล์.xls หรือ .xlsx>'],
    ['ทดสอบก่อนเขียนจริง', 'เติม --dry-run ท้ายคำสั่งเพื่อตรวจสอบความถูกต้องโดยยังไม่แก้ไขไฟล์ใด ๆ'],
    ['ความปลอดภัย', 'ห้ามอัปโหลดไฟล์นี้ขึ้น GitHub สาธารณะ เพราะมีเฉลย D / I / S / C ครบทุกข้อ']
  ];
  var wsGuide = XLSX.utils.aoa_to_sheet(guideRows);
  wsGuide['!cols'] = [{ wch: 28 }, { wch: 85 }];

  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, wsQ, 'คำถาม');
  XLSX.utils.book_append_sheet(wb, wsGuide, 'คู่มือ');
  return wb;
}

/* ---------- คำสั่ง 1: export-questions ---------- */

function exportQuestions(outArg, flags) {
  var items = loadCurrentSource();
  if (!items.length) {
    throw new Error('ไม่พบข้อมูลคำถามใน ' + SRC);
  }
  var fmt = resolveBookType(outArg, flags.format);
  var outPath = outArg
    ? path.resolve(process.cwd(), outArg)
    : path.join(EXPORT_DIR, 'disc-questions-' + dateStamp(false) + fmt.ext);

  if (outArg && flags.format && path.extname(outPath).toLowerCase() !== fmt.ext) {
    outPath = outPath.replace(/\.[^.]+$/, '') + fmt.ext;
  }

  var wb = buildQuestionsWorkbook(items);
  writeWorkbook(wb, outPath, fmt.bookType);

  var v = validateQuestions(items);
  return {
    outPath: outPath,
    format: fmt.label,
    count: items.length,
    pairs: v.pairs
  };
}

/* ---------- คำสั่ง 2: import-questions ---------- */

function parseQuestionsFile(inPath, sheetName) {
  if (!fs.existsSync(inPath)) {
    throw new Error('ไม่พบไฟล์: ' + inPath);
  }
  var wb = XLSX.readFile(inPath);
  var chosenSheet = pickQuestionSheet(wb, sheetName);
  var rows = XLSX.utils.sheet_to_json(wb.Sheets[chosenSheet], { header: 1, defval: '' });

  if (!rows.length) {
    throw new Error('ชีต "' + chosenSheet + '" ไม่มีข้อมูล');
  }

  // หาแถวหัวตาราง (เผื่อมีแถวชื่อเรื่องอยู่ด้านบน)
  var headerRowIdx = -1;
  var colMap = null;
  for (var r = 0; r < Math.min(rows.length, 10); r++) {
    var candidate = mapQuestionColumns(rows[r]);
    if (candidate.q !== -1 && (candidate.D !== -1 || candidate.I !== -1 || candidate.S !== -1 || candidate.C !== -1)) {
      headerRowIdx = r;
      colMap = candidate;
      break;
    }
  }

  if (headerRowIdx === -1) {
    throw new Error(
      'ไม่พบหัวตารางคำถามในชีต "' + chosenSheet + '"\n' +
      'ต้องมีคอลัมน์: คำถาม, D (กระทิง), I (นกอินทรี), S (หนู), C (หมี)'
    );
  }

  var missingCols = [];
  if (colMap.q === -1) missingCols.push('คำถาม');
  STYLES.forEach(function (s) {
    if (colMap[s] === -1) missingCols.push(s);
  });
  if (missingCols.length) {
    throw new Error('หัวตารางขาดคอลัมน์จำเป็น: ' + missingCols.join(', '));
  }

  var items = [];
  for (var i = headerRowIdx + 1; i < rows.length; i++) {
    var row = rows[i] || [];
    var q = String(row[colMap.q] == null ? '' : row[colMap.q]).trim();
    var D = String(row[colMap.D] == null ? '' : row[colMap.D]).trim();
    var I = String(row[colMap.I] == null ? '' : row[colMap.I]).trim();
    var S = String(row[colMap.S] == null ? '' : row[colMap.S]).trim();
    var C = String(row[colMap.C] == null ? '' : row[colMap.C]).trim();
    var tag = colMap.tag !== -1 ? String(row[colMap.tag] == null ? '' : row[colMap.tag]).trim() : '';

    // ข้ามแถวว่างเปล่าทั้งแถว
    if (!q && !D && !I && !S && !C && !tag) continue;

    var item = { q: q, D: D, I: I, S: S, C: C, _row: i + 1 };
    if (tag) item.tag = tag;
    items.push(item);
  }

  return { sheet: chosenSheet, items: items };
}

function summarizeDiff(oldItems, newItems) {
  var oldByQ = {};
  oldItems.forEach(function (e) {
    oldByQ[String(e.q || '').trim()] = e;
  });
  var newByQ = {};
  var unchanged = 0;
  var modified = 0;
  var added = 0;

  newItems.forEach(function (n) {
    newByQ[n.q] = n;
    var o = oldByQ[n.q];
    if (!o) {
      added++;
    } else if (
      String(o.D || '').trim() === n.D &&
      String(o.I || '').trim() === n.I &&
      String(o.S || '').trim() === n.S &&
      String(o.C || '').trim() === n.C &&
      String(o.tag || '').trim() === String(n.tag || '').trim()
    ) {
      unchanged++;
    } else {
      modified++;
    }
  });

  var removed = 0;
  oldItems.forEach(function (o) {
    if (!newByQ[String(o.q || '').trim()]) removed++;
  });

  return {
    oldCount: oldItems.length,
    newCount: newItems.length,
    unchanged: unchanged,
    modified: modified,
    added: added,
    removed: removed
  };
}

function importQuestions(inArg, flags) {
  if (!inArg) {
    throw new Error('กรุณาระบุไฟล์ .xls หรือ .xlsx ที่ต้องการนำเข้า');
  }
  var inPath = path.resolve(process.cwd(), inArg);
  var parsed = parseQuestionsFile(inPath, flags.sheet);
  var validation = validateQuestions(parsed.items);

  if (validation.errors.length) {
    var err = new Error('ไม่ผ่านการตรวจสอบ ' + validation.errors.length + ' จุด — ยังไม่ได้แก้ไขไฟล์ใด ๆ');
    err.validationErrors = validation.errors;
    throw err;
  }

  var oldItems = loadCurrentSource();
  var diff = summarizeDiff(oldItems, parsed.items);

  var cleanItems = parsed.items.map(function (it) {
    var o = { q: it.q, D: it.D, I: it.I, S: it.S, C: it.C };
    if (it.tag) o.tag = it.tag;
    return o;
  });

  var backupPath = null;
  if (!flags['dry-run']) {
    if (!flags['no-backup'] && fs.existsSync(SRC)) {
      ensureDir(BACKUP_DIR);
      backupPath = path.join(BACKUP_DIR, 'questions-' + dateStamp(true) + '.js');
      fs.copyFileSync(SRC, backupPath);
    }
    ensureDir(path.dirname(SRC));
    fs.writeFileSync(SRC, formatQuestionsSource(cleanItems), 'utf8');
  }

  var buildOutput = null;
  if (!flags['dry-run'] && !flags['no-build']) {
    buildOutput = childProcess.execFileSync(process.execPath, [path.join(ROOT, 'build.js')], {
      cwd: ROOT,
      encoding: 'utf8'
    });
  }

  return {
    inPath: inPath,
    sheet: parsed.sheet,
    count: cleanItems.length,
    pairs: validation.pairs,
    warnings: validation.warnings,
    diff: diff,
    dryRun: !!flags['dry-run'],
    backupPath: backupPath,
    built: !flags['dry-run'] && !flags['no-build'],
    buildOutput: buildOutput
  };
}

/* ---------- คำสั่ง 3: template-questions ---------- */

function createQuestionsTemplate(outArg, flags) {
  var fmt = resolveBookType(outArg, flags.format);
  var outPath = outArg
    ? path.resolve(process.cwd(), outArg)
    : path.join(EXPORT_DIR, 'template-questions' + fmt.ext);

  var sample = [
    {
      q: 'ตัวอย่างข้อ 1: เวลาอยู่ในกลุ่มเพื่อนหรือที่ทำงาน คุณมักจะเป็นคนที่...',
      D: 'กล้าตัดสินใจและเป็นคนนำทีม',
      I: 'ชวนคุย สร้างเสียงหัวเราะ และทำให้บรรยากาศคึกคัก',
      S: 'คอยรับฟัง สังเกต และช่วยเหลือคนรอบข้างอยู่เงียบ ๆ',
      C: 'ตรวจสอบความถูกต้องและคอยทักท้วงจุดที่อาจผิดพลาด',
      tag: 'บทบาทในกลุ่ม'
    },
    {
      q: 'ตัวอย่างข้อ 2: เมื่อต้องทำงานร่วมกับผู้อื่น คุณมักจะ...',
      D: 'เป็นผู้นำและตัดสินใจอย่างรวดเร็ว',
      I: 'เป็นมิตรและสร้างบรรยากาศที่สนุกสนาน',
      S: 'ให้ความสำคัญกับความสัมพันธ์และความรู้สึกของทุกคน',
      C: 'ทำงานอย่างเป็นระบบและละเอียดรอบคอบ',
      tag: 'บทบาทในกลุ่ม'
    }
  ];

  var wb = buildQuestionsWorkbook(sample);
  writeWorkbook(wb, outPath, fmt.bookType);
  return { outPath: outPath, format: fmt.label };
}

/* ---------- คำสั่ง 4: template-results ---------- */

function createResultsTemplate(outArg, flags) {
  var fmt = resolveBookType(outArg, flags.format);
  var outPath = outArg
    ? path.resolve(process.cwd(), outArg)
    : path.join(EXPORT_DIR, 'template-results' + fmt.ext);

  var empRows = [
    ['วันเวลาที่ทำ', 'ชื่อ-สกุล', 'แผนก', 'รหัสพนักงาน', 'อีเมล', 'D กระทิง', 'I นกอินทรี', 'S หนู', 'C หมี', 'เวลาที่ใช้ (วินาที)', 'จำนวนข้อ'],
    ['2026-11-24 09:30:00', 'สมชาย ใจดี (ตัวอย่าง)', 'ฝ่ายขาย', 'EMP-0001', 'somchai@company.co.th', 14, 12, 6, 5, 380, 37]
  ];
  var wsEmp = XLSX.utils.aoa_to_sheet(empRows);
  wsEmp['!cols'] = [
    { wch: 20 }, { wch: 24 }, { wch: 16 }, { wch: 14 }, { wch: 26 },
    { wch: 12 }, { wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 18 }, { wch: 12 }
  ];

  var appRows = [
    ['วันเวลาที่ทำ', 'ชื่อ-สกุล', 'ตำแหน่งที่สมัคร', 'D กระทิง', 'I นกอินทรี', 'S หนู', 'C หมี', 'เวลาที่ใช้ (วินาที)', 'จำนวนข้อ'],
    ['2026-11-24 10:15:00', 'สุดา รักงาน (ตัวอย่าง)', 'พนักงานขาย', 10, 15, 7, 5, 420, 37]
  ];
  var wsApp = XLSX.utils.aoa_to_sheet(appRows);
  wsApp['!cols'] = [
    { wch: 20 }, { wch: 24 }, { wch: 22 },
    { wch: 12 }, { wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 18 }, { wch: 12 }
  ];

  var guideRows = [
    ['คู่มือการนำเข้าข้อมูลผลประเมินเข้า Google Sheets (.xls / .xlsx)'],
    [''],
    ['หัวข้อ', 'คำอธิบาย'],
    ['วิธีนำเข้า', 'เปิด Google Sheet → เมนู "แบบประเมิน DISC" → "นำเข้าข้อมูลจาก Excel (.xls / .xlsx)" แล้วเลือกไฟล์นี้'],
    ['ชีต "คำตอบ"', 'สำหรับนำเข้าข้อมูลพนักงานปัจจุบัน (ใช้ชื่อ-สกุล, แผนก, คะแนน D/I/S/C)'],
    ['ชีต "ผู้สมัคร"', 'สำหรับนำเข้าข้อมูลผู้สมัครงาน (ระบบจะเติมคอลัมน์ "ตำแหน่งนี้มองหา", "ประเด็นที่ควรถามเพิ่ม", "จุดที่ควรสังเกต" ให้อัตโนมัติ)'],
    ['คอลัมน์สไตล์หลัก/รอง', 'ไม่ต้องกรอกเองก็ได้ ระบบจะคำนวณจากคะแนน D/I/S/C (หรือจากคอลัมน์ "ลำดับตัวเลือกที่กด") ให้อัตโนมัติ'],
    ['ป้องกันข้อมูลซ้ำ', 'ตอนนำเข้าสามารถติ๊ก "ข้ามรายการที่ซ้ำกับข้อมูลเดิม" ได้'],
    ['ลบแถวตัวอย่าง', 'กรุณาลบแถวตัวอย่างออกก่อนนำเข้าจริง']
  ];
  var wsGuide = XLSX.utils.aoa_to_sheet(guideRows);
  wsGuide['!cols'] = [{ wch: 24 }, { wch: 90 }];

  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, wsEmp, 'คำตอบ');
  XLSX.utils.book_append_sheet(wb, wsApp, 'ผู้สมัคร');
  XLSX.utils.book_append_sheet(wb, wsGuide, 'คู่มือ');

  writeWorkbook(wb, outPath, fmt.bookType);
  return { outPath: outPath, format: fmt.label };
}

/* ---------- คำสั่ง 5: convert (.xls <-> .xlsx) ---------- */

function convertFile(inArg, outArg, flags) {
  if (!inArg) {
    throw new Error('กรุณาระบุไฟล์ต้นทาง (.xls หรือ .xlsx)');
  }
  var inPath = path.resolve(process.cwd(), inArg);
  if (!fs.existsSync(inPath)) {
    throw new Error('ไม่พบไฟล์: ' + inPath);
  }
  var inExt = path.extname(inPath).toLowerCase();
  var defaultFmt = inExt === '.xls' ? 'xlsx' : 'xls';
  var fmt = resolveBookType(outArg, flags.format || (outArg ? '' : defaultFmt));
  var outPath = outArg
    ? path.resolve(process.cwd(), outArg)
    : inPath.replace(/\.[^.]+$/, '') + fmt.ext;

  var wb = XLSX.readFile(inPath);
  writeWorkbook(wb, outPath, fmt.bookType);

  return {
    inPath: inPath,
    outPath: outPath,
    format: fmt.label,
    sheets: wb.SheetNames
  };
}

/* ---------- แสดงวิธีใช้ ---------- */

function printHelp() {
  console.log([
    '',
    '  เครื่องมือนำเข้า–ส่งออกไฟล์ Excel (.xls / .xlsx) — แบบประเมิน DISC',
    '',
    '  วิธีใช้:',
    '    node excel.js export-questions [ไฟล์ปลายทาง] [--format xlsx|xls]',
    '        ส่งออกชุดคำถามปัจจุบันจาก source/questions.js เป็นไฟล์ Excel',
    '',
    '    node excel.js import-questions <ไฟล์.xls|.xlsx> [--sheet ชื่อชีต] [--dry-run] [--no-build]',
    '        นำเข้าชุดคำถามใหม่จาก Excel → ตรวจสอบ → สำรองของเดิม → เขียน source/questions.js → build',
    '',
    '    node excel.js template-questions [ไฟล์ปลายทาง] [--format xlsx|xls]',
    '        สร้างไฟล์ต้นแบบสำหรับกรอกคำถามชุดใหม่',
    '',
    '    node excel.js template-results [ไฟล์ปลายทาง] [--format xlsx|xls]',
    '        สร้างไฟล์ต้นแบบสำหรับนำเข้าข้อมูลผลประเมินเข้า Google Sheets',
    '',
    '    node excel.js convert <ไฟล์ต้นทาง.xls|.xlsx> [ไฟล์ปลายทาง] [--format xlsx|xls]',
    '        แปลงไฟล์ไปมาระหว่าง .xls ↔ .xlsx',
    '',
    '  ตัวอย่าง:',
    '    node excel.js export-questions exports/questions-old.xls',
    '    node excel.js import-questions imports/questions-new.xls --dry-run',
    '    node excel.js import-questions imports/questions-new.xlsx',
    ''
  ].join('\n'));
}

/* ---------- เรียกทำงานจาก Command Line ---------- */

function main() {
  var cmd = process.argv[2];
  if (!cmd || cmd === 'help' || cmd === '--help' || cmd === '-h') {
    printHelp();
    return;
  }

  try {
    var parsed = parseArgs(process.argv.slice(3));
    if (parsed.flags.help) {
      printHelp();
      return;
    }

    if (cmd === 'export-questions') {
      var res = exportQuestions(parsed.positionals[0], parsed.flags);
      console.log('');
      console.log('  ส่งออกชุดคำถามเรียบร้อย');
      console.log('   ไฟล์             ' + res.outPath);
      console.log('   รูปแบบ           ' + res.format);
      console.log('   จำนวนคำถาม       ' + res.count + ' ข้อ');
      console.log('   คู่ความสอดคล้อง   ' + res.pairs + ' คู่');
      console.log('');
    } else if (cmd === 'import-questions') {
      var imp = importQuestions(parsed.positionals[0], parsed.flags);
      console.log('');
      console.log(imp.dryRun ? '  [ทดสอบ --dry-run] ผ่านการตรวจสอบ — ยังไม่ได้แก้ไขไฟล์ใด ๆ' : '  นำเข้าชุดคำถามเรียบร้อย');
      console.log('   ไฟล์ต้นทาง       ' + imp.inPath);
      console.log('   ชีตที่อ่าน        ' + imp.sheet);
      console.log('   จำนวนคำถาม       ' + imp.diff.oldCount + ' -> ' + imp.count + ' ข้อ');
      console.log('   สรุปการเปลี่ยนแปลง  เหมือนเดิม ' + imp.diff.unchanged + ' | แก้ไข ' + imp.diff.modified + ' | เพิ่มใหม่ ' + imp.diff.added + ' | ลบออก ' + imp.diff.removed);
      console.log('   คู่ความสอดคล้อง   ' + imp.pairs + ' คู่');
      if (imp.backupPath) {
        console.log('   สำรองของเดิมไว้ที่  ' + imp.backupPath);
      }
      if (imp.warnings.length) {
        console.log('');
        console.log('  ข้อสังเกต ' + imp.warnings.length + ' จุด:');
        imp.warnings.forEach(function (w) { console.log('   ! ' + w); });
      }
      if (imp.buildOutput) {
        console.log(imp.buildOutput);
      } else {
        console.log('');
      }
    } else if (cmd === 'template-questions') {
      var tq = createQuestionsTemplate(parsed.positionals[0], parsed.flags);
      console.log('');
      console.log('  สร้างไฟล์ต้นแบบคำถามเรียบร้อย');
      console.log('   ไฟล์     ' + tq.outPath);
      console.log('   รูปแบบ   ' + tq.format);
      console.log('');
    } else if (cmd === 'template-results') {
      var tr = createResultsTemplate(parsed.positionals[0], parsed.flags);
      console.log('');
      console.log('  สร้างไฟล์ต้นแบบข้อมูลผลประเมินเรียบร้อย');
      console.log('   ไฟล์     ' + tr.outPath);
      console.log('   รูปแบบ   ' + tr.format);
      console.log('   วิธีใช้   เปิด Google Sheet -> เมนู "แบบประเมิน DISC" -> "นำเข้าข้อมูลจาก Excel"');
      console.log('');
    } else if (cmd === 'convert') {
      var cv = convertFile(parsed.positionals[0], parsed.positionals[1], parsed.flags);
      console.log('');
      console.log('  แปลงไฟล์เรียบร้อย');
      console.log('   ต้นทาง   ' + cv.inPath);
      console.log('   ปลายทาง  ' + cv.outPath + ' (' + cv.format + ')');
      console.log('   ชีต      ' + cv.sheets.join(', '));
      console.log('');
    } else {
      console.error('ไม่รู้จักคำสั่ง: ' + cmd);
      printHelp();
      process.exit(1);
    }
  } catch (err) {
    console.error('\n  ผิดพลาด: ' + err.message + '\n');
    if (err.validationErrors && err.validationErrors.length) {
      err.validationErrors.forEach(function (m) { console.error('   x ' + m); });
      console.error('');
    }
    process.exit(1);
  }
}

if (require.main === module) {
  main();
}

module.exports = {
  exportQuestions: exportQuestions,
  importQuestions: importQuestions,
  parseQuestionsFile: parseQuestionsFile,
  validateQuestions: validateQuestions,
  summarizeDiff: summarizeDiff,
  formatQuestionsSource: formatQuestionsSource,
  createQuestionsTemplate: createQuestionsTemplate,
  createResultsTemplate: createResultsTemplate,
  convertFile: convertFile,
  mapQuestionColumns: mapQuestionColumns,
  resolveBookType: resolveBookType
};
