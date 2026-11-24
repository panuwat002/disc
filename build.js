/* =========================================================================
   ตัวสร้างไฟล์ — สั่งด้วย  node build.js

   อ่าน  source/questions.js   (ไฟล์ต้นฉบับ มีทั้งคำถามและเฉลย)
   เขียน index.html            (เฉพาะคำถาม ไม่มีเฉลย → ขึ้น GitHub ได้)
         apps-script/Code.gs   (เฉพาะเฉลย → อยู่ฝั่ง Google เท่านั้น)

   ทำไมต้องมี: คำถามกับเฉลยต้องอยู่คนละไฟล์เพื่อไม่ให้ผู้สมัครเปิดดูเฉลยได้
   ถ้าแก้มือทั้งสองไฟล์ พลาดเมื่อไหร่คะแนนเพี้ยนโดยไม่มีอะไรฟ้อง
   สคริปต์นี้ทำให้พลาดแบบนั้นไม่ได้
   ========================================================================= */

var fs = require('fs');
var path = require('path');

var ROOT = __dirname;
var SRC = path.join(ROOT, 'source', 'questions.js');
var HTML = path.join(ROOT, 'index.html');
var GS = path.join(ROOT, 'apps-script', 'Code.gs');

var BEGIN = '/* <<< สร้างอัตโนมัติ — ห้ามแก้บรรทัดด้านล่างด้วยมือ';
var END = '/* สร้างอัตโนมัติ >>> */';

var STYLES = ['D', 'I', 'S', 'C'];
var errors = [];
var warnings = [];

/* ---------- อ่านไฟล์ต้นฉบับ ---------- */

if (!fs.existsSync(SRC)) {
  console.error('ไม่พบไฟล์ ' + SRC);
  process.exit(1);
}
var SOURCE = require(SRC);
if (!Array.isArray(SOURCE) || !SOURCE.length) {
  console.error('source/questions.js ต้องมีตัวแปร SOURCE เป็นรายการคำถาม');
  process.exit(1);
}

/* ---------- ตรวจความถูกต้อง ---------- */

var seenQ = {};
var tagged = {};

SOURCE.forEach(function (e, i) {
  var no = 'ข้อ ' + (i + 1);

  if (!e.q || !String(e.q).trim()) {
    errors.push(no + ': ไม่มีข้อความคำถาม (q)');
    return;
  }

  var texts = [];
  STYLES.forEach(function (s) {
    if (!e[s] || !String(e[s]).trim()) {
      errors.push(no + ': ขาดตัวเลือก ' + s);
    } else {
      texts.push(String(e[s]).trim());
    }
  });

  texts.forEach(function (t, n) {
    if (texts.indexOf(t) !== n) {
      errors.push(no + ': มีตัวเลือกซ้ำกัน — "' + t + '"');
    }
  });

  var q = String(e.q).trim();
  if (seenQ[q]) {
    warnings.push(no + ': คำถามซ้ำกับข้อ ' + seenQ[q]);
  } else {
    seenQ[q] = i + 1;
  }

  if (e.tag) {
    var tag = String(e.tag).trim();
    (tagged[tag] = tagged[tag] || []).push(i);
  }
});

Object.keys(tagged).forEach(function (tag) {
  var n = tagged[tag].length;
  if (n !== 2) {
    warnings.push('tag "' + tag + '" ใช้ใน ' + n + ' ข้อ (ต้องเป็น 2 ข้อพอดีถึงจะนำไปวัดความสอดคล้องได้) — ข้ามคู่นี้');
  }
});

if (errors.length) {
  console.error('\n  ไม่ผ่านการตรวจ ' + errors.length + ' จุด — ยังไม่ได้เขียนไฟล์ใด ๆ\n');
  errors.forEach(function (m) { console.error('   x ' + m); });
  console.error('');
  process.exit(1);
}

/* ---------- สลับลำดับตัวเลือกที่เก็บไว้ ----------
   ลำดับที่เขียนลง index.html ต้องไม่ใช่ D,I,S,C เรียงตรง ๆ
   ไม่งั้นคนที่เปิด View Source จะเดาได้ว่าตัวเลือกที่ 1 คือ D เสมอ

   ลำดับนี้คิดมาจากข้อความคำถาม จึงคงที่ตลอด
   แก้ข้อหนึ่งแล้วข้ออื่นไม่ขยับตาม                                        */

function permutations(arr) {
  if (arr.length <= 1) return [arr];
  var out = [];
  arr.forEach(function (x, i) {
    var rest = arr.slice(0, i).concat(arr.slice(i + 1));
    permutations(rest).forEach(function (p) { out.push([x].concat(p)); });
  });
  return out;
}
var PERMS = permutations(STYLES);   // 24 แบบ

function permFor(text) {
  var h = 2166136261;
  for (var i = 0; i < text.length; i++) {
    h ^= text.charCodeAt(i);
    h = Math.imul(h, 16777619);
  }
  return PERMS[Math.abs(h) % PERMS.length];
}

var questions = [];
var answerKey = [];

SOURCE.forEach(function (e) {
  var p = permFor(String(e.q).trim());
  questions.push({
    t: String(e.q).trim(),
    o: p.map(function (s) { return String(e[s]).trim(); })
  });
  answerKey.push(p.join(''));
});

var pairs = [];
Object.keys(tagged).forEach(function (tag) {
  if (tagged[tag].length === 2) pairs.push({ tag: tag, idx: tagged[tag] });
});
pairs.sort(function (a, b) { return a.idx[0] - b.idx[0]; });

/* ---------- เขียนทับระหว่างเครื่องหมาย ---------- */

function replaceBlock(file, body, label) {
  var text = fs.readFileSync(file, 'utf8');
  var a = text.indexOf(BEGIN);
  var b = text.indexOf(END);
  if (a < 0 || b < 0) {
    console.error('ไม่พบเครื่องหมายคร่อมส่วนที่สร้างอัตโนมัติใน ' + label);
    console.error('ต้องมีบรรทัด "' + BEGIN + '" และ "' + END + '"');
    process.exit(1);
  }
  var out = text.slice(0, a) + BEGIN + '\n   แก้ที่ source/questions.js แล้วสั่ง node build.js */\n\n'
    + body + '\n' + END + text.slice(b + END.length);
  fs.writeFileSync(file, out, 'utf8');
}

var esc = JSON.stringify;

/* index.html — คำถามอย่างเดียว ไม่มีเฉลย */
var htmlBody = 'var QUESTIONS = [\n' + questions.map(function (q) {
  return '  { t: ' + esc(q.t) + ',\n    o: [' + q.o.map(esc).join(',\n        ') + '] }';
}).join(',\n') + '\n];\n';

/* Code.gs — เฉลย คู่ความสอดคล้อง และชุดคำถามเริ่มต้นสำหรับสร้างชีต "ชุดคำถาม" */
var gsBody = ''
  + 'var ANSWER_KEY = [\n'
  + answerKey.map(function (k, i) {
      return '  ' + esc(k) + ',' + ' '.repeat(Math.max(1, 3 - String(i + 1).length))
        + '// ข้อ ' + (i + 1) + ' ' + questions[i].t.slice(0, 40);
    }).join('\n')
  + '\n];\n\n'
  + 'var CONSISTENCY_PAIRS = [\n'
  + (pairs.length
      ? pairs.map(function (p) {
          return '  [' + p.idx[0] + ', ' + p.idx[1] + '],   // ' + p.tag
            + ' (ข้อ ' + (p.idx[0] + 1) + ' กับ ข้อ ' + (p.idx[1] + 1) + ')';
        }).join('\n')
      : '  // ยังไม่ได้ใส่ tag ให้ข้อไหนเลย')
  + '\n];\n\n'
  + 'var DEFAULT_QUESTIONS = [\n'
  + SOURCE.map(function (e) {
      var o = {
        q: String(e.q).trim(),
        D: String(e.D).trim(),
        I: String(e.I).trim(),
        S: String(e.S).trim(),
        C: String(e.C).trim()
      };
      if (e.tag && String(e.tag).trim()) o.tag = String(e.tag).trim();
      return '  ' + esc(o);
    }).join(',\n')
  + '\n];\n';

replaceBlock(HTML, htmlBody, 'index.html');
replaceBlock(GS, gsBody, 'apps-script/Code.gs');

/* ---------- สรุป ---------- */

/* รหัสย่อของเฉลยชุดนี้ — ต้องคิดแบบเดียวกับ keyFingerprint_() ใน Code.gs */
function fingerprint(keys) {
  var s = keys.join('');
  var h = 2166136261;
  for (var i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 16777619);
  }
  return ('0000000' + (h >>> 0).toString(16)).slice(-8);
}
var keyId = fingerprint(answerKey);

console.log('');
console.log('  สร้างไฟล์เรียบร้อย');
console.log('');
console.log('   คำถาม            ' + questions.length + ' ข้อ');
console.log('   คู่ความสอดคล้อง   ' + pairs.length + ' คู่');
console.log('   เวลาที่ผู้ตอบใช้   ประมาณ ' + Math.max(1, Math.round(questions.length * 11 / 60)) + ' นาที');
console.log('   รหัสเฉลย          ' + keyId);
console.log('');

if (warnings.length) {
  console.log('  ข้อสังเกต ' + warnings.length + ' จุด (ไฟล์ถูกเขียนแล้ว แต่ควรดู)');
  warnings.forEach(function (m) { console.log('   ! ' + m); });
  console.log('');
}

console.log('  ขั้นต่อไป');
console.log('   1. git add -A && git commit -m "ปรับคำถาม" && git push');
console.log('   2. เปิด apps-script/Code.gs คัดลอกทั้งไฟล์ไปวางใน Apps Script');
console.log('      แล้วกด ดินสอ -> เวอร์ชันใหม่ -> ทำให้ใช้งานได้');
console.log('   3. เปิด URL ของ /exec แล้วดูว่า keyId ตรงกับ ' + keyId + ' หรือยัง');
console.log('      ถ้าไม่ตรง = ข้อ 2 ยังไม่สำเร็จ Google ยังใช้เฉลยเก่าอยู่');
console.log('');
