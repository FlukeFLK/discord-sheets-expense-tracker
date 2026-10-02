// ทดสอบตัวแยกข้อความภาษาไทย (parseEntry_ / extractDate_) ของ apps-script/Code.gs
import fs from 'node:fs'; import vm from 'node:vm'; import assert from 'node:assert';
const ctx = { Utilities: {}, console, Date, JSON, Math, String, Number, Object, RegExp };
vm.createContext(ctx); vm.runInContext(fs.readFileSync(new URL('../apps-script/Code.gs', import.meta.url), 'utf8'), ctx);
const cats = {
  'รายรับ': ['เงินเดือน','โบนัส','ค่าล่วงเวลา (OT)','ฟรีแลนซ์/งานเสริม','ขายของออนไลน์','ดอกเบี้ย/ปันผล','เงินคืนภาษี','รายรับอื่นๆ'],
  'รายจ่าย': ['ค่าอาหาร','ค่าเดินทาง','ค่าที่พัก/ค่าเช่า','ค่าน้ำ-ค่าไฟ','ค่าโทรศัพท์/เน็ต','ของใช้ในบ้าน','ช้อปปิ้ง','สุขภาพ/ยา','ประกันภัย','สังสรรค์/บันเทิง','การศึกษา','ผ่อนชำระ/หนี้','ให้ครอบครัว','ท่องเที่ยว','Subscription','ทำบุญ/บริจาค','รายจ่ายอื่นๆ'],
  'เงินออม': ['เงินฝากออมทรัพย์','เงินสำรองฉุกเฉิน','กองทุน SSF/RMF','หุ้น/ETF','ออมทอง','ออมเพื่อเกษียณ','ออมอื่นๆ'] };
const D = { y: 2026, m: 10, d: 2 };
const k = o => `${o.y}-${String(o.m).padStart(2,'0')}-${String(o.d).padStart(2,'0')}`;
const cases = [
  ['จ่าย 120 อาหาร ข้าวมันไก่', '2026-10-02|รายจ่าย|ค่าอาหาร|120|ข้าวมันไก่'],
  ['จ่าย 850 ไฟ 25/9', '2026-09-25|รายจ่าย|ค่าน้ำ-ค่าไฟ|850|ไฟ'],
  ['เมื่อวาน จ่าย ๖๐ ข้าว', '2026-10-01|รายจ่าย|ค่าอาหาร|60|ข้าว'],
  ['จ่ายค่าไฟ 1,250.50 เดือนกันยา', '2026-10-02|รายจ่าย|ค่าน้ำ-ค่าไฟ|1250.5|เดือนกันยา'],
  ['รับ 1.5k ขายเสื้อมือสอง', '2026-10-02|รายรับ|ขายของออนไลน์|1500|ขายเสื้อมือสอง'],
  ['ออม 2500 ssf', '2026-10-02|เงินออม|กองทุน SSF/RMF|2500|'],
  ['-89 กาแฟ', '2026-10-02|รายจ่าย|ค่าอาหาร|89|กาแฟ'],
];
let n = 0;
for (const [t, exp] of cases) { const p = ctx.parseEntry_(t, D, cats); assert.strictEqual([k(p.ymd), p.type, p.cat, p.amount, p.desc].join('|'), exp, t); n++; }
assert.ok(ctx.parseEntry_('31/2 จ่าย 100 อาหาร', D, cats).error.includes('ไม่ถูกต้อง')); n++;
assert.strictEqual(ctx.parseEntry_('+1', D, cats), null); n++;
assert.strictEqual(ctx.parseEntry_('รับทราบครับ', D, cats), null); n++;
assert.strictEqual(k(ctx.parseEntry_('จ่าย 100 อาหาร 30/12', { y: 2027, m: 1, d: 2 }, cats).ymd), '2026-12-30'); n++;
const x = ctx.extractDate_('25/9/2569', D); assert.strictEqual(k(x.ymd), '2026-09-25'); assert.strictEqual(x.rest, ''); n++;
console.log('parser: ผ่านทั้งหมด ' + n + ' จุด');
