// จำลอง SpreadsheetApp (Sheet/Range) ในหน่วยความจำ สำหรับรันโค้ด Apps Script บน Node.js
import assert from 'node:assert';

export const TZ = 'Asia/Bangkok';

function colToNum(s) { let n = 0; for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }
function parseA1(a1) {
  const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(a1);
  return { r: +m[2], c: colToNum(m[1]), nr: +m[4] - +m[2] + 1, nc: colToNum(m[3]) - colToNum(m[1]) + 1 };
}
class Sheet {
  constructor(name, maxRows) { this.name = name; this.maxRows = maxRows; this.v = new Map(); this.f = new Map(); this.fmt = new Map(); this.copies = []; }
  key(r, c) { return r + ',' + c; }
  get(r, c) { return this.v.has(this.key(r, c)) ? this.v.get(this.key(r, c)) : ''; }
  set(r, c, val) { if (val === '' || val === null) this.v.delete(this.key(r, c)); else this.v.set(this.key(r, c), val); }
  getMaxRows() { return this.maxRows; }
  getLastRow() {
    let last = 0;
    for (const k of [...this.v.keys(), ...this.f.keys()]) last = Math.max(last, +k.split(',')[0]);
    return last;
  }
  insertRowsAfter(after, n) { this.maxRows += n; }
  getRange(a, b, c, d) {
    const g = typeof a === 'string' ? parseA1(a) : { r: a, c: b, nr: c || 1, nc: d || 1 };
    if (g.r + g.nr - 1 > this.maxRows) throw new Error(`range beyond max rows ${g.r + g.nr - 1} > ${this.maxRows}`);
    return new Range(this, g);
  }
}
class Range {
  constructor(sh, g) { this.sh = sh; this.g = g; }
  getValues() {
    const out = [];
    for (let i = 0; i < this.g.nr; i++) { const row = []; for (let j = 0; j < this.g.nc; j++) row.push(this.sh.get(this.g.r + i, this.g.c + j)); out.push(row); }
    return out;
  }
  setValues(vals) {
    assert.strictEqual(vals.length, this.g.nr); assert.strictEqual(vals[0].length, this.g.nc);
    vals.forEach((row, i) => row.forEach((v, j) => {
      // emulate Sheets: a string starting with ' is stored as plain text without the apostrophe
      if (typeof v === 'string' && v.startsWith("'")) v = v.slice(1);
      else if (typeof v === 'string' && v.startsWith('=')) throw new Error('FORMULA INJECTION: ' + v);
      this.sh.set(this.g.r + i, this.g.c + j, v);
    }));
  }
  getFormula() { return this.sh.f.get(this.sh.key(this.g.r, this.g.c)) || ''; }
  setNumberFormat(f) { this.sh.fmt.set(this.sh.key(this.g.r, this.g.c), f); }
  clearContent() { for (let i = 0; i < this.g.nr; i++) for (let j = 0; j < this.g.nc; j++) this.sh.set(this.g.r + i, this.g.c + j, ''); }
  copyTo(dst) {
    this.sh.copies.push([this.g.r, dst.g.r]);
    for (let j = 0; j < this.g.nc; j++) {
      const f = this.sh.f.get(this.sh.key(this.g.r, this.g.c + j));
      if (f) this.sh.f.set(this.sh.key(dst.g.r, dst.g.c + j), f.replace(new RegExp('([A-Z])' + this.g.r + '\\b', 'g'), '$1' + dst.g.r));
      const v = this.sh.get(this.g.r, this.g.c + j); if (v !== '') this.sh.set(dst.g.r, dst.g.c + j, v);
    }
  }
}
export function bkk(y, m, d) { return new Date(`${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}T00:00:00+07:00`); }

export function makeWorkbook(txRowsFormatted = 5005) {
  const setup = new Sheet('Setup', 60);
  const cats = {
    B: ['เงินเดือน', 'โบนัส', 'ค่าล่วงเวลา (OT)', 'ฟรีแลนซ์/งานเสริม', 'ขายของออนไลน์', 'ดอกเบี้ย/ปันผล', 'เงินคืนภาษี', 'รายรับอื่นๆ'],
    D: ['ค่าอาหาร', 'ค่าเดินทาง', 'ค่าที่พัก/ค่าเช่า', 'ค่าน้ำ-ค่าไฟ', 'ค่าโทรศัพท์/เน็ต', 'ของใช้ในบ้าน', 'ช้อปปิ้ง', 'สุขภาพ/ยา',
      'ประกันภัย', 'สังสรรค์/บันเทิง', 'การศึกษา', 'ผ่อนชำระ/หนี้', 'ให้ครอบครัว', 'ท่องเที่ยว', 'Subscription', 'ทำบุญ/บริจาค', 'รายจ่ายอื่นๆ'],
    F: ['เงินฝากออมทรัพย์', 'เงินสำรองฉุกเฉิน', 'กองทุน SSF/RMF', 'หุ้น/ETF', 'ออมทอง', 'ออมเพื่อเกษียณ', 'ออมอื่นๆ'],
  };
  for (const [col, list] of Object.entries(cats)) list.forEach((c, i) => setup.set(6 + i, colToNum(col), c));
  const tx = new Sheet('Transaction', txRowsFormatted);
  for (let r = 6; r <= txRowsFormatted; r++) tx.f.set(tx.key(r, 7), `=IF(COUNTA(B${r}:E${r})=0,"",...)`);
  // ข้อมูลตัวอย่าง 4 แถว ลงวันที่ 1 ของเดือนปัจจุบัน (เวลาไทย) → ทดสอบได้ทุกวัน
  const [y, m] = new Intl.DateTimeFormat('en-CA', { timeZone: TZ }).format(new Date()).split('-').map(Number);
  [[bkk(y, m, 1), 'รายรับ', 'รายรับอื่นๆ', 2000, 'ขายของมือสอง'],
   [bkk(y, m, 1), 'เงินออม', 'หุ้น/ETF', 1384.1, 'DCA ETF'],
   [bkk(y, m, 1), 'รายจ่าย', 'รายจ่ายอื่นๆ', 1049.36, 'ของใช้จิปาถะ'],
   [bkk(y, m, 1), 'รายจ่าย', 'รายจ่ายอื่นๆ', 1049.36, 'ของใช้จิปาถะ']].forEach((row, i) => row.forEach((v, j) => tx.set(6 + i, 2 + j, v)));
  const sheets = { Setup: setup, Transaction: tx };
  const toasts = [];
  return { sheets, toasts, ss: { getSheetByName: n => sheets[n] || null, getSpreadsheetTimeZone: () => TZ, toast: (a, b) => toasts.push(a) } };
}
