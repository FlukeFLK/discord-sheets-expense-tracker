/**
 * 🤖 Discord → บัญชีรายรับรายจ่าย (Google Sheets) — เวอร์ชันคำสั่ง /จ่าย /รับ /ออม
 * -----------------------------------------------------------------------------
 * Discord ⇄ Cloudflare Worker (worker.js) ⇄ สคริปต์นี้ (Web App) ⇄ หน้า Transaction
 *
 * ติดตั้ง (สรุป):
 *   1) วางโค้ดนี้ใน ส่วนขยาย → Apps Script → บันทึก
 *   2) Deploy → New deployment → Web app → Execute as: Me · Who has access: Anyone
 *   3) กลับไปที่ชีต → เมนู 🤖 Discord → ① สร้างรหัสเชื่อมต่อ → ② เชื่อม Worker
 * แก้โค้ดภายหลัง: Deploy → Manage deployments → ✏️ → Version: New version → Deploy (ลิงก์เดิม)
 */

// ============================================================================
// ⚙️ ตั้งค่า (แก้ได้ถ้าโครงสร้างชีตไม่เหมือนเทมเพลต)
// ============================================================================
const CONFIG = {
  TX_SHEET: 'Transaction',          // ชื่อชีตบันทึกรายการ
  SETUP_SHEET: 'Setup',             // ชื่อชีตหมวดหมู่
  FIRST_ROW: 6,                     // แถวแรกของข้อมูลในหน้า Transaction (คอลัมน์ B–F)
  CATEGORY_RANGES: {                // ตำแหน่งรายชื่อหมวดหมู่ในหน้า Setup
    'รายรับ': 'B6:B20',
    'รายจ่าย': 'D6:D30',
    'เงินออม': 'F6:F20',
  },
  TIMEZONE: 'Asia/Bangkok',
  UNDO_LIMIT: 20,                   // จำรายการล่าสุดไว้ให้ /ยกเลิก ได้กี่รายการ
};

// คำย่อ → หมวดหมู่ (เพิ่ม/แก้ได้) ใช้ได้เมื่อหมวดปลายทางมีอยู่ในหน้า Setup
// ระบบดูว่า "คำแรกของรายละเอียด" ขึ้นต้นด้วยคำย่อไหน เช่น ข้าวมันไก่ → ข้าว → ค่าอาหาร
const ALIASES = {
  'รายรับ': {
    'โอที': 'ค่าล่วงเวลา (OT)', 'ค่าจ้าง': 'ฟรีแลนซ์/งานเสริม', 'งานนอก': 'ฟรีแลนซ์/งานเสริม',
    'ขาย': 'ขายของออนไลน์', 'ภาษี': 'เงินคืนภาษี', 'ปันผล': 'ดอกเบี้ย/ปันผล', 'ดอกเบี้ย': 'ดอกเบี้ย/ปันผล',
  },
  'รายจ่าย': {
    'ข้าว': 'ค่าอาหาร', 'กาแฟ': 'ค่าอาหาร', 'ขนม': 'ค่าอาหาร', 'ชานม': 'ค่าอาหาร', 'น้ำดื่ม': 'ค่าอาหาร',
    'ก๋วยเตี๋ยว': 'ค่าอาหาร', 'มื้อ': 'ค่าอาหาร', 'กิน': 'ค่าอาหาร', 'grabfood': 'ค่าอาหาร',
    'lineman': 'ค่าอาหาร', 'foodpanda': 'ค่าอาหาร', 'food': 'ค่าอาหาร',
    'รถ': 'ค่าเดินทาง', 'แท็กซี่': 'ค่าเดินทาง', 'taxi': 'ค่าเดินทาง', 'grab': 'ค่าเดินทาง', 'bolt': 'ค่าเดินทาง',
    'วิน': 'ค่าเดินทาง', 'bts': 'ค่าเดินทาง', 'mrt': 'ค่าเดินทาง', 'น้ำมัน': 'ค่าเดินทาง',
    'ค่าน้ำมัน': 'ค่าเดินทาง', 'ทางด่วน': 'ค่าเดินทาง', 'จอดรถ': 'ค่าเดินทาง',
    'ค่าเช่า': 'ค่าที่พัก/ค่าเช่า', 'เช่า': 'ค่าที่พัก/ค่าเช่า', 'หอพัก': 'ค่าที่พัก/ค่าเช่า', 'คอนโด': 'ค่าที่พัก/ค่าเช่า',
    'ไฟ': 'ค่าน้ำ-ค่าไฟ', 'ประปา': 'ค่าน้ำ-ค่าไฟ', 'ค่าน้ำ': 'ค่าน้ำ-ค่าไฟ',
    'มือถือ': 'ค่าโทรศัพท์/เน็ต', 'โทรศัพท์': 'ค่าโทรศัพท์/เน็ต', 'เน็ต': 'ค่าโทรศัพท์/เน็ต',
    'wifi': 'ค่าโทรศัพท์/เน็ต', 'internet': 'ค่าโทรศัพท์/เน็ต', 'ais': 'ค่าโทรศัพท์/เน็ต',
    'true': 'ค่าโทรศัพท์/เน็ต', 'dtac': 'ค่าโทรศัพท์/เน็ต',
    'ทิชชู่': 'ของใช้ในบ้าน', 'ผงซักฟอก': 'ของใช้ในบ้าน', 'สบู่': 'ของใช้ในบ้าน', 'ยาสีฟัน': 'ของใช้ในบ้าน',
    'ของใช้': 'ของใช้ในบ้าน',
    'เสื้อ': 'ช้อปปิ้ง', 'รองเท้า': 'ช้อปปิ้ง', 'กางเกง': 'ช้อปปิ้ง', 'shopee': 'ช้อปปิ้ง', 'lazada': 'ช้อปปิ้ง',
    'ร้านยา': 'สุขภาพ/ยา', 'ค่ายา': 'สุขภาพ/ยา', 'หาหมอ': 'สุขภาพ/ยา', 'โรงพยาบาล': 'สุขภาพ/ยา',
    'คลินิก': 'สุขภาพ/ยา', 'ทำฟัน': 'สุขภาพ/ยา', 'ตรวจสุขภาพ': 'สุขภาพ/ยา',
    'หนัง': 'สังสรรค์/บันเทิง', 'เหล้า': 'สังสรรค์/บันเทิง', 'เบียร์': 'สังสรรค์/บันเทิง',
    'ปาร์ตี้': 'สังสรรค์/บันเทิง', 'คอนเสิร์ต': 'สังสรรค์/บันเทิง', 'เกม': 'สังสรรค์/บันเทิง',
    'หนังสือ': 'การศึกษา', 'คอร์ส': 'การศึกษา', 'เรียน': 'การศึกษา',
    'ผ่อน': 'ผ่อนชำระ/หนี้', 'บัตรเครดิต': 'ผ่อนชำระ/หนี้',
    'พ่อแม่': 'ให้ครอบครัว',
    'เที่ยว': 'ท่องเที่ยว', 'โรงแรม': 'ท่องเที่ยว', 'ตั๋วเครื่องบิน': 'ท่องเที่ยว',
    'netflix': 'Subscription', 'spotify': 'Subscription', 'youtube': 'Subscription', 'icloud': 'Subscription',
    'บุญ': 'ทำบุญ/บริจาค', 'วัด': 'ทำบุญ/บริจาค',
  },
  'เงินออม': {
    'ฝาก': 'เงินฝากออมทรัพย์', 'สำรอง': 'เงินสำรองฉุกเฉิน', 'กองทุน': 'กองทุน SSF/RMF',
    'dca': 'หุ้น/ETF', 'ทอง': 'ออมทอง', 'เกษียณ': 'ออมเพื่อเกษียณ',
  },
};

// ============================================================================
// ค่าคงที่ภายใน
// ============================================================================
const TYPE_META = {
  'รายรับ': { color: 0x1BAF7A, icon: '💰' },
  'รายจ่าย': { color: 0xEB6834, icon: '💸' },
  'เงินออม': { color: 0x2A78D6, icon: '🏦' },
};
const COLOR = { error: 0xD03B3B, info: 0x4A3AA7, warn: 0xEDA100 };
const TYPE_WORDS = [ // เรียงคำยาวก่อน
  { w: 'รายรับ', type: 'รายรับ' }, { w: 'รายจ่าย', type: 'รายจ่าย' }, { w: 'เงินออม', type: 'เงินออม' },
  { w: 'รับ', type: 'รายรับ' }, { w: 'จ่าย', type: 'รายจ่าย' }, { w: 'ออม', type: 'เงินออม' },
  { w: '+', type: 'รายรับ', sym: true }, { w: '-', type: 'รายจ่าย', sym: true },
];
const REL_DAYS = [['เมื่อวานซืน', -2], ['เมื่อวาน', -1], ['วานนี้', -1], ['วันนี้', 0]];
const TH_MONTH = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
  'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
const TH_MON = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
const AMOUNT_RE = /(?<![A-Za-z0-9.,\/])฿?\s?(\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?)(?:\s?(k|K|พัน))?(?![A-Za-z0-9\/])(?:\s?(?:บาท|฿|thb|THB))?/;

// ============================================================================
// 📋 เมนูในชีต + การตั้งค่า
// ============================================================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🤖 Discord')
    .addItem('① สร้างรหัสเชื่อมต่อ', 'createLinkSecret')
    .addItem('② เชื่อม Worker + ลงทะเบียนคำสั่ง', 'connectWorker')
    .addItem('🔄 อัปเดตหมวดใน Discord', 'registerCommands')
    .addItem('🩺 ตรวจสอบสถานะ', 'statusDiscord')
    .addToUi();
}

/** ① สร้าง (หรือแสดง) รหัสลับที่ใช้คุยกับ Worker */
function createLinkSecret() {
  const props = PropertiesService.getScriptProperties();
  removeOldTriggers_();
  ['BOT_TOKEN', 'LAST_ID', 'OUTBOX', 'WELCOME_PENDING', 'LAST_POLL_AT', 'LAST_POLL_NOTE', 'LAST_ERROR', 'INTENT_WARNED_AT']
    .forEach(function (k) { props.deleteProperty(k); }); // ค่าจากระบบเก่า ไม่ใช้แล้ว
  props.setProperty('SPREADSHEET_ID', SpreadsheetApp.getActiveSpreadsheet().getId());
  let secret = props.getProperty('LINK_SECRET');
  if (!secret) {
    secret = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
    props.setProperty('LINK_SECRET', secret);
  }
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.7">' +
    '<p style="margin:0 0 4px"><b>รหัสเชื่อมต่อ</b> (ใช้เป็นค่า <code>SHARED_SECRET</code> ใน Cloudflare Worker):</p>' +
    '<input readonly value="' + secret + '" onclick="this.select()" style="width:100%;font-size:13px;padding:6px;font-family:monospace">' +
    '<p style="margin:12px 0 4px"><b>APPS_SCRIPT_URL</b> = ลิงก์ Web app ของสคริปต์นี้ (ลงท้าย <code>/exec</code>)</p>' +
    '<p style="margin:0;color:#555">ยังไม่มี? ที่หน้า Apps Script กด <b>Deploy → New deployment</b> → ⚙️ เลือก <b>Web app</b> → ' +
    'Execute as: <b>Me</b> · Who has access: <b>Anyone</b> → Deploy → คัดลอก <b>Web app URL</b></p>' +
    '<p style="margin:12px 0 0;color:#555">คลิกในกล่องแล้วกด Ctrl+C เพื่อคัดลอกรหัส · รหัสนี้เหมือนรหัสผ่าน อย่าแชร์</p>' +
    '</div>').setWidth(560).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, '① รหัสเชื่อมต่อ');
}

/** ② บันทึกลิงก์ Worker แล้วลงทะเบียนคำสั่งใน Discord */
function connectWorker() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('LINK_SECRET')) { ui.alert('กดเมนู ① สร้างรหัสเชื่อมต่อ ก่อนครับ'); return; }
  const old = props.getProperty('WORKER_URL');
  const r = ui.prompt('② เชื่อม Worker',
    'วางลิงก์ Worker จาก Cloudflare (เช่น https://money-bot.ชื่อคุณ.workers.dev)' + (old ? '\n(เว้นว่าง = ใช้ลิงก์เดิม)' : ''),
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const url = (r.getResponseText().trim() || old || '').replace(/\/+$/, '');
  if (!/^https:\/\/[^\s\/]+\.[^\s\/]+$/.test(url)) {
    ui.alert('ลิงก์ไม่ถูกต้อง', 'ต้องเป็นลิงก์หลักของ Worker เช่น https://money-bot.ชื่อคุณ.workers.dev', ui.ButtonSet.OK);
    return;
  }
  props.setProperties({ WORKER_URL: url, SPREADSHEET_ID: SpreadsheetApp.getActiveSpreadsheet().getId() });

  // (ไม่บังคับ) ช่องที่ใช้จด → คำสั่งขึ้นทันทีในเซิร์ฟเวอร์นั้น และถ้าใช้นอกช่องนี้ ตัวเลขจะเห็นเฉพาะคนสั่ง
  const hasCh = props.getProperty('CHANNEL_ID');
  const c = ui.prompt('② ช่อง Discord ที่ใช้จด (ไม่บังคับ)',
    'คลิกขวาที่ชื่อช่องใน Discord → Copy Link แล้ววางที่นี่' + (hasCh ? '\n(เว้นว่าง = ใช้ช่องเดิม)' : '\n(เว้นว่างได้)'),
    ui.ButtonSet.OK_CANCEL);
  if (c.getSelectedButton() === ui.Button.OK && c.getResponseText().trim()) {
    const m = /discord(?:app)?\.com\/channels\/(\d+)\/(\d+)/.exec(c.getResponseText());
    if (m) props.setProperties({ GUILD_ID: m[1], CHANNEL_ID: m[2] });
    else ui.alert('ลิงก์ช่องไม่ถูกต้อง — ข้ามขั้นนี้ไปก่อน (ลิงก์ต้องเป็นแบบ https://discord.com/channels/…/…)');
  }
  registerCommands();
}

/** 🔄 ลงทะเบียน/อัปเดตคำสั่ง (ใช้หลังแก้หมวดหมู่ในหน้า Setup) */
function registerCommands() {
  const p = PropertiesService.getScriptProperties().getProperties();
  if (!p.LINK_SECRET || !p.WORKER_URL) {
    SpreadsheetApp.getUi().alert('ทำเมนู ① และ ② ให้ครบก่อนครับ');
    return;
  }
  const r = fetchJson_(p.WORKER_URL + '/register?key=' + encodeURIComponent(p.LINK_SECRET));
  let head;
  if (r.ok && r.json && r.json.ok) {
    head = '✅ ลงทะเบียนคำสั่งแล้ว: ' + r.json.commands.map(function (c) { return '/' + c; }).join(' ') +
      (r.json.scope === 'global' ? '\n      (ลงทะเบียนแบบทุกเซิร์ฟเวอร์ อาจใช้เวลาสักพักกว่าจะขึ้นใน Discord · ใส่ลิงก์ช่องในเมนู ② เพื่อให้ขึ้นทันที)' : '');
  } else if (r.code === 403) {
    head = '❌ Worker ไม่รับรหัส — SHARED_SECRET ใน Worker ไม่ตรงกับรหัสในเมนู ①';
  } else {
    head = '❌ ลงทะเบียนคำสั่งไม่สำเร็จ: ' + ((r.json && r.json.error) || ('รหัส ' + r.code + ' ' + clip_(r.text, 120)));
  }
  showReport_(head + '\n\n' + buildReport_(), 'ผลการเชื่อมต่อ Discord');
}

/** 🩺 ตรวจทุกจุด */
function statusDiscord() {
  showReport_(buildReport_(), '🩺 สถานะการเชื่อมต่อ Discord');
}

function buildReport_() {
  const p = PropertiesService.getScriptProperties().getProperties();
  const L = [];
  const add = function (state, label, detail) {
    L.push((state === true ? '✅ ' : state === false ? '❌ ' : '⚠️ ') + label + (detail ? '\n      ' + detail : ''));
  };
  const when = function (iso) { return Utilities.formatDate(new Date(iso), CONFIG.TIMEZONE, 'dd/MM HH:mm:ss'); };

  const removed = removeOldTriggers_();
  if (removed) add(null, 'ลบตัวตั้งเวลาของระบบเก่า ' + removed + ' ตัว', 'ระบบใหม่ไม่ต้องใช้แล้ว');
  add(!!p.LINK_SECRET, '① รหัสเชื่อมต่อ', p.LINK_SECRET ? 'สร้างแล้ว' : 'ยังไม่ได้สร้าง → เมนู ① สร้างรหัสเชื่อมต่อ');
  add(!!p.WORKER_URL, '② ลิงก์ Worker', p.WORKER_URL || 'ยังไม่ได้ใส่ → เมนู ② เชื่อม Worker');
  if (p.LAST_API_AT) add(true, 'คำสั่งล่าสุดจาก Discord ' + when(p.LAST_API_AT), p.LAST_API_NOTE || '');

  if (p.LINK_SECRET && p.WORKER_URL) {
    const st = fetchJson_(p.WORKER_URL + '/status?key=' + encodeURIComponent(p.LINK_SECRET));
    if (!st.ok || !st.json) {
      add(false, 'ติดต่อ Worker ไม่ได้', st.code === 403
        ? 'SHARED_SECRET ใน Worker ไม่ตรงกับรหัสในเมนู ① → แก้ใน Cloudflare (Settings → Variables and Secrets) แล้วกด Deploy'
        : 'รหัส ' + st.code + ' — ตรวจลิงก์ Worker และว่าได้กด Deploy โค้ดใน Cloudflare แล้ว');
    } else {
      const s = st.json;
      add(true, 'Worker ทำงาน', p.WORKER_URL);
      const miss = Object.keys(s.vars || {}).filter(function (k) { return !s.vars[k]; });
      add(!miss.length, 'ตัวแปรใน Worker', miss.length ? 'ยังขาด: ' + miss.join(', ') + ' → Settings → Variables and Secrets → Add แล้วกด Deploy' : 'ครบทั้ง 4 ตัว');
      if (s.vars && s.vars.DISCORD_PUBLIC_KEY && !s.publicKeyFormat) {
        add(false, 'DISCORD_PUBLIC_KEY รูปแบบไม่ถูกต้อง', 'ต้องเป็นตัวอักษร 64 ตัว จากช่อง Public Key ในหน้า General Information ของแอป');
      }
      const sh = s.sheet || {};
      add(!!sh.ok, 'Worker → ชีต', sh.ok ? 'เชื่อมได้ (' + sh.sheet + ')'
        : (sh.error || 'ไม่ทราบสาเหตุ') + '\n      → ตรวจ APPS_SCRIPT_URL (ลงท้าย /exec) และตอน Deploy ต้องเลือก Who has access: Anyone');
      if (s.discord) {
        if (!s.discord.ok) {
          add(false, 'Bot Token ใน Worker ใช้ไม่ได้', 'รหัส ' + s.discord.status + ' ' + (s.discord.error || '') + ' → ตรวจ DISCORD_BOT_TOKEN');
        } else {
          add(true, 'แอป Discord', s.discord.appName + ' (ID ' + s.discord.appId + ')');
          add(!!s.discord.endpointOk, 'Interactions Endpoint URL', s.discord.endpointOk ? 'ชี้มาที่ Worker นี้แล้ว'
            : (s.discord.endpointUrl ? 'ตอนนี้ชี้ไปที่ ' + s.discord.endpointUrl + ' (ไม่ใช่ Worker นี้)' : 'ยังไม่ได้ตั้ง') +
              '\n      → Developer Portal → แอปของคุณ → General Information → Interactions Endpoint URL → วางลิงก์ Worker → Save Changes');
          const cmds = s.discord.commands || [];
          add(cmds.length > 0, 'คำสั่งใน Discord', cmds.length ? cmds.map(function (c) { return '/' + c; }).join(' ')
            : 'ยังไม่ได้ลงทะเบียน → เมนู 🔄 อัปเดตหมวดใน Discord' + (s.discord.commandsError ? ' (' + s.discord.commandsError + ')' : ''));
        }
      }
    }
  }
  return L.join('\n');
}

function showReport_(text, title) {
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;font-size:13px">' +
    '<p style="margin:0 0 6px">ถ่ายภาพหน้าจอ หรือคลิกในกล่อง → Ctrl+A → Ctrl+C แล้วส่งให้ผู้ช่วยได้</p>' +
    '<textarea readonly onclick="this.select()" style="width:100%;height:330px;font-size:12.5px;line-height:1.5">' +
    String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;') + '</textarea></div>').setWidth(640).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, title);
}

function fetchJson_(url) {
  try {
    const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    const text = r.getContentText();
    let json = null;
    try { json = JSON.parse(text); } catch (e) { /* ไม่ใช่ JSON */ }
    return { ok: r.getResponseCode() === 200, code: r.getResponseCode(), json: json, text: text };
  } catch (e) {
    return { ok: false, code: 0, json: null, text: e.message };
  }
}

/** ลบตัวตั้งเวลาของระบบเก่า (อ่านข้อความทุกนาที) */
function removeOldTriggers_() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'pollDiscord') { ScriptApp.deleteTrigger(t); n++; }
  });
  return n;
}
/** เผื่อตัวตั้งเวลาเก่ายังค้าง: ทำงานครั้งเดียวแล้วลบตัวเอง */
function pollDiscord() { removeOldTriggers_(); }

// ============================================================================
// 🌐 Web App — รับคำสั่งจาก Worker
// ============================================================================
function doPost(e) {
  let req = {};
  try {
    req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ error: 'รูปแบบข้อมูลไม่ถูกต้อง' });
  }
  const props = PropertiesService.getScriptProperties();
  const secret = props.getProperty('LINK_SECRET');
  if (!secret || req.secret !== secret) {
    return json_({ error: 'รหัสเชื่อมต่อไม่ตรงกัน — SHARED_SECRET ใน Worker ต้องตรงกับรหัสในเมนู ① ของชีต' });
  }
  try {
    return json_(handleApi_(req, props));
  } catch (err) {
    console.error(err);
    return json_({ error: 'ระบบในชีตขัดข้อง: ' + err.message });
  }
}

function doGet() {
  return json_({ ok: true, message: 'Apps Script พร้อมรับคำสั่งจาก Worker (ใช้ POST)' });
}

function handleApi_(req, props) {
  const ss = openSpreadsheet_(props);
  const base = { guildId: props.getProperty('GUILD_ID') || '', channelId: props.getProperty('CHANNEL_ID') || '' };
  if (req.action === 'ping') return Object.assign({ ok: true, sheet: ss.getName() }, base);

  const ctx = makeContext_(ss);
  if (req.action === 'config') return Object.assign({ ok: true, categories: ctx.cats() }, base);

  const today = ymdInTz_(new Date());
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let out, note = '';
  try {
    if (req.action === 'entry') {
      out = apiEntry_(req, ctx, today);
      note = out.note || '';
    } else if (req.action === 'text') {
      const r = processText_(String(req.text || ''), today, ctx);
      out = { embeds: [r ? r.embed : errorEmbed_('ไม่เข้าใจข้อความนี้ — ขึ้นต้นด้วย จ่าย / รับ / ออม เช่น `จ่าย 120 อาหาร ข้าวมันไก่`', req.text)] };
      note = '/จด ' + clip_(req.text, 40);
    } else if (req.action === 'summary') {
      const mo = req.month ? Number(req.month) : today.m;
      const yr = req.year ? normYear_(String(req.year)) : (req.month && mo > today.m ? today.y - 1 : today.y);
      out = { embeds: [summaryEmbed_(ctx, yr, mo)] };
      note = '/สรุป ' + mo + '/' + yr;
    } else if (req.action === 'undo') {
      out = { embeds: [undoLast_(ctx).embed] };
      note = '/ยกเลิก';
    } else {
      out = { error: 'ไม่รู้จักคำสั่ง "' + req.action + '"' };
    }
  } finally {
    lock.releaseLock();
  }
  // ใช้คำสั่งนอกช่องบัญชี → ให้ Worker ตอบแบบเห็นคนเดียว
  if (out.embeds && base.channelId && req.channelId && String(req.channelId) !== base.channelId) out.private = true;
  props.setProperties({ LAST_API_AT: new Date().toISOString(), LAST_API_NOTE: note || String(req.action) });
  delete out.note;
  return out;
}

/** /จ่าย /รับ /ออม — ค่ามาเป็นช่องแยกแล้ว */
function apiEntry_(req, ctx, today) {
  const type = String(req.type || '');
  if (!TYPE_META[type]) return { embeds: [errorEmbed_('ประเภทไม่ถูกต้อง')] };
  const amount = Math.round(Number(req.amount) * 100) / 100;
  if (!(amount > 0)) return { embeds: [errorEmbed_('จำนวนเงินต้องมากกว่า 0')] };

  const cats = ctx.cats()[type] || [];
  let cat = String(req.cat || '').trim();
  let warning = null;
  if (cats.indexOf(cat) < 0) { // หมวดถูกเปลี่ยนชื่อในหน้า Setup หลังลงทะเบียนคำสั่ง
    const c = resolveCategory_(cat, type, cats);
    if (c.error) return { embeds: [errorEmbed_('ไม่พบหมวด "' + cat + '" ในหน้า Setup — กดเมนู 🔄 อัปเดตหมวดใน Discord ในชีต')] };
    warning = 'หมวด "' + cat + '" ไม่มีในหน้า Setup แล้ว จึงบันทึกเป็น "' + c.cat + '" — กดเมนู 🔄 อัปเดตหมวดใน Discord';
    cat = c.cat;
  }
  let ymd = today;
  if (req.date) {
    const d = extractDate_(String(req.date), today);
    if (d.error || d.rest || !d.ymd) {
      return { embeds: [errorEmbed_(d.error || ('วันที่ "' + req.date + '" ไม่ถูกต้อง — ใช้ 25/9, 25/9/2026 หรือ เมื่อวาน'))] };
    }
    ymd = d.ymd;
  }
  const desc = String(req.desc || '').replace(/\s+/g, ' ').trim().slice(0, 200);
  const p = { type: type, amount: amount, cat: cat, desc: desc, ymd: ymd, warning: warning };
  const row = appendTx_(ctx, p);
  pushUndo_({ row: row, date: ymdKey_(ymd), type: type, cat: cat, amount: amount, desc: desc });
  return { embeds: [confirmEmbed_(ctx, p, row)], note: type + ' ฿' + fmt_(amount) + ' ' + cat + ' → แถว ' + row };
}

function openSpreadsheet_(props) {
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { /* web app อาจไม่มี active */ }
  if (!ss && props.getProperty('SPREADSHEET_ID')) ss = SpreadsheetApp.openById(props.getProperty('SPREADSHEET_ID'));
  if (!ss) throw new Error('ไม่พบชีต — กดเมนู ① สร้างรหัสเชื่อมต่อ ในชีตอีกครั้ง');
  return ss;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================================
// 🧠 แปลข้อความ → คำสั่ง หรือ รายการ
// ============================================================================
function processText_(text, msgYMD, ctx) {
  const t = text.replace(/\s+/g, ' ').trim();
  const low = t.toLowerCase();

  if (/^(วิธีใช้|ช่วยเหลือ|help|\?)$/.test(low)) return { embed: helpEmbed_() };
  if (/^(หมวด|หมวดหมู่)$/.test(low)) return { embed: categoriesEmbed_(ctx.cats()) };
  if (/^(ยกเลิก|undo)$/.test(low)) return undoLast_(ctx);
  const sm = /^(สรุป|summary)(?:\s+(\d{1,2})(?:\/(\d{2,4}))?)?$/.exec(low);
  if (sm) {
    const mo = sm[2] ? Number(sm[2]) : msgYMD.m;
    const yr = sm[3] ? normYear_(sm[3]) : (sm[2] && mo > msgYMD.m ? msgYMD.y - 1 : msgYMD.y);
    if (mo < 1 || mo > 12) return { embed: errorEmbed_('เดือนต้องเป็น 1–12 เช่น `สรุป 9`', text) };
    return { embed: summaryEmbed_(ctx, yr, mo) };
  }

  const p = parseEntry_(t, msgYMD, ctx.cats());
  if (!p) return null; // ไม่ใช่รายการ → เงียบไว้
  if (p.error) return { embed: errorEmbed_(p.error, text), react: '❌' };

  const row = appendTx_(ctx, p);
  pushUndo_({ row: row, date: ymdKey_(p.ymd), type: p.type, cat: p.cat, amount: p.amount, desc: p.desc });
  return { embed: confirmEmbed_(ctx, p, row), react: '✅' };
}

/** แยก "ประเภท จำนวนเงิน หมวดหมู่ รายละเอียด" (+ วันที่) — คืน null ถ้าไม่ใช่รายการ */
function parseEntry_(raw, msgYMD, catsByType) {
  // วันที่ (อยู่ตรงไหนก็ได้)
  const dx = extractDate_(raw, msgYMD);
  let s = dx.rest;
  if (!s) return null;
  const ymd = dx.ymd || msgYMD;
  const dateError = dx.error;

  // ประเภท (ต้องขึ้นต้นข้อความ)
  const low = s.toLowerCase();
  let type = null, rest = '', explicit = false, isSym = false;
  for (let i = 0; i < TYPE_WORDS.length; i++) {
    const tw = TYPE_WORDS[i];
    if (low.indexOf(tw.w) !== 0) continue;
    const after = s.slice(tw.w.length);
    if (tw.sym) {
      if (!/^\s?[\d฿]/.test(after)) continue;
      explicit = true; isSym = true;
    } else {
      explicit = after === '' || /^[\s\d฿]/.test(after);
    }
    type = tw.type; rest = after.trim(); break;
  }
  if (!type) return null;
  if (dateError) return { error: dateError, type: type };

  // จำนวนเงิน (ตัวเลขตัวแรก)
  const am = AMOUNT_RE.exec(rest);
  if (!am) return explicit ? { error: 'ไม่พบจำนวนเงิน', type: type } : null;
  const amount = Math.round(parseFloat(am[1].replace(/,/g, '')) * (am[2] ? 1000 : 1) * 100) / 100;
  if (!(amount > 0)) return { error: 'จำนวนเงินต้องมากกว่า 0', type: type };
  const text = (rest.slice(0, am.index) + ' ' + rest.slice(am.index + am[0].length))
    .replace(/\s+/g, ' ').replace(/^[\s:;,|\-–]+/, '').trim();
  if (isSym && !text) return null; // "+1" / "-5" เฉย ๆ ในแชท ไม่นับเป็นรายการ

  // หมวดหมู่ + รายละเอียด
  const c = resolveCategory_(text, type, catsByType[type] || []);
  if (c.error) return { error: c.error, type: type };
  return { type: type, amount: amount, cat: c.cat, desc: c.desc.slice(0, 200), ymd: ymd, how: c.how, warning: c.warning || null };
}

/** ดึงวันที่ออกจากข้อความ: "เมื่อวาน", "25/9", "25/9/2026", "25/9/69" — คืน { ymd, rest, error } */
function extractDate_(raw, msgYMD) {
  let s = String(raw || '').replace(/[๐-๙]/g, function (c) { return String(c.charCodeAt(0) - 0x0E50); })
    .replace(/\s+/g, ' ').trim();
  let ymd = null, error = null;
  for (let i = 0; i < REL_DAYS.length; i++) {
    const re = new RegExp('(^|\\s)' + REL_DAYS[i][0] + '(?=\\s|$)');
    if (re.test(s)) { ymd = addDays_(msgYMD, REL_DAYS[i][1]); s = s.replace(re, ' ').trim(); break; }
  }
  const dm = /(^|\s)(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?(?=\s|$)/.exec(s);
  if (dm) {
    const d = Number(dm[2]), mo = Number(dm[3]);
    let y = dm[4] ? normYear_(dm[4]) : msgYMD.y;
    if (!validYMD_(y, mo, d)) error = 'วันที่ ' + dm[0].trim() + ' ไม่ถูกต้อง (ใช้รูปแบบ วัน/เดือน เช่น 25/9)';
    else {
      if (!dm[4] && ymdKey_({ y: y, m: mo, d: d }) > ymdKey_(msgYMD)) y -= 1; // 30/12 ที่พิมพ์ต้นปี = ปีก่อน
      ymd = { y: y, m: mo, d: d };
    }
    s = (s.slice(0, dm.index) + ' ' + s.slice(dm.index + dm[0].length)).replace(/\s+/g, ' ').trim();
  }
  return { ymd: ymd, rest: s, error: error };
}

function resolveCategory_(text, type, cats) {
  const L = function (x) { return String(x).toLowerCase(); };
  const fallback = cats.filter(function (c) { return c.indexOf('อื่น') >= 0; })[0] || null;
  const t0 = String(text || '').trim();

  if (!t0) {
    return fallback ? { cat: fallback, desc: '', how: 'fallback', warning: 'ไม่ได้ระบุหมวด จึงบันทึกเป็น "' + fallback + '"' }
      : { error: 'ไม่ได้ระบุหมวดหมู่ — พิมพ์ `หมวด` เพื่อดูรายชื่อ' };
  }
  // (ก) ขึ้นต้นด้วยชื่อหมวดเต็ม
  const byLen = cats.slice().sort(function (a, b) { return b.length - a.length; });
  for (let i = 0; i < byLen.length; i++) {
    if (L(t0).indexOf(L(byLen[i])) === 0) return { cat: byLen[i], desc: t0.slice(byLen[i].length).trim(), how: 'exact' };
  }
  const tokens = t0.split(' ');
  // (ข) เลขหมวด
  if (/^\d{1,2}$/.test(tokens[0])) {
    const n = Number(tokens[0]);
    if (n >= 1 && n <= cats.length) return { cat: cats[n - 1], desc: tokens.slice(1).join(' '), how: 'number' };
  }
  // (ค) คำย่อ / (ง) ส่วนหนึ่งของชื่อหมวด — ไล่ทีละคำ
  const al = ALIASES[type] || {};
  const keys = Object.keys(al).sort(function (a, b) { return b.length - a.length; });
  let ambiguous = null;
  for (let i = 0; i < tokens.length; i++) {
    const tok = tokens[i];
    for (let k = 0; k < keys.length; k++) {
      if (L(tok).indexOf(L(keys[k])) === 0 && cats.indexOf(al[keys[k]]) >= 0) {
        return { cat: al[keys[k]], desc: t0, how: 'alias' };
      }
    }
    if (tok.length >= 2) {
      const cands = cats.filter(function (c) { return L(c).indexOf(L(tok)) >= 0; });
      if (cands.length === 1) {
        return { cat: cands[0], desc: tokens.slice(0, i).concat(tokens.slice(i + 1)).join(' '), how: 'partial' };
      }
      if (cands.length > 1 && !ambiguous) ambiguous = { tok: tok, cands: cands };
    }
  }
  if (ambiguous) return { error: '"' + ambiguous.tok + '" ตรงกับหลายหมวด: ' + ambiguous.cands.join(', ') + ' — พิมพ์ให้ชัดขึ้น' };
  if (fallback) return { cat: fallback, desc: t0, how: 'fallback', warning: 'ไม่พบหมวดที่ตรงกับ "' + tokens[0] + '" จึงบันทึกเป็น "' + fallback + '"' };
  return { error: 'ไม่พบหมวด "' + tokens[0] + '" — พิมพ์ `หมวด` เพื่อดูรายชื่อ' };
}

// ============================================================================
// 📄 อ่าน/เขียนชีต
// ============================================================================
function makeContext_(ss) {
  const ctx = { ss: ss, _cats: null, _data: null };
  ctx.cats = function () { if (!ctx._cats) ctx._cats = readCats_(ss); return ctx._cats; };
  ctx.data = function () { if (!ctx._data) ctx._data = readTx_(ss); return ctx._data; };
  return ctx;
}

function txSheet_(ss) {
  const sh = ss.getSheetByName(CONFIG.TX_SHEET);
  if (!sh) throw new Error('ไม่พบชีต "' + CONFIG.TX_SHEET + '"');
  return sh;
}

function readCats_(ss) {
  const sh = ss.getSheetByName(CONFIG.SETUP_SHEET);
  const out = {};
  Object.keys(CONFIG.CATEGORY_RANGES).forEach(function (type) {
    out[type] = !sh ? [] : sh.getRange(CONFIG.CATEGORY_RANGES[type]).getValues()
      .map(function (r) { return String(r[0]).trim(); })
      .filter(function (v) { return v !== ''; });
  });
  return out;
}

function readTx_(ss) {
  const sh = txSheet_(ss);
  const last = sh.getLastRow();
  if (last < CONFIG.FIRST_ROW) return [];
  const tz = ss.getSpreadsheetTimeZone();
  return sh.getRange(CONFIG.FIRST_ROW, 2, last - CONFIG.FIRST_ROW + 1, 4).getValues()
    .filter(function (r) { return r[0] instanceof Date && r[1] !== '' && typeof r[3] === 'number'; })
    .map(function (r) {
      return { ym: Utilities.formatDate(r[0], tz, 'yyyy-MM'), type: String(r[1]), cat: String(r[2]), amt: r[3] };
    });
}

function appendTx_(ctx, e) {
  const data = ctx.data(); // โหลดยอดเดิมก่อนเขียน (ใช้คำนวณยอดรวมเดือน)
  const sh = txSheet_(ctx.ss);
  const first = CONFIG.FIRST_ROW;
  const maxRows = sh.getMaxRows();

  // แถวถัดจากแถวสุดท้ายที่มีข้อมูลในคอลัมน์ B–F
  let row = first;
  if (maxRows >= first) {
    const vals = sh.getRange(first, 2, maxRows - first + 1, 5).getValues();
    for (let i = vals.length - 1; i >= 0; i--) {
      if (vals[i].some(function (v) { return v !== '' && v !== null; })) { row = first + i + 1; break; }
    }
  }
  // แถวเกินที่เตรียมไว้ → เพิ่มแถว และคัดลอกรูปแบบ/Drop-down/สูตรตรวจสอบจากแถวบน
  if (row > maxRows) sh.insertRowsAfter(maxRows, Math.max(500, row - maxRows));
  if (row > first && sh.getRange(row, 7).getFormula() === '') {
    sh.getRange(row - 1, 2, 1, 6).copyTo(sh.getRange(row, 2, 1, 6));
  }

  const date = Utilities.parseDate(ymdKey_(e.ymd), ctx.ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  const desc = /^[=+\-@]/.test(e.desc) ? "'" + e.desc : e.desc; // กันข้อความกลายเป็นสูตร
  sh.getRange(row, 2, 1, 5).setValues([[date, e.type, e.cat, e.amount, desc]]);
  sh.getRange(row, 2).setNumberFormat('dd/mm/yyyy');
  SpreadsheetApp.flush();

  data.push({ ym: ymdKey_(e.ymd).slice(0, 7), type: e.type, cat: e.cat, amt: e.amount });
  return row;
}

function pushUndo_(entry) {
  const props = PropertiesService.getScriptProperties();
  const stack = JSON.parse(props.getProperty('UNDO') || '[]');
  stack.push(entry);
  props.setProperty('UNDO', JSON.stringify(stack.slice(-CONFIG.UNDO_LIMIT)));
}

function undoLast_(ctx) {
  const props = PropertiesService.getScriptProperties();
  const stack = JSON.parse(props.getProperty('UNDO') || '[]');
  const e = stack.pop();
  if (!e) return { embed: infoEmbed_('ไม่มีรายการให้ยกเลิก', 'ยกเลิกได้เฉพาะรายการที่จดผ่าน Discord') };
  props.setProperty('UNDO', JSON.stringify(stack));

  const sh = txSheet_(ctx.ss);
  const v = sh.getRange(e.row, 2, 1, 5).getValues()[0];
  const tz = ctx.ss.getSpreadsheetTimeZone();
  const same = v[0] instanceof Date && Utilities.formatDate(v[0], tz, 'yyyy-MM-dd') === e.date &&
    String(v[1]) === e.type && String(v[2]) === e.cat && Number(v[3]) === e.amount && String(v[4]) === e.desc;
  if (!same) {
    return { embed: errorEmbed_('รายการล่าสุด (แถว ' + e.row + ') ถูกแก้ไขในชีตแล้ว จึงไม่ลบให้ — แก้ในชีตได้โดยตรง') };
  }
  sh.getRange(e.row, 2, 1, 5).clearContent();
  ctx._data = null;
  return {
    embed: {
      title: '↩️ ยกเลิกรายการแล้ว', color: COLOR.info,
      description: e.type + ' ฿' + fmt_(e.amount) + ' · ' + e.cat + (e.desc ? ' · ' + e.desc : '') +
        '\n📅 ' + dateLabel_(parseKey_(e.date)) + ' (แถว ' + e.row + ')',
    },
  };
}

// ============================================================================
// 💬 ข้อความตอบกลับ (Discord embeds)
// ============================================================================
function confirmEmbed_(ctx, p, row) {
  const ym = ymdKey_(p.ymd).slice(0, 7);
  const total = ctx.data().filter(function (r) { return r.ym === ym && r.type === p.type; })
    .reduce(function (a, r) { return a + r.amt; }, 0);
  const meta = TYPE_META[p.type];
  return {
    title: meta.icon + ' ' + p.type + ' ฿' + fmt_(p.amount),
    color: p.warning ? COLOR.warn : meta.color,
    description: '**' + p.cat + '**' + (p.desc ? ' · ' + clip_(p.desc, 300) : '') +
      '\n📅 ' + dateLabel_(p.ymd) + (p.warning ? '\n⚠️ ' + p.warning + ' (`/ยกเลิก` เพื่อลบ)' : ''),
    footer: { text: p.type + ' ' + TH_MON[p.ymd.m - 1] + ' ' + p.ymd.y + ' รวม ฿' + fmt_(total) + ' · แถว ' + row },
  };
}

function errorEmbed_(reason, original) {
  const e = {
    title: '❌ บันทึกไม่ได้', color: COLOR.error,
    description: reason + '\n\nใช้ `/จ่าย` `/รับ` `/ออม` หรือ `/จด จ่าย 120 อาหาร ข้าวมันไก่` · `/จด วิธีใช้` ดูทั้งหมด',
  };
  if (original) e.footer = { text: 'ข้อความ: ' + clip_(original, 150) };
  return e;
}

function infoEmbed_(title, text) {
  return { title: 'ℹ️ ' + title, color: COLOR.info, description: text };
}

function helpEmbed_(title) {
  return {
    title: title || '📒 วิธีจดรายรับ รายจ่าย เงินออม', color: COLOR.info,
    description: [
      '**แบบเลือกทีละช่อง**',
      '`/จ่าย` → ใส่จำนวน → เลือกหมวด → (รายละเอียด) (วันที่)',
      '`/รับ` และ `/ออม` ใช้แบบเดียวกัน',
      '',
      '**แบบพิมพ์รวดเดียว** ด้วย `/จด`',
      '```',
      'จ่าย 120 อาหาร ข้าวมันไก่',
      'รับ 42000 เงินเดือน',
      'ออม 3000 หุ้น DCA',
      'จ่าย 850 ไฟ 25/9',
      '```',
      '• **วันที่**: ไม่ใส่ = วันนี้ · ใส่ `25/9` `25/9/2026` หรือ `เมื่อวาน` ได้',
      '• `/สรุป` ยอดเดือนนี้ (เลือกเดือนได้) · `/ยกเลิก` ลบรายการล่าสุด',
      '• `/จด หมวด` ดูรายชื่อหมวด · เพิ่มหมวดที่หน้า Setup แล้วกดเมนู 🔄 ในชีต',
    ].join('\n'),
  };
}

function categoriesEmbed_(cats) {
  const fields = Object.keys(CONFIG.CATEGORY_RANGES).map(function (type) {
    const list = cats[type] || [];
    return {
      name: TYPE_META[type].icon + ' ' + type,
      value: list.length ? clip_(list.map(function (c, i) { return (i + 1) + '. ' + c; }).join('\n'), 1000) : '-',
      inline: true,
    };
  });
  return {
    title: '📂 หมวดหมู่ (จากหน้า Setup)', color: COLOR.info, fields: fields,
    footer: { text: 'ใน /จด ใช้เลขแทนชื่อได้ เช่น "จ่าย 120 1 ข้าวมันไก่" · เพิ่มหมวดที่หน้า Setup แล้วกดเมนู 🔄 ในชีต' },
  };
}

function summaryEmbed_(ctx, y, m) {
  const ym = y + '-' + pad2_(m);
  const rows = ctx.data().filter(function (r) { return r.ym === ym; });
  const sum = function (type) {
    return rows.filter(function (r) { return r.type === type; }).reduce(function (a, r) { return a + r.amt; }, 0);
  };
  const inc = sum('รายรับ'), exp = sum('รายจ่าย'), sav = sum('เงินออม'), bal = inc - exp - sav;
  const byCat = {};
  rows.filter(function (r) { return r.type === 'รายจ่าย'; })
    .forEach(function (r) { byCat[r.cat] = (byCat[r.cat] || 0) + r.amt; });
  const top = Object.keys(byCat).sort(function (a, b) { return byCat[b] - byCat[a]; }).slice(0, 3)
    .map(function (c, i) { return (i + 1) + '. ' + c + ' ฿' + fmt_(byCat[c]); });
  return {
    title: '📊 สรุปเดือน' + TH_MONTH[m - 1] + ' ' + y, color: bal < 0 ? COLOR.error : COLOR.info,
    fields: [
      { name: '💰 รายรับ', value: '฿' + fmt_(inc), inline: true },
      { name: '💸 รายจ่าย', value: '฿' + fmt_(exp), inline: true },
      { name: '🏦 เงินออม', value: '฿' + fmt_(sav), inline: true },
      { name: '💵 คงเหลือ', value: '฿' + fmt_(bal), inline: true },
      { name: '📈 อัตราการออม', value: inc > 0 ? (Math.round(sav / inc * 1000) / 10) + '%' : '-', inline: true },
      { name: '🧾 จำนวนรายการ', value: String(rows.length), inline: true },
      { name: '🔥 จ่ายมากสุด', value: top.length ? top.join('\n') : '-', inline: false },
    ],
  };
}

// ============================================================================
// 🔧 ตัวช่วย
// ============================================================================
function cmpId_(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return a.length - b.length;
  return a < b ? -1 : a > b ? 1 : 0;
}
function pad2_(n) { return (n < 10 ? '0' : '') + n; }
function ymdKey_(o) { return o.y + '-' + pad2_(o.m) + '-' + pad2_(o.d); }
function parseKey_(s) { const p = s.split('-').map(Number); return { y: p[0], m: p[1], d: p[2] }; }
function dateLabel_(o) { return pad2_(o.d) + '/' + pad2_(o.m) + '/' + o.y; }
function ymdInTz_(date) { return parseKey_(Utilities.formatDate(date, CONFIG.TIMEZONE, 'yyyy-MM-dd')); }
function addDays_(o, n) {
  const t = new Date(Date.UTC(o.y, o.m - 1, o.d) + n * 86400000);
  return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate() };
}
function validYMD_(y, m, d) {
  if (m < 1 || m > 12 || d < 1 || y < 2000 || y > 2100) return false;
  const t = new Date(Date.UTC(y, m - 1, d));
  return t.getUTCMonth() === m - 1 && t.getUTCDate() === d;
}
function normYear_(s) {
  let y = Number(s);
  if (String(s).length <= 2) y = y < 40 ? 2000 + y : 2500 + y - 543; // 26 → 2026, 69 → 2569 → 2026
  else if (y > 2400) y -= 543;                                       // พ.ศ. → ค.ศ.
  return y;
}
function fmt_(n) {
  const neg = n < 0;
  const v = Math.round(Math.abs(n) * 100) / 100;
  const parts = (Number.isInteger(v) ? v.toFixed(0) : v.toFixed(2)).split('.');
  return (neg ? '-' : '') + parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (parts[1] ? '.' + parts[1] : '');
}
function clip_(s, n) { s = String(s); return s.length > n ? s.slice(0, n - 1) + '…' : s; }
