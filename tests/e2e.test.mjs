// End-to-end: คำขอจาก Discord ที่เซ็นด้วย Ed25519 จริง → worker/worker.js → apps-script/Code.gs (doPost) → ชีตจำลอง
import fs from 'node:fs';
import vm from 'node:vm';
import assert from 'node:assert';
import { makeWorkbook, TZ } from './helpers/sheet-mock.mjs';
import worker from '../worker/worker.js';

const GS = fs.readFileSync(new URL('../apps-script/Code.gs', import.meta.url), 'utf8');

let pass = 0;
const ok = (c, m) => { assert.ok(c, m); pass++; };

// ---- Apps Script runtime
function makeGas(wb, initialProps = {}) {
  const props = new Map(Object.entries(initialProps));
  const triggers = [{ getHandlerFunction: () => 'pollDiscord' }]; // leftover from old version
  const urlCache = new Map();
  const ui = { dialogs: [], alerts: [], answers: [] };
  const fmtParts = (date, tz) => Object.fromEntries(new Intl.DateTimeFormat('en-CA', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit' })
    .formatToParts(date).map(x => [x.type, x.value]));
  const ctx = {
    console: { log() {}, warn() {}, error() {} },
    JSON, Math, Date, String, Number, Object, RegExp, Error, encodeURIComponent, parseFloat,
    Utilities: {
      formatDate(d, tz, f) {
        const p = fmtParts(d, tz);
        if (f === 'yyyy-MM-dd') return `${p.year}-${p.month}-${p.day}`;
        if (f === 'yyyy-MM') return `${p.year}-${p.month}`;
        if (f === 'dd/MM HH:mm:ss') return `${p.day}/${p.month} ` + new Intl.DateTimeFormat('en-GB', { timeZone: tz, hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false }).format(d);
        throw new Error('fmt ' + f);
      },
      parseDate(s, tz, f) { assert.strictEqual(tz, TZ); return new Date(s + 'T00:00:00+07:00'); },
      sleep() {},
      getUuid: () => crypto.randomUUID(),
    },
    PropertiesService: { getScriptProperties: () => ({
      getProperty: k => (props.has(k) ? props.get(k) : null), setProperty: (k, v) => props.set(k, String(v)),
      setProperties: o => Object.entries(o).forEach(([k, v]) => props.set(k, String(v))), deleteProperty: k => props.delete(k),
      getProperties: () => Object.fromEntries(props),
    }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, tryLock: () => true, releaseLock() {} }) },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => Object.assign(wb.ss, { getId: () => 'SHEET123', getName: () => 'บัญชีรายรับรายจ่าย' }),
      openById: () => wb.ss, flush() {},
      getUi: () => ({
        Button: { OK: 'OK' }, ButtonSet: { OK: 'OK', OK_CANCEL: 'OKC' },
        prompt: () => { const a = ui.answers.shift(); return { getSelectedButton: () => 'OK', getResponseText: () => a }; },
        alert: (...a) => ui.alerts.push(a.join(' | ')),
        showModalDialog: (h, t) => ui.dialogs.push({ title: t, html: h.html }),
        createMenu: () => ({ addItem() { return this; }, addSeparator() { return this; }, addToUi() {} }),
      }),
    },
    ScriptApp: { getProjectTriggers: () => triggers.slice(), deleteTrigger: t => triggers.splice(triggers.indexOf(t), 1) },
    HtmlService: { createHtmlOutput: h => ({ html: h, setWidth() { return this; }, setHeight() { return this; } }) },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: s => ({ content: s, setMimeType() { return this; } }) },
    UrlFetchApp: { fetch(url) {
      if (!urlCache.has(url)) throw new Error('UrlFetchApp: no prepared response for ' + url);
      const r = urlCache.get(url);
      return { getResponseCode: () => r.status, getContentText: () => r.body };
    } },
  };
  vm.createContext(ctx);
  vm.runInContext(GS, ctx);
  return { ctx, props, triggers, ui, urlCache };
}

// ---- network mock for the Worker's fetch()
const SCRIPT_URL = 'https://script.google.com/macros/s/AKfy-test/exec';
const ORIGIN = 'https://money-bot.example.workers.dev';
const net = { webhook: [], registered: [], discordCalls: [], scriptMode: 'ok', endpoint: ORIGIN, inGuild: true };
let gas; // current Apps Script instance
globalThis.fetch = async (url, opt = {}) => {
  const method = (opt.method || 'GET').toUpperCase();
  if (url === SCRIPT_URL) {
    if (net.scriptMode === 'html') return new Response('<html><body>Sign in – Google Accounts</body></html>', { status: 200 });
    const out = gas.ctx.doPost({ postData: { contents: opt.body } });
    return new Response(out.content, { status: 200 });
  }
  assert.ok(url.startsWith('https://discord.com/api/v10/'), 'unexpected url ' + url);
  assert.ok((opt.headers || {})['User-Agent'].startsWith('DiscordBot ('), 'proper UA');
  const path = url.slice('https://discord.com/api/v10'.length);
  net.discordCalls.push(method + ' ' + path);
  const res = (status, body) => new Response(body === undefined ? '' : JSON.stringify(body), { status });
  if (path.startsWith('/webhooks/')) { net.webhook.push({ method, path, body: JSON.parse(opt.body) }); return res(200, { id: '1' }); }
  assert.strictEqual(opt.headers.Authorization, 'Bot BOT.TOKEN.X');
  if (path === '/applications/@me') return res(200, { id: '999', name: 'Expense Bot', interactions_endpoint_url: net.endpoint });
  if (path === '/users/@me/guilds') return res(200, net.multiGuild ? [{ id: '888', name: 'My Server' }, { id: '777', name: 'Other' }] : [{ id: '888', name: 'My Server' }]);
  if (path === '/applications/999/guilds/888/commands') {
    if (method === 'PUT') {
      if (!net.inGuild) return res(403, { message: 'Missing Access', code: 50001 });
      net.registered = JSON.parse(opt.body); return res(200, net.registered.map((c, i) => Object.assign({ id: String(i) }, c)));
    }
    return res(200, net.registered);
  }
  if (path === '/applications/999/commands' && method === 'GET') return res(200, []);
  throw new Error('unhandled ' + method + ' ' + path);
};

// ---- Discord signing
const kp = await crypto.subtle.generateKey({ name: 'Ed25519' }, true, ['sign', 'verify']);
const pubHex = Buffer.from(await crypto.subtle.exportKey('raw', kp.publicKey)).toString('hex');
const otherKp = await crypto.subtle.generateKey({ name: 'Ed25519' }, true, ['sign', 'verify']);
async function signed(bodyObj, key = kp.privateKey) {
  const body = JSON.stringify(bodyObj), ts = String(Math.floor(Date.now() / 1000));
  const sig = Buffer.from(await crypto.subtle.sign({ name: 'Ed25519' }, key, new TextEncoder().encode(ts + body))).toString('hex');
  return new Request(ORIGIN + '/', { method: 'POST', headers: { 'X-Signature-Ed25519': sig, 'X-Signature-Timestamp': ts, 'Content-Type': 'application/json' }, body });
}
const SECRET = 'a'.repeat(64);
const env = { DISCORD_PUBLIC_KEY: pubHex, DISCORD_BOT_TOKEN: 'BOT.TOKEN.X', APPS_SCRIPT_URL: SCRIPT_URL, SHARED_SECRET: SECRET };
async function call(req, e = env) {
  const waits = [];
  const res = await worker.fetch(req, e, { waitUntil: p => waits.push(p) });
  await Promise.all(waits);
  const t = await res.text();
  let j = null; try { j = JSON.parse(t); } catch {}
  return { status: res.status, text: t, json: j };
}
let seq = 0;
const cmd = (name, options = [], channel = '1234567890123456789') => ({
  type: 2, id: String(++seq), application_id: '999', token: 'tok' + seq, channel_id: channel,
  member: { user: { id: '42', username: 'demo', global_name: 'Demo' } },
  data: { name, type: 1, options: options.map(([n, v]) => ({ name: n, value: v })) },
});
const tx = (wb, r) => wb.sheets.Transaction.getRange(r, 2, 1, 5).getValues()[0];
const ymd = d => new Intl.DateTimeFormat('en-CA', { timeZone: TZ }).format(d);

// =========================================================== setup from the sheet menu
const wb = makeWorkbook();
gas = makeGas(wb, { BOT_TOKEN: 'old', LAST_ID: '1', WEBHOOK_URL: 'https://discord.com/api/webhooks/1/x' }); // ค่าค้างจากระบบเวอร์ชันเก่า
gas.ctx.createLinkSecret();
const secret = gas.props.get('LINK_SECRET');
ok(/^[0-9a-f]{64}$/.test(secret), 'secret is 64 hex');
ok(gas.triggers.length === 0, 'old poll trigger removed');
ok(!gas.props.has('BOT_TOKEN') && !gas.props.has('LAST_ID'), 'old bot token removed from sheet');
ok(gas.ui.dialogs[0].html.includes(secret) && gas.ui.dialogs[0].html.includes('Who has access: <b>Anyone</b>'), 'secret dialog');
gas.ctx.createLinkSecret();
ok(gas.props.get('LINK_SECRET') === secret, 'secret is stable on re-open');
env.SHARED_SECRET = secret;

// =========================================================== Worker basics
ok((await call(new Request(ORIGIN + '/'))).text.includes('ทำงานอยู่'), 'GET /');
ok((await call(new Request(ORIGIN + '/register'))).status === 403, 'register needs key');
ok((await call(new Request(ORIGIN + '/status?key=wrong'))).status === 403, 'status needs key');
ok((await call(new Request(ORIGIN + '/', { method: 'POST', body: '{}' }))).status === 401, 'unsigned → 401');
ok((await call(await signed({ type: 1 }, otherKp.privateKey))).status === 401, 'wrong key signature → 401');
const tampered = await signed({ type: 1 });
const tBody = '{"type":1,"x":1}';
ok((await call(new Request(ORIGIN + '/', { method: 'POST', headers: tampered.headers, body: tBody }))).status === 401, 'tampered body → 401');
const ping = await call(await signed({ type: 1 }));
ok(ping.status === 200 && ping.json.type === 1, 'PING → PONG');
ok((await call(await signed({ type: 1 }), Object.assign({}, env, { DISCORD_PUBLIC_KEY: '' }))).status === 401, 'missing public key → 401');

// =========================================================== register + status via sheet menu
async function prep(path) {
  const u = ORIGIN + path + '?key=' + encodeURIComponent(secret);
  const r = await call(new Request(u));
  gas.urlCache.set(u, { status: r.status, body: r.text });
  return r;
}
gas.ui.answers.push(ORIGIN + '/', 'https://discord.com/channels/888/1234567890123456789');
const reg = await prep('/register');
await prep('/status');
ok(reg.json.ok && reg.json.scope === 'guild' && reg.json.commands.join(',') === 'จ่าย,รับ,ออม,จด,สรุป,ยกเลิก', 'register ok: ' + reg.text);
gas.ctx.connectWorker();
ok(gas.props.get('WORKER_URL') === ORIGIN, 'worker url saved without trailing slash');
ok(gas.props.get('GUILD_ID') === '888' && gas.props.get('CHANNEL_ID') === '1234567890123456789', 'channel link parsed');
net.multiGuild = true; // บอทอยู่หลายเซิร์ฟเวอร์ → ต้องใช้ GUILD_ID จากลิงก์ช่อง
const reg3 = await call(new Request(ORIGIN + '/register?key=' + secret));
ok(reg3.json.ok && reg3.json.scope === 'guild' && reg3.json.guildId === '888', 'guild from channel link');
net.multiGuild = false;
ok(gas.ui.dialogs.at(-1).html.includes('✅ ลงทะเบียนคำสั่งแล้ว: /จ่าย /รับ /ออม /จด /สรุป /ยกเลิก'), 'connect report header');
await prep('/status'); // (mock) คำนวณสถานะใหม่หลังบันทึกลิงก์ช่องแล้ว
gas.ctx.statusDiscord();
const report = gas.ui.dialogs.at(-1).html;
for (const s of ['✅ Worker ทำงาน', '✅ ตัวแปรใน Worker', '✅ Worker → ชีต', 'เชื่อมได้ (บัญชีรายรับรายจ่าย)',
  '✅ แอป Discord', 'Expense Bot (ID 999)', '✅ Interactions Endpoint URL', '✅ คำสั่งใน Discord']) ok(report.includes(s), 'report has ' + s + '\n' + report);

// Discord command schema validity
const NAME_RE = /^[-_\p{L}\p{N}\p{sc=Deva}\p{sc=Thai}]{1,32}$/u;
for (const c of net.registered) {
  ok(NAME_RE.test(c.name) && c.description.length >= 1 && c.description.length <= 100, 'cmd ' + c.name);
  let seenOptional = false;
  for (const o of c.options || []) {
    ok(NAME_RE.test(o.name) && o.description.length <= 100, 'opt ' + o.name);
    if (!o.required) seenOptional = true; else ok(!seenOptional, 'required options first');
    if (o.choices) ok(o.choices.length <= 25 && o.choices.every(x => x.name.length <= 100), 'choices ok');
  }
}
const expenseChoices = net.registered[0].options[1].choices.map(c => c.value);
ok(expenseChoices.length === 17 && expenseChoices[0] === 'ค่าอาหาร', 'expense choices from Setup');

// status: endpoint not set / wrong secret
net.endpoint = '';
await prep('/status');
gas.ctx.statusDiscord();
ok(gas.ui.dialogs.at(-1).html.includes('❌ Interactions Endpoint URL') && gas.ui.dialogs.at(-1).html.includes('ยังไม่ได้ตั้ง'), 'status: endpoint missing');
net.endpoint = ORIGIN + '/';
// note: Discord stores what the user typed; trailing slash must still count as OK
await prep('/status');
gas.ctx.statusDiscord();
ok(gas.ui.dialogs.at(-1).html.includes('✅ Interactions Endpoint URL'), 'status: endpoint with trailing slash ok');
net.endpoint = ORIGIN;

// =========================================================== commands end-to-end
async function run(c) {
  const before = net.webhook.length;
  const r = await call(await signed(c));
  assert.strictEqual(r.json.type, 5, 'deferred');
  return net.webhook.slice(before);
}
let w = await run(cmd('จ่าย', [['จำนวน', 120], ['หมวด', 'ค่าอาหาร'], ['รายละเอียด', 'ข้าวมันไก่']]));
let v = tx(wb, 10);
ok(ymd(v[0]) === ymd(new Date()) && v[1] === 'รายจ่าย' && v[2] === 'ค่าอาหาร' && v[3] === 120 && v[4] === 'ข้าวมันไก่', '/จ่าย row: ' + JSON.stringify(v));
ok(w.length === 1 && w[0].method === 'PATCH' && w[0].path.endsWith('/messages/@original'), 'edits original reply');
ok(w[0].body.embeds[0].title === '💸 รายจ่าย ฿120' && w[0].body.embeds[0].footer.text.includes('แถว 10'), 'confirmation embed: ' + JSON.stringify(w[0].body));
ok(gas.props.get('LAST_API_NOTE') === 'รายจ่าย ฿120 ค่าอาหาร → แถว 10', 'last api note');

w = await run(cmd('รับ', [['จำนวน', 42000], ['หมวด', 'เงินเดือน']]));
ok(tx(wb, 11)[1] === 'รายรับ' && tx(wb, 11)[3] === 42000 && tx(wb, 11)[4] === '', '/รับ row');
w = await run(cmd('ออม', [['จำนวน', 3000.5], ['หมวด', 'หุ้น/ETF'], ['รายละเอียด', 'DCA'], ['วันที่', 'เมื่อวาน']]));
const yest = new Date(Date.now() - 86400000);
ok(tx(wb, 12)[3] === 3000.5 && ymd(tx(wb, 12)[0]) === ymd(yest), '/ออม with date เมื่อวาน');

w = await run(cmd('จด', [['ข้อความ', 'จ่าย 850 ไฟ 25/9']]));
ok(tx(wb, 13)[2] === 'ค่าน้ำ-ค่าไฟ' && ymd(tx(wb, 13)[0]).endsWith('-09-25'), '/จด free text');
w = await run(cmd('จด', [['ข้อความ', 'หมวด']]));
ok(w[0].body.embeds[0].title.startsWith('📂') && tx(wb, 14).every(x => x === ''), '/จด หมวด lists categories');
w = await run(cmd('จด', [['ข้อความ', 'สวัสดี']]));
ok(w[0].body.embeds[0].title === '❌ บันทึกไม่ได้' && w[0].body.embeds[0].description.includes('ไม่เข้าใจ'), '/จด nonsense → explain');

// private use outside the finance channel
w = await run(cmd('จ่าย', [['จำนวน', 59], ['หมวด', 'ค่าอาหาร']], '111'));
ok(tx(wb, 14)[3] === 59, 'recorded even outside channel');
ok(w.length === 2 && w[0].body.content.startsWith('🔒') && w[0].body.embeds.length === 0 && w[1].method === 'POST' && w[1].body.flags === 64, 'outside channel → hidden details');

// errors
w = await run(cmd('รับ', [['จำนวน', 100], ['หมวด', 'เงินเดือน'], ['วันที่', '31/2']]));
ok(w[0].body.embeds[0].title === '❌ บันทึกไม่ได้' && w[0].body.embeds[0].description.includes('31/2') && tx(wb, 15).every(x => x === ''), 'bad date rejected');
w = await run(cmd('รับ', [['จำนวน', 100], ['หมวด', 'เงินเดือน'], ['วันที่', 'พรุ่งนี้']]));
ok(w[0].body.embeds[0].description.includes('ไม่ถูกต้อง') && tx(wb, 15).every(x => x === ''), 'unknown date word rejected');
// category renamed in Setup after registration
wb.sheets.Setup.set(6, 4, 'อาหาร/เครื่องดื่ม');
w = await run(cmd('จ่าย', [['จำนวน', 45], ['หมวด', 'ค่าอาหาร']]));
ok(tx(wb, 15)[2] === 'รายจ่ายอื่นๆ' && w[0].body.embeds[0].description.includes('อัปเดตหมวดใน Discord'), 'renamed category → อื่นๆ + hint: ' + JSON.stringify(tx(wb, 15)) + JSON.stringify(w[0].body.embeds[0]));
wb.sheets.Setup.set(6, 4, 'ค่าอาหาร');

// summary + undo
w = await run(cmd('สรุป'));
const f = Object.fromEntries(w[0].body.embeds[0].fields.map(x => [x.name, x.value]));
ok(w[0].body.embeds[0].title.startsWith('📊 สรุปเดือน') && f['💰 รายรับ'] === '฿44,000', 'summary: ' + JSON.stringify(f));
const sepYear = Number(ymd(tx(wb, 13)[0]).slice(0, 4));
w = await run(cmd('สรุป', [['เดือน', 9], ['ปี', sepYear + 543]]));
ok(w[0].body.embeds[0].title === '📊 สรุปเดือนกันยายน ' + sepYear && w[0].body.embeds[0].fields[1].value === '฿850', 'summary Sep (พ.ศ. → ค.ศ.)');
w = await run(cmd('ยกเลิก'));
ok(tx(wb, 15).every(x => x === '') && w[0].body.embeds[0].title.includes('ยกเลิก'), '/ยกเลิก removes last');

// Apps Script misconfig → clear error in Discord
const savedSecret = env.SHARED_SECRET;
env.SHARED_SECRET = 'wrong';
w = await run(cmd('จ่าย', [['จำนวน', 1], ['หมวด', 'ค่าอาหาร']]));
ok(w[0].body.embeds[0].description.includes('รหัสเชื่อมต่อไม่ตรงกัน'), 'wrong secret explained');
env.SHARED_SECRET = savedSecret;
net.scriptMode = 'html';
w = await run(cmd('จ่าย', [['จำนวน', 1], ['หมวด', 'ค่าอาหาร']]));
ok(w[0].body.embeds[0].description.includes('Who has access: Anyone'), 'not public deployment explained');
await prep('/status');
gas.ctx.statusDiscord();
ok(gas.ui.dialogs.at(-1).html.includes('❌ Worker → ชีต'), 'status shows sheet link problem\n' + gas.ui.dialogs.at(-1).html);
net.scriptMode = 'ok';
ok(tx(wb, 15).every(x => x === ''), 'nothing written on failures');

// bot not in server → register explains
net.inGuild = false;
const reg2 = await call(new Request(ORIGIN + '/register?key=' + secret));
ok(!reg2.json.ok && reg2.json.error.includes('เชิญบอท'), 'register 403 explained');
net.inGuild = true;

// Apps Script direct guards
ok(JSON.parse(gas.ctx.doPost({ postData: { contents: '{"secret":"x","action":"ping"}' } }).content).error.includes('ไม่ตรงกัน'), 'doPost rejects bad secret');
ok(JSON.parse(gas.ctx.doPost({ postData: { contents: 'not json' } }).content).error, 'doPost rejects bad json');
const cfg = JSON.parse(gas.ctx.doPost({ postData: { contents: JSON.stringify({ secret, action: 'config' }) } }).content);
ok(cfg.ok && cfg.guildId === '888' && cfg.categories['เงินออม'].length === 7, 'config action');
ok(!net.discordCalls.some(c => c.includes('/channels/')), 'never calls the blocked channel endpoints');

// ลิงก์ช่องผิดรูปแบบ → แจ้งเตือนและไม่เปลี่ยนค่าเดิม
gas.ui.answers.push('', 'not a link');
gas.ctx.connectWorker();
ok(gas.ui.alerts.some(a => a.includes('ลิงก์ช่องไม่ถูกต้อง')) && gas.props.get('CHANNEL_ID') === '1234567890123456789', 'bad channel link ignored');

console.log(`e2e: ผ่านทั้งหมด ${pass} จุด`);
