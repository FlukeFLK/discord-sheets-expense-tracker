/**
 * 🤖 Cloudflare Worker — Discord slash commands (/จ่าย /รับ /ออม) → Google Sheets
 * -----------------------------------------------------------------------------
 * Discord ส่งคำสั่งมาที่ Worker นี้ → Worker ส่งต่อให้ Apps Script ในชีต → แก้ข้อความตอบกลับใน Discord
 *
 * ตั้งค่าที่ Cloudflare: Worker → Settings → Variables and Secrets (ครบ 4 ตัว แล้วกด Deploy)
 *   DISCORD_PUBLIC_KEY  (Text)    Public Key จาก Discord Developer Portal → แอป → General Information
 *   DISCORD_BOT_TOKEN   (Secret)  Bot Token จากหน้า Bot
 *   APPS_SCRIPT_URL     (Text)    Web app URL ของ Apps Script (ลงท้าย /exec)
 *   SHARED_SECRET       (Secret)  รหัสจากเมนู 🤖 Discord → ① ในชีต
 * แล้วนำลิงก์ Worker นี้ไปใส่ที่ Developer Portal → General Information → Interactions Endpoint URL
 */

const API = 'https://discord.com/api/v10';
const UA = 'DiscordBot (https://workers.cloudflare.com, 1.0)';
const ENTRY = { 'จ่าย': 'รายจ่าย', 'รับ': 'รายรับ', 'ออม': 'เงินออม' };
const MANAGE_SERVER = '32'; // ค่าเริ่มต้น: เฉพาะแอดมินเซิร์ฟเวอร์ใช้คำสั่งได้ (ปรับได้ใน Server Settings → Integrations)

export default {
  async fetch(request, env, ctx) {
    try {
      const url = new URL(request.url);
      if (request.method === 'GET') {
        if (url.pathname === '/register') return denied(url, env) || await register(env);
        if (url.pathname === '/status') return denied(url, env) || await status(env, url.origin);
        return text('✅ Worker บัญชีรายรับรายจ่าย ทำงานอยู่');
      }
      if (request.method !== 'POST') return text('Method not allowed', 405);
      return await interaction(request, env, ctx);
    } catch (e) {
      return text('Worker error: ' + (e && e.message), 500);
    }
  },
};

// ---------------------------------------------------------------- Discord → Worker
async function interaction(request, env, ctx) {
  const sig = request.headers.get('X-Signature-Ed25519');
  const ts = request.headers.get('X-Signature-Timestamp');
  const body = await request.text();
  if (!env.DISCORD_PUBLIC_KEY || !sig || !ts || !(await verifySignature(env.DISCORD_PUBLIC_KEY, sig, ts + body))) {
    return text('invalid request signature', 401);
  }
  const i = JSON.parse(body);
  if (i.type === 1) return json({ type: 1 }); // PING ตอนบันทึก Interactions Endpoint URL
  if (i.type !== 2) return json({ type: 4, data: { content: 'ยังไม่รองรับการกระทำนี้', flags: 64 } });

  const payload = toPayload(i.data || {});
  if (!payload) return json({ type: 4, data: { content: 'ไม่รู้จักคำสั่งนี้', flags: 64 } });

  // ตอบ Discord ทันที ("กำลังคิด…") แล้วค่อยแก้ข้อความเมื่อชีตบันทึกเสร็จ
  ctx.waitUntil(finish(i, env, payload));
  return json({ type: 5 });
}

function toPayload(d) {
  const o = {};
  (d.options || []).forEach(function (x) { o[x.name] = x.value; });
  if (ENTRY[d.name]) {
    return { action: 'entry', type: ENTRY[d.name], amount: o['จำนวน'], cat: o['หมวด'] || '', desc: o['รายละเอียด'] || '', date: o['วันที่'] || '' };
  }
  if (d.name === 'จด') return { action: 'text', text: o['ข้อความ'] || '' };
  if (d.name === 'สรุป') return { action: 'summary', month: o['เดือน'] || null, year: o['ปี'] || null };
  if (d.name === 'ยกเลิก') return { action: 'undo' };
  return null;
}

async function finish(i, env, payload) {
  const user = (i.member && i.member.user) || i.user || {};
  payload.channelId = i.channel_id || (i.channel && i.channel.id) || '';
  payload.user = user.global_name || user.username || '';
  const res = await callSheet(env, payload);
  const embeds = res.embeds && res.embeds.length ? res.embeds : [errorEmbed(res.error || 'ไม่ได้รับผลลัพธ์จากชีต')];
  const hook = API + '/webhooks/' + i.application_id + '/' + i.token;
  if (res.private) {
    // ใช้คำสั่งนอกช่องบัญชี → ข้อความในช่องไม่บอกตัวเลข รายละเอียดเห็นเฉพาะคนสั่ง
    await discord('PATCH', hook + '/messages/@original', { content: '🔒 บันทึกแล้ว — รายละเอียดอยู่ในข้อความที่เห็นเฉพาะคุณ', embeds: [] });
    await discord('POST', hook, { embeds: embeds, flags: 64 });
  } else {
    await discord('PATCH', hook + '/messages/@original', { embeds: embeds, allowed_mentions: { parse: [] } });
  }
}

// ---------------------------------------------------------------- Worker → Apps Script
async function callSheet(env, payload) {
  if (!env.APPS_SCRIPT_URL) return { error: 'ยังไม่ได้ตั้ง APPS_SCRIPT_URL ใน Worker' };
  try {
    const r = await fetch(env.APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(Object.assign({ secret: env.SHARED_SECRET || '' }, payload)),
      redirect: 'follow',
    });
    const t = await r.text();
    try {
      return JSON.parse(t);
    } catch (e) {
      const snippet = t.replace(/<style[\s\S]*?<\/style>/gi, ' ').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 140);
      return { error: 'Apps Script ไม่ได้ตอบเป็นข้อมูล (HTTP ' + r.status + ') — ตรวจว่า Deploy แบบ Who has access: Anyone · ' + snippet };
    }
  } catch (e) {
    return { error: 'ติดต่อ Apps Script ไม่ได้: ' + e.message };
  }
}

// ---------------------------------------------------------------- ลงทะเบียนคำสั่ง
async function register(env) {
  if (!env.DISCORD_BOT_TOKEN) return json({ ok: false, error: 'ยังไม่ได้ตั้ง DISCORD_BOT_TOKEN ใน Worker' });
  const cfg = await callSheet(env, { action: 'config' });
  if (!cfg.ok) return json({ ok: false, step: 'sheet', error: 'อ่านหมวดจากชีตไม่ได้: ' + (cfg.error || '') });
  const app = await discord('GET', API + '/applications/@me', null, env.DISCORD_BOT_TOKEN);
  if (!app.ok) return json({ ok: false, step: 'token', error: 'Bot Token ใช้ไม่ได้ (' + app.status + ') ' + errText(app) });

  let guildId = cfg.guildId || '';
  if (!guildId) {
    const g = await discord('GET', API + '/users/@me/guilds', null, env.DISCORD_BOT_TOKEN);
    if (g.ok && g.json && g.json.length === 1) guildId = g.json[0].id;
  }
  const path = '/applications/' + app.json.id + (guildId ? '/guilds/' + guildId : '') + '/commands';
  const r = await discord('PUT', API + path, buildCommands(cfg.categories || {}), env.DISCORD_BOT_TOKEN);
  if (!r.ok) {
    const hint = r.status === 403 ? ' — บอทยังไม่อยู่ในเซิร์ฟเวอร์ (เชิญบอทก่อน)' : '';
    return json({ ok: false, step: 'discord', status: r.status, error: 'Discord ไม่รับคำสั่ง (' + r.status + ') ' + errText(r) + hint });
  }
  return json({ ok: true, scope: guildId ? 'guild' : 'global', guildId: guildId, commands: r.json.map(function (c) { return c.name; }) });
}

function buildCommands(cats) {
  const uniq = function (a) {
    const seen = {};
    return (a || []).map(function (s) { return String(s).trim().slice(0, 100); })
      .filter(function (s) { if (!s || seen[s]) return false; seen[s] = true; return true; }).slice(0, 25);
  };
  const entry = function (name, type) {
    const list = uniq(cats[type]);
    const cat = { type: 3, name: 'หมวด', description: 'หมวดหมู่' + type, required: true };
    if (list.length) cat.choices = list.map(function (c) { return { name: c, value: c }; });
    return {
      name: name, type: 1, description: 'จด' + type, default_member_permissions: MANAGE_SERVER,
      options: [
        { type: 10, name: 'จำนวน', description: 'จำนวนเงิน (บาท)', required: true, min_value: 0.01 },
        cat,
        { type: 3, name: 'รายละเอียด', description: 'เช่น ข้าวมันไก่ (ไม่ใส่ก็ได้)', required: false, max_length: 200 },
        { type: 3, name: 'วันที่', description: 'ไม่ใส่ = วันนี้ · เช่น 25/9 หรือ เมื่อวาน', required: false, max_length: 20 },
      ],
    };
  };
  return [
    entry('จ่าย', 'รายจ่าย'),
    entry('รับ', 'รายรับ'),
    entry('ออม', 'เงินออม'),
    {
      name: 'จด', type: 1, description: 'จดแบบพิมพ์รวดเดียว เช่น จ่าย 120 อาหาร ข้าวมันไก่', default_member_permissions: MANAGE_SERVER,
      options: [{ type: 3, name: 'ข้อความ', description: 'เช่น จ่าย 120 อาหาร ข้าวมันไก่', required: true, max_length: 300 }],
    },
    {
      name: 'สรุป', type: 1, description: 'สรุปรายรับ รายจ่าย เงินออม ของเดือน', default_member_permissions: MANAGE_SERVER,
      options: [
        { type: 4, name: 'เดือน', description: '1–12 (ไม่ใส่ = เดือนนี้)', required: false, min_value: 1, max_value: 12 },
        { type: 4, name: 'ปี', description: 'ค.ศ. หรือ พ.ศ. (ไม่ใส่ = ปีนี้)', required: false, min_value: 2000, max_value: 2700 },
      ],
    },
    { name: 'ยกเลิก', type: 1, description: 'ลบรายการล่าสุดที่จดผ่าน Discord', default_member_permissions: MANAGE_SERVER },
  ];
}

// ---------------------------------------------------------------- ตรวจสถานะ (เรียกจากเมนูในชีต)
async function status(env, origin) {
  const out = {
    ok: true,
    worker: origin,
    vars: {
      DISCORD_PUBLIC_KEY: !!env.DISCORD_PUBLIC_KEY,
      DISCORD_BOT_TOKEN: !!env.DISCORD_BOT_TOKEN,
      APPS_SCRIPT_URL: !!env.APPS_SCRIPT_URL,
      SHARED_SECRET: !!env.SHARED_SECRET,
    },
    publicKeyFormat: /^[0-9a-f]{64}$/i.test(String(env.DISCORD_PUBLIC_KEY || '').trim()),
  };
  out.sheet = await callSheet(env, { action: 'ping' });
  if (env.DISCORD_BOT_TOKEN) {
    const a = await discord('GET', API + '/applications/@me', null, env.DISCORD_BOT_TOKEN);
    if (!a.ok) {
      out.discord = { ok: false, status: a.status, error: errText(a) };
    } else {
      const ep = String(a.json.interactions_endpoint_url || '').replace(/\/+$/, '');
      out.discord = { ok: true, appId: a.json.id, appName: a.json.name, endpointUrl: ep, endpointOk: ep === origin };
      const gid = out.sheet && out.sheet.guildId;
      const c = await discord('GET', API + '/applications/' + a.json.id + (gid ? '/guilds/' + gid : '') + '/commands', null, env.DISCORD_BOT_TOKEN);
      out.discord.commands = c.ok ? c.json.map(function (x) { return x.name; }) : [];
      if (!c.ok) out.discord.commandsError = c.status + ' ' + errText(c);
    }
  }
  return json(out);
}

// ---------------------------------------------------------------- ตัวช่วย
function denied(url, env) {
  if (env.SHARED_SECRET && url.searchParams.get('key') === env.SHARED_SECRET) return null;
  return json({ ok: false, error: 'key ไม่ถูกต้อง' }, 403);
}

async function discord(method, url, body, token) {
  const headers = { 'User-Agent': UA };
  if (token) headers.Authorization = 'Bot ' + token;
  if (body != null) headers['Content-Type'] = 'application/json';
  const r = await fetch(url, { method: method, headers: headers, body: body == null ? undefined : JSON.stringify(body) });
  const t = await r.text();
  let j = null;
  try { j = t ? JSON.parse(t) : null; } catch (e) { /* ไม่ใช่ JSON */ }
  return { ok: r.ok, status: r.status, json: j, text: t };
}

function errText(r) {
  if (r.json && r.json.message) return r.json.message + (r.json.code ? ' (' + r.json.code + ')' : '');
  return String(r.text || '').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 100);
}

async function verifySignature(publicKeyHex, signatureHex, message) {
  try {
    const keyData = hexToBytes(publicKeyHex.trim());
    let algo = { name: 'Ed25519' };
    let key;
    try {
      key = await crypto.subtle.importKey('raw', keyData, algo, false, ['verify']);
    } catch (e) {
      algo = { name: 'NODE-ED25519', namedCurve: 'NODE-ED25519' }; // รันไทม์รุ่นเก่า
      key = await crypto.subtle.importKey('raw', keyData, algo, false, ['verify']);
    }
    return await crypto.subtle.verify(algo, key, hexToBytes(signatureHex), new TextEncoder().encode(message));
  } catch (e) {
    return false;
  }
}

function hexToBytes(hex) {
  const out = new Uint8Array(hex.length / 2);
  for (let i = 0; i < out.length; i++) out[i] = parseInt(hex.substr(i * 2, 2), 16);
  return out;
}

function errorEmbed(msg) {
  return { title: '❌ บันทึกไม่ได้', color: 0xD03B3B, description: String(msg).slice(0, 3900) };
}

function json(obj, status) {
  return new Response(JSON.stringify(obj), { status: status || 200, headers: { 'Content-Type': 'application/json; charset=utf-8' } });
}

function text(s, status) {
  return new Response(s, { status: status || 200, headers: { 'Content-Type': 'text/plain; charset=utf-8' } });
}
