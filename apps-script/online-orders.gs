/**
 * IZYAN — Online Orders (Google Apps Script web app)
 *
 * What it does
 *  - The Izyan website sends every "Order on WhatsApp" click here.
 *  - The order is saved in a Google Sheet ("Izyan Online Orders" → tab "Online Orders").
 *  - You get an email for every new order.
 *  - Business Manager reads the orders (with your secret key) and marks them Confirmed / Cancelled.
 *
 * One-time setup
 *  1. Go to https://script.google.com → New project → paste this whole file → Save.
 *  2. Select the function "setup" in the toolbar → Run → allow the permissions.
 *     It creates the sheet and emails you your ADMIN KEY (also shown in the log).
 *  3. Deploy → New deployment → type "Web app"
 *       Execute as: Me    |    Who has access: Anyone
 *     → Deploy → copy the Web app URL (ends with /exec).
 *  4. Give that URL to Claude (for the website) and paste URL + ADMIN KEY in
 *     Business Manager → Settings → Online Orders.
 *
 * If you change this code later: Deploy → Manage deployments → Edit (pencil) → Version: New version → Deploy.
 * (That keeps the same URL.)
 */

const ORDERS_SHEET_NAME = 'Online Orders';
const HEADERS = ['Order ID', 'Received At', 'Status', 'Name', 'Phone', 'Address', 'Items', 'Items JSON',
  'Free Gifts', 'Subtotal', 'Discount Code', 'Discount', 'Total', 'Invoice', 'Updated At', 'Note'];
const MAX_ORDERS_PER_PHONE_PER_10_MIN = 5;

// ---------- setup (run once from the editor) ----------
function setup() {
  const props = PropertiesService.getScriptProperties();
  getOrdersSheet_();
  let key = props.getProperty('ADMIN_KEY');
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '').slice(0, 20);
    props.setProperty('ADMIN_KEY', key);
  }
  const ssId = props.getProperty('SHEET_ID');
  const msg = 'IZYAN Online Orders is set up.\n\n' +
    'ADMIN KEY (paste in Business Manager → Settings → Online Orders):\n' + key + '\n\n' +
    'Orders sheet: https://docs.google.com/spreadsheets/d/' + ssId + '/edit\n\n' +
    'Keep this key private — anyone with it can read your orders.';
  Logger.log(msg);
  MailApp.sendEmail(ownerEmail_(), 'IZYAN Online Orders — your admin key', msg);
}

// ---------- web endpoints ----------
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'order') return json_(saveOrder_(body.order || {}));
    if (body.action === 'status') {
      requireKey_(body.key);
      return json_(updateStatus_(body.orderId, body.status, body.invoice, body.note));
    }
    return json_({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    if (p.action === 'list') {
      requireKey_(p.key);
      return json_({ ok: true, orders: listOrders_(Number(p.limit) || 200) });
    }
    return json_({ ok: true, service: 'izyan-online-orders' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

// ---------- orders ----------
function saveOrder_(o) {
  const order = cleanOrder_(o);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getOrdersSheet_();
    // Same order sent twice (customer tapped the button again) → keep only one row
    if (findRow_(sh, order.orderId) > 0) return { ok: true, orderId: order.orderId, duplicate: true };

    const cache = CacheService.getScriptCache();
    const rateKey = 'rate_' + order.phone;
    const count = Number(cache.get(rateKey) || 0);
    if (count >= MAX_ORDERS_PER_PHONE_PER_10_MIN) throw new Error('Too many orders from this number, please try again later');
    cache.put(rateKey, String(count + 1), 600);

    const now = new Date();
    sh.appendRow([
      order.orderId, now, 'New', order.name, order.phone, order.address,
      order.itemsText, JSON.stringify(order.items), order.giftsText,
      order.subtotal, order.discountCode, order.discount, order.total, '', now, ''
    ].map(safeCell_));
  } finally {
    lock.releaseLock();
  }
  try { sendOrderEmail_(order); } catch (err) { Logger.log('Email failed: ' + err); }
  return { ok: true, orderId: order.orderId };
}

function cleanOrder_(o) {
  const str = (v, max) => String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, max);
  const num = v => { const n = Number(v); return isFinite(n) && n >= 0 ? Math.round(n * 100) / 100 : 0; };
  const orderId = str(o.orderId, 40);
  if (!/^WEB-[A-Z0-9-]{4,36}$/.test(orderId)) throw new Error('Invalid order id');
  const name = str(o.name, 80);
  const phone = str(o.phone, 20).replace(/[^\d+]/g, '');
  const address = String(o.address == null ? '' : o.address).trim().slice(0, 400);
  if (!name || phone.replace(/\D/g, '').length < 10 || !address) throw new Error('Name, phone and address are required');
  const items = (Array.isArray(o.items) ? o.items : []).slice(0, 40).map(it => ({
    name: str(it.name, 120), size: str(it.size, 200), qty: Math.max(1, Math.min(99, Math.round(num(it.qty)) || 1)), price: num(it.price)
  })).filter(it => it.name);
  if (!items.length) throw new Error('Order has no items');
  const gifts = (Array.isArray(o.gifts) ? o.gifts : []).slice(0, 10).map(g => ({
    label: str(g.label, 60), qty: Math.max(1, Math.min(99, Math.round(num(g.qty)) || 1))
  })).filter(g => g.label);
  const subtotal = items.reduce((s, it) => s + it.price * it.qty, 0);
  const discount = Math.min(num(o.discount), subtotal);
  return {
    orderId, name, phone, address, items, gifts,
    itemsText: items.map(it => `${it.name} (${it.size}) x${it.qty} = Rs. ${it.price * it.qty}`).join('\n'),
    giftsText: gifts.map(g => `FREE ${g.label} x${g.qty}`).join('\n'),
    subtotal, discountCode: str(o.discountCode, 30), discount, total: subtotal - discount
  };
}

function updateStatus_(orderId, status, invoice, note) {
  if (['New', 'Confirmed', 'Cancelled'].indexOf(status) < 0) throw new Error('Invalid status');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getOrdersSheet_();
    const row = findRow_(sh, String(orderId || ''));
    if (row < 1) throw new Error('Order not found: ' + orderId);
    sh.getRange(row, 3).setValue(status);
    if (invoice != null) sh.getRange(row, 14).setValue(safeCell_(String(invoice).slice(0, 40)));
    sh.getRange(row, 15).setValue(new Date());
    if (note != null) sh.getRange(row, 16).setValue(safeCell_(String(note).slice(0, 300)));
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function listOrders_(limit) {
  const sh = getOrdersSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const n = Math.min(Math.max(limit, 1), 500);
  const start = Math.max(2, last - n + 1);
  const rows = sh.getRange(start, 1, last - start + 1, HEADERS.length).getValues();
  return rows.map(r => {
    let items = [];
    try { items = JSON.parse(r[7] || '[]'); } catch (e) { items = []; }
    return {
      orderId: String(r[0]), receivedAt: r[1] instanceof Date ? r[1].toISOString() : String(r[1]),
      status: String(r[2] || 'New'), name: String(r[3]), phone: String(r[4]), address: String(r[5]),
      items: items, gifts: String(r[8] || ''), subtotal: Number(r[9]) || 0,
      discountCode: String(r[10] || ''), discount: Number(r[11]) || 0, total: Number(r[12]) || 0,
      invoice: String(r[13] || ''), note: String(r[15] || '')
    };
  }).reverse();
}

// ---------- email ----------
function sendOrderEmail_(o) {
  const lines = [
    'New order on the IZYAN website', '',
    'Order ID: ' + o.orderId,
    'Name: ' + o.name,
    'Phone: ' + o.phone,
    'Address: ' + o.address, '',
    'Items:', o.itemsText
  ];
  if (o.giftsText) lines.push('', 'Free gifts:', o.giftsText);
  lines.push('', 'Subtotal: Rs. ' + o.subtotal);
  if (o.discount) lines.push('Discount' + (o.discountCode ? ' (' + o.discountCode + ')' : '') + ': - Rs. ' + o.discount);
  lines.push('Total: Rs. ' + o.total, '',
    'The customer was sent to WhatsApp to send this order. If no WhatsApp message arrives, contact them on ' + o.phone + '.',
    'Open Business Manager → Online Orders to create the bill.');
  MailApp.sendEmail(ownerEmail_(), `🛒 New IZYAN order ${o.orderId} — Rs. ${o.total} — ${o.name}`, lines.join('\n'));
}

// ---------- helpers ----------
function getOrdersSheet_() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  const id = props.getProperty('SHEET_ID');
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create('Izyan Online Orders');
    props.setProperty('SHEET_ID', ss.getId());
  }
  let sh = ss.getSheetByName(ORDERS_SHEET_NAME);
  if (!sh) {
    sh = ss.getSheets()[0];
    sh.setName(ORDERS_SHEET_NAME);
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sh;
}

function findRow_(sh, orderId) {
  const last = sh.getLastRow();
  if (last < 2 || !orderId) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === orderId) return i + 2;
  return -1;
}

function requireKey_(key) {
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  if (!real) throw new Error('Run setup() first');
  if (String(key || '') !== real) throw new Error('Wrong admin key');
}

function ownerEmail_() {
  return Session.getEffectiveUser().getEmail();
}

// Stops text like "=HYPERLINK(...)" from being treated as a formula in the sheet
function safeCell_(v) {
  return typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
