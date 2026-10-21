

// ===== ПАРТНЁРСКАЯ ПРОГРАММА «ПРИВЕДИ ДРУГА» + КУРС ₽/$ + КОМАНДЫ БОТУ =====
// Лист «Партнёры» заполняется сам. Ссылка партнёра: https://t.me/gsm_catalog_bot?startapp=r_<TG ID>
// Новый клиент по ссылке: флакон 62 мл в подарок к первому заказу от MIN_GIFT_ORDER $.
// Пригласившему: флакон 62 мл после того, как вы нажмёте «Оплачен» под заказом друга.
// FRIENDS_FOR_DISCOUNT оплативших друзей → постоянная скидка 3%.
// Команды боту (только от вас): «курс 90», «привязать 15 @ivan», «партнёр @ivan», «партнёры».
// Один раз запустите setupBot — бот проверяет команды и кнопки каждую минуту.
const SHEET_CLIENTS = 'Партнёры';
const CLIENT_HEAD = ['TG ID', 'Имя', 'Ник', 'Пригласил (TG ID)', 'Заказов оформлено', 'Заказов оплачено', 'Друзей оплатили', 'Подарков доступно', 'Подарков получено', 'Скидка 3%', 'Добавлен'];
const ORDER_EXTRA_HEAD = ['TG ID клиента', 'Пригласил (TG ID)', 'Подарков за друзей', 'Скидка 3% $'];
const MIN_GIFT_ORDER = 120;
const FRIENDS_FOR_DISCOUNT = 5;
const DISCOUNT_PCT = 3;

function rub_(n) { return (Math.round(Number(n) / 10) * 10).toLocaleString('ru-RU'); }

// Проверка подписи Telegram: клиент не может подделать свой ID
function verifyInit_(initData) {
  if (!initData) return null;
  try {
    const pairs = String(initData).split('&').map(s => { const i = s.indexOf('='); return [decodeURIComponent(s.slice(0, i)), decodeURIComponent(s.slice(i + 1))]; });
    const hash = (pairs.find(p => p[0] === 'hash') || [])[1];
    if (!hash) return null;
    const dcs = pairs.filter(p => p[0] !== 'hash').sort((a, b) => a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0).map(p => p[0] + '=' + p[1]).join('\n');
    const H = Utilities.MacAlgorithm.HMAC_SHA_256;
    const secret = Utilities.computeHmacSignature(H, Utilities.newBlob(prop_('BOT_TOKEN')).getBytes(), Utilities.newBlob('WebAppData').getBytes());
    const sig = Utilities.computeHmacSignature(H, Utilities.newBlob(dcs).getBytes(), secret);
    const hex = sig.map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
    if (hex !== hash) return null;
    const user = JSON.parse((pairs.find(p => p[0] === 'user') || [])[1] || 'null');
    return user && user.id ? user : null;
  } catch (e) { return null; }
}

function clientsSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_CLIENTS);
  if (!sh) { sh = ss.insertSheet(SHEET_CLIENTS); sh.appendRow(CLIENT_HEAD); sh.setFrozenRows(1); }
  return sh;
}
function clientFind_(id) {
  id = String(id || '').trim().replace(/^@/, '');
  if (!id) return null;
  const sh = clientsSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return null;
  const v = sh.getRange(2, 1, n, 11).getValues();
  const byId = /^\d+$/.test(id);
  const k = v.findIndex(r => byId ? String(r[0]) === id : String(r[2]).toLowerCase() === '@' + id.toLowerCase());
  return k < 0 ? null : { row: k + 2, v: v[k] };
}
function clientSave_(c) { clientsSheet_().getRange(c.row, 1, 1, 11).setValues([c.v]); }
function refLabel_(v) { return String(v[1] || 'без имени') + (v[2] ? ' ' + v[2] : ''); }

// Сколько заказов клиент сделал до запуска программы (по нику/ID в листе «Заказы»)
function oldOrders_(u) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ORDERS);
  if (!sh || sh.getLastRow() < 2) return 0;
  const v = sh.getRange(2, 4, sh.getLastRow() - 1, 1).getValues();
  const nick = u.username ? '@' + String(u.username).toLowerCase() : '';
  return v.filter(r => { const s = String(r[0]).toLowerCase(); return s && (s === nick || s === String(u.id)); }).length;
}

function clientGet_(u) {
  const c = clientFind_(u.id);
  const name = [u.first_name, u.last_name].filter(Boolean).join(' ');
  const nick = u.username ? '@' + u.username : '';
  if (c) { c.v[1] = name || c.v[1]; c.v[2] = nick; return c; }
  const old = oldOrders_(u);
  const sh = clientsSheet_();
  sh.appendRow([String(u.id), name, nick, '', old, old, 0, 0, 0, '', new Date()]);
  return { row: sh.getLastRow(), v: sh.getRange(sh.getLastRow(), 1, 1, 11).getValues()[0] };
}

function allowList_() { try { return JSON.parse(prop_('ALLOW_INVITE') || '[]'); } catch (e) { return []; } }
function canInvite_(v) { return Number(v[5]) > 0 || (v[2] && allowList_().indexOf(String(v[2]).toLowerCase()) >= 0); }

// Данные клиента для мини-аппа + привязка к пригласившему при первом входе по ссылке
function me_(body) {
  const u = verifyInit_(body.initData);
  if (!u) return { ok: false, error: 'auth' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const c = clientGet_(u);
    const refArg = String(body.ref || '').replace(/\D/g, '');
    if (!c.v[3] && Number(c.v[4]) === 0 && refArg && refArg !== String(u.id)) {
      const r = clientFind_(refArg);
      if (r && canInvite_(r.v)) c.v[3] = String(r.v[0]);
    }
    clientSave_(c);
    const ref = c.v[3] ? clientFind_(c.v[3]) : null;
    return {
      ok: true, id: String(u.id),
      newGift: Number(c.v[4]) === 0 && !!c.v[3], refName: ref ? String(ref.v[1] || ref.v[2] || '') : '',
      gifts: Number(c.v[7]) || 0, friends: Number(c.v[6]) || 0, discount: String(c.v[9]) === 'да',
      canInvite: !!canInvite_(c.v), isAdmin: String(u.id) === String(prop_('ADMIN_CHAT_ID')),
      minGift: MIN_GIFT_ORDER, friendsForDiscount: FRIENDS_FOR_DISCOUNT, pct: DISCOUNT_PCT
    };
  } finally { lock.releaseLock(); }
}

// Подарки и скидка в заказе. Решает сервер, а не телефон клиента.
function applyPartner_(body, o) {
  const res = { uid: '', refId: '', refGifts: 0, disc: 0, giftLines: [], notes: [], total: round_(Number(o.total) + Number(o.disc || 0)) };
  const u = verifyInit_(body.initData);
  if (!u) return res;
  res.uid = String(u.id);
  const c = clientGet_(u);
  const gifts = Array.isArray(o.gifts) ? o.gifts : [];
  const nm = g => (g && g.name ? String(g.name).slice(0, 80) : 'флакон 62 мл на выбор (уточнить у клиента)');
  if (Number(c.v[4]) === 0 && c.v[3]) {
    const ref = clientFind_(c.v[3]);
    if (Number(o.base) >= MIN_GIFT_ORDER) {
      res.refId = String(c.v[3]);
      res.giftLines.push(`🎁 · Подарок новому клиенту: ${nm(gifts.find(g => g.type === 'new'))} × 1 = 0 $`);
      res.notes.push(`🤝 Новый клиент по приглашению: ${ref ? refLabel_(ref.v) : res.refId}. После «Оплачен» пригласившему начислится подарок.`);
    } else {
      res.notes.push(`🤝 Клиент по приглашению (${ref ? refLabel_(ref.v) : c.v[3]}), но заказ меньше ${MIN_GIFT_ORDER} $ — подарок не положен.`);
    }
  }
  const avail = Number(c.v[7]) || 0;
  const want = Math.min(avail, gifts.filter(g => g.type === 'ref').length);
  gifts.filter(g => g.type === 'ref').slice(0, want).forEach(g => res.giftLines.push(`🎁 · Подарок за друга: ${nm(g)} × 1 = 0 $`));
  res.refGifts = want;
  if (String(c.v[9]) === 'да') {
    res.disc = round_(res.total * DISCOUNT_PCT / 100);
    res.total = round_(res.total - res.disc);
    res.notes.push(`⭐ Постоянная скидка партнёра ${DISCOUNT_PCT}% учтена.`);
  }
  c.v[4] = Number(c.v[4] || 0) + 1;
  c.v[7] = avail - want;
  c.v[8] = Number(c.v[8] || 0) + want;
  clientSave_(c);
  return res;
}

// Список партнёров — только для вас (раздел «Партнёры» в мини-аппе)
function partners_(body) {
  const u = verifyInit_(body.initData);
  if (!u || String(u.id) !== String(prop_('ADMIN_CHAT_ID'))) return { ok: false, error: 'forbidden' };
  const sh = clientsSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return { ok: true, rows: [] };
  const v = sh.getRange(2, 1, n, 11).getValues();
  const name = {};
  v.forEach(r => name[String(r[0])] = refLabel_(r));
  const rows = v.filter(r => r[3] || Number(r[6]) || Number(r[7]) || Number(r[8])).map(r => ({
    name: refLabel_(r), ref: r[3] ? (name[String(r[3])] || String(r[3])) : '', placed: Number(r[4]) || 0, paid: Number(r[5]) || 0,
    friends: Number(r[6]) || 0, gifts: Number(r[7]) || 0, used: Number(r[8]) || 0, disc: String(r[9]) === 'да'
  }));
  return { ok: true, rows: rows, total: n };
}

function creditRef_(p) {
  p.v[6] = Number(p.v[6] || 0) + 1;
  p.v[7] = Number(p.v[7] || 0) + 1;
  const gotDisc = Number(p.v[6]) >= FRIENDS_FOR_DISCOUNT && String(p.v[9]) !== 'да';
  if (gotDisc) p.v[9] = 'да';
  clientSave_(p);
  sendTg_(p.v[0], '🎁 Ваш друг оплатил заказ — вам подарок: флакон 62 мл на выбор! Выберите его в корзине при следующем заказе.' +
    (gotDisc ? `\n\n⭐ Вы пригласили ${FRIENDS_FOR_DISCOUNT} друзей — теперь у вас постоянная скидка ${DISCOUNT_PCT}% на все заказы!` : ''),
    { inline_keyboard: [[{ text: '🛍 Открыть каталог', web_app: { url: CATALOG_URL } }]] });
}

function markPaid_(oid) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ORDERS);
  const n = sh.getLastRow() - 1;
  if (n < 1) return 'Заказ не найден';
  const v = sh.getRange(2, 1, n, 18).getValues();
  const k = v.findIndex(x => String(x[13]) === oid);
  if (k < 0) return 'Заказ не найден';
  const r = v[k], row = k + 2;
  if (String(r[11]) === 'оплачен') return 'Заказ №' + (row - 1) + ' уже отмечен оплаченным';
  sh.getRange(row, 12).setValue('оплачен');
  const out = ['Заказ №' + (row - 1) + ' отмечен оплаченным'];
  if (r[14]) { const c = clientFind_(r[14]); if (c) { c.v[5] = Number(c.v[5] || 0) + 1; clientSave_(c); } }
  if (r[15]) { const p = clientFind_(r[15]); if (p) { creditRef_(p); out.push('Пригласившему (' + refLabel_(p.v) + ') начислен подарок'); } }
  return out.join('. ');
}

function bindCmd_(num, who) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ORDERS);
  const row = Number(num) + 1;
  if (!sh || row < 2 || row > sh.getLastRow()) return 'Нет заказа №' + num;
  const r = sh.getRange(row, 1, 1, 18).getValues()[0];
  if (!r[14]) return 'В заказе №' + num + ' нет Telegram клиента (сделан не из Telegram или до запуска программы) — привязать нельзя.';
  if (r[15]) return 'Заказ №' + num + ' уже привязан к партнёру.';
  const p = clientFind_(who);
  if (!p) return 'Партнёр ' + who + ' не найден: он должен хотя бы раз открыть мини-апп. Можно указать его TG ID из листа «Партнёры».';
  if (String(p.v[0]) === String(r[14])) return 'Нельзя привязать клиента к самому себе.';
  const c = clientFind_(r[14]);
  if (c && !c.v[3]) { c.v[3] = String(p.v[0]); clientSave_(c); }
  sh.getRange(row, 16).setValue(String(p.v[0]));
  let msg = '✅ Заказ №' + num + ' привязан к ' + refLabel_(p.v) + '. Положите клиенту флакон 62 мл в подарок.';
  if (String(r[11]) === 'оплачен') { creditRef_(p); msg += ' Заказ уже оплачен — партнёру подарок начислен сразу.'; }
  else msg += ' Партнёру подарок начислится, когда нажмёте «Оплачен».';
  return msg;
}

function partnersText_() {
  const sh = clientsSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return 'Партнёров пока нет.';
  const v = sh.getRange(2, 1, n, 11).getValues().filter(r => Number(r[6]) || Number(r[7]) || Number(r[8]));
  if (!v.length) return 'Пока никто не привёл друзей. Клиентов в базе: ' + n + '.';
  v.sort((a, b) => Number(b[6]) - Number(a[6]));
  return '🤝 Партнёры:\n\n' + v.slice(0, 40).map(r => `${refLabel_(r)} — друзей оплатили: ${Number(r[6]) || 0}, подарков доступно: ${Number(r[7]) || 0}, получено: ${Number(r[8]) || 0}` + (String(r[9]) === 'да' ? ', скидка 3%' : '')).join('\n');
}

const BOT_HELP = 'Команды:\n• курс 90 — поменять курс ₽/$ в мини-аппе\n• привязать 15 @ivan — привязать клиента из заказа №15 к партнёру (если он пришёл не по ссылке)\n• партнёр @ivan — разрешить приглашать друзей тому, у кого ещё не было оплаченного заказа в мини-аппе\n• партнёры — список партнёров';

function handleUpdate_(u) {
  const admin = String(prop_('ADMIN_CHAT_ID'));
  const P = PropertiesService.getScriptProperties();
  if (u.callback_query) {
    const q = u.callback_query, d = String(q.data || '');
    if (String(q.from.id) !== admin) { tgCall_('answerCallbackQuery', { callback_query_id: q.id, text: 'Только для продавца' }); return; }
    if (d.indexOf('paid:') === 0) {
      const oid = d.slice(5);
      const res = markPaid_(oid);
      tgCall_('answerCallbackQuery', { callback_query_id: q.id, text: res.slice(0, 190), show_alert: true });
      if (q.message) tgCall_('editMessageReplyMarkup', { chat_id: q.message.chat.id, message_id: q.message.message_id, reply_markup: orderKb_(oid, true) });
    } else tgCall_('answerCallbackQuery', { callback_query_id: q.id });
    return;
  }
  const m = u.message;
  if (!m || !m.text) return;
  const t = m.text.trim(), chat = m.chat.id;
  if (String(m.from.id) !== admin) {
    if (/^\/start/.test(t)) sendTg_(chat, 'Добро пожаловать! Каталог и заказ — по кнопке ниже 👇', { inline_keyboard: [[{ text: '🛍 Открыть каталог', web_app: { url: CATALOG_URL } }]] });
    return;
  }
  let mm;
  if ((mm = t.match(/^\/?курс\s*[:=]?\s*(\d+(?:[.,]\d+)?)/i))) {
    const rate = Number(mm[1].replace(',', '.'));
    if (!(rate >= 20 && rate <= 500)) { sendTg_(chat, 'Не понял курс. Пример: курс 90'); return; }
    P.setProperty('RATE', String(rate));
    P.setProperty('RATE_DATE', Utilities.formatDate(new Date(), POST_TZ, 'dd.MM.yyyy'));
    sendTg_(chat, `✅ Курс обновлён: ${rate} ₽/$. Цены в рублях в мини-аппе пересчитаны (клиенты увидят при следующем открытии).`);
    return;
  }
  if ((mm = t.match(/^\/?привязать\s+№?\s*(\d+)\s+(\S+)/i))) { sendTg_(chat, bindCmd_(mm[1], mm[2])); return; }
  if ((mm = t.match(/^\/?партн[её]р\s+@?([A-Za-z0-9_]{3,})\s*$/i))) {
    const l = allowList_(), nick = '@' + mm[1].toLowerCase();
    if (l.indexOf(nick) < 0) l.push(nick);
    P.setProperty('ALLOW_INVITE', JSON.stringify(l));
    sendTg_(chat, `✅ ${nick} может приглашать друзей. Кнопка «Пригласить друга» появится у него в профиле мини-аппа.`);
    return;
  }
  if (/^\/?партн[её]ры\s*$/i.test(t)) { sendTg_(chat, partnersText_()); return; }
  if (/^\/?(start|help|помощь)/i.test(t)) sendTg_(chat, BOT_HELP);
}

function pollBot() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    const P = PropertiesService.getScriptProperties();
    const off = Number(P.getProperty('BOT_OFFSET') || 0);
    const r = UrlFetchApp.fetch('https://api.telegram.org/bot' + prop_('BOT_TOKEN') + '/getUpdates?timeout=0&offset=' + off +
      '&allowed_updates=' + encodeURIComponent('["message","callback_query"]'), { muteHttpExceptions: true });
    const j = JSON.parse(r.getContentText());
    if (!j.ok) { Logger.log(j.description); return; }
    let next = off;
    j.result.forEach(up => {
      next = up.update_id + 1;
      P.setProperty('BOT_OFFSET', String(next));
      try { handleUpdate_(up); } catch (e) { Logger.log('Ошибка: ' + e); }
    });
  } finally { lock.releaseLock(); }
}

// Запустить один раз: бот начинает принимать команды и кнопку «Оплачен» (проверка каждую минуту)
function setupBot() {
  tgCall_('deleteWebhook', {});
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'pollBot').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('pollBot').timeBased().everyMinutes(1).create();
  clientsSheet_();
  const P = PropertiesService.getScriptProperties();
  if (!P.getProperty('RATE')) { P.setProperty('RATE', '90'); P.setProperty('RATE_DATE', Utilities.formatDate(new Date(), POST_TZ, 'dd.MM.yyyy')); }
  Logger.log('Бот включён: команды и кнопки проверяются каждую минуту. Курс: ' + P.getProperty('RATE'));
}
