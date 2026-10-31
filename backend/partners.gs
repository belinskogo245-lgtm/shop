

// ===== ПАРТНЁРСКАЯ ПРОГРАММА «ПРИВЕДИ ДРУГА» + КУРС ₽/$ + КОМАНДЫ БОТУ =====
// Лист «Партнёры» заполняется сам. Ссылка партнёра: https://t.me/gsm_catalog_bot?startapp=r_<TG ID>
// Новый клиент по ссылке: флакон 62 мл в подарок к первому заказу от MIN_GIFT_ORDER $.
// Пригласившему: флакон 62 мл после того, как вы нажмёте «Оплачен» под первым заказом друга,
// и BONUS_PCT% бонусами с КАЖДОГО оплаченного заказа друга от MIN_GIFT_ORDER $ — навсегда.
// Бонусы ($) копятся в профиле: ими можно оплатить свой заказ (галочка в корзине) или вывести деньгами
// (вы переводите сами и списываете командой «бонус @ivan -10»).
// Команды боту (только от вас): «курс 90», «привязать 15 @ivan», «партнёр @ivan», «партнёры», «бонус @ivan 5», «визитка».
// Один раз запустите setupBot — бот проверяет команды и кнопки каждую минуту.
const SHEET_CLIENTS = 'Партнёры';
const CLIENT_HEAD = ['TG ID', 'Имя', 'Ник', 'Пригласил (TG ID)', 'Заказов оформлено', 'Заказов оплачено', 'Друзей оплатили', 'Подарков доступно', 'Подарков получено', 'Скидка 3% (старое, не используется)', 'Добавлен', 'Бонусы $ (баланс)', 'Бонусов начислено всего $', 'Бонусов списано/выведено $'];
const CLIENT_COLS = 14;
const ORDER_EXTRA_HEAD = ['TG ID клиента', 'Пригласил (TG ID)', 'Подарков за друзей', 'Скидка 3% $', 'Оплачено бонусами $', 'Бонус пригласившему $'];
const MIN_GIFT_ORDER = 120;
const BONUS_PCT = 3;

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
  else if (sh.getRange(1, CLIENT_COLS).getValue() === '') sh.getRange(1, 1, 1, CLIENT_COLS).setValues([CLIENT_HEAD]);
  return sh;
}
function clientFind_(id) {
  id = String(id || '').trim().replace(/^@/, '');
  if (!id) return null;
  const sh = clientsSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return null;
  const v = sh.getRange(2, 1, n, CLIENT_COLS).getValues();
  const byId = /^\d+$/.test(id);
  const k = v.findIndex(r => byId ? String(r[0]) === id : String(r[2]).toLowerCase() === '@' + id.toLowerCase());
  return k < 0 ? null : { row: k + 2, v: v[k] };
}
function clientSave_(c) { clientsSheet_().getRange(c.row, 1, 1, CLIENT_COLS).setValues([c.v]); }
function bonus_(v) { return round_(Number(v[11]) || 0); }
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
  sh.appendRow([String(u.id), name, nick, '', old, old, 0, 0, 0, '', new Date(), 0, 0, 0]);
  return { row: sh.getLastRow(), v: sh.getRange(sh.getLastRow(), 1, 1, CLIENT_COLS).getValues()[0] };
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
      gifts: Number(c.v[7]) || 0, friends: Number(c.v[6]) || 0, discount: false,
      bonus: bonus_(c.v), bonusTotal: round_(Number(c.v[12]) || 0),
      canInvite: !!canInvite_(c.v), isAdmin: String(u.id) === String(prop_('ADMIN_CHAT_ID')),
      minGift: MIN_GIFT_ORDER, pct: BONUS_PCT
    };
  } finally { lock.releaseLock(); }
}

// Подарки и скидка в заказе. Решает сервер, а не телефон клиента.
function applyPartner_(body, o) {
  const res = { uid: '', refId: '', refGifts: 0, disc: 0, bonusUsed: 0, giftLines: [], notes: [], total: round_(Number(o.total) + Number(o.disc || 0) + Number(o.bonus || 0)) };
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
  // Оплата бонусами: списываем сразу при оформлении (чтобы не потратить дважды). Если заказ отменят — верните командой «бонус @ник 5».
  const bal = bonus_(c.v), askB = Number(o.bonus) || 0;
  if (askB > 0 && bal > 0) {
    res.bonusUsed = round_(Math.min(bal, askB, res.total));
    res.total = round_(res.total - res.bonusUsed);
    c.v[11] = round_(bal - res.bonusUsed);
    c.v[13] = round_((Number(c.v[13]) || 0) + res.bonusUsed);
    res.notes.push(`💰 Клиент оплатил бонусами ${res.bonusUsed} $ (остаток бонусов: ${c.v[11]} $). Если заказ отменится — верните: «бонус ${c.v[2] || c.v[0]} ${res.bonusUsed}».`);
  }
  if (c.v[3]) {
    const ref = clientFind_(c.v[3]);
    if (res.total + res.bonusUsed >= MIN_GIFT_ORDER) res.notes.push(`💸 После «Оплачен» пригласившему (${ref ? refLabel_(ref.v) : c.v[3]}) начислится ${BONUS_PCT}% бонусами.`);
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
  const v = sh.getRange(2, 1, n, CLIENT_COLS).getValues();
  const name = {};
  v.forEach(r => name[String(r[0])] = refLabel_(r));
  const rows = v.filter(r => r[3] || Number(r[6]) || Number(r[7]) || Number(r[8]) || Number(r[12])).map(r => ({
    name: refLabel_(r), ref: r[3] ? (name[String(r[3])] || String(r[3])) : '', placed: Number(r[4]) || 0, paid: Number(r[5]) || 0,
    friends: Number(r[6]) || 0, gifts: Number(r[7]) || 0, used: Number(r[8]) || 0, disc: false,
    bonus: round_(Number(r[11]) || 0), bonusTotal: round_(Number(r[12]) || 0)
  }));
  return { ok: true, rows: rows, total: n };
}

// first = первый оплаченный заказ друга (подарок флакон); amt = бонус 3% с этого заказа
function creditRef_(p, first, amt) {
  if (first) { p.v[6] = Number(p.v[6] || 0) + 1; p.v[7] = Number(p.v[7] || 0) + 1; }
  if (amt > 0) { p.v[11] = round_(bonus_(p.v) + amt); p.v[12] = round_((Number(p.v[12]) || 0) + amt); }
  if (!first && !(amt > 0)) return;
  clientSave_(p);
  const t = [];
  if (first) t.push('🎁 Ваш друг оплатил первый заказ — вам подарок: флакон 62 мл на выбор! Выберите его в корзине при следующем заказе.');
  if (amt > 0) t.push(`💰 Начислено ${amt} $ бонусами (${BONUS_PCT}% с заказа друга). Ваш баланс: ${p.v[11]} $.\nБонусами можно оплатить свой заказ в корзине или вывести деньгами — напишите нам.`);
  sendTg_(p.v[0], t.join('\n\n'), { inline_keyboard: [[{ text: '🛍 Открыть каталог', web_app: { url: CATALOG_URL } }]] });
}
function refBonus_(total) { return Number(total) >= MIN_GIFT_ORDER ? round_(Number(total) * BONUS_PCT / 100) : 0; }

function markPaid_(oid) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ORDERS);
  const n = sh.getLastRow() - 1;
  if (n < 1) return 'Заказ не найден';
  const v = sh.getRange(2, 1, n, 20).getValues();
  const k = v.findIndex(x => String(x[13]) === oid);
  if (k < 0) return 'Заказ не найден';
  const r = v[k], row = k + 2;
  if (String(r[11]) === 'оплачен') return 'Заказ №' + (row - 1) + ' уже отмечен оплаченным';
  sh.getRange(row, 12).setValue('оплачен');
  const out = ['Заказ №' + (row - 1) + ' отмечен оплаченным'];
  let refId = '';
  if (r[14]) { const c = clientFind_(r[14]); if (c) { c.v[5] = Number(c.v[5] || 0) + 1; clientSave_(c); refId = String(c.v[3] || ''); } }
  if (r[15]) refId = String(r[15]);
  if (refId) {
    const p = clientFind_(refId);
    if (p) {
      const amt = refBonus_(Number(r[10]) + (Number(r[18]) || 0));
      creditRef_(p, !!r[15], amt);
      if (amt) sh.getRange(row, 20).setValue(amt);
      const got = [r[15] ? 'подарок' : '', amt ? amt + ' $ бонусами' : ''].filter(Boolean).join(' и ');
      if (got) out.push('Пригласившему (' + refLabel_(p.v) + ') начислено: ' + got);
    }
  }
  return out.join('. ');
}

function bindCmd_(num, who) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ORDERS);
  const row = Number(num) + 1;
  if (!sh || row < 2 || row > sh.getLastRow()) return 'Нет заказа №' + num;
  const r = sh.getRange(row, 1, 1, 20).getValues()[0];
  if (!r[14]) return 'В заказе №' + num + ' нет Telegram клиента (сделан не из Telegram или до запуска программы) — привязать нельзя.';
  if (r[15]) return 'Заказ №' + num + ' уже привязан к партнёру.';
  const p = clientFind_(who);
  if (!p) return 'Партнёр ' + who + ' не найден: он должен хотя бы раз открыть мини-апп. Можно указать его TG ID из листа «Партнёры».';
  if (String(p.v[0]) === String(r[14])) return 'Нельзя привязать клиента к самому себе.';
  const c = clientFind_(r[14]);
  if (c && !c.v[3]) { c.v[3] = String(p.v[0]); clientSave_(c); }
  sh.getRange(row, 16).setValue(String(p.v[0]));
  let msg = '✅ Заказ №' + num + ' привязан к ' + refLabel_(p.v) + '. Положите клиенту флакон 62 мл в подарок.';
  if (String(r[11]) === 'оплачен') {
    const amt = refBonus_(Number(r[10]) + (Number(r[18]) || 0));
    creditRef_(p, true, amt); if (amt) sh.getRange(row, 20).setValue(amt);
    msg += ' Заказ уже оплачен — партнёру сразу начислены подарок' + (amt ? ' и ' + amt + ' $ бонусами' : '') + '.';
  }
  else msg += ` Партнёру подарок и ${BONUS_PCT}% бонусами начислятся, когда нажмёте «Оплачен». Дальше ${BONUS_PCT}% будут идти с каждого оплаченного заказа этого клиента.`;
  return msg;
}

function bonusCmd_(who, amt) {
  const c = clientFind_(who);
  if (!c) return 'Клиент ' + who + ' не найден: он должен хотя бы раз открыть мини-апп. Можно указать TG ID из листа «Партнёры».';
  if (!amt) return 'Пример: бонус @ivan 5 или бонус @ivan -10';
  const bal = bonus_(c.v);
  if (amt < 0 && -amt > bal) return 'У ' + refLabel_(c.v) + ' на балансе только ' + bal + ' $.';
  c.v[11] = round_(bal + amt);
  if (amt > 0) c.v[12] = round_((Number(c.v[12]) || 0) + amt); else c.v[13] = round_((Number(c.v[13]) || 0) - amt);
  clientSave_(c);
  sendTg_(c.v[0], amt > 0 ? `💰 Вам начислено ${round_(amt)} $ бонусами. Баланс: ${c.v[11]} $.` : `💸 Списано ${round_(-amt)} $ бонусов (вывод). Баланс: ${c.v[11]} $.`);
  return `✅ ${refLabel_(c.v)}: ${amt > 0 ? '+' : ''}${round_(amt)} $. Баланс бонусов: ${c.v[11]} $.`;
}

function partnersText_() {
  const sh = clientsSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return 'Партнёров пока нет.';
  const v = sh.getRange(2, 1, n, CLIENT_COLS).getValues().filter(r => Number(r[6]) || Number(r[7]) || Number(r[8]) || Number(r[12]));
  if (!v.length) return 'Пока никто не привёл друзей. Клиентов в базе: ' + n + '.';
  v.sort((a, b) => Number(b[6]) - Number(a[6]));
  return '🤝 Партнёры:\n\n' + v.slice(0, 40).map(r => `${refLabel_(r)} — друзей оплатили: ${Number(r[6]) || 0}, подарков доступно: ${Number(r[7]) || 0}, получено: ${Number(r[8]) || 0}, бонусы: ${round_(Number(r[11]) || 0)} $ (всего начислено ${round_(Number(r[12]) || 0)} $)`).join('\n');
}

// «Визитка» — картинка с двумя настоящими кнопками. Её можно переслать в канал или клиенту — кнопки сохраняются.
const CARD_IMG = 'https://belinskogo245-lgtm.github.io/shop/card.jpg';
function sendCard_(chat, call) {
  return (call || tgCall_)('sendPhoto', {
    chat_id: chat, photo: CARD_IMG,
    caption: '🛍 Оптовый каталог парфюмерии\nОтправка по всей России 📦\n\nНажмите кнопку ниже 👇',
    reply_markup: { inline_keyboard: [
      [{ text: '✈️ Открыть в Telegram', url: 'https://t.me/gsm_catalog_bot?startapp' }],
      [{ text: '🌐 Открыть без VPN (браузер)', url: 'https://belinskogo245-lgtm.github.io/shop/' }]
    ] }
  });
}

const BOT_HELP = 'Команды:\n• визитка — картинка с кнопками каталога (перешлите её в канал или клиенту)\n• визитка @канал — сразу опубликовать в канал (бот должен быть админом канала)\n• курс 90 — поменять курс ₽/$ в мини-аппе\n• привязать 15 @ivan — привязать клиента из заказа №15 к партнёру (если он пришёл не по ссылке)\n• партнёр @ivan — разрешить приглашать друзей тому, у кого ещё не было оплаченного заказа в мини-аппе\n• партнёры — список партнёров\n• бонус @ivan 5 — добавить клиенту 5 $ бонусов; бонус @ivan -10 — списать (например, когда перевели ему деньгами)';

function handleUpdate_(u, call, isAdminBot) {
  call = call || tgCall_;
  const send = call === tgCall_ ? sendTg_ : (c, t, k) => sendVia_(call, c, t, k);
  const admin = String(prop_('ADMIN_CHAT_ID'));
  const P = PropertiesService.getScriptProperties();
  if (u.callback_query) {
    const q = u.callback_query, d = String(q.data || '');
    if (String(q.from.id) !== admin) { call('answerCallbackQuery', { callback_query_id: q.id, text: 'Только для продавца' }); return; }
    if (d.indexOf('paid:') === 0) {
      const oid = d.slice(5);
      const res = markPaid_(oid);
      call('answerCallbackQuery', { callback_query_id: q.id, text: res.slice(0, 190), show_alert: true });
      if (q.message) call('editMessageReplyMarkup', { chat_id: q.message.chat.id, message_id: q.message.message_id, reply_markup: orderKb_(oid, true) });
    } else call('answerCallbackQuery', { callback_query_id: q.id });
    return;
  }
  const m = u.message;
  if (!m || !m.text) return;
  const t = m.text.trim(), chat = m.chat.id;
  if (String(m.from.id) !== admin) {
    if (isAdminBot) return; // бот для заказов отвечает только вам
    if (/^\/start/.test(t)) send(chat, 'Добро пожаловать! Каталог и заказ — по кнопке ниже 👇', { inline_keyboard: [[{ text: '🛍 Открыть каталог', web_app: { url: CATALOG_URL } }]] });
    return;
  }
  let mm;
  if ((mm = t.match(/^\/?курс\s*[:=]?\s*(\d+(?:[.,]\d+)?)/i))) {
    const rate = Number(mm[1].replace(',', '.'));
    if (!(rate >= 20 && rate <= 500)) { send(chat, 'Не понял курс. Пример: курс 90'); return; }
    P.setProperty('RATE', String(rate));
    P.setProperty('RATE_DATE', Utilities.formatDate(new Date(), POST_TZ, 'dd.MM.yyyy'));
    send(chat, `✅ Курс обновлён: ${rate} ₽/$. Цены в рублях в мини-аппе пересчитаны (клиенты увидят при следующем открытии).`);
    return;
  }
  if ((mm = t.match(/^\/?привязать\s+№?\s*(\d+)\s+(\S+)/i))) { send(chat, bindCmd_(mm[1], mm[2])); return; }
  if ((mm = t.match(/^\/?партн[её]р\s+@?([A-Za-z0-9_]{3,})\s*$/i))) {
    const l = allowList_(), nick = '@' + mm[1].toLowerCase();
    if (l.indexOf(nick) < 0) l.push(nick);
    P.setProperty('ALLOW_INVITE', JSON.stringify(l));
    send(chat, `✅ ${nick} может приглашать друзей. Кнопка «Пригласить друга» появится у него в профиле мини-аппа.`);
    return;
  }
  if (/^\/?партн[её]ры\s*$/i.test(t)) { send(chat, partnersText_()); return; }
  if ((mm = t.match(/^\/?бонус\S*\s+(\S+)\s+([+-]?\d+(?:[.,]\d+)?)\s*$/i))) { send(chat, bonusCmd_(mm[1], Number(mm[2].replace(',', '.')))); return; }
  if ((mm = t.match(/^\/?визитка(?:\s+(@[A-Za-z0-9_]{4,}))?\s*$/i))) {
    const r = mm[1] ? sendCard_(mm[1]) : sendCard_(chat, call);
    if (mm[1]) send(chat, r && r.ok === false ? '❌ Не получилось опубликовать в ' + mm[1] + ': сделайте бота админом канала с правом публикации.' : '✅ Визитка опубликована в ' + mm[1]);
    return;
  }
  if (/^\/?(start|help|помощь)/i.test(t)) send(chat, BOT_HELP);
}

function pollBot() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    pollOne_(prop_('BOT_TOKEN'), 'BOT_OFFSET', tgCall_, false);
    const at = adminToken_();
    if (at) pollOne_(at, 'ADMIN_BOT_OFFSET', (m, p) => tgApi_(at, m, p), true);
  } finally { lock.releaseLock(); }
}

function pollOne_(token, offKey, call, isAdminBot) {
  const P = PropertiesService.getScriptProperties();
  const off = Number(P.getProperty(offKey) || 0);
  const r = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/getUpdates?timeout=0&offset=' + off +
    '&allowed_updates=' + encodeURIComponent('["message","callback_query"]'), { muteHttpExceptions: true });
  const j = JSON.parse(r.getContentText());
  if (!j.ok) { Logger.log(offKey + ': ' + j.description); return; }
  j.result.forEach(up => {
    P.setProperty(offKey, String(up.update_id + 1));
    try { handleUpdate_(up, call, isAdminBot); } catch (e) { Logger.log('Ошибка: ' + e); }
  });
}

// Запустить один раз: бот начинает принимать команды и кнопку «Оплачен» (проверка каждую минуту)
function setupBot() {
  tgCall_('deleteWebhook', {});
  if (adminToken_()) tgApi_(adminToken_(), 'deleteWebhook', {});
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'pollBot').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('pollBot').timeBased().everyMinutes(1).create();
  clientsSheet_();
  const P = PropertiesService.getScriptProperties();
  if (!P.getProperty('RATE')) { P.setProperty('RATE', '90'); P.setProperty('RATE_DATE', Utilities.formatDate(new Date(), POST_TZ, 'dd.MM.yyyy')); }
  Logger.log('Бот включён: команды и кнопки проверяются каждую минуту. Курс: ' + P.getProperty('RATE'));
}
