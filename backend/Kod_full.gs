/**
 * Бэкенд мини-аппа aromaopt на Google Apps Script (бесплатно).
 * Вставляется в таблицу «Каталог мини-аппа aromaopt»: Расширения → Apps Script.
 *
 * Свойства скрипта (Настройки проекта → Свойства скрипта):
 *   BOT_TOKEN      — токен бота от @BotFather
 *   ADMIN_CHAT_ID  — ваш chat id (получите функцией whoAmI ниже)
 */
const SHEET_ITEMS = 'Товары';
const SHEET_ORDERS = 'Заказы';

function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// GET ?action=catalog → весь каталог из листа «Товары»
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'order') return json_(getOrderById_(e.parameter.id));
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ITEMS);
  const rows = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 13).getValues();
  // столбцы: A id, B Категория, C Бренд, D Название, E Объём, F Пол, G Цена опт, H Цена от 100, I Мин. заказ, J Фото, K В наличии, L Хит, M Описание
  const items = rows
    .filter(r => r[3] && Number(r[6]) > 0)
    .map(r => [r[0], r[1], String(r[3]), String(r[4] || ''), Number(r[6]), String(r[10]).toLowerCase() === 'нет' ? 0 : 1, String(r[9] || ''), String(r[12] || '')]);
  return json_({ ok: true, items, updated: new Date().toISOString(), rate: Number(prop_('RATE')) || 0, rateDate: prop_('RATE_DATE') || '' });
}

// POST {action:'order'|'me'|'partners'} → заказ в лист «Заказы» + сообщение вам в Telegram; данные партнёрки
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'me') return json_(me_(body));
    if (body.action === 'partners') return json_(partners_(body));
    if (body.action === 'orders') return json_(orders_(body));
    if (body.action !== 'order') return json_({ ok: false, error: 'unknown action' });
    const o = body.order;
    if (!o || !o.items || !o.items.length) return json_({ ok: false, error: 'empty order' });

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const ss = SpreadsheetApp.getActive();
      const sh = ss.getSheetByName(SHEET_ORDERS) || ss.insertSheet(SHEET_ORDERS);
      if (sh.getLastRow() === 0) sh.appendRow(['Дата', 'Имя', 'Контакт', 'Telegram', 'Город/ТК', 'Комментарий', 'Позиции', 'Шт.', 'Сумма опт $', 'Наценка $', 'Итого $', 'Статус']);
      const tgu = (o.customer && o.customer.tg) || {};
      const lines = o.items.map(i => i.cur === 'RUB'
        ? `${i.id} · ${i.name} × ${i.qty} = ${round_(i.price * i.qty)} ₽ (акция 62 мл, ${i.price} ₽/шт)`
        : `${i.id} · ${i.name} × ${i.qty} = ${round_(i.price * i.qty)} $`).join('\n');
      if (sh.getRange(1, 13).getValue() === '') sh.getRange(1, 13).setValue('Итого ₽ (акция 62 мл)');
      if (sh.getRange(1, 14).getValue() === '') sh.getRange(1, 14).setValue('Код заказа');
      ORDER_EXTRA_HEAD.forEach((h, k) => { if (sh.getRange(1, 15 + k).getValue() === '') sh.getRange(1, 15 + k).setValue(h); });
      const oid = newOid_();
      const pr = applyPartner_(body, o);
      const allLines = lines + (pr.giftLines.length ? '\n' + pr.giftLines.join('\n') : '');
      sh.appendRow([new Date(), o.customer.name, o.customer.contact, tgu.user ? '@' + tgu.user : (tgu.id || ''), o.customer.city, o.customer.note,
        allLines, o.qty, round_(o.base), round_(o.extra), round_(pr.total), 'новый', o.totalRub || '', oid, pr.uid, pr.refId, pr.refGifts || '', pr.disc || '', pr.bonusUsed || '']);
      const num = sh.getLastRow() - 1;
      const rate = Number(prop_('RATE')) || 0;

      const text =
        `🛒 Новый заказ №${num} из мини-аппа\n\n` +
        `${o.customer.name} · ${o.customer.contact}` + (tgu.user ? ` · @${tgu.user}` : '') + `\n` +
        (o.customer.city ? `Город/ТК: ${o.customer.city}\n` : '') +
        (o.customer.note ? `Комментарий: ${o.customer.note}\n` : '') +
        `\n${allLines}\n\n` +
        `Штук: ${o.qty}` +
        (o.base ? `\nПо опту: ${round_(o.base)} $` : '') +
        (o.extra ? `\nНаценка (ниже 120 $): +${round_(o.extra)} $` : '') +
        (pr.bonusUsed ? `\nОплачено бонусами: −${pr.bonusUsed} $` : '') +
        `\nИтого: ${round_(pr.total)} $` +
        (rate ? ` (≈ ${rub_(pr.total * rate)} ₽ по курсу ${rate})` : '') +
        (pr.notes.length ? '\n\n' + pr.notes.join('\n') : '');
      sendTg_(prop_('ADMIN_CHAT_ID'), text, orderKb_(oid));
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// ===== СТРАНИЦА ЗАКАЗА С ФОТО (order.html) =====
// У каждого заказа секретный код (столбец N «Код заказа»); кнопка «Открыть заказ» открывает order.html?id=код
function newOid_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 16); }
function orderKb_(oid, paid) { return { inline_keyboard: [[{ text: '🛍 Открыть заказ', web_app: { url: CATALOG_URL + 'order.html?id=' + oid } }], [paid ? { text: '✅ Оплачен', callback_data: 'noop' } : { text: '💰 Отметить «Оплачен»', callback_data: 'paid:' + oid }]] }; }

function rowToOrder_(r, n) {
  const items = String(r[6]).split('\n').map(l => {
    const m = l.match(/^(.+?) · (.+) × (\d+) = ([\d.]+) /);
    if (!m) return null;
    const q = Number(m[3]);
    return [m[1].trim(), q, Math.round(Number(m[4]) / q * 100) / 100, m[2].slice(0, 80)];
  }).filter(Boolean);
  const d = r[0] instanceof Date ? Utilities.formatDate(r[0], POST_TZ, 'dd.MM.yyyy HH:mm') : String(r[0]);
  const tg = String(r[3] || '');
  return { n: String(n), d: d, t: Number(r[10]) || null, x: Number(r[9]) || 0, i: items,
    c: { name: String(r[1]), contact: String(r[2]), city: String(r[4]), note: String(r[5]), tgu: tg.charAt(0) === '@' ? tg.slice(1) : '' } };
}

function getOrderById_(id) {
  id = String(id || '');
  if (id.length < 12) return { ok: false };
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ORDERS);
  if (!sh || sh.getLastRow() < 2) return { ok: false };
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, 14).getValues();
  const k = v.findIndex(x => String(x[13]) === id);
  return k < 0 ? { ok: false } : { ok: true, order: rowToOrder_(v[k], k + 1) };
}

// Повторно прислать заказ из листа «Заказы» с кнопкой «Открыть заказ».
// Укажите номер строки в RESEND_ROW и нажмите «Выполнить» (строка 2 — первый заказ).
const RESEND_ROW = 2;
function resendOrder() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ORDERS);
  if (sh.getRange(1, 14).getValue() === '') sh.getRange(1, 14).setValue('Код заказа');
  const r = sh.getRange(RESEND_ROW, 1, 1, 14).getValues()[0];
  let oid = String(r[13] || '');
  if (!oid) { oid = newOid_(); sh.getRange(RESEND_ROW, 14).setValue(oid); }
  const o = rowToOrder_(r, RESEND_ROW - 1);
  const text = '🛒 Заказ из мини-аппа (повторно, с фото)\n\n' + o.c.name + ' · ' + o.c.contact +
    (o.c.city ? '\nГород/ТК: ' + o.c.city : '') + '\nОт: ' + o.d +
    '\n\nПозиций: ' + o.i.length + ', штук: ' + r[7] + '\nИтого: ' + r[10] + ' $';
  sendTg_(prop_('ADMIN_CHAT_ID'), text, orderKb_(oid));
  Logger.log('Отправлено: строка ' + RESEND_ROW + ', код ' + oid);
}

function round_(n) { return Math.round(Number(n) * 100) / 100; }

function sendTg_(chatId, text, kb) {
  UrlFetchApp.fetch(`https://api.telegram.org/bot${prop_('BOT_TOKEN')}/sendMessage`, {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify(Object.assign({ chat_id: chatId, text: text.slice(0, 4000) }, kb ? { reply_markup: kb } : {}))
  });
}

// 1) Напишите своему боту /start в Telegram. 2) Запустите эту функцию. 3) В журнале появится ваш chat id.
function whoAmI() {
  const r = UrlFetchApp.fetch(`https://api.telegram.org/bot${prop_('BOT_TOKEN')}/getUpdates`, { muteHttpExceptions: true });
  const ups = JSON.parse(r.getContentText()).result || [];
  ups.forEach(u => { const m = u.message || {}; if (m.chat) Logger.log(`chat id: ${m.chat.id}  (${m.chat.first_name || ''} @${m.chat.username || ''})`); });
  if (!ups.length) Logger.log('Пусто: сначала напишите боту /start');
}

// Проверка, что сообщения доходят
function testMessage() { sendTg_(prop_('ADMIN_CHAT_ID'), 'Проверка связи: мини-апп aromaopt подключён ✅'); }

// ===== ФОТО: перенос картинок из прайсов в колонку J «Фото» листа «Товары» =====
// Запуск: выберите функцию importPhotos и нажмите «Выполнить». Если фото много,
// скрипт сам продолжит через минуту, пока не перенесёт все (см. журнал).
const PHOTO_SOURCES = [
  { sheet: 'src_25ml',     cat: 'Тестеры 25 мл',          nameCol: 2, imgCol: 1, start: 3 },
  { sheet: 'src_35ml',     cat: 'Тестеры 35 мл',          nameCol: 2, imgCol: 1, start: 3 },
  { sheet: 'src_niche_DJ', cat: 'Селектив и ниша',        nameCol: 2, imgCol: 1, start: 3 },
  { sheet: 'src_turk100',  cat: 'Тестер 100 мл Турция',   nameCol: 2, imgCol: 1, start: 2 },
  { sheet: 'src_arab',     cat: 'Арабская парфюмерия',    nameCol: 2, imgCol: 1, start: 3 },
];
function normName_(s) { return String(s).replace(/\s+/g, ' ').trim().toLowerCase(); }

function importPhotos() {
  const t0 = Date.now();
  const ss = SpreadsheetApp.getActive();
  const P = PropertiesService.getScriptProperties();
  let folder;
  const fid = P.getProperty('PHOTO_FOLDER');
  if (fid) folder = DriveApp.getFolderById(fid);
  else {
    folder = DriveApp.createFolder('aromaopt фото для мини-аппа');
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    P.setProperty('PHOTO_FOLDER', folder.getId());
  }
  const items = ss.getSheetByName(SHEET_ITEMS);
  const n = items.getLastRow() - 1;
  const data = items.getRange(2, 1, n, 10).getValues();
  const idx = {};
  data.forEach((r, i) => { const k = r[1] + '|' + normName_(r[3]); (idx[k] = idx[k] || []).push(i); });
  const photo = data.map(r => [r[9]]);
  let done = 0, miss = 0;

  for (const s of PHOTO_SOURCES) {
    const sh = ss.getSheetByName(s.sheet);
    if (!sh) continue;
    const rowsN = sh.getLastRow() - s.start + 1;
    if (rowsN < 1) continue;
    const vals = sh.getRange(s.start, 1, rowsN, Math.max(s.nameCol, s.imgCol)).getValues();
    for (const row of vals) {
      const img = row[s.imgCol - 1], name = row[s.nameCol - 1];
      if (!name || !img || typeof img.getContentUrl !== 'function') continue;
      const rows = idx[s.cat + '|' + normName_(name)];
      if (!rows) { miss++; continue; }
      if (rows.every(i => photo[i][0])) continue;
      if (Date.now() - t0 > 270000) {
        items.getRange(2, 10, n, 1).setValues(photo);
        Logger.log('Пауза на минуту: перенесено ' + done + ' фото. Скрипт продолжит сам.');
        scheduleNextPhotos_();
        return;
      }
      try {
        const url = img.getContentUrl().replace(/=s\d+/, '=s600');
        const blob = UrlFetchApp.fetch(url).getBlob().setName(normName_(name).slice(0, 60) + '.jpg');
        const f = folder.createFile(blob);
        f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        const link = 'https://lh3.googleusercontent.com/d/' + f.getId() + '=w600';
        rows.forEach(i => photo[i][0] = link);
        done++;
      } catch (e) { miss++; }
    }
  }
  items.getRange(2, 10, n, 1).setValues(photo);
  clearNextPhotos_();
  Logger.log('Готово: перенесено ' + done + ' фото, не найдено/ошибок: ' + miss);
}
function scheduleNextPhotos_() { clearNextPhotos_(); ScriptApp.newTrigger('importPhotos').timeBased().after(60 * 1000).create(); }
function clearNextPhotos_() { ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'importPhotos').forEach(t => ScriptApp.deleteTrigger(t)); }


// ===== ФОТО-2: картинки «поверх таблицы» (Селектив/наборы, Оригинал AD) =====
// Запуск: функция importOverGridPhotos → «Выполнить». Дальше продолжает сама каждую минуту.
const OG_SOURCES = [
  { file: '1DgpgQKWj1HelbIbu9FGGoaiH8yUOJx0JlMZDTTWP_Bw', gid: 1680991244, cat: 'Селектив / наборы',      nameCol: 2, start: 3 },
  { file: '1NIyAobmEBKiCGw_9MuBO8zkL5z75lbjRIp6coJgH7eI', gid: 1764180959, cat: 'Оригинальный парфюм AD', nameCol: 3, start: 2 },
];
const OG_CHUNK = 80;

function photoFolder_() {
  const P = PropertiesService.getScriptProperties();
  const fid = P.getProperty('PHOTO_FOLDER');
  if (fid) return DriveApp.getFolderById(fid);
  const f = DriveApp.createFolder('aromaopt фото для мини-аппа');
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  P.setProperty('PHOTO_FOLDER', f.getId());
  return f;
}

function importOverGridPhotos() {
  const t0 = Date.now();
  const P = PropertiesService.getScriptProperties();
  let st = JSON.parse(P.getProperty('OG_STATE') || '{"src":0,"from":0}');
  const items = SpreadsheetApp.getActive().getSheetByName(SHEET_ITEMS);
  const folder = photoFolder_();

  while (st.src < OG_SOURCES.length) {
    const s = OG_SOURCES[st.src];
    const srcSheet = SpreadsheetApp.openById(s.file).getSheets().filter(x => x.getSheetId() === s.gid)[0];
    const last = srcSheet.getLastRow();
    if (!st.from) st.from = s.start;
    if (st.from > last) { st = { src: st.src + 1, from: 0 }; P.setProperty('OG_STATE', JSON.stringify(st)); continue; }
    const chunk = st.chunk || OG_CHUNK;
    const to = Math.min(last, st.from + chunk - 1);
    const names = srcSheet.getRange(1, s.nameCol, last, 1).getValues().map(r => r[0]);

    // временная копия листа только с нужными строками
    const tmp = SpreadsheetApp.create('tmp_photos_aromaopt');
    const copied = srcSheet.copyTo(tmp);
    tmp.getSheets().filter(x => x.getSheetId() !== copied.getSheetId()).forEach(x => tmp.deleteSheet(x));
    const maxR = copied.getMaxRows();
    if (to < maxR) copied.deleteRows(to + 1, maxR - to);
    if (st.from > 1) copied.deleteRows(1, st.from - 1);
    SpreadsheetApp.flush();
    const exportUrl = 'https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?format=xlsx';
    let files = null;
    try {
      for (let at = 0; at < 3 && !files; at++) {
        if (at) Utilities.sleep(5000);
        const resp = UrlFetchApp.fetch(exportUrl, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
        if (resp.getResponseCode() !== 200) continue;
        try { files = Utilities.unzip(resp.getBlob().setContentType('application/zip')); } catch (e) { files = null; }
      }
    } finally {
      DriveApp.getFileById(tmp.getId()).setTrashed(true);
    }
    if (!files) {
      if (chunk > 10) { st.chunk = Math.max(10, Math.floor(chunk / 2)); Logger.log('Строки ' + st.from + '–' + to + ' не выгрузились, уменьшаю кусок до ' + st.chunk + ' строк'); }
      else { Logger.log('Строки ' + st.from + '–' + to + ' пропущены: Google не отдал файл'); st.from = to + 1; st.chunk = OG_CHUNK; }
      P.setProperty('OG_STATE', JSON.stringify(st));
      if (Date.now() - t0 > 180000) { scheduleOG_(); Logger.log('Пауза на минуту, продолжу сам.'); return; }
      continue;
    }
    const byName = {};
    files.forEach(f => byName[f.getName()] = f);
    const drawName = Object.keys(byName).filter(n => /^xl\/drawings\/drawing\d+\.xml$/.test(n))[0];
    let done = 0, found = 0;
    if (drawName) {
      const rels = {};
      const relsXml = (byName[drawName.replace('drawings/', 'drawings/_rels/') + '.rels'] || { getDataAsString: () => '' }).getDataAsString();
      (relsXml.match(/<Relationship\b[^>]*>/g) || []).forEach(tag => {
        const id = (tag.match(/Id="([^"]+)"/) || [])[1], t = (tag.match(/Target="([^"]+)"/) || [])[1];
        if (id && t) rels[id] = t;
      });
      const data = items.getRange(2, 1, items.getLastRow() - 1, 10).getValues();
      const idx = {};
      data.forEach((r, i) => { const k = r[1] + '|' + normName_(r[3]); (idx[k] = idx[k] || []).push(i); });
      const photo = data.map(r => [r[9]]);
      const anchors = byName[drawName].getDataAsString().split(/<xdr:(?:twoCellAnchor|oneCellAnchor)\b/).slice(1);
      anchors.forEach(a => {
        const rm = a.match(/<xdr:from>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/), em = a.match(/r:embed="([^"]+)"/);
        if (!rm || !em) return;
        found++;
        const row = st.from + Number(rm[1]);
        const name = names[row - 1];
        if (!name) return;
        const rows = idx[s.cat + '|' + normName_(name)];
        if (!rows || rows.every(i => photo[i][0])) return;
        const target = rels[em[1]];
        if (!target) return;
        const img = byName['xl/' + target.replace(/^\.\.\//, '')];
        if (!img) return;
        const f = folder.createFile(img.copyBlob().setName(normName_(name).slice(0, 60) + '_' + target.split('/').pop()));
        f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        const link = 'https://lh3.googleusercontent.com/d/' + f.getId() + '=w600';
        rows.forEach(i => photo[i][0] = link);
        done++;
      });
      items.getRange(2, 10, photo.length, 1).setValues(photo);
    }
    Logger.log(s.cat + ': строки ' + st.from + '–' + to + ', картинок найдено ' + found + ', фото добавлено ' + done);
    st.from = to + 1;
    P.setProperty('OG_STATE', JSON.stringify(st));
    if (Date.now() - t0 > 180000) { scheduleOG_(); Logger.log('Пауза на минуту, продолжу сам.'); return; }
  }
  clearOG_();
  P.deleteProperty('OG_STATE');
  Logger.log('Готово: Селектив/наборы и Оригинал AD обработаны.');
}
function scheduleOG_() { clearOG_(); ScriptApp.newTrigger('importOverGridPhotos').timeBased().after(60 * 1000).create(); }
function clearOG_() { ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'importOverGridPhotos').forEach(t => ScriptApp.deleteTrigger(t)); }


// ===== АВТОПОСТИНГ В TELEGRAM-КАНАЛ (лист «Посты») =====
// Колонки листа «Посты»: A Дата (ГГГГ-ММ-ДД), B Время (ЧЧ:ММ), C Тип, D Текст (HTML), E Фото (ссылки через запятую), F Статус
// Публикуются строки со статусом «ждёт», когда наступили дата и время (по Красноярску).
// Один раз запустите setupPosting — дальше бот проверяет очередь каждые 10 минут.
const SHEET_POSTS = 'Посты';
const POST_CHANNEL = '@parfum_vip_opt';
const CATALOG_URL = 'https://belinskogo245-lgtm.github.io/shop/';
const POST_TZ = 'Asia/Krasnoyarsk';

function postDue() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_POSTS);
  if (!sh || sh.getLastRow() < 2) return;
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
  const now = Utilities.formatDate(new Date(), POST_TZ, 'yyyy-MM-dd HH:mm');
  let sent = 0;
  for (let i = 0; i < vals.length && sent < 3; i++) {
    const r = vals[i];
    if (String(r[5]).trim() !== 'ждёт') continue;
    const day = r[0] instanceof Date ? Utilities.formatDate(r[0], POST_TZ, 'yyyy-MM-dd') : String(r[0]).trim();
    let tm = r[1] instanceof Date ? Utilities.formatDate(r[1], POST_TZ, 'HH:mm') : String(r[1]).trim();
    if (tm.length === 4) tm = '0' + tm;
    if (!day || day + ' ' + tm > now) continue;
    let res;
    try { res = sendPost_(String(r[3]), String(r[4])); } catch (e) { res = String(e); }
    sh.getRange(i + 2, 6).setValue(res === true ? 'опубликован ' + now : 'ошибка: ' + res);
    if (prop_('VK_TOKEN') && VK_ON && VK_TIMES.indexOf(tm) >= 0) {
      let vr;
      try { vr = sendVk_(String(r[3]), String(r[4])); } catch (e) { vr = 'ошибка: ' + e; }
      sh.getRange(i + 2, 7).setValue('ВК: ' + vr);
    }
    sent++;
  }
}

// ===== ВКОНТАКТЕ: те же посты в группу VK_GROUP (ключ сообщества — свойство скрипта VK_TOKEN) =====
const VK_GROUP = 'sibir_optom24';
const VK_ON = true;
const VK_TIMES = ['12:00', '18:00', '20:30']; // в ВК уходят только посты с этим временем // true — публиковать в ВК
const VK_V = '5.199';

function vkApi_(method, params) {
  const p = Object.assign({ access_token: prop_('VK_TOKEN'), v: VK_V }, params);
  const resp = UrlFetchApp.fetch('https://api.vk.com/method/' + method, { method: 'post', payload: p, muteHttpExceptions: true });
  const j = JSON.parse(resp.getContentText());
  if (j.error) throw new Error(j.error.error_code + ' ' + j.error.error_msg);
  return j.response;
}

function vkGroupId_() {
  const P = PropertiesService.getScriptProperties();
  let id = P.getProperty('VK_GROUP_ID_' + VK_GROUP);
  if (!id) {
    const r = vkApi_('groups.getById', { group_id: VK_GROUP });
    id = String((r.groups || r)[0].id);
    P.setProperty('VK_GROUP_ID_' + VK_GROUP, id);
  }
  return id;
}

function htmlToText_(h) {
  return String(h).replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '')
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').trim();
}

// Возвращает 'опубликован с фото', 'опубликован без фото (…)' или бросает ошибку
function sendVk_(html, photos) {
  const gid = vkGroupId_();
  let text = htmlToText_(html).replace(/👇\s*$/, '').trim();
  text += '\n\n🛍 Каталог и заказ: ' + CATALOG_URL;
  const ph = String(photos).split(',').map(s => s.trim()).filter(s => /^https?:/.test(s)).slice(0, 10);
  let att = [], note = '';
  if (ph.length) {
    try {
      const srv = vkApi_('photos.getWallUploadServer', { group_id: gid });
      ph.forEach((u, k) => {
        const up = JSON.parse(UrlFetchApp.fetch(srv.upload_url, { method: 'post', payload: { photo: photoBlob_(u, 'p' + k) }, muteHttpExceptions: true }).getContentText());
        const saved = vkApi_('photos.saveWallPhoto', { group_id: gid, photo: up.photo, server: up.server, hash: up.hash });
        att.push('photo' + saved[0].owner_id + '_' + saved[0].id);
      });
    } catch (e) { att = []; note = ' без фото (' + String(e.message || e).slice(0, 60) + ')'; }
  }
  if (!att.length) att.push(CATALOG_URL);
  vkApi_('wall.post', { owner_id: '-' + gid, from_group: 1, message: text, attachments: att.join(',') });
  return 'опубликован' + (note || (ph.length ? ' с фото' : ''));
}

// Проверка ВК: публикует в группу строку листа «Посты» с номером VK_TEST_ROW
const VK_TEST_ROW = 38;
function testVk() {
  const r = SpreadsheetApp.getActive().getSheetByName(SHEET_POSTS).getRange(VK_TEST_ROW, 1, 1, 6).getValues()[0];
  const res = sendVk_(String(r[3]), String(r[4]));
  SpreadsheetApp.getActive().getSheetByName(SHEET_POSTS).getRange(VK_TEST_ROW, 7).setValue('ВК: ' + res);
  Logger.log('ВК: ' + res);
}

function tgCall_(method, payload) {
  const resp = UrlFetchApp.fetch('https://api.telegram.org/bot' + prop_('BOT_TOKEN') + '/' + method, {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true, payload: JSON.stringify(payload)
  });
  const j = JSON.parse(resp.getContentText());
  return j.ok ? true : (j.description || 'неизвестная ошибка');
}

function tgUpload_(method, fields) {
  const resp = UrlFetchApp.fetch('https://api.telegram.org/bot' + prop_('BOT_TOKEN') + '/' + method, {
    method: 'post', muteHttpExceptions: true, payload: fields
  });
  const j = JSON.parse(resp.getContentText());
  return j.ok ? true : (j.description || 'неизвестная ошибка');
}

// Скачиваем фото сами и загружаем файлом — Телеграм не всегда может скачать ссылку
function photoBlob_(u, name) {
  const r = UrlFetchApp.fetch(u, { muteHttpExceptions: true, followRedirects: true });
  if (r.getResponseCode() !== 200) throw new Error('фото не скачалось: ' + u);
  return r.getBlob().setName(name + '.jpg');
}

function sendPost_(text, photos) {
  const kb = { inline_keyboard: [[{ text: '🛍 Открыть каталог', url: CATALOG_URL }]] };
  const ph = photos.split(',').map(s => s.trim()).filter(s => /^https?:/.test(s));
  if (ph.length > 1) {
    const fields = { chat_id: POST_CHANNEL };
    const media = ph.slice(0, 10).map((u, k) => {
      fields['p' + k] = photoBlob_(u, 'p' + k);
      return k === 0
        ? { type: 'photo', media: 'attach://p0', caption: (text + '\n\n👉 <a href="' + CATALOG_URL + '">Открыть каталог</a>').slice(0, 1024), parse_mode: 'HTML' }
        : { type: 'photo', media: 'attach://p' + k };
    });
    fields.media = JSON.stringify(media);
    return tgUpload_('sendMediaGroup', fields);
  }
  if (ph.length === 1 && text.length <= 1024)
    return tgUpload_('sendPhoto', { chat_id: POST_CHANNEL, photo: photoBlob_(ph[0], 'p0'), caption: text, parse_mode: 'HTML', reply_markup: JSON.stringify(kb) });
  return tgCall_('sendMessage', { chat_id: POST_CHANNEL, text: text, parse_mode: 'HTML', reply_markup: kb, disable_web_page_preview: true });
}

// Запустить один раз: включает проверку очереди постов каждые 10 минут
function setupPosting() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'postDue').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('postDue').timeBased().everyMinutes(10).create();
  Logger.log('Автопостинг включён: очередь проверяется каждые 10 минут.');
}

// Проверка: бот — админ канала и может публиковать
function testChannel() {
  Logger.log(tgCall_('sendMessage', { chat_id: POST_CHANNEL, text: 'Проверка автопостинга ✅ (это сообщение можно удалить)' }));
}


// ===== ПАРТНЁРСКАЯ ПРОГРАММА «ПРИВЕДИ ДРУГА» + КУРС ₽/$ + КОМАНДЫ БОТУ =====
// Лист «Партнёры» заполняется сам. Ссылка партнёра: https://t.me/gsm_catalog_bot?startapp=r_<TG ID>
// Новый клиент по ссылке: флакон 62 мл в подарок к первому заказу от MIN_GIFT_ORDER $.
// Пригласившему: флакон 62 мл после того, как вы нажмёте «Оплачен» под первым заказом друга,
// и BONUS_PCT% бонусами с КАЖДОГО оплаченного заказа друга от MIN_GIFT_ORDER $ — навсегда.
// Бонусы ($) копятся в профиле: ими можно оплатить свой заказ (галочка в корзине) или вывести деньгами
// (вы переводите сами и списываете командой «бонус @ivan -10»).
// Команды боту (только от вас): «курс 90», «привязать 15 @ivan», «партнёры», «бонус @ivan 5», «визитка».
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
// Ссылка для приглашения есть у всех, кто открыл мини-апп (награда всё равно только после оплаты заказа друга)
function canInvite_(v) { return !!(v && v[0]); }

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

const BOT_HELP = 'Команды:\n• курс 90 — поменять курс ₽/$ в мини-аппе\n• привязать 15 @ivan — привязать клиента из заказа №15 к партнёру (если он пришёл не по ссылке)\n• партнёры — список партнёров\n• бонус @ivan 5 — добавить клиенту 5 $ бонусов; бонус @ivan -10 — списать (например, когда перевели ему деньгами)';

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
  if (/^\/?партн[её]ры\s*$/i.test(t)) { sendTg_(chat, partnersText_()); return; }
  if ((mm = t.match(/^\/?бонус\S*\s+(\S+)\s+([+-]?\d+(?:[.,]\d+)?)\s*$/i))) { sendTg_(chat, bonusCmd_(mm[1], Number(mm[2].replace(',', '.')))); return; }
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

// Раздел «Заказы» в мини-аппе — только для владельца (проверка подписи Telegram, как у «Партнёров»)
function orders_(body) {
  const u = verifyInit_(body.initData);
  if (!u || String(u.id) !== String(prop_('ADMIN_CHAT_ID'))) return { ok: false, error: 'forbidden' };
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ORDERS);
  const rate = Number(prop_('RATE')) || 0;
  if (!sh || sh.getLastRow() < 2) return { ok: true, rows: [], rate: rate };
  const last = sh.getLastRow(), first = Math.max(2, last - 99);
  const v = sh.getRange(first, 1, last - first + 1, 14).getValues();
  const rows = v.map((r, k) => ({
    num: first + k - 1,
    date: r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm') : String(r[0]),
    name: String(r[1] || ''), contact: String(r[2] || ''), tg: String(r[3] || ''), city: String(r[4] || ''), note: String(r[5] || ''),
    lines: String(r[6] || ''), qty: Number(r[7]) || 0, base: Number(r[8]) || 0, extra: Number(r[9]) || 0, total: Number(r[10]) || 0,
    status: String(r[11] || ''), oid: String(r[13] || '')
  })).reverse();
  return { ok: true, rows: rows, rate: rate };
}
