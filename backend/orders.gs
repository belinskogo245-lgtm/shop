// Раздел «Заказы» в мини-аппе — только для вас (проверка по подписи Telegram, как у «Партнёров»).
// Отдаёт последние заказы из листа «Заказы» + курс ₽/$ для накладной в рублях.
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
