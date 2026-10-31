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
      sendAdmin_(text, orderKb_(oid));
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

