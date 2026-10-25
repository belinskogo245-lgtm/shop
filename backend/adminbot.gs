// ===== ОТДЕЛЬНЫЙ БОТ ДЛЯ ЗАКАЗОВ =====
// Клиенты пользуются @gsm_catalog_bot, а заказы и кнопка «Оплачен» приходят вам во второй бот.
// Настройка: свойство скрипта ADMIN_BOT_TOKEN = токен второго бота (из BotFather),
// нажмите «Старт» во втором боте и один раз запустите setupBot.
// Пока ADMIN_BOT_TOKEN не задан (или вы не нажали «Старт»), заказы приходят по-старому в @gsm_catalog_bot.
// Команды («курс 90», «визитка», «партнёры»…) работают в обоих ботах.

function adminToken_() { return PropertiesService.getScriptProperties().getProperty('ADMIN_BOT_TOKEN') || ''; }

function tgApi_(token, method, params) {
  const r = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/' + method, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(params || {}), muteHttpExceptions: true
  });
  try { return JSON.parse(r.getContentText()); } catch (e) { return { ok: false, description: String(e) }; }
}

function sendVia_(call, chat, text, kb) {
  const p = { chat_id: chat, text: text, disable_web_page_preview: true };
  if (kb) p.reply_markup = kb;
  return call('sendMessage', p);
}

function sendAdmin_(text, kb) {
  const t = adminToken_(), chat = prop_('ADMIN_CHAT_ID');
  if (t) {
    const r = sendVia_((m, p) => tgApi_(t, m, p), chat, text, kb);
    if (r && r.ok) return r;
    Logger.log('Бот для заказов не смог отправить (нажмите в нём «Старт»): ' + (r && r.description));
  }
  return sendTg_(chat, text, kb);
}

// Проверка: запустите, во втором боте должно прийти тестовое сообщение
function testAdminBot() {
  const r = sendAdmin_('✅ Бот для заказов подключён. Новые заказы будут приходить сюда.');
  Logger.log(JSON.stringify(r));
}
