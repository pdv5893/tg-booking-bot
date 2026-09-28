import 'dotenv/config';
import { Bot, Context, GrammyError, HttpError, InlineKeyboard, InputFile, Keyboard, session, type SessionFlavor } from 'grammy';
import { HttpsProxyAgent } from 'https-proxy-agent';
import { SocksProxyAgent } from 'socks-proxy-agent';
import { config } from './config.js';
import { availableDates, dateKey, formatDate, freeSlots } from './schedule.js';
import {
  addBooking,
  EXCEL_PATH,
  exportToExcel,
  forDate,
  getBooking,
  setStatus,
  upcomingForClient,
  type Booking,
} from './storage.js';

// ---------------------------------------------------------------------------
// Настройка
// ---------------------------------------------------------------------------

const BOT_TOKEN = process.env.BOT_TOKEN;
const ADMIN_CHAT_ID = Number(process.env.ADMIN_CHAT_ID);
if (!BOT_TOKEN) throw new Error('BOT_TOKEN is not set. Copy .env.example to .env and fill it in.');
if (!ADMIN_CHAT_ID) console.warn('ADMIN_CHAT_ID is not set: new bookings will not be sent anywhere. Send /myid to the bot.');

/** Черновик записи, который клиент заполняет по шагам */
interface Draft {
  serviceId?: string;
  date?: string;
  time?: string;
  name?: string;
  phone?: string;
}

interface SessionData {
  /** Какого текстового ответа бот сейчас ждёт от клиента */
  step: 'idle' | 'name' | 'phone';
  draft: Draft;
}

type MyContext = Context & SessionFlavor<SessionData>;

/**
 * Прокси для запросов к Telegram (необязательно). Нужен, если api.telegram.org
 * недоступен из вашей сети. Поддерживаются http(s):// и socks5:// адреса,
 * например PROXY_URL=socks5://127.0.0.1:10808 от локального VPN-клиента.
 */
const PROXY_URL = process.env.PROXY_URL?.trim();
const proxyAgent = PROXY_URL
  ? PROXY_URL.startsWith('socks')
    ? new SocksProxyAgent(PROXY_URL)
    : new HttpsProxyAgent(PROXY_URL)
  : undefined;
if (proxyAgent) console.log(`Using proxy ${PROXY_URL}`);

const bot = new Bot<MyContext>(BOT_TOKEN, {
  client: proxyAgent ? { baseFetchConfig: { agent: proxyAgent, compress: true } } : undefined,
});
bot.use(session({ initial: (): SessionData => ({ step: 'idle', draft: {} }) }));

// ---------------------------------------------------------------------------
// Вспомогательное
// ---------------------------------------------------------------------------

const MENU = {
  book: '📝 Записаться',
  prices: '💰 Услуги и цены',
  contacts: '📍 Контакты',
  my: '📋 Мои записи',
  // Кнопки владельца — видны только в чате ADMIN_CHAT_ID
  today: '📅 Записи на сегодня',
  export: '📊 Выгрузка в Excel',
};

const clientMenu = new Keyboard().text(MENU.book).text(MENU.prices).row().text(MENU.my).text(MENU.contacts).resized();

const adminMenu = new Keyboard()
  .text(MENU.book)
  .text(MENU.prices)
  .row()
  .text(MENU.my)
  .text(MENU.contacts)
  .row()
  .text(MENU.today)
  .text(MENU.export)
  .resized();

/** Защита от HTML-разметки в том, что ввёл пользователь */
function esc(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function rub(n: number): string {
  return `${n.toLocaleString('ru-RU')} ₽`;
}

function findService(id: string | undefined) {
  return config.services.find((s) => s.id === id);
}

function isAdmin(ctx: MyContext): boolean {
  return ctx.chat?.id === ADMIN_CHAT_ID;
}

/** Владелец видит в меню ещё и свои кнопки */
function menuFor(ctx: MyContext): Keyboard {
  return isAdmin(ctx) ? adminMenu : clientMenu;
}

function bookingText(b: Booking): string {
  return (
    `<b>${esc(b.serviceName)}</b> — ${rub(b.price)}\n` +
    `📅 ${formatDate(b.date)}, ${b.time}\n` +
    `👤 ${esc(b.name)}, ${esc(b.phone)}`
  );
}

function resetFlow(ctx: MyContext) {
  ctx.session.step = 'idle';
  ctx.session.draft = {};
}

// ---------------------------------------------------------------------------
// Клавиатуры шагов записи
// ---------------------------------------------------------------------------

function servicesKeyboard(): InlineKeyboard {
  const kb = new InlineKeyboard();
  for (const s of config.services) kb.text(`${s.name} — ${rub(s.price)}`, `svc:${s.id}`).row();
  return kb;
}

function datesKeyboard(): InlineKeyboard {
  const kb = new InlineKeyboard();
  availableDates().forEach((d, i) => {
    kb.text(formatDate(d), `date:${d}`);
    if (i % 3 === 2) kb.row();
  });
  return kb.row().text('← Назад к услугам', 'back:services');
}

function timesKeyboard(date: string): InlineKeyboard {
  const kb = new InlineKeyboard();
  freeSlots(date).forEach((t, i) => {
    kb.text(t, `time:${t}`);
    if (i % 4 === 3) kb.row();
  });
  return kb.row().text('← Назад к датам', 'back:dates');
}

// ---------------------------------------------------------------------------
// Команды и главное меню
// ---------------------------------------------------------------------------

bot.command('start', async (ctx) => {
  resetFlow(ctx);
  await ctx.reply(
    `Здравствуйте! Это бот <b>${esc(config.businessName)}</b>.\n\n` +
      'Здесь можно записаться на услугу, посмотреть цены и свои записи.',
    { parse_mode: 'HTML', reply_markup: menuFor(ctx) }
  );
});

// Помогает владельцу узнать chat id для ADMIN_CHAT_ID (работает и в группе)
bot.command('myid', (ctx) => ctx.reply(`Chat id этого чата: <code>${ctx.chat.id}</code>`, { parse_mode: 'HTML' }));

bot.command('help', async (ctx) => {
  let text =
    '<b>Как пользоваться ботом</b>\n\n' +
    `${MENU.book} — выбрать услугу, дату и время\n` +
    `${MENU.my} — посмотреть или отменить свои записи\n` +
    `${MENU.prices} — список услуг\n` +
    `${MENU.contacts} — адрес и телефон\n\n` +
    '/start — главное меню';
  if (isAdmin(ctx)) {
    text +=
      '\n\n<b>Для владельца</b>\n' +
      '/today — записи на сегодня\n' +
      '/export — все записи файлом Excel\n' +
      'Новые заявки приходят в этот чат с кнопками «Подтвердить» и «Отклонить».';
  }
  await ctx.reply(text, { parse_mode: 'HTML', reply_markup: menuFor(ctx) });
});

bot.hears(MENU.prices, async (ctx) => {
  const lines = config.services.map((s) => `• ${esc(s.name)} — <b>${rub(s.price)}</b> (${s.minutes} мин)`);
  await ctx.reply(`<b>Услуги и цены</b>\n\n${lines.join('\n')}`, { parse_mode: 'HTML' });
});

bot.hears(MENU.contacts, async (ctx) => {
  await ctx.reply(
    `<b>${esc(config.businessName)}</b>\n📍 ${esc(config.address)}\n📞 ${esc(config.phone)}\n🕙 ${esc(config.workingHoursText)}`,
    { parse_mode: 'HTML' }
  );
});

bot.hears(MENU.book, async (ctx) => {
  resetFlow(ctx);
  await ctx.reply('Выберите услугу:', { reply_markup: servicesKeyboard() });
});

bot.hears(MENU.my, async (ctx) => {
  const list = upcomingForClient(ctx.chat.id, dateKey(new Date()));
  if (list.length === 0) {
    await ctx.reply('У вас нет предстоящих записей.', { reply_markup: menuFor(ctx) });
    return;
  }
  for (const b of list) {
    const status = b.status === 'confirmed' ? '✅ подтверждена' : '⏳ ожидает подтверждения';
    await ctx.reply(`${bookingText(b)}\nСтатус: ${status}`, {
      parse_mode: 'HTML',
      reply_markup: new InlineKeyboard().text('Отменить запись', `cl:cancel:${b.id}`),
    });
  }
});

// ---------------------------------------------------------------------------
// Шаги записи: услуга → дата → время → имя → телефон → подтверждение
// ---------------------------------------------------------------------------

bot.callbackQuery(/^svc:(.+)$/, async (ctx) => {
  const service = findService(ctx.match[1]);
  await ctx.answerCallbackQuery();
  if (!service) return;
  ctx.session.draft = { serviceId: service.id };
  if (availableDates().length === 0) {
    await ctx.editMessageText('К сожалению, на ближайшие дни свободного времени нет. Позвоните нам: ' + config.phone);
    return;
  }
  await ctx.editMessageText(`Услуга: <b>${esc(service.name)}</b>\n\nВыберите дату:`, {
    parse_mode: 'HTML',
    reply_markup: datesKeyboard(),
  });
});

bot.callbackQuery(/^date:(\d{4}-\d{2}-\d{2})$/, async (ctx) => {
  await ctx.answerCallbackQuery();
  const service = findService(ctx.session.draft.serviceId);
  if (!service) return ctx.editMessageText('Начните запись заново: нажмите «Записаться».');
  const date = ctx.match[1];
  ctx.session.draft.date = date;
  await ctx.editMessageText(
    `Услуга: <b>${esc(service.name)}</b>\nДата: <b>${formatDate(date)}</b>\n\nВыберите время:`,
    { parse_mode: 'HTML', reply_markup: timesKeyboard(date) }
  );
});

bot.callbackQuery(/^time:(\d{2}:\d{2})$/, async (ctx) => {
  await ctx.answerCallbackQuery();
  const { serviceId, date } = ctx.session.draft;
  const service = findService(serviceId);
  if (!service || !date) return ctx.editMessageText('Начните запись заново: нажмите «Записаться».');

  const time = ctx.match[1];
  // Пока клиент выбирал, время мог занять кто-то другой
  if (!freeSlots(date).includes(time)) {
    await ctx.editMessageText('Это время только что заняли. Выберите другое:', { reply_markup: timesKeyboard(date) });
    return;
  }
  ctx.session.draft.time = time;
  ctx.session.step = 'name';
  await ctx.editMessageText(
    `Услуга: <b>${esc(service.name)}</b>\nДата и время: <b>${formatDate(date)}, ${time}</b>`,
    { parse_mode: 'HTML' }
  );
  await ctx.reply('Как вас зовут?', { reply_markup: { remove_keyboard: true } });
});

bot.callbackQuery('back:services', async (ctx) => {
  await ctx.answerCallbackQuery();
  ctx.session.draft = {};
  await ctx.editMessageText('Выберите услугу:', { reply_markup: servicesKeyboard() });
});

bot.callbackQuery('back:dates', async (ctx) => {
  await ctx.answerCallbackQuery();
  const service = findService(ctx.session.draft.serviceId);
  if (!service) return ctx.editMessageText('Начните запись заново: нажмите «Записаться».');
  await ctx.editMessageText(`Услуга: <b>${esc(service.name)}</b>\n\nВыберите дату:`, {
    parse_mode: 'HTML',
    reply_markup: datesKeyboard(),
  });
});

const phoneKeyboard = new Keyboard().requestContact('📱 Отправить мой номер').resized().oneTime();

/** Нормализует телефон: оставляет цифры и "+" в начале, проверяет длину */
function normalizePhone(raw: string): string | null {
  const cleaned = raw.trim().replace(/[^\d+]/g, '');
  const digits = cleaned.replace(/\D/g, '');
  if (digits.length < 10 || digits.length > 15) return null;
  return cleaned.startsWith('+') ? `+${digits}` : digits;
}

async function askConfirmation(ctx: MyContext) {
  const { serviceId, date, time, name, phone } = ctx.session.draft;
  const service = findService(serviceId);
  if (!service || !date || !time || !name || !phone) return;
  ctx.session.step = 'idle';
  await ctx.reply('Спасибо!', { reply_markup: menuFor(ctx) });
  await ctx.reply(
    `Проверьте запись:\n\n<b>${esc(service.name)}</b> — ${rub(service.price)}\n` +
      `📅 ${formatDate(date)}, ${time}\n👤 ${esc(name)}, ${esc(phone)}`,
    {
      parse_mode: 'HTML',
      reply_markup: new InlineKeyboard().text('✅ Записаться', 'confirm').text('✖ Отмена', 'abort'),
    }
  );
}

bot.on('message:contact', async (ctx) => {
  if (ctx.session.step !== 'phone') return;
  const phone = normalizePhone(ctx.message.contact.phone_number);
  if (!phone) return ctx.reply('Не удалось прочитать номер. Введите его вручную, например +7 900 123-45-67');
  ctx.session.draft.phone = phone;
  await askConfirmation(ctx);
});

bot.on('message:text', async (ctx, next) => {
  const text = ctx.message.text;

  if (ctx.session.step === 'name') {
    const name = text.trim();
    if (name.length < 2 || name.length > 50) return ctx.reply('Введите имя от 2 до 50 символов.');
    ctx.session.draft.name = name;
    ctx.session.step = 'phone';
    await ctx.reply('Оставьте номер телефона для связи: нажмите кнопку ниже или введите вручную.', {
      reply_markup: phoneKeyboard,
    });
    return;
  }

  if (ctx.session.step === 'phone') {
    const phone = normalizePhone(text);
    if (!phone) return ctx.reply('Номер выглядит неверно. Пример: +7 900 123-45-67');
    ctx.session.draft.phone = phone;
    await askConfirmation(ctx);
    return;
  }

  return next();
});

bot.callbackQuery('abort', async (ctx) => {
  await ctx.answerCallbackQuery();
  resetFlow(ctx);
  await ctx.editMessageText('Запись отменена. Вы можете начать заново в любой момент.');
});

bot.callbackQuery('confirm', async (ctx) => {
  await ctx.answerCallbackQuery();
  const { serviceId, date, time, name, phone } = ctx.session.draft;
  const service = findService(serviceId);
  if (!service || !date || !time || !name || !phone) {
    return ctx.editMessageText('Черновик записи устарел. Нажмите «Записаться», чтобы начать заново.');
  }
  // Последняя проверка: время всё ещё свободно?
  if (!freeSlots(date).includes(time)) {
    resetFlow(ctx);
    return ctx.editMessageText('К сожалению, это время уже заняли. Нажмите «Записаться» и выберите другое.');
  }

  const booking = addBooking({
    clientChatId: ctx.chat!.id,
    username: ctx.from.username ?? null,
    name,
    phone,
    serviceName: service.name,
    price: service.price,
    date,
    time,
  });
  resetFlow(ctx);

  await ctx.editMessageText(
    `✅ Вы записаны!\n\n${bookingText(booking)}\n\nМы подтвердим запись в ближайшее время. ` +
      'Посмотреть или отменить её можно в разделе «Мои записи».',
    { parse_mode: 'HTML' }
  );

  if (ADMIN_CHAT_ID) {
    const who = ctx.from.username ? ` (@${esc(ctx.from.username)})` : '';
    await ctx.api.sendMessage(ADMIN_CHAT_ID, `🆕 <b>Новая запись №${booking.id}</b>${who}\n\n${bookingText(booking)}`, {
      parse_mode: 'HTML',
      reply_markup: new InlineKeyboard()
        .text('✅ Подтвердить', `adm:ok:${booking.id}`)
        .text('❌ Отклонить', `adm:no:${booking.id}`),
    });
  }
});

// ---------------------------------------------------------------------------
// Отмена клиентом
// ---------------------------------------------------------------------------

bot.callbackQuery(/^cl:cancel:(\d+)$/, async (ctx) => {
  await ctx.answerCallbackQuery();
  const booking = getBooking(Number(ctx.match[1]));
  if (!booking || booking.clientChatId !== ctx.chat?.id || booking.status === 'cancelled') {
    return ctx.editMessageText('Эта запись уже неактуальна.');
  }
  setStatus(booking.id, 'cancelled');
  await ctx.editMessageText(`Запись отменена:\n\n${bookingText(booking)}`, { parse_mode: 'HTML' });
  if (ADMIN_CHAT_ID) {
    await ctx.api.sendMessage(ADMIN_CHAT_ID, `🚫 Клиент отменил запись №${booking.id}\n\n${bookingText(booking)}`, {
      parse_mode: 'HTML',
    });
  }
});

// ---------------------------------------------------------------------------
// Для владельца (только в чате ADMIN_CHAT_ID)
// ---------------------------------------------------------------------------

bot.callbackQuery(/^adm:(ok|no):(\d+)$/, async (ctx) => {
  if (!isAdmin(ctx)) return ctx.answerCallbackQuery('Недоступно');
  await ctx.answerCallbackQuery();
  const approve = ctx.match[1] === 'ok';
  const booking = getBooking(Number(ctx.match[2]));
  if (!booking) return;
  if (booking.status === 'cancelled') {
    return ctx.editMessageText(`Запись №${booking.id} уже отменена клиентом.\n\n${bookingText(booking)}`, {
      parse_mode: 'HTML',
    });
  }

  setStatus(booking.id, approve ? 'confirmed' : 'cancelled');
  await ctx.editMessageText(
    `${approve ? '✅ Подтверждена' : '❌ Отклонена'} запись №${booking.id}\n\n${bookingText(booking)}`,
    { parse_mode: 'HTML' }
  );

  const clientText = approve
    ? `✅ Ваша запись подтверждена!\n\n${bookingText(booking)}\n\nЖдём вас по адресу: ${esc(config.address)}`
    : `К сожалению, мы не можем принять вас в это время.\n\n${bookingText(booking)}\n\n` +
      'Пожалуйста, выберите другое время через «Записаться» или позвоните нам: ' + esc(config.phone);
  await ctx.api.sendMessage(booking.clientChatId, clientText, { parse_mode: 'HTML' }).catch(() => {
    // Клиент мог заблокировать бота — это не повод падать
  });
});

async function sendToday(ctx: MyContext) {
  if (!isAdmin(ctx)) return;
  const list = forDate(dateKey(new Date()));
  if (list.length === 0) {
    await ctx.reply('На сегодня записей нет.');
    return;
  }
  const lines = list.map((b) => `${b.status === 'confirmed' ? '✅' : '⏳'} ${bookingText(b)}`);
  await ctx.reply(`<b>Записи на сегодня</b>\n\n${lines.join('\n\n')}`, { parse_mode: 'HTML' });
}

async function sendExport(ctx: MyContext) {
  if (!isAdmin(ctx)) return;
  await exportToExcel();
  await ctx.replyWithDocument(new InputFile(EXCEL_PATH, 'bookings.xlsx'), { caption: 'Все записи в Excel' });
}

// Одно и то же доступно и командой, и кнопкой меню
bot.command('today', sendToday);
bot.hears(MENU.today, sendToday);
bot.command('export', sendExport);
bot.hears(MENU.export, sendExport);

// ---------------------------------------------------------------------------
// Запуск
// ---------------------------------------------------------------------------

bot.catch((err) => {
  const e = err.error;
  if (e instanceof GrammyError) console.error('Telegram API error:', e.description);
  else if (e instanceof HttpError) console.error('Network error:', e);
  else console.error('Unexpected error:', e);
});

// Подсказки при вводе "/": клиенты видят только общие команды,
// а в чате владельца появляются ещё и команды для него
const clientCommands = [
  { command: 'start', description: 'Главное меню' },
  { command: 'help', description: 'Как пользоваться ботом' },
];
await bot.api.setMyCommands(clientCommands);
if (ADMIN_CHAT_ID) {
  await bot.api.setMyCommands(
    [
      ...clientCommands,
      { command: 'today', description: 'Записи на сегодня' },
      { command: 'export', description: 'Все записи в Excel' },
    ],
    { scope: { type: 'chat', chat_id: ADMIN_CHAT_ID } }
  );
}
console.log(`Bot for "${config.businessName}" is running. Press Ctrl+C to stop.`);
bot.start();
