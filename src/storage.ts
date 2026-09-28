import fs from 'node:fs';
import path from 'node:path';
import ExcelJS from 'exceljs';

/**
 * Хранилище записей. Для небольшого бизнеса хватает JSON-файла:
 * никакой базы данных ставить не нужно. После каждого изменения
 * бот пересобирает Excel-файл, который владелец получает командой /export.
 */

export type BookingStatus = 'new' | 'confirmed' | 'cancelled';

export interface Booking {
  id: number;
  createdAt: string;
  clientChatId: number;
  username: string | null;
  name: string;
  phone: string;
  serviceName: string;
  price: number;
  date: string; // YYYY-MM-DD
  time: string; // HH:MM
  status: BookingStatus;
}

const DATA_DIR = path.resolve('data');
const JSON_PATH = path.join(DATA_DIR, 'bookings.json');
export const EXCEL_PATH = path.join(DATA_DIR, 'bookings.xlsx');

const STATUS_TEXT: Record<BookingStatus, string> = {
  new: 'Новая',
  confirmed: 'Подтверждена',
  cancelled: 'Отменена',
};

let bookings: Booking[] = load();

function load(): Booking[] {
  if (!fs.existsSync(JSON_PATH)) return [];
  return JSON.parse(fs.readFileSync(JSON_PATH, 'utf8')) as Booking[];
}

function persist(): void {
  fs.mkdirSync(DATA_DIR, { recursive: true });
  // Пишем во временный файл и переименовываем — так файл не испортится при сбое посреди записи
  const tmp = `${JSON_PATH}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(bookings, null, 2));
  fs.renameSync(tmp, JSON_PATH);
  exportToExcel().catch((err) => console.error('Excel export failed:', err));
}

export function isSlotTaken(date: string, time: string): boolean {
  return bookings.some((b) => b.date === date && b.time === time && b.status !== 'cancelled');
}

export function addBooking(data: Omit<Booking, 'id' | 'createdAt' | 'status'>): Booking {
  const booking: Booking = {
    ...data,
    id: bookings.reduce((max, b) => Math.max(max, b.id), 0) + 1,
    createdAt: new Date().toISOString(),
    status: 'new',
  };
  bookings.push(booking);
  persist();
  return booking;
}

export function getBooking(id: number): Booking | undefined {
  return bookings.find((b) => b.id === id);
}

export function setStatus(id: number, status: BookingStatus): Booking | undefined {
  const booking = getBooking(id);
  if (!booking) return undefined;
  booking.status = status;
  persist();
  return booking;
}

function sortByDateTime(a: Booking, b: Booking): number {
  return `${a.date} ${a.time}`.localeCompare(`${b.date} ${b.time}`);
}

/** Предстоящие активные записи клиента */
export function upcomingForClient(chatId: number, todayKey: string): Booking[] {
  return bookings
    .filter((b) => b.clientChatId === chatId && b.status !== 'cancelled' && b.date >= todayKey)
    .sort(sortByDateTime);
}

/** Активные записи на дату — для команды /today */
export function forDate(dateKey: string): Booking[] {
  return bookings.filter((b) => b.date === dateKey && b.status !== 'cancelled').sort(sortByDateTime);
}

export async function exportToExcel(): Promise<void> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Записи');
  ws.columns = [
    { header: '№', key: 'id', width: 6 },
    { header: 'Дата', key: 'date', width: 12 },
    { header: 'Время', key: 'time', width: 8 },
    { header: 'Услуга', key: 'serviceName', width: 24 },
    { header: 'Цена, ₽', key: 'price', width: 10 },
    { header: 'Имя', key: 'name', width: 20 },
    { header: 'Телефон', key: 'phone', width: 18 },
    { header: 'Telegram', key: 'username', width: 18 },
    { header: 'Статус', key: 'status', width: 14 },
    { header: 'Создана', key: 'createdAt', width: 20 },
  ];
  ws.getRow(1).font = { bold: true };
  ws.views = [{ state: 'frozen', ySplit: 1 }];

  for (const b of [...bookings].sort(sortByDateTime)) {
    ws.addRow({
      ...b,
      username: b.username ? `@${b.username}` : '',
      status: STATUS_TEXT[b.status],
      createdAt: new Date(b.createdAt).toLocaleString('ru-RU'),
    });
  }

  fs.mkdirSync(DATA_DIR, { recursive: true });
  await wb.xlsx.writeFile(EXCEL_PATH);
}
