import { config } from './config.js';
import { isSlotTaken } from './storage.js';

const WEEKDAYS = ['Вс', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'];

function pad(n: number): string {
  return String(n).padStart(2, '0');
}

/** Дата в формате YYYY-MM-DD по местному времени сервера */
export function dateKey(d: Date): string {
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

/** "2026-09-29" → "Вт 29.09" */
export function formatDate(key: string): string {
  const [y, m, d] = key.split('-').map(Number);
  const date = new Date(y, m - 1, d);
  return `${WEEKDAYS[date.getDay()]} ${pad(d)}.${pad(m)}`;
}

function toMinutes(hhmm: string): number {
  const [h, m] = hhmm.split(':').map(Number);
  return h * 60 + m;
}

function fromMinutes(total: number): string {
  return `${pad(Math.floor(total / 60))}:${pad(total % 60)}`;
}

/** Все слоты дня по сетке из настроек */
function allSlots(): string[] {
  const slots: string[] = [];
  for (let t = toMinutes(config.firstSlot); t <= toMinutes(config.lastSlot); t += config.slotStepMinutes) {
    slots.push(fromMinutes(t));
  }
  return slots;
}

/** Свободные слоты на дату: без занятых и без тех, что уже слишком близко по времени */
export function freeSlots(key: string): string[] {
  const now = new Date();
  const isToday = key === dateKey(now);
  const nowMinutes = now.getHours() * 60 + now.getMinutes();

  return allSlots().filter((slot) => {
    if (isSlotTaken(key, slot)) return false;
    if (isToday && toMinutes(slot) < nowMinutes + config.minMinutesBeforeStart) return false;
    return true;
  });
}

/** Даты в пределах daysAhead дней: рабочие и с хотя бы одним свободным слотом */
export function availableDates(): string[] {
  const result: string[] = [];
  const today = new Date();
  for (let i = 0; i < config.daysAhead; i++) {
    const d = new Date(today.getFullYear(), today.getMonth(), today.getDate() + i);
    if (config.daysOff.includes(d.getDay())) continue;
    const key = dateKey(d);
    if (freeSlots(key).length > 0) result.push(key);
  }
  return result;
}
