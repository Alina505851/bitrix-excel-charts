/**
 * Средний цикл сделки в CRM: разница календарных дат по локальным полуночам.
 * Один и тот же календарный день создания и закрытия → 0 (сделка закрыта в тот же день).
 */
export function calendarDaysBetweenLocalMidnight(from: Date, to: Date): number {
  const a = new Date(from);
  a.setHours(0, 0, 0, 0);
  const b = new Date(to);
  b.setHours(0, 0, 0, 0);
  return Math.round((b.getTime() - a.getTime()) / 86_400_000);
}
