/**
 * Числа в таблицах, KPI и экспорте: без группировки разрядов (без лишних
 * разделителей в строке; в ru-RU группировка даёт узкий пробел между разрядами).
 */
export function formatDataInt(n: number): string {
  return new Intl.NumberFormat("ru-RU", {
    useGrouping: false,
    maximumFractionDigits: 0,
    minimumFractionDigits: 0,
  }).format(n);
}

export function formatDataDecimal(
  n: number,
  maximumFractionDigits: number,
  minimumFractionDigits = 0,
): string {
  return new Intl.NumberFormat("ru-RU", {
    useGrouping: false,
    maximumFractionDigits,
    minimumFractionDigits,
  }).format(n);
}

/** Доля в процентах, одна значащая цифра после запятой при необходимости. */
export function formatDataPercent1(n: number): string {
  return `${formatDataDecimal(n, 1, 0)}%`;
}

/** Подписи осей графиков: компактно, без группировки тысяч. */
export function formatAxisCompact(n: number): string {
  return new Intl.NumberFormat("ru-RU", {
    notation: "compact",
    compactDisplay: "short",
    useGrouping: false,
    maximumFractionDigits: 1,
  }).format(n);
}
