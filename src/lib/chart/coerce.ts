const RU_DATE = /^(\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/;
/** дд/мм/гггг или дд-мм-гггг — тот же порядок, что и в RU_DATE с точками */
const EU_SLASH_OR_DASH_DATE =
  /^(\d{1,2})[/-](\d{1,2})[/-](\d{2}|\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/;
const ISO_DATE =
  /^(\d{4})-(\d{2})-(\d{2})(?:[T\s](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/;
const ISO_TIMESTAMP_TZ =
  /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:\d{2})$/i;

function expandYear(y: string): number {
  if (y.length === 4) {
    return Number(y);
  }
  const yy = Number(y);
  return yy >= 70 ? 1900 + yy : 2000 + yy;
}

function buildStrictLocalDate(
  year: number,
  month: number,
  day: number,
  hour: number,
  minute: number,
  second: number,
): Date | null {
  if (
    !Number.isInteger(year) ||
    !Number.isInteger(month) ||
    !Number.isInteger(day) ||
    !Number.isInteger(hour) ||
    !Number.isInteger(minute) ||
    !Number.isInteger(second)
  ) {
    return null;
  }
  if (
    month < 1 ||
    month > 12 ||
    day < 1 ||
    day > 31 ||
    hour < 0 ||
    hour > 23 ||
    minute < 0 ||
    minute > 59 ||
    second < 0 ||
    second > 59
  ) {
    return null;
  }
  const dt = new Date(year, month - 1, day, hour, minute, second);
  if (Number.isNaN(dt.getTime())) {
    return null;
  }
  if (
    dt.getFullYear() !== year ||
    dt.getMonth() !== month - 1 ||
    dt.getDate() !== day ||
    dt.getHours() !== hour ||
    dt.getMinutes() !== minute ||
    dt.getSeconds() !== second
  ) {
    return null;
  }
  return dt;
}

function excelSerialToDate(value: number): Date | null {
  if (!Number.isFinite(value)) {
    return null;
  }
  // Ограничиваем диапазон Excel serial, чтобы обычные числовые ID
  // (например 134, 522, 100500) не интерпретировались как даты.
  if (value < 20_000 || value > 80_000) {
    return null;
  }
  // Excel serial date (1900 system): days since 1899-12-30, fractional part is time.
  const ms = Math.round((value - 25569) * 86400 * 1000);
  const dt = new Date(ms);
  if (Number.isNaN(dt.getTime())) {
    return null;
  }
  // Календарный день в локальной зоне (как в ячейке Excel), без «сдвига» времени суток.
  return new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
}

export function formatDateYmdLocal(dt: Date): string {
  const y = dt.getFullYear();
  const m = String(dt.getMonth() + 1).padStart(2, "0");
  const d = String(dt.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

/** дд.мм.гггг по локальным календарным компонентам (как в Excel/Битрикс), без toLocaleDateString. */
export function formatDateDdMmYyyyRuLocal(dt: Date): string {
  const d = String(dt.getDate()).padStart(2, "0");
  const m = String(dt.getMonth() + 1).padStart(2, "0");
  const y = String(dt.getFullYear());
  return `${d}.${m}.${y}`;
}

/**
 * Дата и время для подписей и ячеек: дд.мм.гггг, при ненулевом времени — чч:мм или чч:мм:сс.
 */
export function formatDateTimeDdMmYyyyRuLocal(dt: Date): string {
  const d = String(dt.getDate()).padStart(2, "0");
  const mo = String(dt.getMonth() + 1).padStart(2, "0");
  const y = String(dt.getFullYear());
  const hh = dt.getHours();
  const mm = dt.getMinutes();
  const ss = dt.getSeconds();
  const msec = dt.getMilliseconds();
  const base = `${d}.${mo}.${y}`;
  if (hh === 0 && mm === 0 && ss === 0 && msec === 0) {
    return base;
  }
  const h = String(hh).padStart(2, "0");
  const min = String(mm).padStart(2, "0");
  if (ss === 0 && msec === 0) {
    return `${base} ${h}:${min}`;
  }
  return `${base} ${h}:${min}:${String(ss).padStart(2, "0")}`;
}

export function tryParseNumber(value: unknown): number | null {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  if (typeof value !== "string") {
    return null;
  }
  const normalized = value.replace(/\s/g, "").replace(",", ".");
  if (normalized === "" || normalized === "-") {
    return null;
  }
  const n = Number(normalized);
  return Number.isFinite(n) ? n : null;
}

/**
 * Первое число — день, второе — месяц (дд.мм / дд/мм / дд-мм), как в типичной выгрузке Excel/Битрикс.
 */
function tryParseDmyMatch(m: RegExpMatchArray): Date | null {
  const [, dayStr, monthStr, y, hh = "0", mm = "0", ss = "0"] = m;
  const year = expandYear(y);
  const h = Number(hh);
  const min = Number(mm);
  const sec = Number(ss);
  return buildStrictLocalDate(
    year,
    Number(monthStr),
    Number(dayStr),
    h,
    min,
    sec,
  );
}

/**
 * Ячейка .xlsx с `cellDates: true` даёт `Date`, у которого локальные месяц и день
 * соответствуют порядку мм/дд (US Excel), а не дд.мм как в выгрузке Битрикс24.
 * Собираем календарную дату в интерпретации дд.мм: день и месяц меняются местами.
 */
export function reinterpretExcelDateAsRuDayMonth(d: Date): Date {
  if (Number.isNaN(d.getTime())) {
    return d;
  }
  const y = d.getFullYear();
  return new Date(y, d.getDate() - 1, d.getMonth() + 1);
}

export function tryParseDate(value: unknown): Date | null {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value;
  }
  if (typeof value === "number") {
    return excelSerialToDate(value);
  }
  if (typeof value !== "string") {
    return null;
  }
  const s = value.trim();
  if (!s) {
    return null;
  }
  const ru = s.match(RU_DATE);
  if (ru) {
    return tryParseDmyMatch(ru);
  }
  const eu = s.match(EU_SLASH_OR_DASH_DATE);
  if (eu) {
    return tryParseDmyMatch(eu);
  }
  const iso = s.match(ISO_DATE);
  if (iso) {
    const [, y, m, d, hh = "0", mm = "0", ss = "0"] = iso;
    return buildStrictLocalDate(
      Number(y),
      Number(m),
      Number(d),
      Number(hh),
      Number(mm),
      Number(ss),
    );
  }
  if (ISO_TIMESTAMP_TZ.test(s)) {
    const head = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
    if (head) {
      const tailTime = /T(\d{1,2}):(\d{2})(?::(\d{2}))?/.exec(s);
      const hh = tailTime ? Number(tailTime[1]) : 0;
      const min = tailTime ? Number(tailTime[2]) : 0;
      const sec = tailTime && tailTime[3] ? Number(tailTime[3]) : 0;
      return buildStrictLocalDate(
        Number(head[1]),
        Number(head[2]),
        Number(head[3]),
        hh,
        min,
        sec,
      );
    }
    const dt = new Date(s);
    return Number.isNaN(dt.getTime()) ? null : dt;
  }
  // Не используем Date.parse для неоднозначных строк вроде 09.04.26.
  return null;
}

/** Подпись для пустой ячейки при группировке; совпадает с `groupLabel` для null/"". */
export const EMPTY_GROUP_LABEL = "(пусто)" as const;

export function groupLabel(value: unknown): string {
  if (value === null || value === undefined || value === "") {
    return EMPTY_GROUP_LABEL;
  }
  if (value instanceof Date) {
    return formatDateTimeDdMmYyyyRuLocal(value);
  }
  return String(value);
}
