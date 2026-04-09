const RU_DATE = /^(\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/;
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
  return Number.isNaN(dt.getTime()) ? null : dt;
}

export function formatDateYmdLocal(dt: Date): string {
  const y = dt.getFullYear();
  const m = String(dt.getMonth() + 1).padStart(2, "0");
  const d = String(dt.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
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
    const [, d, m, y, hh = "0", mm = "0", ss = "0"] = ru;
    return buildStrictLocalDate(
      expandYear(y),
      Number(m),
      Number(d),
      Number(hh),
      Number(mm),
      Number(ss),
    );
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
    const dt = new Date(s);
    return Number.isNaN(dt.getTime()) ? null : dt;
  }
  // Не используем Date.parse для неоднозначных строк вроде 09.04.26.
  return null;
}

export function groupLabel(value: unknown): string {
  if (value === null || value === undefined || value === "") {
    return "(пусто)";
  }
  if (value instanceof Date) {
    return formatDateYmdLocal(value);
  }
  return String(value);
}
