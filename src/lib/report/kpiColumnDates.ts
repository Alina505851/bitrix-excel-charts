import type { ColumnMeta } from "@/lib/types";
import { formatDateDdMmYyyyRuLocal, tryParseDate } from "@/lib/chart/coerce";

/**
 * Одна дата из ячейки колонки: для каждой строки не больше одного значения.
 * Нет «всей колонкой выкинуть числа»: при смеси строк и серий Excel учитываются оба типа.
 * Для колонки с inferredType `string` число в ячейке не считаем датой (часто ID в диапазоне serial).
 */
export function collectParsedDatesForKpiColumn(
  rows: Record<string, unknown>[],
  columnKey: string,
  columnMeta?: ColumnMeta | null,
): Date[] {
  const skipNumberAsExcelSerial =
    columnMeta?.inferredType === "string";
  const out: Date[] = [];
  for (const row of rows) {
    const raw = row[columnKey];
    if (raw instanceof Date && !Number.isNaN(raw.getTime())) {
      // Только локальный календарный день (как в дд.мм в файле), без сдвига из‑за времени суток.
      out.push(
        new Date(raw.getFullYear(), raw.getMonth(), raw.getDate()),
      );
      continue;
    }
    if (typeof raw === "string") {
      const trimmed = raw.trim();
      if (trimmed === "") {
        continue;
      }
      const d = tryParseDate(raw);
      if (d && !Number.isNaN(d.getTime())) {
        out.push(d);
      }
      continue;
    }
    if (
      typeof raw === "number" &&
      Number.isFinite(raw) &&
      !skipNumberAsExcelSerial
    ) {
      const d = tryParseDate(raw);
      if (d && !Number.isNaN(d.getTime())) {
        out.push(d);
      }
    }
  }
  return out;
}

/** Ключ для сравнения только календарных дней в локальной зоне (без времени суток). */
function localCalendarDayKey(d: Date): number {
  return (
    d.getFullYear() * 10_000 +
    (d.getMonth() + 1) * 100 +
    d.getDate()
  );
}

/**
 * Самая ранняя и поздняя календарная дата (локально), без сдвига из‑за времени внутри дня.
 */
export function minMaxFromDateList(
  dates: Date[],
): { min: Date; max: Date } | null {
  if (dates.length === 0) {
    return null;
  }
  let minKey = localCalendarDayKey(dates[0]!);
  let maxKey = minKey;
  let minSrc = dates[0]!;
  let maxSrc = dates[0]!;
  for (const d of dates) {
    const k = localCalendarDayKey(d);
    if (k < minKey) {
      minKey = k;
      minSrc = d;
    }
    if (k > maxKey) {
      maxKey = k;
      maxSrc = d;
    }
  }
  return {
    min: new Date(
      minSrc.getFullYear(),
      minSrc.getMonth(),
      minSrc.getDate(),
    ),
    max: new Date(
      maxSrc.getFullYear(),
      maxSrc.getMonth(),
      maxSrc.getDate(),
    ),
  };
}

/** Подпись для KPI: дд.мм.гггг (как в выгрузке Excel/Битрикс). */
export function formatKpiDateRu(d: Date): string {
  return formatDateDdMmYyyyRuLocal(d);
}

/** Первая/последняя дата в колонке — тот же порядок дд.мм.гггг, что и в файле. */
export function formatKpiRecordDateRu(d: Date): string {
  return formatDateDdMmYyyyRuLocal(d);
}
