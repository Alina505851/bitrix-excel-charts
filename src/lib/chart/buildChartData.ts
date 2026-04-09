import type {
  ChartConfig,
  ChartFilter,
  ColumnMeta,
  DateGranularity,
} from "@/lib/types";
import {
  tryParseNumber,
  tryParseDate,
  groupLabel,
  formatDateYmdLocal,
} from "@/lib/chart/coerce";

const RU_MONTHS_SHORT = [
  "янв.",
  "февр.",
  "мар.",
  "апр.",
  "мая",
  "июн.",
  "июл.",
  "авг.",
  "сент.",
  "окт.",
  "нояб.",
  "дек.",
] as const;

export function metaByKeyMap(columns: ColumnMeta[]): Map<string, ColumnMeta> {
  return new Map(columns.map((c) => [c.key, c]));
}

export function rowValueAsYmd(raw: unknown): string | null {
  if (raw instanceof Date && !Number.isNaN(raw.getTime())) {
    return formatDateYmdLocal(raw);
  }
  const d = tryParseDate(raw);
  return d ? formatDateYmdLocal(d) : null;
}

function ymdToCalendarQuarterKey(ymd: string): string | null {
  const [yStr, mStr] = ymd.split("-");
  const y = Number(yStr);
  const mo = Number(mStr);
  if (!Number.isFinite(y) || !Number.isFinite(mo) || mo < 1 || mo > 12) {
    return null;
  }
  const q = Math.ceil(mo / 3);
  return `${y}-Q${q}`;
}

/** Уникальные месяцы YYYY-MM по колонке (отсортированы). */
export function distinctCalendarMonthsFromColumn(
  rows: Record<string, unknown>[],
  columnKey: string,
  maxItems = 144,
): string[] {
  const s = new Set<string>();
  for (const row of rows) {
    const ymd = rowValueAsYmd(row[columnKey]);
    if (ymd) {
      s.add(ymd.slice(0, 7));
    }
  }
  const sorted = Array.from(s).sort();
  return sorted.length > maxItems ? sorted.slice(-maxItems) : sorted;
}

/** Уникальные кварталы YYYY-Qn по колонке (отсортированы по году и кварталу). */
export function distinctCalendarQuartersFromColumn(
  rows: Record<string, unknown>[],
  columnKey: string,
  maxItems = 48,
): string[] {
  const s = new Set<string>();
  for (const row of rows) {
    const ymd = rowValueAsYmd(row[columnKey]);
    if (!ymd) {
      continue;
    }
    const qk = ymdToCalendarQuarterKey(ymd);
    if (qk) {
      s.add(qk);
    }
  }
  const sorted = Array.from(s).sort((a, b) => {
    const ma = /^(\d+)-Q([1-4])$/.exec(a);
    const mb = /^(\d+)-Q([1-4])$/.exec(b);
    if (!ma || !mb) {
      return a.localeCompare(b);
    }
    const ya = Number(ma[1]);
    const yb = Number(mb[1]);
    if (ya !== yb) {
      return ya - yb;
    }
    return Number(ma[2]) - Number(mb[2]);
  });
  return sorted.length > maxItems ? sorted.slice(-maxItems) : sorted;
}

export function formatCalendarMonthRu(ym: string): string {
  const m = /^(\d{4})-(\d{2})$/.exec(ym);
  if (!m) {
    return ym;
  }
  const year = Number(m[1]);
  const month = Number(m[2]);
  if (month < 1 || month > 12) {
    return ym;
  }
  const label = RU_MONTHS_SHORT[month - 1] ?? String(month);
  return `${label} ${String(year)}`;
}

export function formatCalendarQuarterRu(yq: string): string {
  const m = /^(\d{4})-Q([1-4])$/.exec(yq);
  if (!m) {
    return yq;
  }
  return `${m[2]} кв. ${m[1]}`;
}

export function applyFilters(
  rows: Record<string, unknown>[],
  filters: ChartFilter[],
): Record<string, unknown>[] {
  if (filters.length === 0) {
    return rows;
  }
  return rows.filter((row) => {
    for (const f of filters) {
      if (!f.columnKey) {
        continue;
      }
      const raw = row[f.columnKey];
      if (f.values && f.values.length > 0) {
        const s = groupLabel(raw);
        if (!f.values.includes(s)) {
          return false;
        }
      }
      if (f.min !== undefined || f.max !== undefined) {
        const n = tryParseNumber(raw);
        if (n === null) {
          return false;
        }
        if (f.min !== undefined && n < f.min) {
          return false;
        }
        if (f.max !== undefined && n > f.max) {
          return false;
        }
      }
      if (f.dateFrom !== undefined || f.dateTo !== undefined) {
        const ymd = rowValueAsYmd(raw);
        if (ymd === null) {
          return false;
        }
        if (f.dateFrom !== undefined && ymd < f.dateFrom) {
          return false;
        }
        if (f.dateTo !== undefined && ymd > f.dateTo) {
          return false;
        }
      }
      if (f.calendarMonths && f.calendarMonths.length > 0) {
        const ymd = rowValueAsYmd(raw);
        if (ymd === null) {
          return false;
        }
        const ym = ymd.slice(0, 7);
        if (!f.calendarMonths.includes(ym)) {
          return false;
        }
      }
      if (f.calendarQuarters && f.calendarQuarters.length > 0) {
        const ymd = rowValueAsYmd(raw);
        if (ymd === null) {
          return false;
        }
        const qk = ymdToCalendarQuarterKey(ymd);
        if (qk === null || !f.calendarQuarters.includes(qk)) {
          return false;
        }
      }
    }
    return true;
  });
}

/** Строки после всех фильтров, кроме условия с индексом `skipIndex` (для подбора значений в UI). */
export function applyFiltersExcept(
  rows: Record<string, unknown>[],
  filters: ChartFilter[],
  skipIndex: number,
): Record<string, unknown>[] {
  const subset = filters.filter((_, i) => i !== skipIndex);
  return applyFilters(rows, subset);
}

function formatMonthRu(year: number, month1to12: number): string {
  const m = RU_MONTHS_SHORT[month1to12 - 1] ?? String(month1to12);
  return `${m} ${year}`;
}

type GroupAccum = {
  count: number;
  sums: Record<string, number>;
  counts: Record<string, number>;
  /** Непустые значения в колонке Y (для подсчёта ID и т.п., без суммирования). */
  filled: Record<string, number>;
  /** Подпись категории на оси X (для месяца — «сент. 2025»). */
  xLabel: string;
};

function yCellNonempty(raw: unknown): boolean {
  if (raw === null || raw === undefined) {
    return false;
  }
  if (typeof raw === "number") {
    return Number.isFinite(raw);
  }
  if (typeof raw === "string") {
    const t = raw.trim();
    return t !== "" && t !== "-";
  }
  if (raw instanceof Date) {
    return !Number.isNaN(raw.getTime());
  }
  return true;
}

export type ChartRow = Record<string, string | number>;

export function buildAggregatedChartRows(
  rows: Record<string, unknown>[],
  config: ChartConfig,
  columns: ColumnMeta[],
): { data: ChartRow[]; warnings: string[] } {
  const warnings: string[] = [];
  const { xKey, yKeys, aggregation } = config;
  if (!xKey || yKeys.length === 0) {
    return { data: [], warnings };
  }

  const colMap = metaByKeyMap(columns);
  const xMeta = colMap.get(xKey);
  if (!xMeta) {
    warnings.push("Ось X не найдена в данных.");
    return { data: [], warnings };
  }

  const filtered = applyFilters(rows, config.filters);
  if (filtered.length === 0) {
    warnings.push("После фильтров не осталось строк.");
    return { data: [], warnings };
  }

  const groups = new Map<string, GroupAccum>();
  const granularity = config.dateGranularity ?? "day";

  for (const row of filtered) {
    const gx = normalizeXValue(
      row[xKey],
      xMeta.inferredType,
      granularity,
    );
    const key = gx.key;
    let g = groups.get(key);
    if (!g) {
      g = { count: 0, sums: {}, counts: {}, filled: {}, xLabel: gx.label };
      groups.set(key, g);
    }
    g.count += 1;
    for (const yk of yKeys) {
      if (!g.sums[yk]) {
        g.sums[yk] = 0;
        g.counts[yk] = 0;
        g.filled[yk] = 0;
      }
      const rawY = row[yk];
      const num = tryParseNumber(rawY);
      if (num !== null) {
        g.sums[yk] += num;
        g.counts[yk] += 1;
      }
      if (yCellNonempty(rawY)) {
        g.filled[yk] += 1;
      }
    }
  }

  const sortKeys = Array.from(groups.keys()).sort((a, b) =>
    a.localeCompare(b, "ru", { numeric: true }),
  );

  const data: ChartRow[] = sortKeys.map((k) => {
    const g = groups.get(k)!;
    const point: ChartRow = { name: g.xLabel };
    for (const yk of yKeys) {
      const sum = g.sums[yk] ?? 0;
      const c = g.counts[yk] ?? 0;
      let v = 0;
      if (aggregation === "sum") {
        v = sum;
      } else if (aggregation === "avg") {
        v = c > 0 ? sum / c : 0;
      } else if (aggregation === "count_nonempty") {
        v = g.filled[yk] ?? 0;
      } else {
        v = g.count;
      }
      point[yk] = v;
    }
    return point;
  });

  if (config.cumulative && yKeys.length > 0 && config.chartType !== "pie") {
    const runs: Record<string, number> = Object.fromEntries(
      yKeys.map((yk) => [yk, 0]),
    );
    for (const point of data) {
      for (const yk of yKeys) {
        const add = typeof point[yk] === "number" ? point[yk] : 0;
        runs[yk] += add;
        point[yk] = runs[yk];
      }
    }
  }

  return { data, warnings };
}

function normalizeXValue(
  raw: unknown,
  inferred: ColumnMeta["inferredType"],
  dateGranularity: DateGranularity,
): { key: string; label: string } {
  const parsed =
    inferred === "date"
      ? raw instanceof Date
        ? raw
        : tryParseDate(raw)
      : tryParseDate(raw);

  if (parsed && !Number.isNaN(parsed.getTime())) {
    const y = parsed.getFullYear();
    const mo = parsed.getMonth() + 1;
    if (dateGranularity === "month") {
      const key = `${y}-${String(mo).padStart(2, "0")}`;
      return { key, label: formatMonthRu(y, mo) };
    }
    if (dateGranularity === "quarter") {
      const ymd = formatDateYmdLocal(parsed);
      const qk = ymdToCalendarQuarterKey(ymd);
      if (qk) {
        return { key: qk, label: formatCalendarQuarterRu(qk) };
      }
    }
    if (dateGranularity === "year") {
      const ys = String(y);
      return { key: ys, label: ys };
    }
    if (inferred === "date" || dateGranularity === "day") {
      const ymd = formatDateYmdLocal(parsed);
      return { key: ymd, label: ymd };
    }
  }

  if (inferred === "date") {
    const d = raw instanceof Date ? raw : tryParseDate(raw);
    if (d) {
      const ymd = formatDateYmdLocal(d);
      return { key: ymd, label: ymd };
    }
  }

  const lab = groupLabel(raw);
  return { key: lab, label: lab };
}

export function distinctStringValues(
  rows: Record<string, unknown>[],
  columnKey: string,
  limit = 200,
): string[] {
  const set = new Set<string>();
  for (const row of rows) {
    set.add(groupLabel(row[columnKey]));
    if (set.size >= limit) {
      break;
    }
  }
  return Array.from(set).sort((a, b) => a.localeCompare(b, "ru"));
}
