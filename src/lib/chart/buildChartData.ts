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
  EMPTY_GROUP_LABEL,
  formatDateYmdLocal,
  formatDateDdMmYyyyRuLocal,
} from "@/lib/chart/coerce";
import {
  isDealStageLikeColumn,
  sortDealPipelineStageKeys,
  getEffectiveDealStageOrder,
  sortDealStageKeysDefault,
} from "@/lib/chart/dealStageOrder";

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
      const hasWhitelist = f.values && f.values.length > 0;
      const hasFragmentPick =
        (f.valuesContainAll && f.valuesContainAll.length > 0) ||
        (f.valuesContainAny && f.valuesContainAny.length > 0);
      const hasBlacklist = f.excludeValues && f.excludeValues.length > 0;
      if (hasWhitelist || hasFragmentPick || hasBlacklist) {
        const s = groupLabel(raw);
        const lower = s.toLowerCase();
        if (hasWhitelist && f.values && !f.values.includes(s)) {
          return false;
        }
        if (!hasWhitelist && hasFragmentPick) {
          if (
            f.valuesContainAll?.length &&
            !f.valuesContainAll.every((frag) =>
              lower.includes(frag.toLowerCase()),
            )
          ) {
            return false;
          }
          if (
            f.valuesContainAny?.length &&
            !f.valuesContainAny.some((frag) =>
              lower.includes(frag.toLowerCase()),
            )
          ) {
            return false;
          }
        }
        if (hasBlacklist && f.excludeValues?.includes(s)) {
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

/** Сравнимое время начала периода для ключей оси X из normalizeXValue (день / месяц / квартал / год). */
function timelineKeyToSortTime(s: string): number | null {
  const ymd = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (ymd) {
    return Date.UTC(Number(ymd[1]), Number(ymd[2]) - 1, Number(ymd[3]));
  }
  const ym = /^(\d{4})-(\d{2})$/.exec(s);
  if (ym) {
    return Date.UTC(Number(ym[1]), Number(ym[2]) - 1, 1);
  }
  const yq = /^(\d{4})-Q([1-4])$/i.exec(s);
  if (yq) {
    const mo = (Number(yq[2]) - 1) * 3;
    return Date.UTC(Number(yq[1]), mo, 1);
  }
  const yOnly = /^(\d{4})$/.exec(s);
  if (yOnly) {
    return Date.UTC(Number(yOnly[1]), 0, 1);
  }
  return null;
}

function compareChronologicalXKeys(a: string, b: string): number {
  const ta = timelineKeyToSortTime(a);
  const tb = timelineKeyToSortTime(b);
  if (ta !== null && tb !== null) {
    return ta - tb;
  }
  if (ta !== null) {
    return -1;
  }
  if (tb !== null) {
    return 1;
  }
  return a.localeCompare(b, "ru", { numeric: true });
}

function shouldSortXKeysChronologically(
  xMeta: ColumnMeta,
  rawKeys: string[],
): boolean {
  if (xMeta.inferredType === "date") {
    return true;
  }
  if (rawKeys.length === 0) {
    return false;
  }
  return rawKeys.every((k) => timelineKeyToSortTime(k) !== null);
}

export function buildAggregatedChartRows(
  rows: Record<string, unknown>[],
  config: ChartConfig,
  columns: ColumnMeta[],
  /** Порядок подписей стадий сделки (из UI); только для колонок стадии, см. `isDealStageLikeColumn`. */
  dealStageOrder?: string[] | null,
  /**
   * Если задано (YYYY-MM-DD), строки, у которых ось X распознаётся как календарная дата позже
   * этого дня, не учитываются (графики компаний и сделок).
   */
  maxXCellCalendarDayYmd?: string,
): { data: ChartRow[]; warnings: string[] } {
  const warnings: string[] = [];
  const { xKey, yKeys, aggregation } = config;
  if (yKeys.length === 0) {
    return { data: [], warnings };
  }
  if (
    config.ySourceKeys &&
    config.ySourceKeys.length !== yKeys.length
  ) {
    warnings.push("Длина ySourceKeys не совпадает с yKeys — источники серий проигнорированы.");
  }
  const ySourceKeys =
    config.ySourceKeys && config.ySourceKeys.length === yKeys.length
      ? config.ySourceKeys
      : undefined;
  if (
    config.yAggregations &&
    config.yAggregations.length !== yKeys.length
  ) {
    warnings.push("Длина yAggregations не совпадает с yKeys — для части серий используется общий aggregation.");
  }

  const colMap = metaByKeyMap(columns);

  if (config.literalAvgBars && config.literalAvgBars.length > 0) {
    for (const bar of config.literalAvgBars) {
      if (!yKeys.includes(bar.seriesKey)) {
        warnings.push(
          "literalAvgBars: seriesKey должен совпадать с одним из yKeys.",
        );
        return { data: [], warnings };
      }
    }
    const filtered = applyFilters(rows, config.filters);
    if (filtered.length === 0) {
      warnings.push("После фильтров не осталось строк.");
      return { data: [], warnings };
    }
    const data: ChartRow[] = [];
    for (const lit of config.literalAvgBars) {
      const sub = applyFilters(filtered, lit.filters);
      let sum = 0;
      let c = 0;
      for (const row of sub) {
        const num = tryParseNumber(row[lit.valueKey]);
        if (num !== null && Number.isFinite(num)) {
          sum += num;
          c += 1;
        }
      }
      const avg = c > 0 ? sum / c : 0;
      const point: ChartRow = { name: lit.label };
      for (const yk of yKeys) {
        point[yk] = yk === lit.seriesKey ? avg : 0;
      }
      data.push(point);
    }
    return { data, warnings };
  }

  if (config.literalPieComplement) {
    const lpc = config.literalPieComplement;
    if (config.chartType !== "pie" || !xKey) {
      warnings.push(
        "literalPieComplement: нужны chartType «pie» и непустой xKey.",
      );
      return { data: [], warnings };
    }
    if (lpc.repeatMatchAny.length === 0) {
      warnings.push("literalPieComplement: задайте хотя бы одно условие в repeatMatchAny.");
      return { data: [], warnings };
    }
    const filtered = applyFilters(rows, config.filters);
    const nonEmpty = filtered.filter(
      (row) => groupLabel(row[xKey]) !== EMPTY_GROUP_LABEL,
    );
    const yk = yKeys[0]!;
    let repeatCount = 0;
    let nonRepeatCount = 0;
    for (const row of nonEmpty) {
      const isRepeat = lpc.repeatMatchAny.some(
        (f) => applyFilters([row], [f]).length === 1,
      );
      if (isRepeat) {
        repeatCount += 1;
      } else {
        nonRepeatCount += 1;
      }
    }
    const data: ChartRow[] = [
      { name: lpc.repeatLabel, [yk]: repeatCount },
      { name: lpc.nonRepeatLabel, [yk]: nonRepeatCount },
    ];
    return { data, warnings };
  }

  if (!xKey) {
    return { data: [], warnings };
  }

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
  const rowsForAggregation =
    config.countDistinctByKey && config.countDistinctByKey.trim()
      ? (() => {
          const distinctKey = config.countDistinctByKey!;
          const seen = new Set<string>();
          const out: Record<string, unknown>[] = [];
          for (const row of filtered) {
            const raw = row[distinctKey];
            const id = raw == null ? "" : String(raw).trim();
            if (!id) {
              out.push(row);
              continue;
            }
            if (seen.has(id)) {
              continue;
            }
            seen.add(id);
            out.push(row);
          }
          return out;
        })()
      : filtered;

  const groups = new Map<string, GroupAccum>();
  const granularity = config.dateGranularity ?? "day";

  for (const row of rowsForAggregation) {
    if (maxXCellCalendarDayYmd) {
      const cellYmd = rowValueAsYmd(row[xKey]);
      if (cellYmd && cellYmd > maxXCellCalendarDayYmd) {
        continue;
      }
    }
    const gx = normalizeXValue(
      row[xKey],
      xMeta.inferredType,
      granularity,
    );
    if (gx.key === EMPTY_GROUP_LABEL) {
      continue;
    }
    const key = gx.key;
    let g = groups.get(key);
    if (!g) {
      g = { count: 0, sums: {}, counts: {}, filled: {}, xLabel: gx.label };
      groups.set(key, g);
    }
    g.count += 1;
    for (let i = 0; i < yKeys.length; i++) {
      const yk = yKeys[i]!;
      const sourceKey = ySourceKeys ? ySourceKeys[i]! : yk;
      // Нельзя использовать !g.sums[yk]: при сумме 0 сбрасывались counts/filled (ломались avg и count_nonempty).
      if (g.sums[yk] === undefined) {
        g.sums[yk] = 0;
        g.counts[yk] = 0;
        g.filled[yk] = 0;
      }
      const rawY = row[sourceKey];
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

  const rawKeys = Array.from(groups.keys());
  const effectiveStageOrder = getEffectiveDealStageOrder(dealStageOrder);
  const sortKeys =
    isDealStageLikeColumn(xMeta) && effectiveStageOrder?.length
      ? sortDealPipelineStageKeys(rawKeys, effectiveStageOrder)
      : isDealStageLikeColumn(xMeta)
        ? sortDealStageKeysDefault(rawKeys)
        : shouldSortXKeysChronologically(xMeta, rawKeys)
          ? [...rawKeys].sort(compareChronologicalXKeys)
          : rawKeys.sort((a, b) => a.localeCompare(b, "ru", { numeric: true }));
  const sortKeysClampedByReportDate =
    maxXCellCalendarDayYmd && shouldSortXKeysChronologically(xMeta, rawKeys)
      ? (() => {
          const maxTime = timelineKeyToSortTime(maxXCellCalendarDayYmd);
          if (maxTime === null) {
            return sortKeys;
          }
          return sortKeys.filter((k) => {
            const kt = timelineKeyToSortTime(k);
            return kt !== null && kt <= maxTime;
          });
        })()
      : sortKeys;

  const data: ChartRow[] = sortKeysClampedByReportDate.map((k) => {
    const g = groups.get(k)!;
    const point: ChartRow = { name: g.xLabel };
    for (let i = 0; i < yKeys.length; i++) {
      const yk = yKeys[i]!;
      const agg = config.yAggregations?.[i] ?? aggregation;
      const sum = g.sums[yk] ?? 0;
      const c = g.counts[yk] ?? 0;
      let v = 0;
      if (agg === "sum") {
        v = sum;
      } else if (agg === "avg") {
        v = c > 0 ? sum / c : 0;
      } else if (agg === "count_nonempty") {
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
      for (let i = 0; i < yKeys.length; i++) {
        const yk = yKeys[i]!;
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
      const label = formatDateDdMmYyyyRuLocal(parsed);
      return { key: ymd, label };
    }
  }

  if (inferred === "date") {
    const d = raw instanceof Date ? raw : tryParseDate(raw);
    if (d) {
      const ymd = formatDateYmdLocal(d);
      return { key: ymd, label: formatDateDdMmYyyyRuLocal(d) };
    }
  }

  const lab = groupLabel(raw);
  return { key: lab, label: lab };
}

export function distinctStringValues(
  rows: Record<string, unknown>[],
  columnKey: string,
  limit = 200,
  columnMeta?: ColumnMeta,
  dealStageOrder?: string[] | null,
): string[] {
  const set = new Set<string>();
  for (const row of rows) {
    set.add(groupLabel(row[columnKey]));
    if (set.size >= limit) {
      break;
    }
  }
  const arr = Array.from(set);
  if (columnMeta && isDealStageLikeColumn(columnMeta)) {
    const effectiveStageOrder = getEffectiveDealStageOrder(dealStageOrder);
    return effectiveStageOrder?.length
      ? sortDealPipelineStageKeys(arr, effectiveStageOrder)
      : sortDealStageKeysDefault(arr);
  }
  return arr.sort((a, b) => a.localeCompare(b, "ru"));
}
