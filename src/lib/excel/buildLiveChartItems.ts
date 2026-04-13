import { buildAggregatedChartRows } from "@/lib/chart/buildChartData";
import { tryParseDate } from "@/lib/chart/coerce";
import { seriesLegendLabel } from "@/lib/chart/seriesLegendLabel";
import type { ChartConfig, ChartFilter, TabularData } from "@/lib/types";
import {
  ENTITY_BLOCKS,
  chartsForEntity,
  normalizeConfigForTabular,
} from "@/lib/predefinedCharts";
import { REPORT_INCLUDES_THROUGH_YMD } from "@/lib/report/reportThroughDate";

export type LiveChartSeries = {
  name: string;
  values: number[];
  sourceHeader: string;
  sourceColumnIndex?: number;
  countDistinctBySourceColumnIndex?: number;
  aggregation: ChartConfig["aggregation"];
  cumulative?: boolean;
};

export type LiveChartItem = {
  entityId: string;
  /** ID пресета из predefinedCharts — для серверного Excel (подписи pie и т.п.). */
  chartId: string;
  blockLabel: string;
  title: string;
  sheetName: string;
  chartType: "column" | "bar" | "line" | "area" | "pie";
  xHeader: string;
  xColumnIndex?: number;
  xIsDate?: boolean;
  /** Верхняя граница календарной даты для формул Excel (YYYY-MM-DD, включительно). */
  reportThroughYmd?: string;
  filters?: Array<{
    columnHeader: string;
    values?: string[];
    excludeValues?: string[];
  }>;
  dateGranularity?: ChartConfig["dateGranularity"];
  linkToSourcePossible?: boolean;
  chartDataEmpty?: boolean;
  resolveError?: string;
  categories: string[];
  series: LiveChartSeries[];
};

const CHART_IDS_IGNORE_STAGE_GLOBAL_FILTERS = new Set<string>([
  "deals_cumulative_count_by_month_area",
]);

function isStageLikeFilter(tabular: TabularData, f: ChartFilter): boolean {
  const meta = tabular.columns.find((c) => c.key === f.columnKey);
  const h = (meta?.header || f.columnKey || "").toLowerCase();
  return /стад|stage|воронк|pipeline/.test(h);
}

function withGlobalFilters(
  cfg: ChartConfig,
  global: ChartFilter[],
  tabular: TabularData,
  chartId: string,
): ChartConfig {
  const globalPrepared = CHART_IDS_IGNORE_STAGE_GLOBAL_FILTERS.has(chartId)
    ? global.filter((f) => !isStageLikeFilter(tabular, f))
    : global;
  return {
    ...cfg,
    filters: [...globalPrepared, ...cfg.filters],
  };
}

function columnTitle(data: TabularData, key: string): string {
  return data.columns.find((c) => c.key === key)?.header ?? key;
}

function xAxisLooksDateLike(
  tabular: TabularData,
  xKey: string | undefined,
  blockId: string,
): boolean {
  if (!xKey) {
    return false;
  }
  const xMeta = tabular.columns.find((c) => c.key === xKey);
  if (!xMeta) {
    return false;
  }
  if (xMeta.inferredType === "date") {
    return true;
  }
  if (xMeta.inferredType !== "string" && xMeta.inferredType !== "unknown") {
    return false;
  }
  const h = xMeta.header.trim().toLowerCase();
  if (
    /дата|date|time|period|закрыт|создан|изменен|отправк|план|deadline|calendar/i.test(
      h,
    )
  ) {
    return true;
  }
  // Для сделок часто тип колонки определяется как string, но значения - даты.
  if (blockId === "deals") {
    for (const v of xMeta.sampleValues.slice(0, 12)) {
      if (tryParseDate(v) != null) {
        return true;
      }
    }
  }
  return false;
}

function periodLegendLabel(cfg: ChartConfig, tabular: TabularData): string {
  const xMeta = cfg.xKey ? tabular.columns.find((c) => c.key === cfg.xKey) : null;
  if (xMeta?.inferredType !== "date") {
    return "категория";
  }
  if (cfg.dateGranularity === "month") {
    return "месяц";
  }
  if (cfg.dateGranularity === "quarter") {
    return "квартал";
  }
  if (cfg.dateGranularity === "year") {
    return "год";
  }
  return "день";
}

function toXlsxChartType(
  chartType: ChartConfig["chartType"],
): LiveChartItem["chartType"] {
  if (chartType === "bar") {
    return "bar";
  }
  if (chartType === "line") {
    return "line";
  }
  if (chartType === "area") {
    return "area";
  }
  if (chartType === "pie") {
    return "pie";
  }
  return "column";
}

/**
 * Те же определения графиков, что и для ExcelJS-экспорта и UI.
 * Вызывать в браузере перед POST /api/export-charts, чтобы агрегация шла по
 * исходному tabular (типы дат/чисел), а не по JSON после req.json() на сервере.
 */
export function buildLiveChartItems(
  tabular: TabularData,
  enabledChartIds: Set<string>,
  globalFilters: ChartFilter[],
  dealStageOrder?: string[] | null,
): LiveChartItem[] {
  const out: LiveChartItem[] = [];

  for (const block of ENTITY_BLOCKS) {
    for (const spec of chartsForEntity(block.id)) {
      if (!enabledChartIds.has(spec.id)) {
        continue;
      }
      const resolved = spec.resolve(tabular.columns);
      if (!resolved.ok) {
        out.push({
          entityId: block.id,
          chartId: spec.id,
          blockLabel: block.label,
          title: spec.title,
          sheetName: `${block.label} — ${spec.title}`,
          chartType: "column",
          xHeader: "",
          chartDataEmpty: true,
          resolveError: resolved.error,
          categories: [],
          series: [],
        });
        continue;
      }
      const cfg = normalizeConfigForTabular(
        tabular,
        withGlobalFilters(resolved.config, globalFilters, tabular, spec.id),
      );
      const reportChartThroughYmd =
        block.id === "deals" || block.id === "companies"
          ? REPORT_INCLUDES_THROUGH_YMD
          : undefined;
      const hasLiteralBars =
        Boolean(cfg.literalAvgBars && cfg.literalAvgBars.length > 0);
      if ((!cfg.xKey && !hasLiteralBars) || cfg.yKeys.length === 0) {
        out.push({
          entityId: block.id,
          chartId: spec.id,
          blockLabel: block.label,
          title: spec.title,
          sheetName: `${block.label} — ${spec.title}`,
          chartType: "column",
          xHeader: "",
          chartDataEmpty: true,
          resolveError:
            "Не удалось сопоставить колонки осей с текущим листом (проверьте заголовки).",
          categories: [],
          series: [],
        });
        continue;
      }
      const { data } = buildAggregatedChartRows(
        tabular.rows,
        cfg,
        tabular.columns,
        dealStageOrder,
        reportChartThroughYmd,
      );
      const chartDataEmpty = data.length === 0;
      const categories = chartDataEmpty ? [] : data.map((row) => String(row.name));
      const period = periodLegendLabel(cfg, tabular);
      const countDistinctBySourceColumnIndex = cfg.countDistinctByKey
        ? (() => {
            const idx = tabular.columns.findIndex(
              (c) => c.key === cfg.countDistinctByKey,
            );
            return idx >= 0 ? idx : undefined;
          })()
        : undefined;
      const series = cfg.yKeys.map((k, i) => {
        const srcKey = cfg.ySourceKeys?.[i] ?? k;
        const agg = cfg.yAggregations?.[i] ?? cfg.aggregation;
        const sourceColumnIndex = tabular.columns.findIndex((c) => c.key === srcKey);
        return {
          name: `${seriesLegendLabel(columnTitle(tabular, srcKey), agg, block.id)} · ${period}`,
          sourceHeader: columnTitle(tabular, srcKey),
          sourceColumnIndex: sourceColumnIndex >= 0 ? sourceColumnIndex : undefined,
          countDistinctBySourceColumnIndex,
          aggregation: agg,
          cumulative: cfg.cumulative,
          values: chartDataEmpty
            ? []
            : data.map((row) => {
                const v = row[k];
                return typeof v === "number" && Number.isFinite(v) ? v : Number(v) || 0;
              }),
        };
      });
      const xMeta = cfg.xKey
        ? tabular.columns.find((c) => c.key === cfg.xKey)
        : undefined;
      const xIsDate = xAxisLooksDateLike(tabular, cfg.xKey, block.id);
      const dateGranularity = cfg.dateGranularity ?? "day";
      const filters = cfg.filters
        .map((f) => {
          const header = columnTitle(tabular, f.columnKey);
          const values =
            f.values && f.values.length > 0 ? [...f.values] : undefined;
          const excludeValues =
            f.excludeValues && f.excludeValues.length > 0
              ? [...f.excludeValues]
              : undefined;
          if (!values && !excludeValues) {
            return null;
          }
          return { columnHeader: header, values, excludeValues };
        })
        .filter((x): x is NonNullable<typeof x> => x != null);
      /**
       * В Excel подставляются формулы COUNTIFS/SUMIFS от листа «Исходные данные» (день, месяц,
       * текстовая ось по точному совпадению), если нет фильтров. Иначе в Python — формула =n.
       */
      /**
       * Формулы COUNTIFS/SUMIFS от «Исходные данные» возможны без глобальных/локальных
       * фильтров (иначе совпадение с агрегацией в JS не гарантируется) и при известной оси X.
       */
      const linkToSourcePossible = Boolean(cfg.xKey) && cfg.filters.length === 0;

      out.push({
        entityId: block.id,
        chartId: spec.id,
        blockLabel: block.label,
        title: spec.title,
        sheetName: `${block.label} — ${spec.title}`,
        chartType: toXlsxChartType(cfg.chartType),
        xHeader: cfg.xKey ? columnTitle(tabular, cfg.xKey) : "Категория",
        xColumnIndex: cfg.xKey
          ? (() => {
              const idx = tabular.columns.findIndex((c) => c.key === cfg.xKey);
              return idx >= 0 ? idx : undefined;
            })()
          : undefined,
        xIsDate,
        reportThroughYmd: reportChartThroughYmd,
        filters,
        dateGranularity,
        linkToSourcePossible,
        chartDataEmpty,
        categories,
        series,
      });
    }
  }

  return out;
}
