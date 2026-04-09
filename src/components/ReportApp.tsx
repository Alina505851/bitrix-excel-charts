"use client";

import type { CSSProperties } from "react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  exportChartsToExcelFile,
  type ExcelExportTheme,
} from "@/lib/excel/exportChartsExcel";
import { parseWorkbookFromBuffer } from "@/lib/excel/parseWorkbook";
import {
  applyFilters,
  applyFiltersExcept,
  buildAggregatedChartRows,
  distinctCalendarMonthsFromColumn,
  distinctCalendarQuartersFromColumn,
  distinctStringValues,
  formatCalendarMonthRu,
  formatCalendarQuarterRu,
} from "@/lib/chart/buildChartData";
import type {
  AggregationMode,
  ChartConfig,
  ChartFilter,
  DateGranularity,
  TabularData,
} from "@/lib/types";
import {
  ENTITY_BLOCKS,
  chartsForEntity,
  companyCreatedDateColumn,
  normalizeConfigForTabular,
  responsibleColumn,
  suggestEntityAndCharts,
  suggestEntityFromFileName,
} from "@/lib/predefinedCharts";
import { formatDateYmdLocal, groupLabel, tryParseDate } from "@/lib/chart/coerce";
import {
  chartPaletteForEntity,
  COMPANIES_RESPONSIBLE_PIE_COLORS,
} from "@/lib/chart/chartPalettes";
import {
  CHART_AXES_DARK_SURFACE,
  CHART_SURFACE_OPTIONS,
  CHART_SURFACE_STORAGE_KEY,
  chartShellClass,
  chartUsesDarkSurface,
  isChartSurfaceId,
} from "@/lib/chart/chartSurface";
import type { ChartSurfaceId } from "@/lib/chart/chartSurface";
import type {
  EntityBlockId,
  PredefinedChartSpec,
} from "@/lib/predefinedCharts";
import {
  Area,
  AreaChart,
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  ComposedChart,
  Legend,
  LabelList,
  Line,
  LineChart,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
  type TooltipContentProps,
} from "recharts";

const MAX_FILE_BYTES = 20 * 1024 * 1024;
const PREVIEW_ROWS = 120;
const EXPORT_THEME_OPTIONS: { id: ExcelExportTheme; label: string }[] = [
  { id: "classic", label: "Classic" },
  { id: "vivid", label: "Vivid" },
  { id: "minimal", label: "Minimal" },
];

type ChartAxesTheme = { axis: string; grid: string; tick: string };

type PieLegendListOpts = {
  maxHeight?: number;
  maxWidth?: number | string;
};

function pieLegendListContent(
  payload:
    | readonly { value?: unknown; color?: string; payload?: unknown }[]
    | undefined,
  tickColor: string,
  total: number,
  opts?: PieLegendListOpts,
) {
  if (!payload?.length) {
    return null;
  }
  const listMaxH = opts?.maxHeight ?? 352;
  const listMaxW = opts?.maxWidth ?? 272;
  const listWidth: number | string =
    typeof listMaxW === "string" ? listMaxW : listMaxW;
  return (
    <ul
      style={{
        listStyle: "none",
        margin: 0,
        padding: "6px 4px 0 0",
        maxHeight: listMaxH,
        overflowY: "auto",
        overflowX: "hidden",
        width: listWidth,
        maxWidth: "100%",
        boxSizing: "border-box",
        textAlign: "left",
      }}
    >
      {payload.map((entry, i) => (
        <li
          key={`${String(entry.value)}-${i}`}
          style={{
            display: "flex",
            alignItems: "flex-start",
            gap: 8,
            marginBottom: 8,
            fontSize: 12,
            lineHeight: 1.4,
            color: tickColor,
            fontWeight: 500,
          }}
        >
          <span
            aria-hidden
            style={{
              width: 10,
              height: 10,
              borderRadius: 999,
              background: entry.color ?? "#94a3b8",
              flexShrink: 0,
              marginTop: 3,
              boxShadow: "inset 0 0 0 1px rgba(255,255,255,0.35)",
            }}
          />
          <span style={{ wordBreak: "break-word", overflowWrap: "anywhere" }}>
            {(() => {
              const name =
                entry.value != null && entry.value !== ""
                  ? String(entry.value)
                  : "—";
              const raw =
                entry.payload != null &&
                typeof entry.payload === "object" &&
                "value" in entry.payload
                  ? (entry.payload as { value?: unknown }).value
                  : null;
              const value = typeof raw === "number" && Number.isFinite(raw) ? raw : 0;
              const pct = total > 0 ? (value / total) * 100 : 0;
              return `${name} — ${formatTooltipNumber(value)} (${pct.toFixed(1)}%)`;
            })()}
          </span>
        </li>
      ))}
    </ul>
  );
}

function chartAxesTheme(entity: EntityBlockId | undefined): ChartAxesTheme {
  switch (entity) {
    case "deals":
      return {
        axis: "#64748b",
        grid: "rgba(168, 85, 247, 0.2)",
        tick: "#5b21b6",
      };
    case "contacts":
      return {
        axis: "#64748b",
        grid: "rgba(16, 185, 129, 0.19)",
        tick: "#065f46",
      };
    case "companies":
      return {
        axis: "#64748b",
        grid: "rgba(79, 70, 229, 0.2)",
        tick: "#4338ca",
      };
    case "quotes":
      return {
        axis: "#64748b",
        grid: "rgba(244, 63, 94, 0.2)",
        tick: "#9f1239",
      };
    default:
      return {
        axis: "#64748b",
        grid: "rgba(6, 182, 212, 0.18)",
        tick: "#0e7490",
      };
  }
}

function tooltipChrome(entity: EntityBlockId | undefined): CSSProperties {
  const accent =
    entity === "deals"
      ? "139, 92, 246"
      : entity === "contacts"
        ? "20, 184, 166"
        : entity === "companies"
          ? "99, 102, 241"
          : entity === "quotes"
            ? "225, 29, 72"
            : "14, 165, 233";
  return {
    borderRadius: "12px",
    fontSize: "13px",
    padding: "10px 14px",
    boxShadow: `0 12px 40px -10px rgba(${accent}, 0.22), 0 4px 16px rgba(15, 23, 42, 0.06)`,
    border: `1px solid rgba(${accent}, 0.22)`,
    backgroundColor: "rgba(255, 255, 255, 0.96)",
    backdropFilter: "blur(8px)",
  };
}

function formatTooltipNumber(v: number): string {
  if (!Number.isFinite(v)) {
    return String(v);
  }
  if (Math.abs(v - Math.round(v)) < 1e-9) {
    return Math.round(v).toLocaleString("ru-RU");
  }
  return v.toLocaleString("ru-RU", {
    maximumFractionDigits: 2,
    minimumFractionDigits: 0,
  });
}

/** Подпись метрики в легенде и тултипе: заголовок колонки из выгрузки + режим агрегации. */
function metricLegendLabel(
  tabular: TabularData,
  columnKey: string,
  aggregation: AggregationMode,
  periodLabel: string,
  cumulative?: boolean,
): string {
  const header =
    tabular.columns.find((c) => c.key === columnKey)?.header?.trim() ||
    columnKey;
  const headerLower = header.toLowerCase();
  let aggLabel = "сумма";
  if (cumulative) {
    aggLabel = "накопительно";
  } else {
    switch (aggregation) {
      case "count":
        aggLabel = "количество записей";
        break;
      case "count_nonempty":
        if (headerLower.includes("id") || headerLower.includes("ид")) {
          return `Кол-во компаний · количество компаний · ${periodLabel}`;
        }
        aggLabel = "количество непустых значений";
        break;
      case "avg":
        aggLabel = "среднее значение";
        break;
      case "sum":
      default:
        aggLabel = "сумма";
        break;
    }
  }
  return `${header} · ${aggLabel} · ${periodLabel}`;
}

/** Tailwind: полные строки классов для корректного JIT */
const ENTITY_CARD_BORDER: Record<EntityBlockId, string> = {
  leads: "border-l-4 border-l-sky-500",
  deals: "border-l-4 border-l-violet-500",
  contacts: "border-l-4 border-l-teal-500",
  companies: "border-l-4 border-l-indigo-500",
  quotes: "border-l-4 border-l-rose-500",
};

const ENTITY_SECTION_BADGE: Record<EntityBlockId, string> = {
  leads:
    "bg-sky-100 text-sky-800 dark:bg-sky-950/80 dark:text-sky-200 dark:ring-sky-800/60",
  deals:
    "bg-violet-100 text-violet-800 dark:bg-violet-950/80 dark:text-violet-200 dark:ring-violet-800/60",
  contacts:
    "bg-teal-100 text-teal-800 dark:bg-teal-950/80 dark:text-teal-200 dark:ring-teal-800/60",
  companies:
    "bg-indigo-100 text-indigo-900 dark:bg-indigo-950/80 dark:text-indigo-200 dark:ring-indigo-800/60",
  quotes:
    "bg-rose-100 text-rose-900 dark:bg-rose-950/80 dark:text-rose-100 dark:ring-rose-800/60",
};

/** Выбранный график в списке — акцент цветом раздела CRM */
const ENTITY_CHART_ROW_SELECTED: Record<EntityBlockId, string> = {
  leads:
    "border-sky-300/90 bg-gradient-to-br from-sky-50 to-white shadow-md shadow-sky-600/10 ring-2 ring-sky-400/30 dark:border-sky-600 dark:from-sky-950/50 dark:to-zinc-900 dark:ring-sky-500/25",
  deals:
    "border-violet-300/90 bg-gradient-to-br from-violet-50 to-white shadow-md shadow-violet-600/10 ring-2 ring-violet-400/30 dark:border-violet-600 dark:from-violet-950/50 dark:to-zinc-900 dark:ring-violet-500/25",
  contacts:
    "border-teal-300/90 bg-gradient-to-br from-teal-50 to-white shadow-md shadow-teal-600/10 ring-2 ring-teal-400/30 dark:border-teal-600 dark:from-teal-950/50 dark:to-zinc-900 dark:ring-teal-500/25",
  companies:
    "border-indigo-300/90 bg-gradient-to-br from-indigo-50 to-violet-50/50 shadow-md shadow-indigo-600/12 ring-2 ring-indigo-400/30 dark:border-indigo-600 dark:from-indigo-950/50 dark:to-zinc-900 dark:ring-indigo-500/25",
  quotes:
    "border-rose-300/90 bg-gradient-to-br from-rose-50 to-white shadow-md shadow-rose-600/10 ring-2 ring-rose-400/30 dark:border-rose-600 dark:from-rose-950/45 dark:to-zinc-900 dark:ring-rose-500/25",
};

/** Карточка одного графика — градиент и бэйдж под раздел CRM */
const ENTITY_CHART_CARD: Record<
  EntityBlockId,
  {
    wrap: string;
    head: string;
    title: string;
    badge: string;
    letter: string;
  }
> = {
  leads: {
    wrap: "rounded-2xl border border-sky-300/70 bg-gradient-to-br from-sky-100/70 via-sky-50/35 to-white p-5 shadow-lg shadow-sky-500/10 ring-1 ring-sky-200/60 dark:border-sky-700/55 dark:from-sky-950/35 dark:via-zinc-900/88 dark:to-zinc-950 dark:ring-sky-500/22",
    head: "mb-4 flex items-center gap-3 border-b border-sky-200/80 pb-3 dark:border-sky-800/35",
    title:
      "text-sm font-semibold tracking-tight text-sky-950 dark:text-sky-50",
    badge:
      "flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-gradient-to-br from-cyan-400 via-sky-500 to-indigo-600 text-xs font-bold text-white shadow-lg shadow-cyan-500/30 ring-1 ring-white/30",
    letter: "Л",
  },
  deals: {
    wrap: "rounded-2xl border border-violet-300/70 bg-gradient-to-br from-violet-100/65 via-fuchsia-50/30 to-white p-5 shadow-lg shadow-violet-500/12 ring-1 ring-violet-200/55 dark:border-violet-700/55 dark:from-violet-950/35 dark:via-zinc-900/88 dark:to-zinc-950 dark:ring-fuchsia-500/20",
    head: "mb-4 flex items-center gap-3 border-b border-violet-100/85 pb-3 dark:border-violet-800/32",
    title:
      "text-sm font-semibold tracking-tight text-violet-950 dark:text-violet-50",
    badge:
      "flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-gradient-to-br from-fuchsia-500 via-violet-600 to-indigo-700 text-xs font-bold text-white shadow-lg shadow-fuchsia-500/28 ring-1 ring-white/25",
    letter: "С",
  },
  contacts: {
    wrap: "rounded-2xl border border-emerald-300/65 bg-gradient-to-br from-emerald-100/55 via-teal-50/28 to-white p-5 shadow-lg shadow-emerald-500/10 ring-1 ring-emerald-200/50 dark:border-emerald-800/55 dark:from-emerald-950/32 dark:via-zinc-900/88 dark:to-zinc-950 dark:ring-emerald-500/18",
    head: "mb-4 flex items-center gap-3 border-b border-emerald-200/80 pb-3 dark:border-emerald-800/35",
    title:
      "text-sm font-semibold tracking-tight text-emerald-950 dark:text-emerald-50",
    badge:
      "flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-gradient-to-br from-emerald-400 via-teal-500 to-cyan-600 text-xs font-bold text-white shadow-lg shadow-emerald-500/28 ring-1 ring-white/25",
    letter: "К",
  },
  companies: {
    wrap: "rounded-2xl border border-indigo-300/65 bg-gradient-to-br from-indigo-100/55 via-violet-50/38 to-cyan-50/22 p-5 shadow-lg shadow-indigo-500/12 ring-1 ring-indigo-200/55 dark:border-indigo-800/55 dark:from-indigo-950/32 dark:via-zinc-900/85 dark:to-zinc-950 dark:ring-indigo-500/20",
    head: "mb-4 flex items-center gap-3 border-b border-indigo-200/70 pb-3 dark:border-indigo-800/45",
    title:
      "text-sm font-semibold tracking-tight text-indigo-950 dark:text-indigo-50",
    badge:
      "flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-gradient-to-br from-indigo-500 via-violet-600 to-cyan-500 text-xs font-bold text-white shadow-lg shadow-indigo-500/28 ring-1 ring-white/25",
    letter: "ЮЛ",
  },
  quotes: {
    wrap: "rounded-2xl border border-rose-300/70 bg-gradient-to-br from-rose-100/60 via-pink-50/30 to-white p-5 shadow-lg shadow-rose-500/12 ring-1 ring-rose-200/55 dark:border-rose-800/55 dark:from-rose-950/32 dark:via-zinc-900/88 dark:to-zinc-950 dark:ring-rose-500/20",
    head: "mb-4 flex items-center gap-3 border-b border-rose-100/85 pb-3 dark:border-rose-800/35",
    title:
      "text-sm font-semibold tracking-tight text-rose-950 dark:text-rose-50",
    badge:
      "flex h-9 min-w-9 shrink-0 items-center justify-center rounded-xl bg-gradient-to-br from-rose-500 via-pink-500 to-fuchsia-600 px-1.5 text-[10px] font-bold leading-none text-white shadow-lg shadow-rose-500/30 ring-1 ring-white/25",
    letter: "КП",
  },
};

const ENTITY_CHART_ERR: Record<EntityBlockId, string> = {
  leads:
    "rounded-2xl border border-amber-200/90 bg-gradient-to-br from-amber-50/95 to-sky-50/20 p-4 shadow-sm ring-1 ring-amber-100/45 dark:border-amber-900/55 dark:from-amber-950/42 dark:to-zinc-900/92 dark:ring-amber-900/25",
  deals:
    "rounded-2xl border border-amber-200/90 bg-gradient-to-br from-amber-50/95 to-violet-50/20 p-4 shadow-sm ring-1 ring-amber-100/45 dark:border-amber-900/55 dark:from-amber-950/42 dark:to-zinc-900/92 dark:ring-amber-900/25",
  contacts:
    "rounded-2xl border border-amber-200/90 bg-gradient-to-br from-amber-50/95 to-teal-50/20 p-4 shadow-sm ring-1 ring-amber-100/45 dark:border-amber-900/55 dark:from-amber-950/42 dark:to-zinc-900/92 dark:ring-amber-900/25",
  companies:
    "rounded-2xl border border-indigo-300/75 bg-gradient-to-br from-indigo-50/95 to-violet-50/45 p-4 shadow-sm ring-1 ring-indigo-200/50 dark:border-indigo-800/65 dark:from-indigo-950/48 dark:to-zinc-900/92 dark:ring-indigo-900/28",
  quotes:
    "rounded-2xl border border-amber-200/90 bg-gradient-to-br from-amber-50/95 to-rose-50/30 p-4 shadow-sm ring-1 ring-amber-100/45 dark:border-amber-900/55 dark:from-amber-950/42 dark:to-zinc-900/92 dark:ring-rose-900/22",
};

const ENTITY_PICKER_SELECTED: Record<EntityBlockId, string> = {
  leads:
    "border-sky-400 bg-gradient-to-br from-sky-50 via-white to-cyan-50/30 shadow-lg shadow-sky-500/20 ring-2 ring-sky-400/50 dark:border-sky-500 dark:from-sky-950/45 dark:via-zinc-900 dark:to-sky-950/25 dark:ring-sky-500/45",
  deals:
    "border-violet-400 bg-gradient-to-br from-violet-50 via-white to-fuchsia-50/25 shadow-lg shadow-violet-500/18 ring-2 ring-violet-400/45 dark:border-violet-500 dark:from-violet-950/45 dark:via-zinc-900 dark:to-violet-950/25 dark:ring-violet-500/40",
  contacts:
    "border-emerald-400 bg-gradient-to-br from-emerald-50 via-white to-teal-50/25 shadow-lg shadow-emerald-500/18 ring-2 ring-emerald-400/45 dark:border-emerald-500 dark:from-emerald-950/40 dark:via-zinc-900 dark:to-emerald-950/22 dark:ring-emerald-500/40",
  companies:
    "border-indigo-400 bg-gradient-to-br from-indigo-50 via-white to-violet-50/20 shadow-lg shadow-indigo-500/18 ring-2 ring-indigo-400/45 dark:border-indigo-500 dark:from-indigo-950/42 dark:via-zinc-900 dark:to-indigo-950/22 dark:ring-indigo-500/40",
  quotes:
    "border-rose-400 bg-gradient-to-br from-rose-50 via-white to-pink-50/25 shadow-lg shadow-rose-500/18 ring-2 ring-rose-400/45 dark:border-rose-500 dark:from-rose-950/42 dark:via-zinc-900 dark:to-rose-950/22 dark:ring-rose-500/40",
};

const ENTITY_PICKER_IDLE =
  "border-zinc-200/75 bg-white/90 shadow-sm hover:border-zinc-300 hover:shadow-md active:scale-[0.99] dark:border-zinc-700/75 dark:bg-zinc-900/55 dark:hover:border-zinc-600 dark:hover:bg-zinc-900/75";

const ENTITY_CHECKBOX_ACCENT: Record<EntityBlockId, string> = {
  leads:
    "accent-sky-600 focus-visible:ring-sky-500/55 dark:accent-sky-500",
  deals:
    "accent-violet-600 focus-visible:ring-violet-500/55 dark:accent-violet-500",
  contacts:
    "accent-emerald-600 focus-visible:ring-emerald-500/55 dark:accent-emerald-500",
  companies:
    "accent-indigo-600 focus-visible:ring-indigo-500/55 dark:accent-indigo-500",
  quotes: "accent-rose-600 focus-visible:ring-rose-500/55 dark:accent-rose-500",
};

const SECTION_LABEL =
  "text-[10px] font-bold uppercase tracking-[0.2em] text-zinc-500 dark:text-zinc-400";

const INPUT_FIELD =
  "w-full rounded-2xl border-0 bg-white/95 px-3.5 py-2.5 text-sm text-zinc-900 shadow-[0_1px_2px_rgb(0_0_0/0.04)] ring-1 ring-zinc-200/85 transition-[box-shadow,ring-color] placeholder:text-zinc-400 hover:ring-zinc-300/80 focus:outline-none focus:ring-2 focus:ring-sky-500/45 dark:bg-zinc-900/90 dark:text-zinc-100 dark:ring-zinc-600/85 dark:placeholder:text-zinc-500 dark:hover:ring-zinc-500/70 dark:focus:ring-sky-400/45";

const BTN_PRIMARY =
  "inline-flex items-center justify-center gap-2 rounded-2xl bg-gradient-to-r from-cyan-500 via-sky-600 to-indigo-600 px-4 py-2.5 text-sm font-semibold text-white shadow-[0_10px_28px_-6px_rgb(2_132_199/0.45)] transition-[filter,box-shadow,transform] hover:brightness-[1.03] hover:shadow-[0_14px_32px_-8px_rgb(79_70_229/0.4)] active:scale-[0.99] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-cyan-400 focus-visible:ring-offset-2 focus-visible:ring-offset-white disabled:pointer-events-none disabled:opacity-45 dark:focus-visible:ring-offset-zinc-900";

const BTN_GHOST =
  "inline-flex items-center justify-center gap-2 rounded-2xl border border-zinc-200/90 bg-white/90 px-4 py-2.5 text-sm font-medium text-zinc-800 shadow-sm backdrop-blur-md transition-[border-color,background-color,box-shadow,transform] hover:border-zinc-300 hover:bg-white hover:shadow-md active:scale-[0.99] disabled:pointer-events-none disabled:opacity-45 dark:border-zinc-600 dark:bg-zinc-800/90 dark:text-zinc-100 dark:hover:bg-zinc-800";

const BTN_EXPORT =
  "inline-flex items-center justify-center gap-2 rounded-2xl border border-emerald-400/50 bg-gradient-to-r from-emerald-500 via-teal-500 to-cyan-600 px-4 py-2.5 text-sm font-semibold text-white shadow-[0_10px_28px_-6px_rgb(16_185_129/0.42)] transition-[filter,box-shadow,transform] hover:brightness-[1.03] hover:shadow-[0_14px_32px_-8px_rgb(20_184_166/0.38)] active:scale-[0.99] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-emerald-400 focus-visible:ring-offset-2 focus-visible:ring-offset-white disabled:pointer-events-none disabled:opacity-45 disabled:shadow-none dark:focus-visible:ring-offset-zinc-900";

function chartCountLabel(n: number): string {
  if (n === 0) {
    return "нет выбранных графиков";
  }
  const m100 = n % 100;
  const m10 = n % 10;
  if (m100 >= 11 && m100 <= 14) {
    return `${n} графиков`;
  }
  if (m10 === 1) {
    return `${n} график`;
  }
  if (m10 >= 2 && m10 <= 4) {
    return `${n} графика`;
  }
  return `${n} графиков`;
}

function ExportDock({
  exportEnabled,
  exportButtonTitle,
  onExport,
  chartCount,
  filteredRowCount,
}: {
  exportEnabled: boolean;
  exportButtonTitle: string;
  onExport: () => void;
  chartCount: number;
  filteredRowCount: number;
}) {
  return (
    <div
      className="fixed bottom-0 left-0 right-0 z-40 border-t border-emerald-400/35 bg-gradient-to-b from-white/98 to-emerald-50/40 pb-[max(0.85rem,env(safe-area-inset-bottom))] pt-3.5 shadow-[0_-12px_48px_-16px_rgba(16,185,129,0.22)] backdrop-blur-xl dark:border-emerald-800/50 dark:from-zinc-950/98 dark:to-emerald-950/25 dark:shadow-[0_-12px_48px_-12px_rgba(0,0,0,0.55)]"
      role="region"
      aria-label="Выгрузка отчёта в Excel"
    >
      <div className="mx-auto flex max-w-[1600px] flex-col gap-3 px-4 sm:flex-row sm:items-center sm:justify-between sm:gap-4 sm:px-6 lg:px-8">
        <div className="min-w-0 flex-1">
          <p className="text-[15px] font-bold tracking-tight text-emerald-900 dark:text-emerald-100">
            Выгрузка отчёта в Excel
          </p>
          <p className="mt-0.5 text-xs leading-snug text-zinc-600 dark:text-zinc-400">
            {chartCount > 0 ? (
              <>
                {chartCountLabel(chartCount)} ·{" "}
                {filteredRowCount.toLocaleString("ru-RU")} строк после фильтров
              </>
            ) : (
              "Отметьте графики в боковой панели — затем нажмите кнопку справа"
            )}
          </p>
        </div>
        <button
          type="button"
          className={
            (exportEnabled ? BTN_EXPORT : BTN_GHOST) +
            " w-full min-h-[3rem] shrink-0 px-7 py-3.5 text-base font-bold shadow-[0_12px_32px_-12px_rgba(16,185,129,0.45)] sm:w-auto sm:min-w-[min(100%,15rem)]"
          }
          disabled={!exportEnabled}
          title={exportButtonTitle}
          onClick={onExport}
        >
          <svg
            className={
              exportEnabled ? "h-5 w-5 text-white/95" : "h-5 w-5 text-zinc-500"
            }
            fill="none"
            viewBox="0 0 24 24"
            stroke="currentColor"
            strokeWidth="1.75"
            aria-hidden
          >
            <path
              strokeLinecap="round"
              strokeLinejoin="round"
              d="M3 16.5v2.25A2.25 2.25 0 0 0 5.25 21h13.5A2.25 2.25 0 0 0 21 18.75V16.5M16.5 12 12 16.5m0 0L7.5 12m4.5 4.5V3"
            />
          </svg>
          Скачать Excel
        </button>
      </div>
    </div>
  );
}

function ReportSetupSteps({
  entitySelected,
  chartsSelected,
}: {
  entitySelected: boolean;
  chartsSelected: boolean;
}) {
  const pill = (opts: {
    done: boolean;
    current: boolean;
    n: string;
    label: string;
  }) => {
    const { done, current, n, label } = opts;
    return (
      <li
        className={[
          "inline-flex items-center gap-1.5 rounded-full border px-2.5 py-1.5 text-[11px] font-semibold leading-none tracking-tight transition-colors",
          done
            ? "border-emerald-300/75 bg-emerald-500/10 text-emerald-900 shadow-sm dark:border-emerald-700/50 dark:bg-emerald-500/12 dark:text-emerald-100"
            : current
              ? "border-sky-400/70 bg-sky-500/12 text-sky-950 ring-1 ring-sky-400/35 dark:border-sky-600/60 dark:bg-sky-500/14 dark:text-sky-100 dark:ring-sky-500/25"
              : "border-zinc-200/75 bg-zinc-50/80 text-zinc-500 dark:border-zinc-700 dark:bg-zinc-900/45 dark:text-zinc-500",
        ].join(" ")}
      >
        <span
          className={[
            "flex h-4 min-w-[1rem] shrink-0 items-center justify-center rounded-full text-[9px] font-bold text-white",
            done
              ? "bg-emerald-500"
              : current
                ? "bg-sky-500"
                : "bg-zinc-400 dark:bg-zinc-600",
          ].join(" ")}
          aria-hidden
        >
          {done ? "✓" : n}
        </span>
        {label}
      </li>
    );
  };

  return (
    <nav aria-label="Шаги настройки отчёта" className="relative">
      <ol className="flex flex-wrap items-center gap-x-1 gap-y-2">
        {pill({ done: true, current: false, n: "1", label: "Файл" })}
        <li
          aria-hidden
          className="select-none px-0.5 text-zinc-300 dark:text-zinc-600"
        >
          →
        </li>
        {pill({
          done: entitySelected,
          current: !entitySelected,
          n: "2",
          label: "Сущность",
        })}
        <li
          aria-hidden
          className="select-none px-0.5 text-zinc-300 dark:text-zinc-600"
        >
          →
        </li>
        {pill({
          done: chartsSelected,
          current: entitySelected && !chartsSelected,
          n: "3",
          label: "Графики",
        })}
      </ol>
    </nav>
  );
}

function monthKeyForCompanyTrend(d: Date): string {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

/**
 * Средний помесячный % прироста числа записей с датой создания:
 * по каждой паре соседних календарных месяцев от минимального к максимальному;
 * месяцы без записей считаются как 0.
 */
/** Сколько разных непустых значений в колонке «Ответственный» (после фильтров). */
function countDistinctResponsible(
  rows: Record<string, unknown>[],
  columnKey: string | null,
): number | null {
  if (!columnKey) {
    return null;
  }
  if (rows.length === 0) {
    return 0;
  }
  const distinct = new Set<string>();
  for (const row of rows) {
    const lab = groupLabel(row[columnKey]);
    if (lab !== "(пусто)" && lab.trim() !== "") {
      distinct.add(lab);
    }
  }
  return distinct.size;
}

function averageMonthOverMonthGrowthPercent(dates: Date[]): number | null {
  if (dates.length === 0) {
    return null;
  }
  const byMonth = new Map<string, number>();
  for (const d of dates) {
    const k = monthKeyForCompanyTrend(d);
    byMonth.set(k, (byMonth.get(k) ?? 0) + 1);
  }
  let minT = dates[0]!.getTime();
  let maxT = dates[0]!.getTime();
  for (const d of dates) {
    const t = d.getTime();
    if (t < minT) {
      minT = t;
    }
    if (t > maxT) {
      maxT = t;
    }
  }
  const start = new Date(minT);
  const end = new Date(maxT);
  let y = start.getFullYear();
  let mo = start.getMonth();
  const yEnd = end.getFullYear();
  const moEnd = end.getMonth();
  const counts: number[] = [];
  while (y < yEnd || (y === yEnd && mo <= moEnd)) {
    const k = `${y}-${String(mo + 1).padStart(2, "0")}`;
    counts.push(byMonth.get(k) ?? 0);
    mo += 1;
    if (mo > 11) {
      mo = 0;
      y += 1;
    }
  }
  if (counts.length < 2) {
    return null;
  }
  const pcts: number[] = [];
  for (let i = 1; i < counts.length; i++) {
    const prev = counts[i - 1]!;
    const cur = counts[i]!;
    pcts.push((100 * (cur - prev)) / Math.max(prev, 1));
  }
  return pcts.reduce((a, b) => a + b, 0) / pcts.length;
}

function mergeGlobalFilters(
  cfg: ChartConfig,
  global: ChartFilter[],
): ChartConfig {
  return {
    ...cfg,
    filters: [...global, ...cfg.filters],
  };
}

function activationHintFromResolveError(error: string): string {
  const normalized = error.trim().replace(/\.$/, "");
  return `Для активации нужно: ${normalized}.`;
}

function feasibleChartIdsForEntity(
  entity: EntityBlockId,
  tabular: TabularData,
): string[] {
  return chartsForEntity(entity)
    .filter((s) => s.resolve(tabular.columns).ok)
    .map((s) => s.id);
}

function autoSelectableChartIdsForEntity(
  entity: EntityBlockId,
  tabular: TabularData,
): string[] {
  const ids: string[] = [];
  for (const spec of chartsForEntity(entity)) {
    const resolved = spec.resolve(tabular.columns);
    if (!resolved.ok) {
      continue;
    }
    const cfg = normalizeConfigForTabular(tabular, resolved.config);
    if (!cfg.xKey || cfg.yKeys.length === 0) {
      continue;
    }
    const built = buildAggregatedChartRows(tabular.rows, cfg, tabular.columns);
    if (built.data.length > 0) {
      ids.push(spec.id);
    }
  }
  return ids;
}

export default function ReportApp() {
  const bufferRef = useRef<ArrayBuffer | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const [tabular, setTabular] = useState<TabularData | null>(null);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [fileLoading, setFileLoading] = useState(false);
  const [dragActive, setDragActive] = useState(false);

  const [enabledChartIds, setEnabledChartIds] = useState<Set<string>>(
    () => new Set(),
  );
  /** Сущность выгрузки: сначала выбирается она, затем графики из выпадающего списка */
  const [importEntity, setImportEntity] = useState<EntityBlockId | null>(null);
  const [globalFilters, setGlobalFilters] = useState<ChartFilter[]>([]);
  const [exportMessage, setExportMessage] = useState<string | null>(null);
  const [chartSurface, setChartSurface] = useState<ChartSurfaceId>("default");
  const [exportTheme, setExportTheme] = useState<ExcelExportTheme>("classic");

  useEffect(() => {
    try {
      const raw = localStorage.getItem(CHART_SURFACE_STORAGE_KEY);
      if (raw && isChartSurfaceId(raw)) {
        setChartSurface(raw);
      }
    } catch {
      /* ignore */
    }
  }, []);

  useEffect(() => {
    try {
      localStorage.setItem(CHART_SURFACE_STORAGE_KEY, chartSurface);
    } catch {
      /* ignore */
    }
  }, [chartSurface]);

  const setChartEnabled = useCallback((id: string, on: boolean) => {
    setEnabledChartIds((prev) => {
      const next = new Set(prev);
      if (on) {
        next.add(id);
      } else {
        next.delete(id);
      }
      return next;
    });
  }, []);

  const applyEntitySelection = useCallback(
    (entity: EntityBlockId | null) => {
      setImportEntity(entity);
      if (!entity || !tabular) {
        setEnabledChartIds(new Set());
        return;
      }
      setEnabledChartIds(new Set(feasibleChartIdsForEntity(entity, tabular)));
    },
    [tabular],
  );

  const applyFileBuffer = useCallback(
    (buffer: ArrayBuffer, name: string, sheet?: string) => {
      setLoadError(null);
      const result = parseWorkbookFromBuffer(buffer, sheet);
      if (!result.ok) {
        setTabular(null);
        setLoadError(result.error);
        setEnabledChartIds(new Set());
        setImportEntity(null);
        setGlobalFilters([]);
        return;
      }
      bufferRef.current = buffer;
      setFileName(name);
      setTabular(result.data);
      setGlobalFilters([]);
      setExportMessage(null);
      const inferredByName = suggestEntityFromFileName(name);
      const suggestion = inferredByName
        ? {
            entity: inferredByName,
            chartIds: autoSelectableChartIdsForEntity(inferredByName, result.data),
          }
        : suggestEntityAndCharts(result.data.columns);
      if (suggestion) {
        setImportEntity(suggestion.entity);
        const autoIds = autoSelectableChartIdsForEntity(
          suggestion.entity,
          result.data,
        );
        setEnabledChartIds(new Set(autoIds));
      } else {
        setImportEntity(null);
        setEnabledChartIds(new Set());
      }
    },
    [],
  );

  useEffect(() => {
    if (!tabular || !importEntity) {
      return;
    }
    const feasible = new Set(feasibleChartIdsForEntity(importEntity, tabular));
    setEnabledChartIds((prev) => {
      const next = new Set([...prev].filter((id) => feasible.has(id)));
      if (next.size === prev.size) {
        return prev;
      }
      return next;
    });
  }, [tabular, importEntity]);

  const onPickFile = useCallback(
    (file: File | null) => {
      if (!file) {
        return;
      }
      const lower = file.name.toLowerCase();
      if (!lower.endsWith(".xlsx") && !lower.endsWith(".xls")) {
        setLoadError("Поддерживаются только файлы .xlsx и .xls.");
        return;
      }
      if (file.size > MAX_FILE_BYTES) {
        setLoadError(
          `Файл слишком больший (максимум ${(MAX_FILE_BYTES / 1024 / 1024).toFixed(0)} МБ).`,
        );
        return;
      }
      setLoadError(null);
      setFileLoading(true);
      file
        .arrayBuffer()
        .then((buf) => {
          try {
            applyFileBuffer(buf, file.name);
          } catch {
            setLoadError(
              "Ошибка при разборе файла. Проверьте, что это корректная книга Excel.",
            );
          }
        })
        .catch(() => {
          setLoadError(
            "Не удалось прочитать файл. Закройте его в Excel и попробуйте снова.",
          );
        })
        .finally(() => {
          setFileLoading(false);
        });
    },
    [applyFileBuffer],
  );

  const previewColumns = tabular?.columns.slice(0, 12) ?? [];

  const filteredRows = useMemo(() => {
    if (!tabular) {
      return [];
    }
    return applyFilters(tabular.rows, globalFilters);
  }, [tabular, globalFilters]);

  /** Сводка для выгрузки компаний: всего строк, первая/последняя дата, среднее в месяц. */
  const companySummaryKpi = useMemo(() => {
    if (!tabular || importEntity !== "companies") {
      return null;
    }
    const total = filteredRows.length;
    const respMeta = responsibleColumn(tabular.columns);
    const responsibleKey = respMeta?.key ?? null;
    const activeResponsiblesCount = countDistinctResponsible(
      filteredRows,
      responsibleKey,
    );
    const hasResponsibleColumn = !!respMeta;
    const dateMeta = companyCreatedDateColumn(tabular.columns);
    const emptyDates = {
      firstRecordDate: null as string | null,
      lastRecordDate: null as string | null,
      growthRateAvgPercent: null as number | null,
    };
    const kpiBase = {
      total,
      activeResponsiblesCount,
      hasResponsibleColumn,
    };
    if (!dateMeta || total === 0) {
      return {
        ...kpiBase,
        avgPerMonth: null as number | null,
        hasDateColumn: !!dateMeta,
        ...emptyDates,
      };
    }
    const dateKey = dateMeta.key;
    const dates: Date[] = [];
    for (const row of filteredRows) {
      const raw = row[dateKey];
      let d: Date | null = null;
      if (raw instanceof Date && !Number.isNaN(raw.getTime())) {
        d = raw;
      } else {
        d = tryParseDate(raw);
      }
      if (d && !Number.isNaN(d.getTime())) {
        dates.push(d);
      }
    }
    if (dates.length === 0) {
      return {
        ...kpiBase,
        avgPerMonth: null as number | null,
        hasDateColumn: true,
        ...emptyDates,
      };
    }
    let minT = dates[0]!.getTime();
    let maxT = dates[0]!.getTime();
    for (const d of dates) {
      const t = d.getTime();
      if (t < minT) {
        minT = t;
      }
      if (t > maxT) {
        maxT = t;
      }
    }
    const minD = new Date(minT);
    const maxD = new Date(maxT);
    const monthSpan =
      (maxD.getFullYear() - minD.getFullYear()) * 12 +
      (maxD.getMonth() - minD.getMonth()) +
      1;
    const avgPerMonth = total / Math.max(1, monthSpan);
    const dateFmt = (dt: Date) =>
      dt.toLocaleDateString("ru-RU", {
        day: "2-digit",
        month: "2-digit",
        year: "numeric",
      });
    const growthRateAvgPercent = averageMonthOverMonthGrowthPercent(dates);
    return {
      ...kpiBase,
      avgPerMonth,
      hasDateColumn: true,
      firstRecordDate: dateFmt(minD),
      lastRecordDate: dateFmt(maxD),
      growthRateAvgPercent,
    };
  }, [tabular, importEntity, filteredRows]);

  const onExportChartsExcel = useCallback(async () => {
    if (!tabular) {
      return;
    }
    const base =
      fileName?.replace(/\.(xlsx|xls)$/i, "").trim() || "Отчет";
    const res = await exportChartsToExcelFile(
      tabular,
      enabledChartIds,
      globalFilters,
      {
        fileNameBase: base,
        sourceFileName: fileName ?? undefined,
        theme: exportTheme,
      },
    );
    setExportMessage(
      res.ok
        ? `Файл сохранён: ${res.fileName} (${res.sheetsWritten} лист.)`
        : res.error,
    );
  }, [tabular, fileName, enabledChartIds, globalFilters, exportTheme]);

  const exportEnabled =
    tabular != null &&
    enabledChartIds.size > 0 &&
    tabular.rows.length > 0;

  const exportButtonTitle = useMemo(() => {
    if (exportEnabled) {
      return "Скачать Excel: сводка, таблицы и диаграммы по выбранным отчётам и фильтрам";
    }
    if (!tabular || tabular.rows.length === 0) {
      return "В файле нет строк с данными — загрузите другой файл";
    }
    return "Выберите сущность выгрузки и отметьте хотя бы один график в списке слева";
  }, [exportEnabled, tabular?.rows.length]);

  return (
    <div className="app-shell-bg min-h-full">
      <a href="#main-content" className="skip-link">
        К основному содержимому
      </a>
      <header className="sticky top-0 z-30 border-b border-zinc-200/50 bg-white/70 shadow-[0_12px_40px_-18px_rgb(15_23_42/0.12)] backdrop-blur-2xl dark:border-zinc-800/60 dark:bg-zinc-950/75 dark:shadow-[0_12px_40px_-16px_rgb(0_0_0/0.45)]">
        <div
          className="pointer-events-none absolute inset-x-0 bottom-0 h-px bg-gradient-to-r from-transparent via-sky-400/50 to-transparent dark:via-indigo-400/40"
          aria-hidden
        />
        <div className="relative mx-auto flex max-w-[1600px] flex-wrap items-center justify-between gap-4 px-4 py-3.5 sm:px-6 lg:px-8">
          <div className="flex min-w-0 items-center gap-3.5 sm:gap-4">
            <div
              className="relative flex h-11 w-11 shrink-0 items-center justify-center rounded-2xl bg-gradient-to-br from-cyan-400 via-sky-600 to-indigo-700 text-[10px] font-bold uppercase tracking-[0.08em] text-white shadow-[0_8px_20px_-4px_rgb(2_132_199/0.5)] ring-1 ring-white/40 dark:from-sky-500 dark:via-indigo-600 dark:to-violet-700 dark:ring-white/12 sm:h-12 sm:w-12 sm:text-[11px]"
              aria-hidden
            >
              B24
              <span className="absolute -bottom-0.5 -right-0.5 flex h-3.5 w-3.5 items-center justify-center rounded-full bg-emerald-500 ring-2 ring-white dark:ring-zinc-950" />
            </div>
            <div className="min-w-0">
              <p className={`${SECTION_LABEL} mb-0.5 hidden sm:block`}>
                Локальные отчёты
              </p>
              <h1 className="truncate text-lg font-semibold tracking-tight text-zinc-900 dark:text-zinc-50 sm:text-xl">
                Отчёты по выгрузкам Excel
              </h1>
              <p className="mt-0.5 text-xs leading-snug text-zinc-500 dark:text-zinc-400 sm:text-sm">
                Лиды, сделки, КП, контакты и компании — без отправки на сервер
              </p>
            </div>
          </div>
          {tabular && fileName ? (
            <div className="flex max-w-full min-w-0 items-center gap-2 sm:max-w-[min(100%,22rem)]">
              <span className="flex min-w-0 items-center gap-2 truncate rounded-2xl border border-zinc-200/80 bg-zinc-50/95 py-2 pl-3 pr-3.5 text-xs font-medium text-zinc-700 shadow-[inset_0_1px_0_rgb(255_255_255/0.65)] dark:border-zinc-700/90 dark:bg-zinc-900/85 dark:text-zinc-200 dark:shadow-none">
                <svg
                  className="h-3.5 w-3.5 shrink-0 text-zinc-400"
                  fill="none"
                  viewBox="0 0 24 24"
                  stroke="currentColor"
                  strokeWidth="2"
                  aria-hidden
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    d="M19.5 14.25v-2.625a3.375 3.375 0 0 0-3.375-3.375h-1.5A1.125 1.125 0 0 1 13.5 7.125v-1.5a3.375 3.375 0 0 0-3.375-3.375H8.25m0 12.75h7.5m-7.5 3H12M10.5 2.25H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 0 0-9-9Z"
                  />
                </svg>
                <span className="truncate">{fileName}</span>
              </span>
            </div>
          ) : (
            <span className="inline-flex items-center gap-1.5 rounded-full border border-emerald-200/90 bg-emerald-50/95 px-3.5 py-1.5 text-xs font-semibold text-emerald-800 shadow-sm dark:border-emerald-800/70 dark:bg-emerald-950/80 dark:text-emerald-200">
              <span
                className="h-2 w-2 shrink-0 rounded-full bg-emerald-500 shadow-[0_0_0_3px_rgba(16,185,129,0.25)]"
                aria-hidden
              />
              Без отправки на сервер
            </span>
          )}
        </div>
      </header>

      <main
        id="main-content"
        tabIndex={-1}
        className={`mx-auto max-w-[1600px] scroll-mt-28 space-y-11 px-4 py-9 outline-none sm:px-6 sm:py-11 lg:px-8 ${tabular ? "pb-32 sm:pb-28" : ""}`}
      >
        <section
          aria-labelledby="upload-title"
          aria-describedby="upload-hint"
          aria-busy={fileLoading}
          className={[
            "group relative min-h-[272px] overflow-hidden rounded-[1.75rem] border p-10 text-center shadow-[0_20px_50px_-28px_rgb(15_23_42/0.1)] ring-1 ring-zinc-200/40 transition-[border-color,box-shadow,background-color,transform] duration-300 sm:min-h-[288px] sm:p-14 dark:ring-zinc-700/40",
            dragActive
              ? "scale-[1.01] border-sky-400/80 bg-gradient-to-br from-sky-50 via-cyan-50/90 to-indigo-50/80 shadow-[0_24px_56px_-24px_rgb(14_165_233/0.35)] ring-sky-300/50 dark:border-sky-500/70 dark:from-sky-950/50 dark:via-cyan-950/35 dark:to-indigo-950/40 dark:ring-sky-500/35"
              : "border-zinc-200/90 bg-gradient-to-br from-white/90 via-white/70 to-sky-50/30 backdrop-blur-sm dark:border-zinc-700/80 dark:from-zinc-900/75 dark:via-zinc-900/55 dark:to-indigo-950/20",
          ].join(" ")}
        >
          {/*
            Нативный выбор файла: input не с display:none (иначе .click() и диалог
            могут не работать в части браузеров). Прозрачный слой на весь блок —
            клик попадает в input напрямую.
          */}
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.xls,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            className="absolute inset-0 z-20 h-full w-full cursor-pointer opacity-0 disabled:cursor-wait"
            id="excel-upload"
            disabled={fileLoading}
            aria-label="Загрузить файл Excel (.xlsx или .xls)"
            onDragEnter={(e) => {
              e.preventDefault();
              setDragActive(true);
            }}
            onDragOver={(e) => {
              e.preventDefault();
              e.dataTransfer.dropEffect = "copy";
              setDragActive(true);
            }}
            onDragLeave={(e) => {
              e.preventDefault();
              const related = e.relatedTarget as Node | null;
              if (related && e.currentTarget.contains(related)) {
                return;
              }
              setDragActive(false);
            }}
            onDrop={(e) => {
              e.preventDefault();
              setDragActive(false);
              const f = e.dataTransfer.files[0];
              onPickFile(f ?? null);
            }}
            onChange={(e) => {
              const f = e.target.files?.[0] ?? null;
              onPickFile(f);
              e.target.value = "";
            }}
          />
          <div className="pointer-events-none absolute -right-20 -top-24 h-56 w-56 rounded-full bg-cyan-400/20 blur-3xl dark:bg-cyan-500/14" />
          <div className="pointer-events-none absolute -bottom-16 -left-12 h-44 w-44 rounded-full bg-fuchsia-400/14 blur-3xl dark:bg-violet-500/12" />
          <div className="pointer-events-none absolute left-1/2 top-1/2 h-64 w-64 -translate-x-1/2 -translate-y-1/2 rounded-full bg-indigo-400/8 blur-3xl dark:bg-indigo-500/10" />
          {fileLoading && (
            <div className="pointer-events-auto absolute inset-0 z-30 flex items-center justify-center bg-white/88 backdrop-blur-md dark:bg-zinc-950/85">
              <span className="flex items-center gap-2 text-sm font-medium text-sky-800 dark:text-sky-300">
                <span className="inline-block size-4 animate-pulse rounded-full bg-sky-500/80" />
                Чтение файла…
              </span>
            </div>
          )}
          <div className="pointer-events-none relative z-10 mx-auto flex min-h-[240px] w-full max-w-2xl flex-col items-center justify-center gap-5 px-2 sm:min-h-[260px]">
            <span className="flex h-[4.25rem] w-[4.25rem] items-center justify-center rounded-2xl bg-gradient-to-br from-cyan-100/95 via-sky-50 to-indigo-100/90 text-sky-600 shadow-[0_12px_28px_-8px_rgb(14_165_233/0.25)] ring-1 ring-cyan-200/70 dark:from-sky-950 dark:via-indigo-950 dark:to-violet-950 dark:text-cyan-400 dark:ring-cyan-800/45">
              <svg
                className="h-8 w-8"
                fill="none"
                viewBox="0 0 24 24"
                stroke="currentColor"
                strokeWidth="1.5"
                aria-hidden
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  d="M3 16.5v2.25A2.25 2.25 0 0 0 5.25 21h13.5A2.25 2.25 0 0 0 21 18.75V16.5m-13.5-9L12 3m0 0 4.5 4.5M12 3v13.5"
                />
              </svg>
            </span>
            <div className="space-y-2">
              <h2
                id="upload-title"
                className="block text-lg font-semibold tracking-tight text-zinc-900 dark:text-zinc-50 sm:text-xl"
              >
                Перетащите выгрузку Excel
              </h2>
              <p
                id="upload-hint"
                className="mx-auto block max-w-md text-sm leading-relaxed text-zinc-500 dark:text-zinc-400"
              >
                Или нажмите в этой области — разбор только в браузере, данные не
                уходят в сеть
              </p>
            </div>
            <span className={BTN_PRIMARY + " pointer-events-none px-8 py-3"}>
              Выбрать файл .xlsx / .xls
            </span>
            <span className="max-w-md text-xs leading-relaxed text-zinc-500 dark:text-zinc-500">
              До {(MAX_FILE_BYTES / 1024 / 1024).toFixed(0)} МБ · колонки
              сопоставляются с типичными заголовками Битрикс24
            </span>
          </div>
        </section>

        {loadError && (
          <div
            className="flex gap-4 rounded-2xl border border-red-200/90 bg-red-50/95 px-5 py-4 text-sm text-red-950 shadow-md shadow-red-900/5 dark:border-red-900/70 dark:bg-red-950/45 dark:text-red-100"
            role="alert"
          >
            <span
              className="mt-0.5 inline-flex h-8 w-8 shrink-0 items-center justify-center rounded-xl bg-red-100 text-sm font-bold text-red-700 dark:bg-red-900/80 dark:text-red-200"
              aria-hidden
            >
              !
            </span>
            <p className="min-w-0 leading-relaxed">{loadError}</p>
          </div>
        )}

        {tabular && (
          <>
          <div className="grid gap-8 lg:grid-cols-[minmax(0,400px)_1fr] lg:items-start lg:gap-10">
            <aside
              className="glass-surface relative flex flex-col gap-6 overflow-hidden rounded-3xl p-6 lg:sticky lg:top-28 lg:max-h-[calc(100vh-8rem)] lg:overflow-y-auto lg:overscroll-contain"
              aria-label="Настройки отчёта: файл, сущность, графики и фильтры"
            >
              <div
                className="pointer-events-none absolute inset-y-5 left-0 w-1 rounded-r-full bg-gradient-to-b from-cyan-400 via-indigo-500 to-violet-500 opacity-95 shadow-[2px_0_14px_rgba(6,182,212,0.28)] dark:opacity-80"
                aria-hidden
              />
              <div className="relative">
                <ReportSetupSteps
                  entitySelected={importEntity != null}
                  chartsSelected={enabledChartIds.size > 0}
                />
              </div>
              <div className="relative space-y-4 border-b border-zinc-200/60 pb-5 dark:border-zinc-700/70">
                <p className={SECTION_LABEL}>Источник данных</p>
                <div className="space-y-1">
                  <span className="text-xs text-zinc-500 dark:text-zinc-400">
                    Файл
                  </span>
                  <p className="truncate rounded-2xl border border-zinc-200/60 bg-zinc-50/90 px-3.5 py-2.5 text-sm font-medium text-zinc-900 shadow-[inset_0_1px_0_rgb(255_255_255/0.7)] dark:border-zinc-700/70 dark:bg-zinc-800/70 dark:text-zinc-100">
                    {fileName}
                  </p>
                </div>
                {companySummaryKpi && (
                  <div
                    className="grid gap-2 rounded-xl border border-indigo-200/80 bg-indigo-50/75 px-3.5 py-3 text-sm dark:border-indigo-800/50 dark:bg-indigo-950/35"
                    role="region"
                    aria-label="Сводка по компаниям"
                  >
                    <div className="flex items-baseline justify-between gap-4">
                      <span className="text-zinc-600 dark:text-zinc-400">
                        Всего компаний
                      </span>
                      <span className="tabular-nums text-base font-semibold text-zinc-900 dark:text-zinc-50">
                        {companySummaryKpi.total.toLocaleString("ru-RU")}
                      </span>
                    </div>
                    <div className="flex items-baseline justify-between gap-4 border-t border-indigo-200/70 pt-2 dark:border-indigo-800/45">
                      <span
                        className="text-zinc-600 dark:text-zinc-400"
                        title="Сколько разных менеджеров в колонке «Ответственный» по строкам после фильтров (пустые ячейки не считаются)."
                      >
                        Активных ответственных
                      </span>
                      <span className="tabular-nums text-base font-semibold text-zinc-900 dark:text-zinc-50">
                        {companySummaryKpi.activeResponsiblesCount != null
                          ? companySummaryKpi.activeResponsiblesCount.toLocaleString(
                              "ru-RU",
                            )
                          : "—"}
                      </span>
                    </div>
                    <div className="flex items-baseline justify-between gap-4">
                      <span
                        className="text-zinc-600 dark:text-zinc-400"
                        title="Минимальная дата в колонке создания по строкам после фильтров (при неоднозначных датах — среди успешно разобранных значений)."
                      >
                        Дата первой записи
                      </span>
                      <span className="shrink-0 text-right text-sm font-semibold tabular-nums text-zinc-900 dark:text-zinc-50">
                        {companySummaryKpi.firstRecordDate ?? "—"}
                      </span>
                    </div>
                    <div className="flex items-baseline justify-between gap-4">
                      <span
                        className="text-zinc-600 dark:text-zinc-400"
                        title="Максимальная дата в колонке создания по строкам после фильтров."
                      >
                        Дата последней записи
                      </span>
                      <span className="shrink-0 text-right text-sm font-semibold tabular-nums text-zinc-900 dark:text-zinc-50">
                        {companySummaryKpi.lastRecordDate ?? "—"}
                      </span>
                    </div>
                    <div className="flex items-baseline justify-between gap-4 border-t border-indigo-200/70 pt-2 dark:border-indigo-800/45">
                      <span
                        className="text-zinc-600 dark:text-zinc-400"
                        title="Строк после фильтров, делённое на число календарных месяцев от минимальной до максимальной даты в колонке даты создания (как в выгрузке Битрикс24)."
                      >
                        Среднее в месяц
                      </span>
                      <span className="tabular-nums text-base font-semibold text-zinc-900 dark:text-zinc-50">
                        {companySummaryKpi.avgPerMonth != null
                          ? companySummaryKpi.avgPerMonth.toLocaleString(
                              "ru-RU",
                              {
                                maximumFractionDigits: 1,
                                minimumFractionDigits: 0,
                              },
                            )
                          : "—"}
                      </span>
                    </div>
                    <div className="flex items-baseline justify-between gap-4">
                      <span
                        className="text-zinc-600 dark:text-zinc-400"
                        title="Средний помесячный прирост в % к предыдущему календарному месяцу по числу компаний с разобранной датой создания; пустые месяцы в диапазоне считаются как 0. Нужно минимум два календарных месяца в периоде."
                      >
                        Темп роста компаний
                      </span>
                      <span className="shrink-0 text-right text-sm font-semibold tabular-nums text-zinc-900 dark:text-zinc-50">
                        {companySummaryKpi.growthRateAvgPercent != null
                          ? `${companySummaryKpi.growthRateAvgPercent.toLocaleString("ru-RU", {
                              maximumFractionDigits: 1,
                              minimumFractionDigits: 0,
                            })}%`
                          : "—"}
                      </span>
                    </div>
                    {companySummaryKpi.total > 0 &&
                      companySummaryKpi.avgPerMonth == null &&
                      companySummaryKpi.hasDateColumn && (
                        <p className="text-[11px] leading-snug text-violet-800/90 dark:text-violet-200/90">
                          Не удалось разобрать даты в колонке создания — среднее
                          недоступно.
                        </p>
                      )}
                    {!companySummaryKpi.hasResponsibleColumn && (
                      <p className="text-[11px] leading-snug text-violet-800/90 dark:text-violet-200/90">
                        Нет колонки «Ответственный» — число активных менеджеров
                        недоступно.
                      </p>
                    )}
                    {!companySummaryKpi.hasDateColumn && (
                      <p className="text-[11px] leading-snug text-violet-800/90 dark:text-violet-200/90">
                        Нет колонки «Дата создания» — добавьте её в выгрузку для
                        расчёта среднего в месяц.
                      </p>
                    )}
                  </div>
                )}
              </div>

              <div className="relative space-y-5">
                <fieldset className="space-y-3">
                  <legend
                    id="entity-pick-legend"
                    className={`${SECTION_LABEL} mb-0`}
                  >
                    Сущность выгрузки
                  </legend>
                  <p className="text-xs leading-relaxed text-zinc-600 dark:text-zinc-400">
                    После загрузки файла раздел и графики подбираются по колонкам:
                    включаются только те пресеты, для которых хватает полей. Раздел
                    выбирается там, где таких графиков больше всего — при
                    необходимости смените вручную и отметьте другие графики. Для
                    отчёта по нескольким разделам переключайте сущность и включайте
                    нужные пункты в каждом блоке.
                  </p>
                  <div
                    className="grid grid-cols-1 gap-2.5 sm:grid-cols-2"
                    role="radiogroup"
                    aria-labelledby="entity-pick-legend"
                  >
                    {ENTITY_BLOCKS.map((b) => {
                      const selected = importEntity === b.id;
                      const chrome = ENTITY_CHART_CARD[b.id];
                      return (
                        <button
                          key={b.id}
                          type="button"
                          role="radio"
                          aria-checked={selected}
                          className={[
                            "flex w-full flex-col rounded-2xl border p-3 text-left transition-all duration-200",
                            selected
                              ? ENTITY_PICKER_SELECTED[b.id]
                              : ENTITY_PICKER_IDLE,
                          ].join(" ")}
                          onClick={() => applyEntitySelection(b.id)}
                        >
                          <div className="flex items-start gap-3">
                            <span
                              className={`${chrome.badge} !h-11 !min-h-[2.75rem] !w-11 !min-w-[2.75rem] shrink-0`}
                              aria-hidden
                            >
                              {chrome.letter}
                            </span>
                            <div className="min-w-0 flex-1 pt-0.5">
                              <span className="block text-sm font-bold tracking-tight text-zinc-900 dark:text-zinc-50">
                                {b.label}
                              </span>
                              <span className="mt-1 line-clamp-2 text-[11px] leading-snug text-zinc-500 dark:text-zinc-400">
                                {b.description}
                              </span>
                            </div>
                          </div>
                        </button>
                      );
                    })}
                  </div>
                  {importEntity != null && (
                    <button
                      type="button"
                      className="text-xs font-medium text-zinc-500 underline decoration-zinc-300/90 underline-offset-[3px] transition-colors hover:text-zinc-800 dark:text-zinc-400 dark:decoration-zinc-600 dark:hover:text-zinc-200"
                      onClick={() => applyEntitySelection(null)}
                    >
                      Снять выбор сущности
                    </button>
                  )}
                </fieldset>

                {importEntity ? (
                  (() => {
                    const block = ENTITY_BLOCKS.find(
                      (b) => b.id === importEntity,
                    )!;
                    const chrome = ENTITY_CHART_CARD[block.id];
                    const chartRows = chartsForEntity(importEntity)
                      .map((spec) => {
                        const resolved = spec.resolve(tabular.columns);
                        return { spec, resolved };
                      })
                      .sort((a, b) => {
                        if (a.resolved.ok === b.resolved.ok) {
                          return 0;
                        }
                        return a.resolved.ok ? -1 : 1;
                      });
                    return (
                      <div
                        className={`overflow-hidden rounded-2xl border border-zinc-200/65 bg-gradient-to-b from-white/98 to-zinc-50/35 shadow-[0_12px_32px_-20px_rgb(15_23_42/0.15)] dark:border-zinc-700/75 dark:from-zinc-900/92 dark:to-zinc-950/55 dark:shadow-[0_16px_40px_-20px_rgb(0_0_0/0.45)] ${ENTITY_CARD_BORDER[block.id]}`}
                      >
                        <div className="border-b border-zinc-200/70 bg-gradient-to-r from-zinc-50/95 to-transparent px-3.5 py-3 dark:border-zinc-800/80 dark:from-zinc-800/35">
                          <div className="flex items-center gap-3">
                            <span
                              className={`${chrome.badge} !h-9 !w-9 !text-[10px]`}
                              aria-hidden
                            >
                              {chrome.letter}
                            </span>
                            <div className="min-w-0 flex-1">
                              <h3 className="text-sm font-bold tracking-tight text-zinc-900 dark:text-zinc-50">
                                {block.label}
                              </h3>
                              <p className="line-clamp-2 text-[11px] leading-snug text-zinc-500 dark:text-zinc-400">
                                {block.description}
                              </p>
                            </div>
                          </div>
                        </div>
                        <div className="p-3 sm:p-3.5">
                          <p
                            className={`${SECTION_LABEL} mb-3`}
                            id="chart-pick-label"
                          >
                            Графики в отчёт
                          </p>
                          <ul
                            className="flex max-h-[min(420px,52vh)] flex-col gap-2 overflow-y-auto overscroll-contain pr-0.5 [scrollbar-width:thin]"
                            role="group"
                            aria-labelledby="chart-pick-label"
                          >
                            {chartRows.map(({ spec, resolved }) => {
                              const on = enabledChartIds.has(spec.id);
                              const available = resolved.ok;
                              const cbId = `chart-cb-${spec.id}`;
                              return (
                                <li key={spec.id}>
                                  <label
                                    htmlFor={cbId}
                                    className={[
                                      "group relative flex gap-3 overflow-hidden rounded-xl border p-3 transition-[border-color,box-shadow,background-color,transform,opacity] duration-200 sm:p-3.5",
                                      !available
                                        ? "cursor-not-allowed border-zinc-200/65 bg-zinc-100/75 opacity-80 dark:border-zinc-700/65 dark:bg-zinc-900/45"
                                        : on
                                        ? ENTITY_CHART_ROW_SELECTED[
                                            importEntity
                                          ]
                                        : "cursor-pointer border-zinc-200/75 bg-white/90 hover:border-zinc-300/90 hover:bg-white hover:shadow-sm active:scale-[0.995] dark:border-zinc-700/75 dark:bg-zinc-900/50 dark:hover:border-zinc-600 dark:hover:bg-zinc-900/68",
                                    ].join(" ")}
                                  >
                                    <input
                                      id={cbId}
                                      type="checkbox"
                                      className={[
                                        "mt-0.5 size-[1.125rem] shrink-0 rounded-md border-zinc-300 bg-white focus:outline-none focus-visible:ring-2 focus-visible:ring-offset-2 focus-visible:ring-offset-white dark:border-zinc-600 dark:bg-zinc-800 dark:focus-visible:ring-offset-zinc-900",
                                        ENTITY_CHECKBOX_ACCENT[importEntity],
                                      ].join(" ")}
                                      checked={available && on}
                                      disabled={!available}
                                      onChange={(e) =>
                                        setChartEnabled(
                                          spec.id,
                                          e.target.checked,
                                        )
                                      }
                                    />
                                    <div className="min-w-0 flex-1">
                                      <div className="flex flex-wrap items-start justify-between gap-2">
                                        <span className="text-sm font-semibold leading-snug text-zinc-900 dark:text-zinc-100">
                                          {spec.title}
                                        </span>
                                        {resolved.ok ? (
                                          <span className="shrink-0 rounded-full bg-emerald-500/12 px-2 py-0.5 text-[10px] font-semibold uppercase tracking-wide text-emerald-800 dark:bg-emerald-500/20 dark:text-emerald-200">
                                            Ок
                                          </span>
                                        ) : (
                                          <span className="shrink-0 rounded-full bg-amber-500/12 px-2 py-0.5 text-[10px] font-semibold uppercase tracking-wide text-amber-900 dark:bg-amber-500/18 dark:text-amber-100">
                                            Нет данных
                                          </span>
                                        )}
                                      </div>
                                      <p className="mt-1.5 line-clamp-2 text-xs leading-snug text-zinc-600 dark:text-zinc-400">
                                        {spec.description}
                                      </p>
                                      {!resolved.ok && (
                                        <p className="mt-2 text-[11px] leading-snug text-amber-800/95 dark:text-amber-200/95">
                                          {activationHintFromResolveError(
                                            resolved.error,
                                          )}
                                        </p>
                                      )}
                                    </div>
                                  </label>
                                </li>
                              );
                            })}
                          </ul>
                          <p className="mt-3 text-[11px] leading-snug text-zinc-500 dark:text-zinc-400">
                            Для отчёта по нескольким разделам смените сущность
                            выше и отметьте графики в каждом блоке.
                          </p>
                          {chartRows.every(
                            ({ spec, resolved }) =>
                              !resolved.ok || !enabledChartIds.has(spec.id),
                          ) && (
                            <p className="mt-2 rounded-lg bg-zinc-100/90 px-2.5 py-2 text-center text-xs italic text-zinc-600 dark:bg-zinc-800/80 dark:text-zinc-400">
                              Ни одного графика не выбрано — отметьте хотя бы
                              один пункт выше.
                            </p>
                          )}
                        </div>
                      </div>
                    );
                  })()
                ) : (
                  <div className="rounded-2xl border border-dashed border-zinc-300/85 bg-gradient-to-br from-zinc-50/90 to-white/60 px-5 py-8 text-center dark:border-zinc-600/85 dark:from-zinc-900/40 dark:to-zinc-950/35">
                    <div className="mx-auto mb-3 flex h-12 w-12 items-center justify-center rounded-2xl bg-zinc-200/90 text-zinc-500 shadow-inner dark:bg-zinc-800 dark:text-zinc-400">
                      <svg
                        className="h-6 w-6"
                        fill="none"
                        viewBox="0 0 24 24"
                        stroke="currentColor"
                        strokeWidth="1.5"
                        aria-hidden
                      >
                        <path
                          strokeLinecap="round"
                          strokeLinejoin="round"
                          d="M3.75 6A2.25 2.25 0 0 1 6 3.75h2.25A2.25 2.25 0 0 1 10.5 6v2.25a2.25 2.25 0 0 1-2.25 2.25H6a2.25 2.25 0 0 1-2.25-2.25V6ZM9.75 6A2.25 2.25 0 0 1 12 3.75h2.25A2.25 2.25 0 0 1 16.5 6v2.25a2.25 2.25 0 0 1-2.25 2.25H12a2.25 2.25 0 0 1-2.25-2.25V6ZM15.75 6A2.25 2.25 0 0 1 18 3.75h2.25A2.25 2.25 0 0 1 22.5 6v2.25a2.25 2.25 0 0 1-2.25 2.25H18a2.25 2.25 0 0 1-2.25-2.25V6ZM3.75 15.75A2.25 2.25 0 0 1 6 13.5h2.25a2.25 2.25 0 0 1 2.25 2.25V18a2.25 2.25 0 0 1-2.25 2.25H6A2.25 2.25 0 0 1 3.75 18v-2.25ZM9.75 15.75A2.25 2.25 0 0 1 12 13.5h2.25a2.25 2.25 0 0 1 2.25 2.25V18a2.25 2.25 0 0 1-2.25 2.25H12a2.25 2.25 0 0 1-2.25-2.25v-2.25ZM15.75 15.75A2.25 2.25 0 0 1 18 13.5h2.25a2.25 2.25 0 0 1 2.25 2.25V18a2.25 2.25 0 0 1-2.25 2.25H18a2.25 2.25 0 0 1-2.25-2.25v-2.25Z"
                        />
                      </svg>
                    </div>
                    <p className="text-sm font-semibold text-zinc-800 dark:text-zinc-200">
                      Выберите сущность выше
                    </p>
                    <p className="mx-auto mt-1 max-w-xs text-xs leading-relaxed text-zinc-500 dark:text-zinc-400">
                      Здесь появится список графиков, которые можно включить в
                      Excel-отчёт.
                    </p>
                  </div>
                )}
              </div>

              <FilterBlock
                tabular={tabular}
                filters={globalFilters}
                onChange={setGlobalFilters}
              />

              <div className="space-y-2">
                <label
                  className={`block ${SECTION_LABEL}`}
                  htmlFor="chart-surface-select"
                >
                  Фон области графиков
                </label>
                <p className="text-[11px] leading-snug text-zinc-500 dark:text-zinc-400">
                  Цвет подложки под диаграммами. Сохраняется в этом браузере.
                </p>
                <select
                  id="chart-surface-select"
                  className={INPUT_FIELD}
                  value={chartSurface}
                  onChange={(e) => {
                    const v = e.target.value;
                    if (isChartSurfaceId(v)) {
                      setChartSurface(v);
                    }
                  }}
                >
                  {CHART_SURFACE_OPTIONS.map((o) => (
                    <option key={o.id} value={o.id}>
                      {o.label}
                    </option>
                  ))}
                </select>
              </div>

              <div className="space-y-2">
                <label
                  className={`block ${SECTION_LABEL}`}
                  htmlFor="export-theme-select"
                >
                  Тема графиков в Excel
                </label>
                <p className="text-[11px] leading-snug text-zinc-500 dark:text-zinc-400">
                  Влияет на палитру и оформление всех диаграмм в выгрузке.
                </p>
                <select
                  id="export-theme-select"
                  className={INPUT_FIELD}
                  value={exportTheme}
                  onChange={(e) => {
                    const v = e.target.value as ExcelExportTheme;
                    if (EXPORT_THEME_OPTIONS.some((o) => o.id === v)) {
                      setExportTheme(v);
                    }
                  }}
                >
                  {EXPORT_THEME_OPTIONS.map((o) => (
                    <option key={o.id} value={o.id}>
                      {o.label}
                    </option>
                  ))}
                </select>
              </div>

              <div className="relative rounded-2xl border border-zinc-200/65 bg-zinc-50/90 px-3.5 py-3 text-xs text-zinc-600 shadow-[inset_0_1px_0_rgb(255_255_255/0.65)] dark:border-zinc-700/80 dark:bg-zinc-800/55 dark:text-zinc-400 dark:shadow-none">
                <p>
                  Строк в файле:{" "}
                  <strong className="text-zinc-900 dark:text-zinc-200">
                    {tabular.rows.length}
                  </strong>
                </p>
                <p className="mt-1">
                  После фильтров:{" "}
                  <strong className="text-zinc-900 dark:text-zinc-200">
                    {filteredRows.length}
                  </strong>
                </p>
                {tabular.rows.length >= 50000 && (
                  <span className="mt-1 block text-amber-700 dark:text-amber-300">
                    Показаны первые 50 000 строк.
                  </span>
                )}
              </div>
            </aside>

            <div className="flex min-w-0 flex-col gap-8">
              <div className="flex flex-col gap-4 rounded-3xl border border-zinc-200/70 bg-gradient-to-br from-white/95 via-white/90 to-zinc-50/40 p-5 shadow-[0_20px_48px_-26px_rgb(15_23_42/0.14)] backdrop-blur-sm dark:border-zinc-800/80 dark:from-zinc-900/92 dark:via-zinc-900/85 dark:to-zinc-950/70 dark:shadow-[0_24px_56px_-24px_rgb(0_0_0/0.5)] sm:flex-row sm:items-center sm:justify-between sm:p-6">
                <div className="space-y-1.5">
                  <p className={`${SECTION_LABEL} sm:hidden`}>Рабочая область</p>
                  <h2 className="text-base font-semibold tracking-tight text-zinc-900 dark:text-zinc-50 sm:text-lg">
                    Просмотр и экспорт
                  </h2>
                  <p className="max-w-md text-xs leading-relaxed text-zinc-500 dark:text-zinc-400">
                    Графики по отмеченным отчётам и превью таблицы. В Excel —
                    первая страница с общей информацией и по одному листу на
                    каждый график (таблица и цветная диаграмма по текущим данным
                    и фильтрам).
                  </p>
                  <p className="mt-2 max-w-md text-[11px] font-medium leading-snug text-emerald-800/90 dark:text-emerald-300/95">
                    Кнопка выгрузки закреплена внизу экрана — можно не
                    возвращаться к началу страницы.
                  </p>
                </div>
                <div className="flex flex-col items-stretch gap-3 sm:items-end">
                  <button
                    type="button"
                    className={
                      (exportEnabled ? BTN_EXPORT : BTN_GHOST) +
                      " w-full sm:w-auto disabled:pointer-events-none disabled:opacity-45"
                    }
                    disabled={!exportEnabled}
                    title={exportButtonTitle}
                    onClick={onExportChartsExcel}
                  >
                    <svg
                      className={
                        exportEnabled ? "h-4 w-4 text-white/90" : "h-4 w-4 text-zinc-500"
                      }
                      fill="none"
                      viewBox="0 0 24 24"
                      stroke="currentColor"
                      strokeWidth="1.5"
                      aria-hidden
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        d="M3 16.5v2.25A2.25 2.25 0 0 0 5.25 21h13.5A2.25 2.25 0 0 0 21 18.75V16.5M16.5 12 12 16.5m0 0L7.5 12m4.5 4.5V3"
                      />
                    </svg>
                    Скачать в Excel
                  </button>
                  {exportMessage && (
                    <p
                      className={
                        exportMessage.startsWith("Файл сохранён")
                          ? "max-w-md rounded-xl border border-emerald-200/80 bg-emerald-50/90 px-3 py-2 text-right text-xs leading-relaxed text-emerald-900 dark:border-emerald-900/50 dark:bg-emerald-950/40 dark:text-emerald-100"
                          : "max-w-md rounded-xl border border-red-200/80 bg-red-50/90 px-3 py-2 text-right text-xs leading-relaxed text-red-900 dark:border-red-900/50 dark:bg-red-950/40 dark:text-red-100"
                      }
                      role="status"
                    >
                      {exportMessage}
                    </p>
                  )}
                </div>
              </div>

              {tabular.rows.length === 0 ? (
                <div className="rounded-3xl border border-dashed border-zinc-300/90 bg-white/70 py-20 text-center dark:border-zinc-600 dark:bg-zinc-900/40">
                  <div className="mx-auto mb-3 flex h-12 w-12 items-center justify-center rounded-2xl bg-zinc-100 text-zinc-400 dark:bg-zinc-800 dark:text-zinc-500">
                    <svg
                      className="h-6 w-6"
                      fill="none"
                      viewBox="0 0 24 24"
                      stroke="currentColor"
                      strokeWidth="1.5"
                      aria-hidden
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        d="M19.5 14.25v-2.625a3.375 3.375 0 0 0-3.375-3.375h-1.5A1.125 1.125 0 0 1 13.5 7.125v-1.5a3.375 3.375 0 0 0-3.375-3.375H8.25m0 12.75h7.5m-7.5 3H12M10.5 2.25H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 0 0-9-9Z"
                      />
                    </svg>
                  </div>
                  <p className="text-sm font-medium text-zinc-600 dark:text-zinc-300">
                    В файле нет строк с данными
                  </p>
                  <p className="mx-auto mt-1 max-w-sm text-xs text-zinc-500 dark:text-zinc-400">
                    Загрузите другой файл.
                  </p>
                </div>
              ) : filteredRows.length === 0 ? (
                <div
                  className="flex gap-3 rounded-2xl border border-amber-200/90 bg-amber-50/95 px-4 py-3.5 text-sm text-amber-950 shadow-sm dark:border-amber-900/70 dark:bg-amber-950/40 dark:text-amber-100"
                  role="status"
                >
                  <span
                    className="mt-0.5 inline-flex h-5 w-5 shrink-0 items-center justify-center rounded-full bg-amber-200 text-xs font-bold text-amber-900 dark:bg-amber-900 dark:text-amber-200"
                    aria-hidden
                  >
                    ⚠
                  </span>
                  <p className="min-w-0 leading-relaxed">
                    Текущие фильтры отсекли все строки. Ослабьте условия в
                    боковой панели или сбросьте их — графики и превью покажут
                    данные снова.
                  </p>
                </div>
              ) : (
                ENTITY_BLOCKS.map((block) => {
                  const specs = chartsForEntity(block.id).filter((s) =>
                    enabledChartIds.has(s.id),
                  );
                  if (specs.length === 0) {
                    return null;
                  }
                  return (
                    <section
                      key={block.id}
                      className="space-y-5 overflow-hidden rounded-3xl border border-zinc-200/70 bg-gradient-to-b from-white/92 to-zinc-50/35 p-6 shadow-[0_20px_50px_-28px_rgb(15_23_42/0.11)] backdrop-blur-[2px] dark:border-zinc-800/85 dark:from-zinc-900/88 dark:to-zinc-950/55 dark:shadow-[0_24px_56px_-26px_rgb(0_0_0/0.48)]"
                    >
                      <div className="flex flex-wrap items-center gap-3 border-b border-zinc-200/55 pb-4 dark:border-zinc-800/80">
                        <span
                          className={`inline-flex items-center rounded-full px-3 py-1 text-xs font-semibold ring-1 ring-inset ${ENTITY_SECTION_BADGE[block.id]}`}
                        >
                          {block.label}
                        </span>
                        <h2 className="text-base font-semibold tracking-tight text-zinc-900 dark:text-zinc-100">
                          Графики раздела
                        </h2>
                      </div>
                      <div className="flex flex-col gap-6">
                        {specs.map((spec) => (
                          <PredefinedChartCard
                            key={spec.id}
                            specTitle={spec.title}
                            tabular={tabular}
                            globalFilters={globalFilters}
                            chartSurface={chartSurface}
                            spec={spec}
                          />
                        ))}
                      </div>
                    </section>
                  );
                })
              )}

              {enabledChartIds.size === 0 && filteredRows.length > 0 && (
                <div className="rounded-3xl border border-dashed border-zinc-300/85 bg-gradient-to-b from-white/80 to-zinc-50/40 py-20 text-center dark:border-zinc-600 dark:from-zinc-900/50 dark:to-zinc-950/30">
                  <div className="mx-auto mb-4 flex h-14 w-14 items-center justify-center rounded-2xl bg-zinc-100 text-zinc-400 dark:bg-zinc-800 dark:text-zinc-500">
                    <svg
                      className="h-7 w-7"
                      fill="none"
                      viewBox="0 0 24 24"
                      stroke="currentColor"
                      strokeWidth="1.5"
                      aria-hidden
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        d="M3 13.125C3 12.504 3.504 12 4.125 12h2.25c.621 0 1.125.504 1.125 1.125v6.75C7.5 20.496 6.996 21 6.375 21h-2.25A1.125 1.125 0 0 1 3 19.875v-6.75ZM9.75 8.625c0-.621.504-1.125 1.125-1.125h2.25c.621 0 1.125.504 1.125 1.125v11.25c0 .621-.504 1.125-1.125 1.125h-2.25a1.125 1.125 0 0 1-1.125-1.125V8.625ZM16.5 4.125c0-.621.504-1.125 1.125-1.125h2.25C20.496 3 21 3.504 21 4.125v15.75c0 .621-.504 1.125-1.125 1.125h-2.25a1.125 1.125 0 0 1-1.125-1.125V4.125Z"
                      />
                    </svg>
                  </div>
                  <p className="mx-auto max-w-sm text-sm font-medium leading-relaxed text-zinc-700 dark:text-zinc-300">
                    Выберите графики в боковой панели
                  </p>
                  <p className="mx-auto mt-1 max-w-xs text-xs leading-relaxed text-zinc-500 dark:text-zinc-400">
                    Сначала сущность выгрузки, затем отметьте графики
                    галочками — визуализации появятся здесь.
                  </p>
                </div>
              )}

              <div className="overflow-hidden rounded-3xl border border-zinc-200/70 bg-white/95 shadow-[0_16px_42px_-26px_rgb(15_23_42/0.11)] backdrop-blur-sm dark:border-zinc-800/85 dark:bg-zinc-900/72 dark:shadow-[0_20px_48px_-26px_rgb(0_0_0/0.42)]">
                <div className="flex flex-col gap-1 border-b border-zinc-200/65 bg-gradient-to-r from-zinc-50/98 to-white px-4 py-3.5 sm:flex-row sm:items-center sm:justify-between dark:border-zinc-800/80 dark:from-zinc-800/65 dark:to-zinc-900/88">
                  <h3 className="text-sm font-semibold tracking-tight text-zinc-900 dark:text-zinc-100">
                    Превью данных
                  </h3>
                  <span className="text-xs text-zinc-500 dark:text-zinc-400">
                    До {PREVIEW_ROWS} строк · после фильтров:{" "}
                    <strong className="font-semibold text-zinc-700 dark:text-zinc-300">
                      {filteredRows.length}
                    </strong>
                  </span>
                </div>
                <div className="max-h-[min(520px,55vh)] overflow-auto">
                  <table className="min-w-full text-left text-xs">
                    <thead className="sticky top-0 z-10 shadow-[0_1px_0_rgb(0_0_0/0.06)] dark:shadow-[0_1px_0_rgb(255_255_255/0.06)]">
                      <tr className="border-b border-zinc-200/90 bg-zinc-50/98 backdrop-blur-sm dark:border-zinc-700 dark:bg-zinc-800/95">
                        {previewColumns.map((c) => (
                          <th
                            key={c.key}
                            className="whitespace-nowrap px-3.5 py-3 font-semibold text-zinc-700 dark:text-zinc-200"
                          >
                            {c.header}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-zinc-100 dark:divide-zinc-800/80">
                      {filteredRows.length === 0 ? (
                        <tr>
                          <td
                            colSpan={Math.max(previewColumns.length, 1)}
                            className="px-3 py-10 text-center text-sm text-zinc-500 dark:text-zinc-400"
                          >
                            Нет строк, подходящих под фильтры.
                          </td>
                        </tr>
                      ) : (
                        filteredRows
                          .slice(0, PREVIEW_ROWS)
                          .map((row, ri) => (
                            <tr
                              key={ri}
                              className="bg-white odd:bg-zinc-50/35 transition-colors hover:bg-sky-50/45 dark:bg-transparent dark:odd:bg-zinc-800/28 dark:hover:bg-sky-950/20"
                            >
                              {previewColumns.map((c) => (
                                <td
                                  key={c.key}
                                  className="max-w-56 truncate px-3.5 py-2 text-zinc-700 dark:text-zinc-300"
                                >
                                  {formatCell(row[c.key])}
                                </td>
                              ))}
                            </tr>
                          ))
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          </div>
          <ExportDock
            exportEnabled={exportEnabled}
            exportButtonTitle={exportButtonTitle}
            onExport={onExportChartsExcel}
            chartCount={enabledChartIds.size}
            filteredRowCount={filteredRows.length}
          />
          </>
        )}
      </main>
    </div>
  );
}

function PredefinedChartCard({
  spec,
  specTitle,
  tabular,
  globalFilters,
  chartSurface,
}: {
  spec: PredefinedChartSpec;
  specTitle: string;
  tabular: TabularData;
  globalFilters: ChartFilter[];
  chartSurface: ChartSurfaceId;
}) {
  const resolved = spec.resolve(tabular.columns);
  const config =
    resolved.ok
      ? normalizeConfigForTabular(
          tabular,
          mergeGlobalFilters(resolved.config, globalFilters),
        )
      : null;

  if (!resolved.ok || !config) {
    return (
      <section className={ENTITY_CHART_ERR[spec.entity]}>
        <h2 className="text-sm font-semibold text-amber-950 dark:text-amber-100">
          {specTitle}
        </h2>
        <p className="mt-1.5 text-sm leading-relaxed text-amber-900/92 dark:text-amber-200/92">
          {resolved.ok ? "Нет конфигурации." : resolved.error}
        </p>
      </section>
    );
  }

  const chrome = ENTITY_CHART_CARD[spec.entity];

  return (
    <section className={chrome.wrap}>
      <div className={chrome.head}>
        <span className={chrome.badge} aria-hidden>
          {chrome.letter}
        </span>
        <h2 className={chrome.title}>{specTitle}</h2>
      </div>
      <ChartCanvas
        tabular={tabular}
        config={config}
        entity={spec.entity}
        chartId={spec.id}
        chartSurface={chartSurface}
      />
    </section>
  );
}

function ChartCanvas({
  tabular,
  config,
  entity,
  chartId,
  chartSurface,
}: {
  tabular: TabularData;
  config: ChartConfig;
  entity?: EntityBlockId;
  chartId?: string;
  chartSurface: ChartSurfaceId;
}) {
  const xMetaForChart = config.xKey
    ? tabular.columns.find((c) => c.key === config.xKey)
    : undefined;
  const isTemporalChart =
    Boolean(config.xKey) &&
    xMetaForChart?.inferredType === "date" &&
    config.chartType !== "pie";

  const [periodOverride, setPeriodOverride] = useState<DateGranularity | null>(
    null,
  );

  const presetDateGranularity: DateGranularity =
    config.dateGranularity ?? "day";
  const effectiveDateGranularity =
    periodOverride ?? presetDateGranularity;

  const configForRows = useMemo(
    (): ChartConfig => ({
      ...config,
      dateGranularity: effectiveDateGranularity,
    }),
    [config, effectiveDateGranularity],
  );

  const { data: chartRows, warnings } = useMemo(
    () => buildAggregatedChartRows(tabular.rows, configForRows, tabular.columns),
    [tabular.rows, tabular.columns, configForRows],
  );

  const e = entity ?? "leads";
  const palette = chartPaletteForEntity(entity);
  const axes = chartUsesDarkSurface(chartSurface)
    ? CHART_AXES_DARK_SURFACE
    : chartAxesTheme(entity);
  const tickStyle = { fontSize: 11, fill: axes.tick };
  const tooltipContentStyle = tooltipChrome(entity);

  const legendStyle = {
    paddingTop: 18,
    fontSize: "12px",
    fontWeight: 500,
    color: axes.tick,
  };

  const chartShellClassResolved = chartShellClass(e, chartSurface);
  const isCompanyTypePie =
    chartId === "companies_count_by_company_type" && config.chartType === "pie";
  const isResponsibleDistributionPie =
    chartId === "companies_count_by_responsible" &&
    config.chartType === "pie";
  const isHasDealPie =
    chartId === "companies_percent_with_deal" &&
    config.chartType === "pie";
  const pieRows = chartRows.map((r) => ({
    name: String(r.name),
    value: Number(r[config.yKeys[0]!] ?? 0),
  }));
  const pieColorFor = useCallback(
    (name: string, index: number) => {
      if (isResponsibleDistributionPie) {
        return COMPANIES_RESPONSIBLE_PIE_COLORS[
          index % COMPANIES_RESPONSIBLE_PIE_COLORS.length
        ]!;
      }
      if (!isHasDealPie) {
        return palette[index % palette.length]!;
      }
      const n = name.trim().toLowerCase();
      if (
        n === "да" ||
        n === "yes" ||
        n === "true" ||
        n === "1" ||
        n.includes("есть")
      ) {
        return "#16a34a";
      }
      if (
        n === "нет" ||
        n === "no" ||
        n === "false" ||
        n === "0" ||
        n.includes("нет")
      ) {
        return "#ef4444";
      }
      return ["#0ea5e9", "#f59e0b", "#a855f7", "#14b8a6"][index % 4]!;
    },
    [isResponsibleDistributionPie, isHasDealPie, palette],
  );

  const yLegendByKey = useMemo(() => {
    const xMeta = config.xKey
      ? tabular.columns.find((c) => c.key === config.xKey)
      : null;
    const periodLabel =
      xMeta?.inferredType === "date"
        ? effectiveDateGranularity === "month"
          ? "месяц"
          : effectiveDateGranularity === "quarter"
            ? "квартал"
            : effectiveDateGranularity === "year"
              ? "год"
              : "день"
        : "категория";
    const m = new Map<string, string>();
    for (const key of config.yKeys) {
      m.set(
        key,
        metricLegendLabel(
          tabular,
          key,
          config.aggregation,
          periodLabel,
          config.cumulative,
        ),
      );
    }
    return m;
  }, [
    tabular,
    config.yKeys,
    config.aggregation,
    config.cumulative,
    config.xKey,
    effectiveDateGranularity,
  ]);

  const pieIssue =
    config.chartType === "pie" && config.yKeys.length !== 1
      ? "Круговая диаграмма требует одной метрики."
      : null;
  const pieTotal =
    config.chartType === "pie"
      ? chartRows.reduce((acc, row) => {
          const v = Number(row[config.yKeys[0]!] ?? 0);
          return acc + (Number.isFinite(v) ? v : 0);
        }, 0)
      : 0;

  const pieLegendPayload = useMemo(
    () =>
      pieRows.map((row, i) => ({
        value: row.name,
        color: pieColorFor(row.name, i),
        payload: { name: row.name, value: row.value },
      })),
    [pieRows, pieColorFor],
  );

  const chartReady =
    config.xKey &&
    config.yKeys.length > 0 &&
    !pieIssue &&
    chartRows.length > 0;

  const manyBars = chartRows.length > 8;
  const veryManyBars = chartRows.length > 12;
  const isMonthlyTimeline =
    effectiveDateGranularity === "month" ||
    effectiveDateGranularity === "quarter";
  const showPointLabels = isMonthlyTimeline && chartRows.length <= 18;
  const areaIsSingleSeries =
    config.chartType === "area" && config.yKeys.length === 1;
  const areaChartTopMargin = showPointLabels ? 42 : 28;
  const renderAreaPointLabel = useCallback(
    (props: { x?: number; y?: number; value?: unknown }) => {
      const { x, y, value } = props;
      if (x == null || y == null) {
        return null;
      }
      const t =
        typeof value === "number"
          ? formatTooltipNumber(value)
          : String(value ?? "");
      const dark = chartUsesDarkSurface(chartSurface);
      return (
        <text
          x={x}
          y={y}
          textAnchor="middle"
          dy={-14}
          fontSize={11}
          fontWeight={600}
        >
          <tspan
            stroke={dark ? "rgba(15,23,42,0.9)" : "rgba(255,255,255,0.96)"}
            strokeWidth={3}
            strokeLinejoin="round"
            paintOrder="stroke"
            fill={dark ? "#f8fafc" : "#0f172a"}
          >
            {t}
          </tspan>
        </text>
      );
    },
    [chartSurface],
  );
  const barBottom =
    chartReady && config.chartType === "bar"
      ? veryManyBars
        ? 60
        : manyBars
          ? 52
          : 14
      : 14;
  const xAxisAngle =
    chartReady && config.chartType === "bar" && manyBars ? -32 : 0;

  const renderTooltip = ({
    active,
    payload,
    label,
  }: TooltipContentProps) => {
    if (!active || !payload?.length) {
      return null;
    }
    const first = payload[0];
    const pieSliceName =
      config.chartType === "pie" &&
      first?.payload != null &&
      typeof first.payload === "object" &&
      first.payload !== null &&
      "name" in first.payload
        ? String((first.payload as { name: unknown }).name)
        : null;
    const headerText =
      pieSliceName != null && pieSliceName !== ""
        ? pieSliceName
        : label != null && String(label) !== ""
          ? String(label)
          : "";
    return (
      <div
        style={{
          ...tooltipContentStyle,
          pointerEvents: "none" as const,
        }}
      >
        {headerText !== "" && (
          <p
            style={{
              fontWeight: 700,
              marginBottom: 8,
              color: "#0f172a",
              margin: "0 0 8px 0",
            }}
          >
            {headerText}
          </p>
        )}
        <ul
          style={{
            margin: 0,
            padding: 0,
            listStyle: "none",
            display: "flex",
            flexDirection: "column",
            gap: 4,
          }}
        >
          {payload.map((e, i) => {
            const v = e.value;
            const text =
              typeof v === "number" && Number.isFinite(v)
                ? formatTooltipNumber(v)
                : v == null
                  ? "—"
                  : String(v);
            const y0 = config.yKeys[0];
            const seriesName =
              config.chartType === "pie" && y0
                ? (yLegendByKey.get(y0) ?? y0)
                : e.name != null && e.name !== ""
                  ? String(e.name)
                  : `Серия ${i + 1}`;
            return (
              <li
                key={i}
                style={{ color: "#334155", margin: 0, fontSize: 13 }}
              >
                <span style={{ color: e.color ?? "#64748b", marginRight: 8 }}>
                  ●
                </span>
                {seriesName}:{" "}
                <strong style={{ fontWeight: 600 }}>{text}</strong>
              </li>
            );
          })}
        </ul>
      </div>
    );
  };

  return (
    <div className="space-y-3">
      {(warnings.length > 0 || pieIssue) && (
        <div className="space-y-1 rounded-xl border border-violet-200/85 bg-violet-50/95 px-3 py-2.5 text-sm text-violet-950 shadow-sm dark:border-violet-800/65 dark:bg-violet-950/40 dark:text-violet-100">
          {pieIssue && <p>{pieIssue}</p>}
          {warnings.map((w) => (
            <p key={w}>{w}</p>
          ))}
        </div>
      )}
      <div
        className={
          config.chartType === "pie"
            ? `${chartShellClassResolved} min-w-0`
            : chartShellClassResolved
        }
      >
        {!chartReady && (
          <div
            className="flex h-[308px] items-center justify-center px-6 text-center text-sm leading-relaxed text-zinc-500 dark:text-zinc-400"
            style={{ color: `${axes.tick}99` }}
          >
            Недостаточно данных для этого графика. Проверьте фильтры и колонки
            в выгрузке.
          </div>
        )}
        {chartReady && config.chartType === "pie" && (
          <div className="flex w-full min-w-0 flex-col items-center justify-center gap-6 px-2 py-3 sm:px-3 lg:flex-row lg:items-center lg:justify-center lg:gap-8">
            <div className="flex shrink-0 items-center justify-center overflow-visible">
              <div className="aspect-square h-[min(340px,calc(100vw-2.75rem))] w-[min(340px,calc(100vw-2.75rem))] max-h-[360px] max-w-[360px] sm:h-[min(360px,calc(100vw-3rem))] sm:w-[min(360px,calc(100vw-3rem))] sm:max-h-[384px] sm:max-w-[384px]">
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart margin={{ top: 12, right: 12, bottom: 12, left: 12 }}>
                    <defs>
                      {pieRows.map((row, i) => {
                        const c = pieColorFor(row.name, i);
                        const id = `pieSlice-${e}-${i}`;
                        if (isResponsibleDistributionPie) {
                          return (
                            <radialGradient
                              key={id}
                              id={id}
                              cx="36%"
                              cy="32%"
                              r="92%"
                            >
                              <stop offset="0%" stopColor={c} stopOpacity={1} />
                              <stop
                                offset="45%"
                                stopColor={c}
                                stopOpacity={0.98}
                              />
                              <stop
                                offset="100%"
                                stopColor={c}
                                stopOpacity={0.94}
                              />
                            </radialGradient>
                          );
                        }
                        return (
                          <radialGradient
                            key={id}
                            id={id}
                            cx="42%"
                            cy="38%"
                            r="78%"
                          >
                            <stop offset="0%" stopColor={c} stopOpacity={1} />
                            <stop
                              offset="55%"
                              stopColor={c}
                              stopOpacity={0.92}
                            />
                            <stop
                              offset="100%"
                              stopColor={c}
                              stopOpacity={0.78}
                            />
                          </radialGradient>
                        );
                      })}
                    </defs>
                    <Pie
                      data={pieRows}
                      dataKey="value"
                      nameKey="name"
                      cx="50%"
                      cy="50%"
                      innerRadius={44}
                      outerRadius={
                        isCompanyTypePie || isHasDealPie ? 118 : 102
                      }
                      paddingAngle={
                        isResponsibleDistributionPie ? 1.2 : 2.4
                      }
                      cornerRadius={5}
                      label={
                        isCompanyTypePie || isHasDealPie
                          ? false
                          : ({
                              cx,
                              cy,
                              name,
                              midAngle,
                              innerRadius: ir,
                              outerRadius: or,
                              percent,
                              value,
                            }) => {
                              if (
                                cx == null ||
                                cy == null ||
                                midAngle == null ||
                                ir == null ||
                                or == null ||
                                percent == null
                              ) {
                                return null;
                              }
                              const RADIAN = Math.PI / 180;
                              const radius =
                                Number(ir) +
                                (Number(or) - Number(ir)) * 0.58;
                              const x =
                                Number(cx) +
                                radius * Math.cos(-midAngle * RADIAN);
                              const y =
                                Number(cy) +
                                radius * Math.sin(-midAngle * RADIAN);
                              const pct = (percent * 100).toFixed(0);
                              const count =
                                typeof value === "number" &&
                                Number.isFinite(value)
                                  ? formatTooltipNumber(value)
                                  : "0";
                              return (
                                <text
                                  x={x}
                                  y={y}
                                  fill={axes.tick}
                                  textAnchor={
                                    x > Number(cx) ? "start" : "end"
                                  }
                                  dominantBaseline="central"
                                  fontSize={11}
                                  fontWeight={700}
                                >
                                  {`${count} · ${pct}%`}
                                </text>
                              );
                            }
                      }
                      labelLine={
                        isCompanyTypePie ||
                        isHasDealPie ||
                        isResponsibleDistributionPie
                          ? false
                          : {
                              stroke: axes.axis,
                              strokeWidth: 1,
                              opacity: 0.75,
                            }
                      }
                    >
                      {chartRows.map((_, i) => (
                        <Cell
                          key={`slice-${i}`}
                          fill={`url(#pieSlice-${e}-${i})`}
                          stroke={
                            isResponsibleDistributionPie
                              ? "rgba(255,255,255,0.98)"
                              : "rgba(255,255,255,0.92)"
                          }
                          strokeWidth={
                            isResponsibleDistributionPie ? 3 : 2.5
                          }
                        />
                      ))}
                    </Pie>
                    <Tooltip content={renderTooltip as never} />
                  </PieChart>
                </ResponsiveContainer>
              </div>
            </div>
            <div
              className="flex w-full max-w-sm flex-col justify-center lg:w-64 lg:max-w-[17rem] lg:shrink-0"
              style={{
                paddingTop: 0,
                color: legendStyle.color,
                fontWeight: legendStyle.fontWeight,
                fontSize: legendStyle.fontSize,
              }}
            >
              {pieLegendListContent(
                pieLegendPayload,
                axes.tick,
                pieTotal,
                { maxHeight: 360, maxWidth: "100%" },
              )}
            </div>
          </div>
        )}
        {chartReady && config.chartType === "bar" && (
          <ResponsiveContainer width="100%" height={384}>
            <BarChart
              data={chartRows}
              margin={{ top: 18, right: 18, left: 6, bottom: barBottom }}
              barCategoryGap="14%"
            >
              <defs>
                {config.yKeys.map((_, i) => {
                  const c = palette[i % palette.length]!;
                  const id = `barGrad-${e}-${i}`;
                  return (
                    <linearGradient
                      key={id}
                      id={id}
                      x1="0"
                      y1="0"
                      x2="0"
                      y2="1"
                    >
                      <stop offset="0%" stopColor={c} stopOpacity={1} />
                      <stop offset="45%" stopColor={c} stopOpacity={0.92} />
                      <stop offset="100%" stopColor={c} stopOpacity={0.58} />
                    </linearGradient>
                  );
                })}
              </defs>
              <CartesianGrid
                strokeDasharray="4 8"
                stroke={axes.grid}
                vertical={false}
              />
              <XAxis
                dataKey="name"
                tick={tickStyle}
                stroke={axes.axis}
                tickLine={false}
                axisLine={{ stroke: axes.axis, strokeOpacity: 0.35 }}
                angle={xAxisAngle}
                textAnchor={xAxisAngle ? "end" : "middle"}
                height={xAxisAngle ? 72 : undefined}
                interval={manyBars ? 0 : "preserveStartEnd"}
                tickMargin={8}
              />
              <YAxis
                tick={tickStyle}
                stroke={axes.axis}
                width={52}
                tickLine={false}
                axisLine={false}
                tickFormatter={(v) =>
                  typeof v === "number"
                    ? v.toLocaleString("ru-RU", { notation: "compact" })
                    : String(v)
                }
              />
              <Tooltip
                content={renderTooltip as never}
                cursor={{ fill: "rgba(99, 102, 241, 0.06)" }}
              />
              <Legend
                iconType="circle"
                iconSize={9}
                wrapperStyle={legendStyle}
              />
              {config.yKeys.map((k, i) => (
                <Bar
                  key={k}
                  dataKey={k}
                  name={yLegendByKey.get(k)}
                  fill={`url(#barGrad-${e}-${i})`}
                  radius={[10, 10, 0, 0]}
                  maxBarSize={58}
                />
              ))}
            </BarChart>
          </ResponsiveContainer>
        )}
        {chartReady && config.chartType === "line" && (
          <ResponsiveContainer width="100%" height={384}>
            <LineChart
              data={chartRows}
              margin={{ top: 18, right: 18, left: 6, bottom: 12 }}
            >
              <CartesianGrid
                strokeDasharray="4 8"
                stroke={axes.grid}
                vertical={false}
              />
              <XAxis
                dataKey="name"
                tick={tickStyle}
                stroke={axes.axis}
                tickLine={false}
                axisLine={{ stroke: axes.axis, strokeOpacity: 0.35 }}
                interval={isMonthlyTimeline ? 0 : "preserveStartEnd"}
                tickMargin={8}
              />
              <YAxis
                tick={tickStyle}
                stroke={axes.axis}
                width={52}
                tickLine={false}
                axisLine={false}
                tickFormatter={(v) =>
                  typeof v === "number"
                    ? v.toLocaleString("ru-RU", { notation: "compact" })
                    : String(v)
                }
              />
              <Tooltip content={renderTooltip as never} />
              <Legend
                iconType="circle"
                iconSize={9}
                wrapperStyle={legendStyle}
              />
              {config.yKeys.map((k, i) => (
                <Line
                  key={k}
                  type="monotone"
                  dataKey={k}
                  name={yLegendByKey.get(k)}
                  stroke={palette[i % palette.length]!}
                  strokeWidth={3}
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  dot={{
                    r: 5,
                    strokeWidth: 2,
                    fill: palette[i % palette.length]!,
                    stroke: "#fff",
                  }}
                  activeDot={{
                    r: 8,
                    strokeWidth: 2,
                    fill: palette[i % palette.length]!,
                    stroke: "#fff",
                  }}
                >
                  {showPointLabels && (
                    <LabelList
                      dataKey={k}
                      position="top"
                      offset={8}
                      fill={axes.tick}
                      fontSize={11}
                      formatter={(v: unknown) =>
                        typeof v === "number" ? formatTooltipNumber(v) : String(v ?? "")
                      }
                    />
                  )}
                </Line>
              ))}
            </LineChart>
          </ResponsiveContainer>
        )}
        {chartReady && config.chartType === "area" && (
          <ResponsiveContainer width="100%" height={400}>
            {areaIsSingleSeries ? (
              <ComposedChart
                data={chartRows}
                margin={{
                  top: areaChartTopMargin,
                  right: 20,
                  left: 8,
                  bottom: 14,
                }}
              >
                <defs>
                  {config.yKeys.map((_, i) => {
                    const c = palette[i % palette.length]!;
                    const id = `areaGrad-${e}-${i}`;
                    return (
                      <linearGradient
                        key={id}
                        id={id}
                        x1="0"
                        y1="0"
                        x2="0"
                        y2="1"
                      >
                        <stop offset="0%" stopColor={c} stopOpacity={0.82} />
                        <stop offset="28%" stopColor={c} stopOpacity={0.42} />
                        <stop offset="62%" stopColor={c} stopOpacity={0.16} />
                        <stop offset="100%" stopColor={c} stopOpacity={0.02} />
                      </linearGradient>
                    );
                  })}
                </defs>
                <CartesianGrid
                  strokeDasharray="3 7"
                  stroke={axes.grid}
                  strokeOpacity={0.55}
                  vertical={false}
                />
                <XAxis
                  dataKey="name"
                  tick={tickStyle}
                  stroke={axes.axis}
                  tickLine={false}
                  axisLine={{ stroke: axes.axis, strokeOpacity: 0.35 }}
                  interval={isMonthlyTimeline ? 0 : "preserveStartEnd"}
                  tickMargin={8}
                />
                <YAxis
                  tick={tickStyle}
                  stroke={axes.axis}
                  width={56}
                  tickLine={false}
                  axisLine={false}
                  tickFormatter={(v) =>
                    typeof v === "number"
                      ? v.toLocaleString("ru-RU", { notation: "compact" })
                      : String(v)
                  }
                />
                <Tooltip content={renderTooltip as never} />
                <Legend
                  iconType="circle"
                  iconSize={9}
                  wrapperStyle={legendStyle}
                />
                {config.yKeys.map((k, i) => {
                  const c = palette[i % palette.length]!;
                  return (
                    <Area
                      key={`area-fill-${k}`}
                      type="monotone"
                      dataKey={k}
                      stroke="none"
                      strokeWidth={0}
                      fill={`url(#areaGrad-${e}-${i})`}
                      fillOpacity={1}
                      tooltipType="none"
                      legendType="none"
                      isAnimationActive="auto"
                    />
                  );
                })}
                {config.yKeys.map((k, i) => {
                  const c = palette[i % palette.length]!;
                  return (
                    <Line
                      key={`area-line-${k}`}
                      type="monotone"
                      dataKey={k}
                      name={yLegendByKey.get(k)}
                      stroke={c}
                      strokeWidth={2.75}
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      dot={{
                        r: 5,
                        strokeWidth: 2,
                        fill: c,
                        stroke: "#fff",
                      }}
                      activeDot={{
                        r: 8,
                        strokeWidth: 2,
                        fill: c,
                        stroke: "#fff",
                      }}
                    >
                      {showPointLabels && (
                        <LabelList
                          dataKey={k}
                          position="top"
                          content={renderAreaPointLabel as never}
                        />
                      )}
                    </Line>
                  );
                })}
              </ComposedChart>
            ) : (
              <AreaChart
                data={chartRows}
                margin={{
                  top: areaChartTopMargin,
                  right: 20,
                  left: 8,
                  bottom: 14,
                }}
              >
                <defs>
                  {config.yKeys.map((_, i) => {
                    const c = palette[i % palette.length]!;
                    const id = `areaGrad-${e}-${i}`;
                    return (
                      <linearGradient
                        key={id}
                        id={id}
                        x1="0"
                        y1="0"
                        x2="0"
                        y2="1"
                      >
                        <stop offset="0%" stopColor={c} stopOpacity={0.78} />
                        <stop offset="30%" stopColor={c} stopOpacity={0.38} />
                        <stop offset="65%" stopColor={c} stopOpacity={0.14} />
                        <stop offset="100%" stopColor={c} stopOpacity={0.025} />
                      </linearGradient>
                    );
                  })}
                </defs>
                <CartesianGrid
                  strokeDasharray="3 7"
                  stroke={axes.grid}
                  strokeOpacity={0.55}
                  vertical={false}
                />
                <XAxis
                  dataKey="name"
                  tick={tickStyle}
                  stroke={axes.axis}
                  tickLine={false}
                  axisLine={{ stroke: axes.axis, strokeOpacity: 0.35 }}
                  interval={isMonthlyTimeline ? 0 : "preserveStartEnd"}
                  tickMargin={8}
                />
                <YAxis
                  tick={tickStyle}
                  stroke={axes.axis}
                  width={56}
                  tickLine={false}
                  axisLine={false}
                  tickFormatter={(v) =>
                    typeof v === "number"
                      ? v.toLocaleString("ru-RU", { notation: "compact" })
                      : String(v)
                  }
                />
                <Tooltip content={renderTooltip as never} />
                <Legend
                  iconType="circle"
                  iconSize={9}
                  wrapperStyle={legendStyle}
                />
                {config.yKeys.map((k, i) => (
                  <Area
                    key={k}
                    type="monotone"
                    dataKey={k}
                    name={yLegendByKey.get(k)}
                    stackId="1"
                    stroke={palette[i % palette.length]!}
                    strokeWidth={2.5}
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    fill={`url(#areaGrad-${e}-${i})`}
                    fillOpacity={1}
                    dot={
                      showPointLabels
                        ? {
                            r: 4.5,
                            strokeWidth: 2,
                            fill: palette[i % palette.length]!,
                            stroke: "#fff",
                          }
                        : false
                    }
                    activeDot={{
                      r: 7,
                      strokeWidth: 2,
                      fill: palette[i % palette.length]!,
                      stroke: "#fff",
                    }}
                  >
                    {showPointLabels && (
                      <LabelList
                        dataKey={k}
                        position="top"
                        offset={10}
                        content={renderAreaPointLabel as never}
                      />
                    )}
                  </Area>
                ))}
              </AreaChart>
            )}
          </ResponsiveContainer>
        )}
        {isTemporalChart && chartReady && (
          <div
            className="flex flex-wrap items-center gap-2 border-t border-zinc-200/55 px-2 pb-2 pt-3 dark:border-zinc-600/45"
            role="group"
            aria-label="Шаг группировки по времени"
          >
            <span className="text-[11px] font-semibold uppercase tracking-wide text-zinc-500 dark:text-zinc-400">
              Период
            </span>
            <div className="flex flex-wrap gap-1.5">
              {(
                [
                  { id: "day" as const, label: "По дням" },
                  { id: "month" as const, label: "По месяцам" },
                  { id: "quarter" as const, label: "По кварталам" },
                  { id: "year" as const, label: "По годам" },
                ] as const
              ).map(({ id, label }) => {
                const active = effectiveDateGranularity === id;
                return (
                  <button
                    key={id}
                    type="button"
                    onClick={() => setPeriodOverride(id)}
                    className={
                      active
                        ? "rounded-xl border border-sky-400/85 bg-sky-100/95 px-3 py-1.5 text-xs font-semibold text-sky-950 shadow-sm transition-[box-shadow,transform] hover:brightness-[1.02] active:scale-[0.98] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-sky-500/55 dark:border-sky-500/55 dark:bg-sky-950/50 dark:text-sky-50"
                        : "rounded-xl border border-zinc-200/85 bg-white/90 px-3 py-1.5 text-xs font-medium text-zinc-700 transition-[background-color,border-color,transform] hover:border-zinc-300 hover:bg-zinc-50 active:scale-[0.98] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-sky-500/40 dark:border-zinc-600 dark:bg-zinc-800/85 dark:text-zinc-200 dark:hover:border-zinc-500 dark:hover:bg-zinc-800"
                    }
                  >
                    {label}
                  </button>
                );
              })}
            </div>
            {periodOverride !== null && (
              <button
                type="button"
                onClick={() => setPeriodOverride(null)}
                className={
                  BTN_GHOST +
                  " ml-auto shrink-0 px-3 py-1.5 text-xs font-medium text-zinc-600 dark:text-zinc-300"
                }
              >
                Как в отчёте
              </button>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

function formatCell(v: unknown): string {
  if (v instanceof Date) {
    return formatDateYmdLocal(v);
  }
  if (typeof v === "string") {
    const d = tryParseDate(v);
    if (d) {
      return formatDateYmdLocal(d);
    }
  }
  if (v === null || v === undefined) {
    return "";
  }
  return String(v);
}

type NumberRangeFilterProps = {
  initialMin?: number;
  initialMax?: number;
  onCommit: (min: number | undefined, max: number | undefined) => void;
};

function NumberRangeFilter({
  initialMin,
  initialMax,
  onCommit,
}: NumberRangeFilterProps) {
  const [minS, setMinS] = useState(() =>
    initialMin !== undefined ? String(initialMin) : "",
  );
  const [maxS, setMaxS] = useState(() =>
    initialMax !== undefined ? String(initialMax) : "",
  );

  return (
    <div className="flex gap-2">
      <input
        type="number"
        className={INPUT_FIELD + " w-1/2"}
        placeholder="Мин"
        value={minS}
        onChange={(e) => setMinS(e.target.value)}
        onBlur={() => {
          onCommit(
            minS === "" ? undefined : Number(minS),
            maxS === "" ? undefined : Number(maxS),
          );
        }}
      />
      <input
        type="number"
        className={INPUT_FIELD + " w-1/2"}
        placeholder="Макс"
        value={maxS}
        onChange={(e) => setMaxS(e.target.value)}
        onBlur={() => {
          onCommit(
            minS === "" ? undefined : Number(minS),
            maxS === "" ? undefined : Number(maxS),
          );
        }}
      />
    </div>
  );
}

type FilterBlockProps = {
  tabular: TabularData;
  filters: ChartFilter[];
  onChange: (f: ChartFilter[]) => void;
};

function FilterBlock({ tabular, filters, onChange }: FilterBlockProps) {
  const patchFilter = useCallback(
    (index: number, patch: Partial<ChartFilter>) => {
      onChange(
        filters.map((f, i) => (i === index ? { ...f, ...patch } : f)),
      );
    },
    [filters, onChange],
  );

  const setColumnKey = useCallback(
    (index: number, nextKey: string) => {
      patchFilter(index, {
        columnKey: nextKey,
        values: undefined,
        min: undefined,
        max: undefined,
        dateFrom: undefined,
        dateTo: undefined,
        calendarMonths: undefined,
        calendarQuarters: undefined,
      });
    },
    [patchFilter],
  );

  return (
    <div className="space-y-3 border-t border-zinc-200/90 pt-5 dark:border-zinc-800">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <span className={`block ${SECTION_LABEL}`}>Фильтры данных</span>
        <button
          type="button"
          className={
            BTN_GHOST +
            " shrink-0 px-3 py-1.5 text-xs font-medium text-sky-700 dark:text-sky-300"
          }
          onClick={() => onChange([...filters, { columnKey: "" }])}
        >
          + Условие
        </button>
      </div>
      <p className="text-xs text-zinc-500 dark:text-zinc-400">
        Условия объединяются по «И» и действуют на все графики и превью
        таблицы. Для колонок с датой можно задать календарный диапазон и
        дополнительно отметить нужные месяцы или кварталы.
      </p>
      {filters.length === 0 ? (
        <p className="rounded-lg bg-zinc-50/80 px-3 py-2 text-xs text-zinc-500 dark:bg-zinc-800/50 dark:text-zinc-400">
          Нет условий — нажмите «+ Условие», выберите колонку и задайте
          значения.
        </p>
      ) : (
        <ul className="flex flex-col gap-3">
          {filters.map((filter, index) => (
            <li key={`filter-${index}`}>
              <FilterConditionRow
                tabular={tabular}
                filter={filter}
                filterIndex={index}
                allFilters={filters}
                onSetColumnKey={(k) => setColumnKey(index, k)}
                onPatch={(patch) => patchFilter(index, patch)}
                onRemove={() =>
                  onChange(filters.filter((_, i) => i !== index))
                }
              />
            </li>
          ))}
        </ul>
      )}
      {filters.length > 0 && (
        <button
          type="button"
          className="text-xs font-medium text-zinc-500 underline decoration-zinc-300 underline-offset-2 hover:text-zinc-700 dark:text-zinc-400 dark:hover:text-zinc-200"
          onClick={() => onChange([])}
        >
          Сбросить все фильтры
        </button>
      )}
    </div>
  );
}

function FilterConditionRow({
  tabular,
  filter,
  filterIndex,
  allFilters,
  onSetColumnKey,
  onPatch,
  onRemove,
}: {
  tabular: TabularData;
  filter: ChartFilter;
  filterIndex: number;
  allFilters: ChartFilter[];
  onSetColumnKey: (columnKey: string) => void;
  onPatch: (patch: Partial<ChartFilter>) => void;
  onRemove: () => void;
}) {
  const columnKey = filter.columnKey;
  const meta = tabular.columns.find((c) => c.key === columnKey);

  const baseRows = useMemo(
    () => applyFiltersExcept(tabular.rows, allFilters, filterIndex),
    [tabular.rows, allFilters, filterIndex],
  );

  const stringChoices = useMemo(() => {
    if (!meta || !columnKey) {
      return [] as string[];
    }
    if (meta.inferredType === "number" || meta.inferredType === "date") {
      return [];
    }
    return distinctStringValues(baseRows, columnKey, 200);
  }, [baseRows, columnKey, meta]);

  const calendarMonthOptions = useMemo(() => {
    if (!columnKey || meta?.inferredType !== "date") {
      return [] as string[];
    }
    return distinctCalendarMonthsFromColumn(baseRows, columnKey);
  }, [baseRows, columnKey, meta?.inferredType]);

  const calendarQuarterOptions = useMemo(() => {
    if (!columnKey || meta?.inferredType !== "date") {
      return [] as string[];
    }
    return distinctCalendarQuartersFromColumn(baseRows, columnKey);
  }, [baseRows, columnKey, meta?.inferredType]);

  return (
    <div className="rounded-2xl border border-zinc-200/75 bg-white/90 p-3 shadow-[0_4px_16px_-8px_rgb(15_23_42/0.1)] dark:border-zinc-700/80 dark:bg-zinc-900/55 dark:shadow-none">
      <div className="mb-2 flex items-start gap-2">
        <select
          className={INPUT_FIELD + " min-w-0 flex-1"}
          value={columnKey}
          onChange={(e) => onSetColumnKey(e.target.value)}
          aria-label="Колонка для фильтра"
        >
          <option value="">— колонка —</option>
          {tabular.columns.map((c) => (
            <option key={c.key} value={c.key}>
              {c.header}
            </option>
          ))}
        </select>
        <button
          type="button"
          className="shrink-0 rounded-lg border border-zinc-200 px-2.5 py-2 text-xs font-medium text-zinc-600 transition-colors hover:bg-zinc-50 dark:border-zinc-600 dark:text-zinc-300 dark:hover:bg-zinc-800"
          onClick={onRemove}
          title="Удалить условие"
          aria-label="Удалить условие"
        >
          ×
        </button>
      </div>
      {!columnKey || !meta ? (
        <p className="text-xs text-zinc-500 dark:text-zinc-400">
          Выберите колонку.
        </p>
      ) : meta.inferredType === "date" ? (
        <div className="flex flex-col gap-3">
          <div className="flex flex-col gap-2 sm:flex-row">
            <label className="flex flex-1 flex-col gap-1 text-xs text-zinc-500 dark:text-zinc-400">
              С даты
              <input
                type="date"
                className={INPUT_FIELD}
                value={filter.dateFrom ?? ""}
                onChange={(e) =>
                  onPatch({
                    dateFrom: e.target.value || undefined,
                    values: undefined,
                    min: undefined,
                    max: undefined,
                  })
                }
              />
            </label>
            <label className="flex flex-1 flex-col gap-1 text-xs text-zinc-500 dark:text-zinc-400">
              По дату
              <input
                type="date"
                className={INPUT_FIELD}
                value={filter.dateTo ?? ""}
                onChange={(e) =>
                  onPatch({
                    dateTo: e.target.value || undefined,
                    values: undefined,
                    min: undefined,
                    max: undefined,
                  })
                }
              />
            </label>
          </div>
          <p className="text-[11px] leading-snug text-zinc-500 dark:text-zinc-400">
            Диапазон дат, выбранные месяцы и кварталы применяются вместе (логика
            «И»): остаются строки, которые подходят под все указанные условия.
          </p>
          {calendarMonthOptions.length > 0 && (
            <div className="space-y-1.5">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <span className="text-xs font-medium text-zinc-600 dark:text-zinc-300">
                  По месяцам
                </span>
                {(filter.calendarMonths?.length ?? 0) > 0 && (
                  <button
                    type="button"
                    className="text-[11px] font-medium text-sky-700 underline decoration-sky-300/80 underline-offset-2 hover:text-sky-900 dark:text-sky-300 dark:hover:text-sky-100"
                    onClick={() => onPatch({ calendarMonths: undefined })}
                  >
                    Снять выбор месяцев
                  </button>
                )}
              </div>
              <div className="max-h-36 space-y-1 overflow-y-auto rounded-lg border border-zinc-200/80 bg-zinc-50/50 p-2 text-xs dark:border-zinc-700 dark:bg-zinc-950/40">
                {calendarMonthOptions.map((ym) => {
                  const active = filter.calendarMonths?.includes(ym) ?? false;
                  return (
                    <label
                      key={ym}
                      className="flex cursor-pointer items-center gap-2"
                    >
                      <input
                        type="checkbox"
                        checked={active}
                        className="size-3.5 shrink-0 rounded border-zinc-300 text-sky-600 dark:border-zinc-600"
                        onChange={(e) => {
                          const on = e.target.checked;
                          const set = new Set(filter.calendarMonths ?? []);
                          if (on) {
                            set.add(ym);
                          } else {
                            set.delete(ym);
                          }
                          const next = Array.from(set).sort();
                          onPatch({
                            calendarMonths:
                              next.length > 0 ? next : undefined,
                          });
                        }}
                      />
                      <span className="truncate">
                        {formatCalendarMonthRu(ym)}
                      </span>
                    </label>
                  );
                })}
              </div>
            </div>
          )}
          {calendarQuarterOptions.length > 0 && (
            <div className="space-y-1.5">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <span className="text-xs font-medium text-zinc-600 dark:text-zinc-300">
                  По кварталам
                </span>
                {(filter.calendarQuarters?.length ?? 0) > 0 && (
                  <button
                    type="button"
                    className="text-[11px] font-medium text-sky-700 underline decoration-sky-300/80 underline-offset-2 hover:text-sky-900 dark:text-sky-300 dark:hover:text-sky-100"
                    onClick={() => onPatch({ calendarQuarters: undefined })}
                  >
                    Снять выбор кварталов
                  </button>
                )}
              </div>
              <div className="max-h-32 space-y-1 overflow-y-auto rounded-lg border border-zinc-200/80 bg-zinc-50/50 p-2 text-xs dark:border-zinc-700 dark:bg-zinc-950/40">
                {calendarQuarterOptions.map((yq) => {
                  const active =
                    filter.calendarQuarters?.includes(yq) ?? false;
                  return (
                    <label
                      key={yq}
                      className="flex cursor-pointer items-center gap-2"
                    >
                      <input
                        type="checkbox"
                        checked={active}
                        className="size-3.5 shrink-0 rounded border-zinc-300 text-sky-600 dark:border-zinc-600"
                        onChange={(e) => {
                          const on = e.target.checked;
                          const set = new Set(filter.calendarQuarters ?? []);
                          if (on) {
                            set.add(yq);
                          } else {
                            set.delete(yq);
                          }
                          const next = Array.from(set).sort((a, b) => {
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
                          onPatch({
                            calendarQuarters:
                              next.length > 0 ? next : undefined,
                          });
                        }}
                      />
                      <span className="truncate">
                        {formatCalendarQuarterRu(yq)}
                      </span>
                    </label>
                  );
                })}
              </div>
            </div>
          )}
          {calendarMonthOptions.length === 0 &&
            calendarQuarterOptions.length === 0 && (
              <p className="text-[11px] text-zinc-500 dark:text-zinc-400">
                В данных нет распознанных дат в этой колонке — месяцы и кварталы
                недоступны.
              </p>
            )}
        </div>
      ) : meta.inferredType === "number" ? (
        <NumberRangeFilter
          key={`${filterIndex}-${columnKey}`}
          initialMin={filter.min}
          initialMax={filter.max}
          onCommit={(min, max) => {
            onPatch({
              min,
              max,
              values: undefined,
              dateFrom: undefined,
              dateTo: undefined,
              calendarMonths: undefined,
              calendarQuarters: undefined,
            });
          }}
        />
      ) : (
        <div className="max-h-40 space-y-1 overflow-y-auto rounded-lg border border-zinc-200/80 bg-zinc-50/50 p-2 text-xs dark:border-zinc-700 dark:bg-zinc-950/40">
          {stringChoices.length === 0 ? (
            <p className="px-1 py-2 text-zinc-500 dark:text-zinc-400">
              Нет значений (проверьте другие фильтры или тип колонки).
            </p>
          ) : (
            stringChoices.slice(0, 80).map((val) => {
              const active = filter.values?.includes(val) ?? false;
              return (
                <label
                  key={val}
                  className="flex cursor-pointer items-center gap-2"
                >
                  <input
                    type="checkbox"
                    checked={active}
                    className="size-3.5 shrink-0 rounded border-zinc-300 text-sky-600 dark:border-zinc-600"
                    onChange={(e) => {
                      const on = e.target.checked;
                      const set = new Set(filter.values ?? []);
                      if (on) {
                        set.add(val);
                      } else {
                        set.delete(val);
                      }
                      onPatch({
                        values: Array.from(set),
                        min: undefined,
                        max: undefined,
                        dateFrom: undefined,
                        dateTo: undefined,
                        calendarMonths: undefined,
                        calendarQuarters: undefined,
                      });
                    }}
                  />
                  <span className="truncate">{val}</span>
                </label>
              );
            })
          )}
        </div>
      )}
    </div>
  );
}
