import type { EntityBlockId } from "@/lib/predefinedCharts";

/** «Компании» / единый стиль Excel-выгрузки: первый цвет как Accent1 (#0070C0), далее разнообразие рядов. */
export const COMPANY_CHART_COLORS = [
  "#0070C0",
  "#06b6d4",
  "#16a34a",
  "#f59e0b",
  "#ef4444",
  "#0891b2",
  "#db2777",
  "#0d9488",
  "#22c55e",
  "#0ea5e9",
  "#e11d48",
  "#eab308",
];

/** Круговая «Доля суммы по стадиям» (сделки): разнесённые по кругу оттенки, не одна гамма. */
export const DEALS_STAGE_SUM_PIE_COLORS = [
  "#6366f1",
  "#0070C0",
  "#0891b2",
  "#0d9488",
  "#16a34a",
  "#65a30d",
  "#ca8a04",
  "#ea580c",
  "#dc2626",
  "#db2777",
  "#c026d3",
  "#9333ea",
  "#7c3aed",
  "#e11d48",
  "#4f46e5",
];

/** Круговая «Распределение по ответственным»: насыщенные контрастные секторы. */
export const COMPANIES_RESPONSIBLE_PIE_COLORS = [
  "#0070C0",
  "#0891b2",
  "#0d9488",
  "#16a34a",
  "#65a30d",
  "#ca8a04",
  "#ea580c",
  "#dc2626",
  "#db2777",
  "#c026d3",
  "#7c3aed",
  "#4f46e5",
  "#e11d48",
];

export const ENTITY_CHART_PALETTES: Record<EntityBlockId, string[]> = {
  leads: [
    "#22d3ee",
    "#06b6d4",
    "#0ea5e9",
    "#38bdf8",
    "#3b82f6",
    "#2563eb",
    "#60a5fa",
    "#6366f1",
    "#2dd4bf",
    "#0284c7",
    "#818cf8",
    "#7dd3fc",
  ],
  /** Сделки — те же цвета, что и «Компании» (графики и PNG в Excel как в эталоне). */
  deals: COMPANY_CHART_COLORS,
  contacts: [
    "#10b981",
    "#22c55e",
    "#34d399",
    "#4ade80",
    "#14b8a6",
    "#2dd4bf",
    "#059669",
    "#65a30d",
    "#84cc16",
    "#0d9488",
    "#16a34a",
    "#047857",
  ],
  companies: COMPANY_CHART_COLORS,
  quotes: [
    "#fb7185",
    "#f43f5e",
    "#ec4899",
    "#e11d48",
    "#f472b6",
    "#be123c",
    "#fda4af",
    "#db2777",
    "#f87171",
    "#ef4444",
    "#9f1239",
    "#d946ef",
  ],
  generic: [
    "#64748b",
    "#475569",
    "#0ea5e9",
    "#06b6d4",
    "#14b8a6",
    "#8b5cf6",
    "#a855f7",
    "#d946ef",
    "#ec4899",
    "#f43f5e",
    "#f59e0b",
    "#84cc16",
  ],
};

export function chartPaletteForEntity(
  entity: EntityBlockId | undefined,
): string[] {
  return ENTITY_CHART_PALETTES[entity ?? "leads"];
}

/** Единый яркий цикл для Chart.js (экспорт PNG в Excel). */
export const CHART_EXPORT_COLOR_CYCLE = [
  "#06b6d4",
  "#3b82f6",
  "#6366f1",
  "#8b5cf6",
  "#d946ef",
  "#ec4899",
  "#f43f5e",
  "#22c55e",
  "#eab308",
  "#14b8a6",
  "#0ea5e9",
  "#a855f7",
];
