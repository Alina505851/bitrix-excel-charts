import type { EntityBlockId } from "@/lib/predefinedCharts";

/** «Компании»: яркие разные оттенки без оранжевого (синий, бирюза, зелёный, пурпур, розовый). */
export const COMPANY_CHART_COLORS = [
  "#2563eb",
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

/** Круговая «Распределение по ответственным»: насыщенные контрастные секторы. */
export const COMPANIES_RESPONSIBLE_PIE_COLORS = [
  "#2563eb",
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
  deals: [
    "#a855f7",
    "#9333ea",
    "#d946ef",
    "#c026d3",
    "#e879f9",
    "#8b5cf6",
    "#c084fc",
    "#a21caf",
    "#7c3aed",
    "#f0abfc",
    "#6d28d9",
    "#db2777",
  ],
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
