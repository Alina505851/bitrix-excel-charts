import type { EntityBlockId } from "@/lib/predefinedCharts";

/** Фон области построения диаграммы (внутри карточки графика). */
export type ChartSurfaceId =
  | "default"
  | "white"
  | "zinc"
  | "sky"
  | "violet"
  | "emerald"
  | "dark";

export const CHART_SURFACE_OPTIONS: {
  id: ChartSurfaceId;
  label: string;
}[] = [
  { id: "default", label: "По разделу CRM (как сейчас)" },
  { id: "white", label: "Белый" },
  { id: "zinc", label: "Нейтральный серый" },
  { id: "sky", label: "Голубой" },
  { id: "violet", label: "Фиолетовый" },
  { id: "emerald", label: "Изумрудный" },
  { id: "dark", label: "Тёмный" },
];

export const CHART_SURFACE_STORAGE_KEY = "b24-chart-surface";

export function isChartSurfaceId(v: string): v is ChartSurfaceId {
  return CHART_SURFACE_OPTIONS.some((o) => o.id === v);
}

const SHELL_BASE =
  "min-h-[340px] rounded-xl py-3 px-0.5 ring-1 ring-inset transition-[background-color] duration-200";

/** Пресеты, общие для всех сущностей (не «по разделу»). */
const CHART_SURFACE_PRESETS: Record<
  Exclude<ChartSurfaceId, "default">,
  string
> = {
  white: `${SHELL_BASE} bg-white ring-zinc-200/65 dark:bg-zinc-950 dark:ring-zinc-700/55`,
  zinc: `${SHELL_BASE} bg-gradient-to-b from-zinc-50 via-white to-zinc-100/40 ring-zinc-200/55 dark:from-zinc-900 dark:via-zinc-950 dark:to-zinc-900 dark:ring-zinc-700/45`,
  sky: `${SHELL_BASE} bg-gradient-to-b from-sky-50 via-cyan-50/35 to-white ring-sky-200/50 dark:from-sky-950/45 dark:via-zinc-950 dark:to-sky-950/25 dark:ring-sky-800/40`,
  violet: `${SHELL_BASE} bg-gradient-to-b from-violet-50 via-fuchsia-50/30 to-white ring-violet-200/50 dark:from-violet-950/40 dark:via-zinc-950 dark:to-violet-950/25 dark:ring-violet-800/38`,
  emerald: `${SHELL_BASE} bg-gradient-to-b from-emerald-50 via-teal-50/28 to-white ring-emerald-200/48 dark:from-emerald-950/35 dark:via-zinc-950 dark:to-emerald-950/22 dark:ring-emerald-800/38`,
  dark: `${SHELL_BASE} bg-gradient-to-b from-zinc-700 via-zinc-800 to-zinc-900 ring-zinc-500/45 dark:from-zinc-800 dark:via-zinc-900 dark:to-zinc-950 dark:ring-zinc-600/40`,
};

/** Темы осей для тёмного фона графика — светлые подписи и сетка. */
export const CHART_AXES_DARK_SURFACE: {
  axis: string;
  grid: string;
  tick: string;
} = {
  axis: "#94a3b8",
  grid: "rgba(148, 163, 184, 0.14)",
  tick: "#e2e8f0",
};

const ENTITY_CHART_SHELL: Record<EntityBlockId, string> = {
  leads:
    "min-h-[340px] rounded-xl bg-gradient-to-b from-sky-100/55 via-sky-50/25 to-white py-3 px-0.5 ring-1 ring-inset ring-sky-200/40 dark:from-sky-950/28 dark:via-sky-950/12 dark:to-zinc-950/50 dark:ring-sky-500/18",
  deals:
    "min-h-[340px] rounded-xl bg-gradient-to-b from-violet-100/50 via-violet-50/22 to-white py-3 px-0.5 ring-1 ring-inset ring-violet-200/42 dark:from-violet-950/28 dark:via-violet-950/12 dark:to-zinc-950/50 dark:ring-fuchsia-500/16",
  contacts:
    "min-h-[340px] rounded-xl bg-gradient-to-b from-emerald-100/45 via-teal-50/22 to-white py-3 px-0.5 ring-1 ring-inset ring-emerald-200/38 dark:from-emerald-950/26 dark:via-teal-950/10 dark:to-zinc-950/50 dark:ring-emerald-500/16",
  companies:
    "min-h-[340px] rounded-xl bg-gradient-to-b from-indigo-100/45 via-violet-50/30 to-cyan-50/20 py-3 px-0.5 ring-1 ring-inset ring-indigo-200/42 dark:from-indigo-950/24 dark:via-violet-950/12 dark:to-zinc-950/52 dark:ring-indigo-500/18",
  quotes:
    "min-h-[340px] rounded-xl bg-gradient-to-b from-rose-100/50 via-rose-50/20 to-white py-3 px-0.5 ring-1 ring-inset ring-rose-200/42 dark:from-rose-950/26 dark:via-rose-950/10 dark:to-zinc-950/50 dark:ring-rose-500/16",
};

export function chartShellClass(
  entity: EntityBlockId,
  surface: ChartSurfaceId,
): string {
  if (surface === "default") {
    return ENTITY_CHART_SHELL[entity];
  }
  return CHART_SURFACE_PRESETS[surface];
}

export function chartUsesDarkSurface(surface: ChartSurfaceId): boolean {
  return surface === "dark";
}
