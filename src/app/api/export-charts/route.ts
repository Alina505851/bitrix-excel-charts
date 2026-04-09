import { randomUUID } from "node:crypto";
import { promises as fs } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { NextResponse } from "next/server";
import { applyFilters, buildAggregatedChartRows } from "@/lib/chart/buildChartData";
import { tryParseDate } from "@/lib/chart/coerce";
import type { ChartConfig, ChartFilter, TabularData } from "@/lib/types";
import {
  ENTITY_BLOCKS,
  chartsForEntity,
  normalizeConfigForTabular,
} from "@/lib/predefinedCharts";

export const runtime = "nodejs";
const execFileAsync = promisify(execFile);

type LiveChartSeries = {
  name: string;
  values: number[];
  sourceHeader: string;
  aggregation: ChartConfig["aggregation"];
  cumulative?: boolean;
};

type LiveChartItem = {
  entityId: string;
  blockLabel: string;
  title: string;
  sheetName: string;
  chartType: "column" | "bar" | "line" | "area" | "pie";
  xHeader: string;
  xIsDate?: boolean;
  dateGranularity?: ChartConfig["dateGranularity"];
  linkToSourcePossible?: boolean;
  categories: string[];
  series: LiveChartSeries[];
};

type ExportChartsRequest = {
  tabular: TabularData;
  enabledChartIds: string[];
  globalFilters: ChartFilter[];
  fileNameBase?: string;
  sourceFileName?: string;
  theme?: "classic" | "vivid" | "minimal";
};

type LiveSourceData = {
  headers: string[];
  rows: Array<Array<string | number>>;
};

const INVALID_EXPORT_BASE = /[<>:"/\\|?*\u0000-\u001f]/g;

function sanitizeExportFileBase(raw: string): string {
  const s = raw
    .replace(INVALID_EXPORT_BASE, "_")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/_+/g, "_");
  return s.slice(0, 180) || "Отчет";
}

function formatExportTimestampLocal(): string {
  const d = new Date();
  const p = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}_${p(d.getHours())}-${p(d.getMinutes())}-${p(d.getSeconds())}`;
}

function buildExportDownloadFileName(fileNameBase?: string): string {
  const stripped = fileNameBase?.replace(/\.xlsx?$/i, "").trim() ?? "";
  const base = sanitizeExportFileBase(stripped);
  return `${base}_графики_${formatExportTimestampLocal()}.xlsx`;
}

function withGlobalFilters(cfg: ChartConfig, global: ChartFilter[]): ChartConfig {
  return {
    ...cfg,
    filters: [...global, ...cfg.filters],
  };
}

function columnTitle(data: TabularData, key: string): string {
  return data.columns.find((c) => c.key === key)?.header ?? key;
}

function seriesLegendLabel(header: string, aggregation: ChartConfig["aggregation"]): string {
  const h = header.trim();
  const lower = h.toLowerCase();
  let agg = "сумма";
  if (aggregation === "count_nonempty") {
    if (lower.includes("id") || lower.includes("ид")) {
      return "Кол-во компаний · количество компаний";
    }
    agg = "количество непустых значений";
  }
  if (aggregation === "count") {
    agg = "количество записей";
  }
  if (aggregation === "avg") {
    agg = "среднее значение";
  }
  if (aggregation === "sum") {
    agg = "сумма";
  }
  return `${h} · ${agg}`;
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

function buildChartItems(
  tabular: TabularData,
  enabledChartIds: Set<string>,
  globalFilters: ChartFilter[],
): LiveChartItem[] {
  const out: LiveChartItem[] = [];

  for (const block of ENTITY_BLOCKS) {
    for (const spec of chartsForEntity(block.id)) {
      if (!enabledChartIds.has(spec.id)) {
        continue;
      }
      const resolved = spec.resolve(tabular.columns);
      if (!resolved.ok) {
        continue;
      }
      const cfg = normalizeConfigForTabular(
        tabular,
        withGlobalFilters(resolved.config, globalFilters),
      );
      if (!cfg.xKey || cfg.yKeys.length === 0) {
        continue;
      }
      const { data } = buildAggregatedChartRows(tabular.rows, cfg, tabular.columns);
      if (data.length === 0) {
        continue;
      }

      const categories = data.map((row) => String(row.name));
      const period = periodLegendLabel(cfg, tabular);
      const series = cfg.yKeys.map((k) => ({
        name: `${seriesLegendLabel(columnTitle(tabular, k), cfg.aggregation)} · ${period}`,
        sourceHeader: columnTitle(tabular, k),
        aggregation: cfg.aggregation,
        cumulative: cfg.cumulative,
        values: data.map((row) => {
          const v = row[k];
          return typeof v === "number" && Number.isFinite(v) ? v : Number(v) || 0;
        }),
      }));
      out.push({
        entityId: block.id,
        blockLabel: block.label,
        title: spec.title,
        sheetName: `${block.label} — ${spec.title}`,
        chartType: toXlsxChartType(cfg.chartType),
        xHeader: columnTitle(tabular, cfg.xKey),
        xIsDate:
          (tabular.columns.find((c) => c.key === cfg.xKey)?.inferredType ?? "unknown") ===
          "date",
        dateGranularity: cfg.dateGranularity,
        linkToSourcePossible: cfg.filters.length === 0,
        categories,
        series,
      });
    }
  }

  return out;
}

function buildSourceData(tabular: TabularData): LiveSourceData {
  const headers = tabular.columns.map((c) => c.header || c.key);
  const keys = tabular.columns.map((c) => c.key);
  const isCreatedDateColumn = tabular.columns.map((c) =>
    /дата\s*создания|created[_\s-]*at/i.test(c.header || c.key),
  );
  const formatRuDate = (d: Date) => {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${day}.${m}.${y}`;
  };
  const rows = tabular.rows.map((row) =>
    keys.map((k, idx) => {
      const v = row[k];
      if (v === null || v === undefined) {
        return "";
      }
      if (isCreatedDateColumn[idx]) {
        const d = v instanceof Date ? v : tryParseDate(v);
        if (d) {
          return formatRuDate(d);
        }
      }
      if (typeof v === "number" && Number.isFinite(v)) {
        return v;
      }
      if (v instanceof Date) {
        return formatRuDate(v);
      }
      return String(v);
    }),
  );
  return { headers, rows };
}

async function generateWorkbookViaPython(payload: object): Promise<Buffer> {
  const runId = randomUUID();
  const dir = tmpdir();
  const payloadPath = join(dir, `b24-live-charts-${runId}.json`);
  const outputPath = join(dir, `b24-live-charts-${runId}.xlsx`);
  const scriptPath = join(
    process.cwd(),
    "src/lib/excel/server/export_live_charts.py",
  );

  await fs.writeFile(payloadPath, JSON.stringify(payload), "utf8");
  try {
    await execFileAsync("python3", [scriptPath, payloadPath, outputPath], {
      timeout: 120000,
      maxBuffer: 1024 * 1024 * 4,
    });
    return await fs.readFile(outputPath);
  } finally {
    await Promise.allSettled([
      fs.unlink(payloadPath),
      fs.unlink(outputPath),
    ]);
  }
}

export async function POST(req: Request) {
  try {
    const body = (await req.json()) as ExportChartsRequest;
    if (!body || !Array.isArray(body.enabledChartIds) || body.enabledChartIds.length === 0) {
      return NextResponse.json(
        { error: "Не выбран ни один график для экспорта." },
        { status: 400 },
      );
    }
    if (!body.tabular?.columns || !body.tabular?.rows) {
      return NextResponse.json(
        { error: "Некорректные данные таблицы для экспорта." },
        { status: 400 },
      );
    }

    const chartItems = buildChartItems(
      body.tabular,
      new Set(body.enabledChartIds),
      body.globalFilters ?? [],
    );
    if (chartItems.length === 0) {
      return NextResponse.json(
        { error: "Нет данных для построения живых диаграмм по выбранным графикам." },
        { status: 400 },
      );
    }

    const fileName = buildExportDownloadFileName(body.fileNameBase);
    const payload = {
      theme: body.theme ?? "classic",
      summary: {
        createdAt: new Date().toLocaleString("ru-RU", {
          dateStyle: "short",
          timeStyle: "medium",
        }),
        sourceFileName: body.sourceFileName ?? "",
        activeSheet: body.tabular.activeSheet,
        rowsTotal: body.tabular.rows.length,
        rowsFiltered: applyFilters(body.tabular.rows, body.globalFilters ?? []).length,
        columnsTotal: body.tabular.columns.length,
      },
      sourceData: buildSourceData(body.tabular),
      charts: chartItems,
    };
    const data = await generateWorkbookViaPython(payload);

    return new NextResponse(data, {
      status: 200,
      headers: {
        "content-type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "content-disposition": `attachment; filename*=UTF-8''${encodeURIComponent(fileName)}`,
        "x-export-filename": encodeURIComponent(fileName),
        "x-sheets-written": String(chartItems.length + 1),
      },
    });
  } catch {
    return NextResponse.json(
      { error: "Не удалось сформировать Excel с диаграммами." },
      { status: 500 },
    );
  }
}
