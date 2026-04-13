import { randomUUID } from "node:crypto";
import { promises as fs } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { NextResponse } from "next/server";
import {
  buildLiveChartItems,
  type LiveChartItem,
} from "@/lib/excel/buildLiveChartItems";
import { buildSourceData, type SourceDataMatrix } from "@/lib/excel/buildSourceData";
import type { ChartFilter, TabularData } from "@/lib/types";

export const runtime = "nodejs";
/** Лимит времени выполнения route (сек.), в т.ч. для вызова Python на Vercel. */
export const maxDuration = 120;
const execFileAsync = promisify(execFile);

type ExportChartsSummary = {
  createdAt?: string;
  sourceFileName?: string;
  activeSheet?: string;
  columnsTotal?: number;
  mainInfo?: { label: string; value: string | number }[];
};

type ExportChartsRequest = {
  theme?: "classic" | "vivid" | "minimal";
  fileNameBase?: string;
  sourceFileName?: string;
  mainInfo?: { label: string; value: string | number }[];
  /** Предсобрано в браузере: те же series/categories, что в UI (без потери типов после JSON). */
  charts?: LiveChartItem[];
  sourceData?: SourceDataMatrix;
  summary?: ExportChartsSummary;
  /** Fallback: пересборка на сервере (tabular после req.json() может отличаться от клиента). */
  tabular?: TabularData;
  enabledChartIds?: string[];
  globalFilters?: ChartFilter[];
  dealStageOrder?: string[];
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

function isPrebuiltChartsBody(body: ExportChartsRequest): boolean {
  return (
    Array.isArray(body.charts) &&
    body.charts.length > 0 &&
    body.sourceData != null &&
    Array.isArray(body.sourceData.headers) &&
    Array.isArray(body.sourceData.rows)
  );
}

function wrapPythonExecError(err: unknown): Error {
  const e = err as { message?: string; stderr?: Buffer };
  const stderr = e.stderr ? String(e.stderr).trim() : "";
  const text = stderr.slice(0, 500) || e.message || String(err);
  return new Error(text);
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
  const execOpts = {
    timeout: 120000,
    maxBuffer: 1024 * 1024 * 4,
  };
  const run = (cmd: string) =>
    execFileAsync(cmd, [scriptPath, payloadPath, outputPath], execOpts);

  try {
    try {
      await run("python3");
    } catch (e3) {
      if ((e3 as NodeJS.ErrnoException).code !== "ENOENT") {
        throw wrapPythonExecError(e3);
      }
      try {
        await run("python");
      } catch (e2) {
        if ((e2 as NodeJS.ErrnoException).code !== "ENOENT") {
          throw wrapPythonExecError(e2);
        }
        throw new Error(
          "NO_PYTHON: не найдены команды python3 и python. Установите Python 3 и пакет xlsxwriter (pip install xlsxwriter).",
        );
      }
    }
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
    if (!body) {
      return NextResponse.json(
        { error: "Пустое тело запроса." },
        { status: 400 },
      );
    }

    const prebuilt = isPrebuiltChartsBody(body);
    let chartItems: LiveChartItem[];
    let sourceSheetCount: number;

    if (prebuilt) {
      chartItems = body.charts as LiveChartItem[];
      sourceSheetCount = body.sourceData!.headers.length > 0 ? 1 : 0;
    } else {
      if (!Array.isArray(body.enabledChartIds) || body.enabledChartIds.length === 0) {
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

      chartItems = buildLiveChartItems(
        body.tabular,
        new Set(body.enabledChartIds),
        body.globalFilters ?? [],
        body.dealStageOrder?.length ? body.dealStageOrder : null,
      );
      if (chartItems.length === 0) {
        return NextResponse.json(
          { error: "Нет данных для построения живых диаграмм по выбранным графикам." },
          { status: 400 },
        );
      }
      sourceSheetCount = body.tabular.columns.length ? 1 : 0;
    }

    const fileName = buildExportDownloadFileName(body.fileNameBase);

    const mainInfoMerged =
      (Array.isArray(body.summary?.mainInfo) && body.summary!.mainInfo) ||
      (Array.isArray(body.mainInfo) ? body.mainInfo : []);

    const payload = {
      theme: body.theme ?? "classic",
      summary: {
        createdAt:
          body.summary?.createdAt ??
          new Date().toLocaleString("ru-RU", {
            dateStyle: "short",
            timeStyle: "medium",
          }),
        sourceFileName: body.summary?.sourceFileName ?? body.sourceFileName ?? "",
        activeSheet:
          body.summary?.activeSheet ?? (prebuilt ? "" : body.tabular!.activeSheet),
        columnsTotal:
          typeof body.summary?.columnsTotal === "number"
            ? body.summary.columnsTotal
            : prebuilt
              ? 0
              : body.tabular!.columns.length,
        mainInfo: mainInfoMerged,
      },
      sourceData: prebuilt ? body.sourceData! : buildSourceData(body.tabular!),
      charts: chartItems,
    };
    const data = await generateWorkbookViaPython(payload);

    const bodyBytes =
      data instanceof Uint8Array ? data : new Uint8Array(data);
    return new NextResponse(bodyBytes, {
      status: 200,
      headers: {
        "content-type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "content-disposition": `attachment; filename*=UTF-8''${encodeURIComponent(fileName)}`,
        "x-export-filename": encodeURIComponent(fileName),
        "x-sheets-written": String(chartItems.length + 1 + sourceSheetCount),
      },
    });
  } catch (err) {
    console.error("[export-charts]", err);
    const raw =
      err instanceof Error ? err.message : "Не удалось сформировать Excel с диаграммами.";
    const error = raw.length > 500 ? `${raw.slice(0, 500)}…` : raw;
    return NextResponse.json({ error }, { status: 500 });
  }
}
