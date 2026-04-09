import ExcelJS from "exceljs";
import {
  applyFilters,
  buildAggregatedChartRows,
} from "@/lib/chart/buildChartData";
import type { ChartConfig, ChartFilter, TabularData } from "@/lib/types";
import {
  ENTITY_BLOCKS,
  chartsForEntity,
  normalizeConfigForTabular,
  type PredefinedChartSpec,
} from "@/lib/predefinedCharts";
import { renderAggregatedChartPng } from "@/lib/excel/chartToPng";

export type ExcelExportTheme = "classic" | "vivid" | "minimal";

const SHEET_INVALID = /[:\\/?*[\]]/g;

/** Символы, недопустимые в имени файла Windows / macOS */
const INVALID_EXPORT_BASE = /[<>:"/\\|?*\u0000-\u001f]/g;

function sanitizeExportFileBase(raw: string): string {
  const s = raw
    .replace(INVALID_EXPORT_BASE, "_")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/_+/g, "_");
  return s.slice(0, 180) || "Отчет";
}

/** Локальная метка времени для имени файла (без двоеточий). */
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

/** ARGB для заливок и шрифтов в Excel */
const XL = {
  white: "FFFFFFFF",
  ink: "FF0F172A",
  inkMuted: "FF64748B",
  sky700: "FF0369A1",
  sky500: "FF0EA5E9",
  sky100: "FFE0F2FE",
  sky50: "FFF0F9FF",
  slate100: "FFF1F5F9",
  slate200: "FFE2E8F0",
  slate50: "FFF8FAFC",
  emerald700: "FF047857",
  rose50: "FFFFF1F2",
  rose700: "FFBE123C",
  border: "FFE2E8F0",
} as const;

function colLetter(zeroBasedIndex: number): string {
  let n = zeroBasedIndex + 1;
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function thinBorder(
  color: string = XL.border,
): Partial<ExcelJS.Borders> {
  const b = { style: "thin" as const, color: { argb: color } };
  return { top: b, left: b, bottom: b, right: b };
}

function fillSolid(argb: string): ExcelJS.Fill {
  return {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb },
  };
}

function styleSummarySheet(
  ws: ExcelJS.Worksheet,
  tableRowStart: number,
  chartRows: number,
): void {
  ws.columns = [
    { width: 26 },
    { width: 42 },
    { width: 52 },
  ];

  ws.mergeCells(`A1:C1`);
  const t = ws.getCell("A1");
  t.value = "Общая информация";
  t.font = { size: 20, bold: true, color: { argb: XL.white } };
  t.fill = fillSolid(XL.sky700);
  t.alignment = { vertical: "middle", horizontal: "center" };
  ws.getRow(1).height = 40;

  const metaEnd = tableRowStart - 2;
  for (let r = 3; r <= metaEnd; r++) {
    const row = ws.getRow(r);
    if (!row.getCell(1).value && !row.getCell(2).value) {
      continue;
    }
    const bg = r % 2 === 1 ? XL.slate50 : XL.white;
    row.getCell(1).font = { size: 11, color: { argb: XL.inkMuted }, bold: true };
    row.getCell(1).alignment = { vertical: "middle" };
    row.getCell(1).fill = fillSolid(bg);
    row.getCell(2).font = { size: 11, color: { argb: XL.ink } };
    row.getCell(2).alignment = { vertical: "middle", wrapText: true };
    row.getCell(2).fill = fillSolid(bg);
    row.getCell(3).fill = fillSolid(bg);
    for (let col = 1; col <= 3; col++) {
      row.getCell(col).border = thinBorder();
    }
  }

  const headerR = ws.getRow(tableRowStart);
  ["Раздел", "График", "Статус"].forEach((_, i) => {
    const c = headerR.getCell(i + 1);
    c.font = { bold: true, size: 11, color: { argb: XL.white } };
    c.fill = fillSolid(XL.sky500);
    c.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    c.border = thinBorder(XL.sky700);
  });
  headerR.height = 26;

  for (let i = 0; i < chartRows; i++) {
    const r = ws.getRow(tableRowStart + 1 + i);
    const zebra = i % 2 === 0 ? XL.sky50 : XL.white;
    r.eachCell({ includeEmpty: true }, (cell, col) => {
      if (col <= 3) {
        cell.fill = fillSolid(zebra);
        cell.font = { size: 11, color: { argb: XL.ink } };
        cell.alignment = { vertical: "top", wrapText: true };
        cell.border = thinBorder();
      }
    });
    r.height = 22;
  }

  ws.views = [{ state: "frozen", ySplit: tableRowStart }];
}

function styleChartTitleRow(
  ws: ExcelJS.Worksheet,
  colCount: number,
  title: string,
): void {
  if (colCount < 1) {
    return;
  }
  const end = colLetter(colCount - 1);
  ws.mergeCells(`A1:${end}1`);
  const c = ws.getCell("A1");
  c.value = title;
  c.font = { size: 14, bold: true, color: { argb: XL.ink } };
  c.fill = fillSolid(XL.sky100);
  c.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
  ws.getRow(1).height = 32;
}

function styleDataTableBlock(
  ws: ExcelJS.Worksheet,
  headerRowIndex: number,
  colCount: number,
  dataRowCount: number,
): void {
  for (let c = 0; c < colCount; c++) {
    ws.getColumn(c + 1).width = c === 0 ? 28 : 16;
  }

  const hr = ws.getRow(headerRowIndex);
  for (let c = 1; c <= colCount; c++) {
    const cell = hr.getCell(c);
    cell.font = { bold: true, size: 11, color: { argb: XL.white } };
    cell.fill = fillSolid(XL.sky500);
    cell.alignment = {
      vertical: "middle",
      horizontal: c === 1 ? "left" : "right",
      wrapText: true,
    };
    cell.border = thinBorder(XL.sky700);
  }
  hr.height = 24;

  for (let i = 0; i < dataRowCount; i++) {
    const r = ws.getRow(headerRowIndex + 1 + i);
    const zebra = i % 2 === 0 ? XL.white : XL.slate50;
    for (let c = 1; c <= colCount; c++) {
      const cell = r.getCell(c);
      cell.fill = fillSolid(zebra);
      cell.font = { size: 11, color: { argb: XL.ink } };
      cell.alignment = {
        vertical: "middle",
        horizontal: c === 1 ? "left" : "right",
        wrapText: true,
      };
      cell.border = thinBorder();
      if (c > 1 && typeof cell.value === "number") {
        cell.numFmt = "#,##0.##";
      }
    }
  }

  const end = colLetter(colCount - 1);
  ws.views = [{ state: "frozen", ySplit: headerRowIndex }];
}

function styleMessageMerged(
  ws: ExcelJS.Worksheet,
  row: number,
  colSpan: number,
  message: string,
  opts: { kind: "warn" | "muted" | "error" },
): void {
  const end = colLetter(colSpan - 1);
  ws.mergeCells(`A${row}:${end}${row}`);
  const cell = ws.getCell(`A${row}`);
  cell.value = message;
  const colors =
    opts.kind === "error"
      ? { fg: XL.rose700, bg: XL.rose50 }
      : opts.kind === "warn"
        ? { fg: XL.ink, bg: "FFFFFBEB" }
        : { fg: XL.inkMuted, bg: XL.slate100 };
  cell.font = { size: 11, color: { argb: colors.fg }, bold: opts.kind === "error" };
  cell.fill = fillSolid(colors.bg);
  cell.alignment = { vertical: "middle", wrapText: true, indent: 1 };
  ws.getRow(row).height = 28;
}

function styleChartSectionLabel(ws: ExcelJS.Worksheet, row: number, colSpan: number): void {
  const end = colLetter(colSpan - 1);
  ws.mergeCells(`A${row}:${end}${row}`);
  const c = ws.getCell(`A${row}`);
  c.value = "Диаграмма";
  c.font = { size: 11, bold: true, color: { argb: XL.sky700 } };
  c.fill = fillSolid(XL.sky100);
  c.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
  ws.getRow(row).height = 22;
}

function styleResolveErrorSheet(
  ws: ExcelJS.Worksheet,
  title: string,
  errorText: string,
): void {
  const span = 4;
  styleChartTitleRow(ws, span, title);
  styleMessageMerged(ws, 3, span, "График не построен", { kind: "error" });
  styleMessageMerged(ws, 4, span, errorText, { kind: "muted" });
  for (let c = 1; c <= span; c++) {
    ws.getColumn(c).width = c === 1 ? 22 : 24;
  }
}

function finishEmptySheetStyles(
  ws: ExcelJS.Worksheet,
  title: string,
  note: string,
  headerRow: string[],
): void {
  const colCount = headerRow.length;
  const span = Math.max(4, colCount);
  styleChartTitleRow(ws, span, title);
  styleMessageMerged(ws, 3, span, "Нет данных", { kind: "warn" });
  styleMessageMerged(ws, 4, span, note, { kind: "muted" });
  styleDataTableBlock(ws, 6, colCount, 1);
}

function withGlobalFilters(
  cfg: ChartConfig,
  global: ChartFilter[],
): ChartConfig {
  return {
    ...cfg,
    filters: [...global, ...cfg.filters],
  };
}

function columnTitle(data: TabularData, key: string): string {
  return data.columns.find((c) => c.key === key)?.header ?? key;
}

function sanitizeSheetName(raw: string, used: Set<string>): string {
  let s = raw.replace(SHEET_INVALID, " ").replace(/\s+/g, " ").trim();
  s = s.slice(0, 31) || "График";
  let base = s;
  let n = 2;
  while (used.has(s)) {
    const suffix = ` ${n}`;
    base = base.slice(0, 31 - suffix.length);
    s = (base + suffix).slice(0, 31);
    n++;
  }
  used.add(s);
  return s;
}

function addMatrix(ws: ExcelJS.Worksheet, aoa: (string | number)[][]) {
  for (const row of aoa) {
    ws.addRow(row);
  }
}

type BlockInfo = (typeof ENTITY_BLOCKS)[number];

type PreparedChartSheet =
  | {
      kind: "data";
      block: BlockInfo;
      spec: PredefinedChartSpec;
      cfg: ChartConfig;
      dataRows: ReturnType<typeof buildAggregatedChartRows>["data"];
      headerRow: string[];
    }
  | {
      kind: "empty";
      block: BlockInfo;
      spec: PredefinedChartSpec;
      cfg: ChartConfig;
      warnings: string[];
      headerRow: string[];
    }
  | {
      kind: "resolve_error";
      block: BlockInfo;
      spec: PredefinedChartSpec;
      error: string;
    };

function listEnabledCharts(
  enabledChartIds: Set<string>,
): { block: BlockInfo; spec: PredefinedChartSpec }[] {
  const out: { block: BlockInfo; spec: PredefinedChartSpec }[] = [];
  for (const block of ENTITY_BLOCKS) {
    for (const spec of chartsForEntity(block.id)) {
      if (enabledChartIds.has(spec.id)) {
        out.push({ block, spec });
      }
    }
  }
  return out;
}

function prepareChartSheets(
  tabular: TabularData,
  enabledChartIds: Set<string>,
  globalFilters: ChartFilter[],
): PreparedChartSheet[] {
  const prepared: PreparedChartSheet[] = [];

  for (const { block, spec } of listEnabledCharts(enabledChartIds)) {
    const resolved = spec.resolve(tabular.columns);
    if (!resolved.ok) {
      prepared.push({
        kind: "resolve_error",
        block,
        spec,
        error: resolved.error,
      });
      continue;
    }

    const cfg = normalizeConfigForTabular(
      tabular,
      withGlobalFilters(resolved.config, globalFilters),
    );

    if (!cfg.xKey || cfg.yKeys.length === 0) {
      prepared.push({
        kind: "resolve_error",
        block,
        spec,
        error:
          "Не удалось сопоставить колонки осей с текущим листом (проверьте заголовки).",
      });
      continue;
    }

    const headerRow = [
      columnTitle(tabular, cfg.xKey),
      ...cfg.yKeys.map((k) => columnTitle(tabular, k)),
    ];

    const { data, warnings } = buildAggregatedChartRows(
      tabular.rows,
      cfg,
      tabular.columns,
    );

    if (data.length === 0) {
      prepared.push({
        kind: "empty",
        block,
        spec,
        cfg,
        warnings,
        headerRow,
      });
    } else {
      prepared.push({
        kind: "data",
        block,
        spec,
        cfg,
        dataRows: data,
        headerRow,
      });
    }
  }

  return prepared;
}

function summaryStatusText(p: PreparedChartSheet): string {
  switch (p.kind) {
    case "data":
      return "Готово";
    case "empty": {
      const w = p.warnings.filter(Boolean).join("; ");
      return w ? `Нет данных — ${w.slice(0, 160)}` : "Нет данных";
    }
    case "resolve_error":
      return `Не построен — ${p.error.slice(0, 200)}`;
    default:
      return "";
  }
}

export type ExportChartsResult =
  | { ok: true; fileName: string; sheetsWritten: number; buffer?: ArrayBuffer }
  | { ok: false; error: string };

function toArrayBuffer(buf: ExcelJS.Buffer): ArrayBuffer {
  if (buf instanceof ArrayBuffer) {
    return buf;
  }
  const u8 = buf as Uint8Array;
  const slice = u8.buffer.slice(
    u8.byteOffset,
    u8.byteOffset + u8.byteLength,
  );
  return slice as ArrayBuffer;
}

function downloadBlob(fileName: string, data: ExcelJS.Buffer) {
  const blob = new Blob([data], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  a.click();
  URL.revokeObjectURL(url);
}

async function exportChartsViaApi(
  tabular: TabularData,
  enabledChartIds: Set<string>,
  globalFilters: ChartFilter[],
  fileNameBase?: string,
  sourceFileName?: string,
  theme: ExcelExportTheme = "classic",
): Promise<ExportChartsResult> {
  try {
    const res = await fetch("/api/export-charts", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        tabular,
        enabledChartIds: [...enabledChartIds],
        globalFilters,
        fileNameBase,
        sourceFileName,
        theme,
      }),
    });
    if (!res.ok) {
      let error = "Ошибка экспорта Excel с диаграммами.";
      try {
        const payload = (await res.json()) as { error?: string };
        if (payload?.error) {
          error = payload.error;
        }
      } catch {
        /* ignore */
      }
      return { ok: false, error };
    }
    const blob = await res.blob();
    const encodedName = res.headers.get("x-export-filename");
    const fileName = encodedName
      ? decodeURIComponent(encodedName)
      : buildExportDownloadFileName(fileNameBase);
    const a = document.createElement("a");
    const url = URL.createObjectURL(blob);
    a.href = url;
    a.download = fileName;
    a.click();
    URL.revokeObjectURL(url);
    const sheetsWritten = Number(res.headers.get("x-sheets-written") || "0") || 0;
    return { ok: true, fileName, sheetsWritten };
  } catch {
    return {
      ok: false,
      error: "Не удалось скачать файл с диаграммами. Проверьте соединение и повторите.",
    };
  }
}

async function appendChartImageIfPossible(
  workbook: ExcelJS.Workbook,
  worksheet: ExcelJS.Worksheet,
  tabular: TabularData,
  p: PreparedChartSheet & { kind: "data" },
  colSpan: number,
): Promise<void> {
  const labels = p.dataRows.map((row) => String(row.name));
  const series = p.cfg.yKeys.map((k) => ({
    label: columnTitle(tabular, k),
    values: p.dataRows.map((row) => {
      const v = row[k];
      return typeof v === "number" && Number.isFinite(v) ? v : Number(v) || 0;
    }),
  }));

  const span = Math.max(4, colSpan);

  const base64 = await renderAggregatedChartPng({
    chartType: p.cfg.chartType,
    labels,
    series,
    entity: p.block.id,
  });
  if (!base64) {
    worksheet.addRow([]);
    const r = worksheet.rowCount;
    styleMessageMerged(worksheet, r, span, "Диаграмма не сформирована", {
      kind: "muted",
    });
    return;
  }

  const imageId = workbook.addImage({
    base64,
    extension: "png",
  });

  worksheet.addRow([]);
  const labelRow = worksheet.rowCount;
  styleChartSectionLabel(worksheet, labelRow, span);
  worksheet.addRow([]);
  const anchorRow = worksheet.rowCount;
  worksheet.addImage(imageId, {
    tl: { col: 0, row: anchorRow },
    ext: { width: 760, height: 420 },
  });
}

export async function exportChartsToExcelFile(
  tabular: TabularData,
  enabledChartIds: Set<string>,
  globalFilters: ChartFilter[],
  options?: {
    fileNameBase?: string;
    sourceFileName?: string;
    theme?: ExcelExportTheme;
    /** Для тестов: не скачивать файл, вернуть буфер книги. */
    returnBuffer?: boolean;
  },
): Promise<ExportChartsResult> {
  if (enabledChartIds.size === 0) {
    return {
      ok: false,
      error:
        "Не выбран ни один график: отметьте нужные отчёты галочками в боковой панели.",
    };
  }

  if (
    !options?.returnBuffer &&
    typeof window !== "undefined" &&
    typeof document !== "undefined"
  ) {
    return exportChartsViaApi(
      tabular,
      enabledChartIds,
      globalFilters,
      options?.fileNameBase,
      options?.sourceFileName,
      options?.theme ?? "classic",
    );
  }

  const prepared = prepareChartSheets(
    tabular,
    enabledChartIds,
    globalFilters,
  );

  if (prepared.length === 0) {
    return {
      ok: false,
      error:
        "Не удалось подготовить листы: проверьте соответствие выгрузки типу графиков.",
    };
  }

  const filteredRowCount = applyFilters(tabular.rows, globalFilters).length;

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Отчёты Excel · Битрикс";
  workbook.created = new Date();
  const usedSheetNames = new Set<string>();

  const summaryStamp = new Date().toLocaleString("ru-RU", {
    dateStyle: "short",
    timeStyle: "medium",
  });

  const summaryAoa: (string | number)[][] = [
    ["Общая информация"],
    [],
    ["Дата формирования", summaryStamp],
    ...(options?.sourceFileName
      ? ([["Исходный файл", options.sourceFileName]] as (string | number)[][])
      : []),
    ["Строк на листе", tabular.rows.length],
    ["Строк после фильтров", filteredRowCount],
    ["Графиков в файле", prepared.length],
    [],
    ["Раздел", "График", "Статус"],
    ...prepared.map((p) => [
      p.block.label,
      p.spec.title,
      summaryStatusText(p),
    ]),
  ];

  const summaryName = sanitizeSheetName("Общая информация", usedSheetNames);
  const wsSummary = workbook.addWorksheet(summaryName);
  addMatrix(wsSummary, summaryAoa);
  const summaryTableHeaderRow = 9 + (options?.sourceFileName ? 1 : 0);
  styleSummarySheet(wsSummary, summaryTableHeaderRow, prepared.length);

  let sheetsWritten = 1;

  for (const p of prepared) {
    const sheetTitle = `${p.block.label} — ${p.spec.title}`;
    const sheetName = sanitizeSheetName(sheetTitle, usedSheetNames);
    const ws = workbook.addWorksheet(sheetName);
    sheetsWritten++;

    const title = `${p.block.label}: ${p.spec.title}`;

    if (p.kind === "resolve_error") {
      addMatrix(ws, [[title], [], ["График не построен"], [p.error]]);
      styleResolveErrorSheet(ws, title, p.error);
      continue;
    }

    if (p.kind === "empty") {
      const note =
        p.warnings.filter(Boolean).join("; ") ||
        "Нет строк (проверьте фильтры и данные).";
      const h = p.headerRow;
      addMatrix(ws, [
        [title],
        [],
        ["Нет данных"],
        [note],
        [],
        h,
        ["—", ...h.slice(1).map(() => "—")],
      ]);
      finishEmptySheetStyles(ws, title, note, h);
      continue;
    }

    const dataRows = p.dataRows.map((row) => [
      row.name,
      ...p.cfg.yKeys.map((k) => row[k] as number | string),
    ]);

    const colCount = p.headerRow.length;
    addMatrix(ws, [[title], [], p.headerRow, ...dataRows]);
    styleChartTitleRow(ws, Math.max(4, colCount), title);
    styleDataTableBlock(ws, 3, colCount, p.dataRows.length);

    await appendChartImageIfPossible(workbook, ws, tabular, p, colCount);
  }

  const fileName = buildExportDownloadFileName(options?.fileNameBase);

  const rawBuffer = await workbook.xlsx.writeBuffer();
  const arrayBuffer = toArrayBuffer(rawBuffer);

  if (options?.returnBuffer) {
    return { ok: true, fileName, sheetsWritten, buffer: arrayBuffer };
  }

  if (typeof window !== "undefined" && typeof document !== "undefined") {
    downloadBlob(fileName, rawBuffer);
  }

  return { ok: true, fileName, sheetsWritten };
}
