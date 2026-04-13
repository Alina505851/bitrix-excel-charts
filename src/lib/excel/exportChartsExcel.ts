import ExcelJS from "exceljs";
import { buildAggregatedChartRows } from "@/lib/chart/buildChartData";
import { seriesLegendLabel } from "@/lib/chart/seriesLegendLabel";
import { buildLiveChartItems } from "@/lib/excel/buildLiveChartItems";
import { buildSourceData } from "@/lib/excel/buildSourceData";
import type { ExcelMainInfoLine } from "@/lib/excel/exportMainInfo";
import type {
  AggregationMode,
  ChartConfig,
  ChartFilter,
  TabularData,
} from "@/lib/types";
import {
  ENTITY_BLOCKS,
  chartsForEntity,
  normalizeConfigForTabular,
  type EntityBlockId,
  type PredefinedChartSpec,
} from "@/lib/predefinedCharts";
import { REPORT_INCLUDES_THROUGH_YMD } from "@/lib/report/reportThroughDate";

export type ExcelExportTheme = "classic" | "vivid" | "minimal";

export type { ExcelMainInfoLine } from "@/lib/excel/exportMainInfo";

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
  mainInfoLines?: ExcelMainInfoLine[],
): void {
  const sectionBg = "FFE0E7FF";
  const sectionFg = "FF4338CA";
  const mainCount = mainInfoLines?.length ?? 0;

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

  if (mainCount > 0) {
    ws.mergeCells("A3:C3");
    const sec = ws.getCell("A3");
    sec.value = "Основные показатели";
    sec.font = { bold: true, size: 12, color: { argb: sectionFg } };
    sec.fill = fillSolid(sectionBg);
    sec.border = thinBorder();
    sec.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
    ws.getRow(3).height = 28;
    for (let i = 0; i < mainCount; i++) {
      const r = 4 + i;
      const line = mainInfoLines![i]!;
      const row = ws.getRow(r);
      const bg = i % 2 === 0 ? XL.slate50 : XL.white;
      row.getCell(1).font = { size: 11, color: { argb: XL.inkMuted }, bold: true };
      row.getCell(1).alignment = { vertical: "middle" };
      row.getCell(1).fill = fillSolid(bg);
      row.getCell(1).border = thinBorder();
      const c2 = row.getCell(2);
      c2.font = { size: 11, color: { argb: XL.ink } };
      c2.alignment = { vertical: "middle", wrapText: true };
      c2.fill = fillSolid(bg);
      c2.border = thinBorder();
      if (typeof line.value === "number" && Number.isFinite(line.value)) {
        c2.numFmt = summaryNumFmtForLabel(line.label);
      }
      row.getCell(3).fill = fillSolid(bg);
      row.getCell(3).border = thinBorder();
    }
  }

  const metaStart = mainCount > 0 ? 4 + mainCount + 1 : 3;
  const metaEnd = tableRowStart - 2;
  for (let r = metaStart; r <= metaEnd; r++) {
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
    }
  }

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

function styleExportSectionHeading(
  ws: ExcelJS.Worksheet,
  row: number,
  colSpan: number,
  title: string,
): void {
  const end = colLetter(colSpan - 1);
  ws.mergeCells(`A${row}:${end}${row}`);
  const c = ws.getCell(`A${row}`);
  c.value = title;
  c.font = { size: 11, bold: true, color: { argb: XL.sky700 } };
  c.fill = fillSolid(XL.sky100);
  c.alignment = { vertical: "middle", horizontal: "left", indent: 1, wrapText: true };
  ws.getRow(row).height = 22;
}

function buildMethodologyNarrative(
  spec: PredefinedChartSpec,
  tabular: TabularData,
  cfg: ChartConfig,
): string {
  const parts: string[] = [spec.description.trim()];
  const xLabel = cfg.xKey ? columnTitle(tabular, cfg.xKey) : "";
  parts.push(
    `Ось категорий (X): «${xLabel}». Сначала применяются глобальные фильтры отчёта и фильтры пресета графика; затем строки группируются по подписи категории на оси X, и для каждой группы считаются значения серий в соответствии с режимом агрегации (сумма, среднее, число записей и т.д.). Таблица на листе совпадает с данными диаграммы.`,
  );
  if (cfg.dateGranularity && cfg.xKey) {
    const xMeta = tabular.columns.find((c) => c.key === cfg.xKey);
    if (xMeta?.inferredType === "date") {
      parts.push(
        `Шаг группировки по времени на оси X: ${cfg.dateGranularity} (календарные интервалы).`,
      );
    }
  }
  if (cfg.cumulative) {
    parts.push(
      "Включён накопительный итог: значение в точке включает сумму по предыдущим периодам на оси времени.",
    );
  }
  if (cfg.chartType === "pie") {
    parts.push(
      "Круговая диаграмма: сектор соответствует категории на оси; величина пропорциональна метрике относительно суммы по всем категориям выборки.",
    );
  }
  return parts.join("\n\n");
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

function exportSeriesColumnHeader(
  tabular: TabularData,
  cfg: ChartConfig,
  index: number,
  entity: EntityBlockId,
): string {
  const srcKey = cfg.ySourceKeys?.[index] ?? cfg.yKeys[index]!;
  const agg = cfg.yAggregations?.[index] ?? cfg.aggregation;
  return seriesLegendLabel(columnTitle(tabular, srcKey), agg, entity);
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

/** Числовой формат Excel: рубли с разрядами; счётчики — целые с разрядами. */
const EXCEL_NUM_FMT_RUB = '# ##0,00 "₽"';
const EXCEL_NUM_FMT_COUNT = "# ##0";
const EXCEL_NUM_FMT_PCT = "0.0%";
const EXCEL_NUM_FMT_DAYS = "0.0";

/** Как на листе Python: деньги / % / дни / счётчики по подписи строки KPI. */
function summaryNumFmtForLabel(label: string): string {
  const l = label.toLowerCase();
  if (/(конверсия|темп роста)/.test(l)) {
    return EXCEL_NUM_FMT_PCT;
  }
  if (/цикл/.test(l)) {
    return EXCEL_NUM_FMT_DAYS;
  }
  if (
    /(руб|₽|сумм|выруч|доход|чек|стоим|оплат|средн|денег)/.test(l)
  ) {
    return EXCEL_NUM_FMT_RUB;
  }
  return EXCEL_NUM_FMT_COUNT;
}

function numFmtForSeriesAggregation(agg: AggregationMode): string {
  return agg === "sum" || agg === "avg" ? EXCEL_NUM_FMT_RUB : EXCEL_NUM_FMT_COUNT;
}

function applyChartDataNumberFormats(
  ws: ExcelJS.Worksheet,
  cfg: ChartConfig,
  dataRowCount: number,
  dataStartRow1Based: number,
): void {
  if (dataRowCount <= 0) {
    return;
  }
  for (let r = 0; r < dataRowCount; r++) {
    const row = ws.getRow(dataStartRow1Based + r);
    for (let j = 0; j < cfg.yKeys.length; j++) {
      const agg = cfg.yAggregations?.[j] ?? cfg.aggregation;
      row.getCell(j + 2).numFmt = numFmtForSeriesAggregation(agg);
    }
  }
}

function sourceHeaderLooksLikeMoney(header: string): boolean {
  const h = header.trim().toLowerCase();
  return /сумма|amount|руб|₽|стоим|цена|выруч|оплат|бюджет|чек|доход|денег/i.test(
    h,
  );
}

function applySourceSheetNumberFormats(
  ws: ExcelJS.Worksheet,
  headers: string[],
  dataRowCount: number,
  dataStartRow1Based: number,
): void {
  if (dataRowCount <= 0 || headers.length === 0) {
    return;
  }
  for (let r = 0; r < dataRowCount; r++) {
    const row = ws.getRow(dataStartRow1Based + r);
    for (let c = 0; c < headers.length; c++) {
      const cell = row.getCell(c + 1);
      const v = cell.value;
      const isFormula =
        v !== null &&
        typeof v === "object" &&
        "formula" in v &&
        typeof (v as { formula?: unknown }).formula === "string";
      if (isFormula || typeof v === "number") {
        cell.numFmt = sourceHeaderLooksLikeMoney(headers[c] ?? "")
          ? EXCEL_NUM_FMT_RUB
          : EXCEL_NUM_FMT_COUNT;
      }
    }
  }
}

/** Число в ячейке — формула =n (пересчитываемая ячейка, тот же результат). */
function cellValueForExport(
  cell: string | number,
): string | number | ExcelJS.CellFormulaValue {
  if (typeof cell === "number" && Number.isFinite(cell)) {
    return {
      formula: `=${cell}`,
      result: cell,
    };
  }
  return cell;
}

function addMatrix(ws: ExcelJS.Worksheet, aoa: (string | number)[][]) {
  for (const row of aoa) {
    ws.addRow(row.map((cell) => cellValueForExport(cell)));
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

function addMethodologyAndLegendBlocks(
  ws: ExcelJS.Worksheet,
  tabular: TabularData,
  p: PreparedChartSheet & { kind: "data" },
  span: number,
): void {
  const end = colLetter(span - 1);
  ws.addRow([]);
  const rHead = ws.rowCount;
  styleExportSectionHeading(ws, rHead, span, "Методика расчёта");
  ws.addRow([]);
  const rBody = ws.rowCount;
  ws.mergeCells(`A${rBody}:${end}${rBody}`);
  const narrative = buildMethodologyNarrative(p.spec, tabular, p.cfg);
  const bodyCell = ws.getCell(`A${rBody}`);
  bodyCell.value = narrative;
  bodyCell.font = { size: 10, color: { argb: XL.ink } };
  bodyCell.alignment = { vertical: "top", wrapText: true, indent: 1 };
  bodyCell.fill = fillSolid(XL.slate50);
  bodyCell.border = thinBorder();
  const approxLines = narrative.split(/\r?\n/).reduce((acc, line) => {
    return acc + Math.max(1, Math.ceil(line.length / 88));
  }, 0);
  ws.getRow(rBody).height = Math.min(240, Math.max(56, approxLines * 13 + 16));

  ws.addRow([]);
  const rLegTitle = ws.rowCount;
  styleExportSectionHeading(
    ws,
    rLegTitle,
    span,
    "Легенда серий (к диаграмме и таблице)",
  );
  ws.addRow([]);
  const rHdr = ws.rowCount;
  ws.getRow(rHdr).getCell(1).value = "Обозначение";
  ws.getRow(rHdr).getCell(2).value = "Описание метрики";
  if (span > 2) {
    ws.mergeCells(`B${rHdr}:${end}${rHdr}`);
  }
  for (let c = 1; c <= span; c++) {
    const cell = ws.getRow(rHdr).getCell(c);
    cell.font = { bold: true, size: 10, color: { argb: XL.white } };
    cell.fill = fillSolid(XL.sky500);
    cell.border = thinBorder(XL.sky700);
    cell.alignment = {
      vertical: "middle",
      horizontal: "left",
      wrapText: true,
      indent: c === 1 ? 1 : 0,
    };
  }
  ws.getRow(rHdr).height = 22;

  for (let i = 0; i < p.cfg.yKeys.length; i++) {
    ws.addRow([]);
    const rr = ws.rowCount;
    const desc = exportSeriesColumnHeader(
      tabular,
      p.cfg,
      i,
      p.block.id,
    );
    ws.getRow(rr).getCell(1).value = `Серия ${i + 1}`;
    ws.getRow(rr).getCell(2).value = desc;
    if (span > 2) {
      ws.mergeCells(`B${rr}:${end}${rr}`);
    }
    const zebra = i % 2 === 0 ? XL.white : XL.slate50;
    for (let c = 1; c <= span; c++) {
      const cell = ws.getRow(rr).getCell(c);
      cell.fill = fillSolid(zebra);
      cell.border = thinBorder();
      cell.font = { size: 10, color: { argb: XL.ink } };
      cell.alignment = {
        vertical: "top",
        wrapText: true,
        indent: c === 1 ? 1 : 0,
      };
    }
    ws.getRow(rr).height = Math.min(
      120,
      18 + Math.ceil(desc.length / 70) * 12,
    );
  }

  ws.getColumn(1).width = Math.max(ws.getColumn(1).width ?? 8, 18);
  ws.getColumn(2).width = Math.max(ws.getColumn(2).width ?? 8, 56);
}

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
  dealStageOrder?: string[] | null,
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
      withGlobalFilters(resolved.config, globalFilters, tabular, spec.id),
    );

    const reportChartThroughYmd =
      block.id === "deals" || block.id === "companies"
        ? REPORT_INCLUDES_THROUGH_YMD
        : undefined;

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
      ...cfg.yKeys.map((_, i) =>
        exportSeriesColumnHeader(tabular, cfg, i, block.id),
      ),
    ];

    const { data, warnings } = buildAggregatedChartRows(
      tabular.rows,
      cfg,
      tabular.columns,
      dealStageOrder,
      reportChartThroughYmd,
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
  | {
      ok: true;
      fileName: string;
      sheetsWritten: number;
      buffer?: ArrayBuffer;
      /** Встроенные диаграммы Excel — только при успешной генерации через Python (/api/export-charts). */
      chartsEmbedded: boolean;
      /** Если графики не встроены, но файл скачан (запасной путь ExcelJS). */
      warning?: string;
    }
  | { ok: false; error: string };

function enrichExportChartsErrorMessage(message: string): string {
  const m = message.toLowerCase();
  if (m.includes("no module named 'xlsxwriter'") || m.includes("no module named xlsxwriter")) {
    return `${message.trim()} Установите: pip install xlsxwriter`;
  }
  if (message.includes("NO_PYTHON")) {
    return "Не найден Python (python3 / python). Встроенные диаграммы в Excel формируются скриптом на Python 3 с пакетом xlsxwriter.";
  }
  return message.trim();
}

function fallbackChartsWarning(apiError: string): string {
  const hint = enrichExportChartsErrorMessage(apiError);
  return `Без встроенных диаграмм (в файле — таблицы и сводка). Не удалось вызвать генератор с графиками: ${hint}`;
}

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
  dealStageOrder?: string[] | null,
  mainInfo?: ExcelMainInfoLine[],
): Promise<ExportChartsResult> {
  try {
    const apiUrl =
      typeof window !== "undefined"
        ? new URL("/api/export-charts", window.location.origin).toString()
        : "/api/export-charts";
    const charts = buildLiveChartItems(
      tabular,
      enabledChartIds,
      globalFilters,
      dealStageOrder?.length ? dealStageOrder : null,
    );
    if (charts.length === 0) {
      return {
        ok: false,
        error:
          "Нет данных для построения диаграмм по выбранным отчётам — проверьте сущность и графики.",
      };
    }
    const sourceData = buildSourceData(tabular);
    const summary = {
      createdAt: new Date().toLocaleString("ru-RU", {
        dateStyle: "short",
        timeStyle: "medium",
      }),
      sourceFileName: sourceFileName ?? "",
      activeSheet: tabular.activeSheet,
      columnsTotal: tabular.columns.length,
      mainInfo: mainInfo?.length ? mainInfo : [],
    };
    const res = await fetch(apiUrl, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        charts,
        sourceData,
        summary,
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
    return { ok: true, fileName, sheetsWritten, chartsEmbedded: true };
  } catch {
    return {
      ok: false,
      error: "Не удалось скачать файл с диаграммами. Проверьте соединение и повторите.",
    };
  }
}

/** Методика и легенда под таблицей. */
function appendMethodologyAndLegendForDataSheet(
  worksheet: ExcelJS.Worksheet,
  tabular: TabularData,
  p: PreparedChartSheet & { kind: "data" },
  colSpan: number,
): void {
  const span = Math.max(4, colSpan);
  addMethodologyAndLegendBlocks(worksheet, tabular, p, span);
}

export async function exportChartsToExcelFile(
  tabular: TabularData,
  enabledChartIds: Set<string>,
  globalFilters: ChartFilter[],
  options?: {
    fileNameBase?: string;
    sourceFileName?: string;
    theme?: ExcelExportTheme;
    /** Порядок стадий сделки (как в боковой панели). */
    dealStageOrder?: string[] | null;
    /** Для тестов: не скачивать файл, вернуть буфер книги. */
    returnBuffer?: boolean;
    /** Блок «Основные показатели» на первом листе (как в боковой сводке). */
    mainInfo?: ExcelMainInfoLine[];
  },
): Promise<ExportChartsResult> {
  if (enabledChartIds.size === 0) {
    return {
      ok: false,
      error:
        "Не выбран ни один график: отметьте нужные отчёты галочками в боковой панели.",
    };
  }

  const isBrowser =
    typeof window !== "undefined" && typeof document !== "undefined";
  const tryApiFirst = !options?.returnBuffer && isBrowser;

  let apiErrorForFallback: string | null = null;
  if (tryApiFirst) {
    const apiResult = await exportChartsViaApi(
      tabular,
      enabledChartIds,
      globalFilters,
      options?.fileNameBase,
      options?.sourceFileName,
      options?.theme ?? "classic",
      options?.dealStageOrder,
      options?.mainInfo,
    );
    if (apiResult.ok) {
      return apiResult;
    }
    apiErrorForFallback = apiResult.error;
    /* Сервер недоступен или Python/xlsxwriter не собрали файл — выгрузка через ExcelJS
       (таблицы и методика; встроенные диаграммы Excel только при успешном ответе /api). */
  }

  const prepared = prepareChartSheets(
    tabular,
    enabledChartIds,
    globalFilters,
    options?.dealStageOrder,
  );

  if (prepared.length === 0) {
    return {
      ok: false,
      error:
        "Не удалось подготовить листы: проверьте соответствие выгрузки типу графиков.",
    };
  }

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Отчёты Excel · Битрикс";
  workbook.created = new Date();
  const usedSheetNames = new Set<string>();

  const summaryStamp = new Date().toLocaleString("ru-RU", {
    dateStyle: "short",
    timeStyle: "medium",
  });

  const mainLines = options?.mainInfo ?? [];
  const summaryAoa: (string | number)[][] = [
    ["Общая информация"],
    [],
  ];
  if (mainLines.length > 0) {
    summaryAoa.push(["Основные показатели", "", ""]);
    for (const line of mainLines) {
      summaryAoa.push([line.label, line.value, ""]);
    }
    summaryAoa.push([]);
  }
  summaryAoa.push(
    ["Дата формирования", summaryStamp, ""],
    ...(options?.sourceFileName
      ? ([["Исходный файл", options.sourceFileName, ""]] as (string | number)[][])
      : []),
    ["Графиков в файле", prepared.length, ""],
    [],
    ["Раздел", "График", "Статус"],
    ...prepared.map((p) => [
      p.block.label,
      p.spec.title,
      summaryStatusText(p),
    ]),
  );

  const summaryName = sanitizeSheetName("Общая информация", usedSheetNames);
  const wsSummary = workbook.addWorksheet(summaryName);
  addMatrix(wsSummary, summaryAoa);
  const chartTableHeaderIdx = summaryAoa.findIndex(
    (r) => r[0] === "Раздел" && r[1] === "График",
  );
  const summaryTableHeaderRow =
    chartTableHeaderIdx >= 0 ? chartTableHeaderIdx + 1 : 9;
  styleSummarySheet(wsSummary, summaryTableHeaderRow, prepared.length, mainLines);

  let sheetsWritten = 1;

  const { headers: srcHeaders, rows: srcRows } = buildSourceData(tabular);
  if (srcHeaders.length > 0) {
    const srcSheetName = sanitizeSheetName("Исходные данные", usedSheetNames);
    const wsSrc = workbook.addWorksheet(srcSheetName);
    const srcColCount = srcHeaders.length;
    addMatrix(wsSrc, [
      ["Исходные данные"],
      [],
      srcHeaders,
      ...srcRows,
    ]);
    styleChartTitleRow(wsSrc, srcColCount, "Исходные данные");
    styleDataTableBlock(wsSrc, 3, srcColCount, srcRows.length);
    applySourceSheetNumberFormats(wsSrc, srcHeaders, srcRows.length, 4);
    sheetsWritten++;
  }

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
    applyChartDataNumberFormats(ws, p.cfg, p.dataRows.length, 4);

    appendMethodologyAndLegendForDataSheet(ws, tabular, p, colCount);
  }

  const fileName = buildExportDownloadFileName(options?.fileNameBase);

  const rawBuffer = await workbook.xlsx.writeBuffer();
  const arrayBuffer = toArrayBuffer(rawBuffer);

  const chartsEmbedded = false;
  const warning =
    tryApiFirst && apiErrorForFallback
      ? fallbackChartsWarning(apiErrorForFallback)
      : undefined;

  if (options?.returnBuffer) {
    return {
      ok: true,
      fileName,
      sheetsWritten,
      buffer: arrayBuffer,
      chartsEmbedded,
      ...(warning ? { warning } : {}),
    };
  }

  if (typeof window !== "undefined" && typeof document !== "undefined") {
    downloadBlob(fileName, rawBuffer);
  }

  return {
    ok: true,
    fileName,
    sheetsWritten,
    chartsEmbedded,
    ...(warning ? { warning } : {}),
  };
}
