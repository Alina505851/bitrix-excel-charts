import * as XLSX from "xlsx";
import { tryParseDate } from "@/lib/chart/coerce";
import type { ColumnMeta, InferredType, TabularData } from "@/lib/types";

const SAMPLE_LIMIT = 50;
const TYPE_SAMPLE_ROWS = 400;

function dedupeHeaders(rawHeaders: unknown[]): string[] {
  const counts = new Map<string, number>();
  return rawHeaders.map((cell, index) => {
    const base =
      cell !== null && cell !== undefined && String(cell).trim() !== ""
        ? String(cell).trim()
        : `column_${index + 1}`;
    const n = (counts.get(base) ?? 0) + 1;
    counts.set(base, n);
    return n === 1 ? base : `${base}_${n}`;
  });
}

function tryParseNumber(value: unknown): number | null {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  if (typeof value !== "string") {
    return null;
  }
  const normalized = value.replace(/\s/g, "").replace(",", ".");
  if (normalized === "" || normalized === "-") {
    return null;
  }
  const n = Number(normalized);
  return Number.isFinite(n) ? n : null;
}

function inferTypeFromValues(values: unknown[]): InferredType {
  let numberHits = 0;
  let dateHits = 0;
  let total = 0;
  for (const v of values) {
    if (v === null || v === undefined || v === "") {
      continue;
    }
    total++;
    if (v instanceof Date || tryParseDate(v)) {
      dateHits++;
      continue;
    }
    if (tryParseNumber(v) !== null) {
      numberHits++;
    }
  }
  if (total === 0) {
    return "unknown";
  }
  if (dateHits / total >= 0.65) {
    return "date";
  }
  if (numberHits / total >= 0.65) {
    return "number";
  }
  return "string";
}

export function inferColumnTypes(
  columns: { key: string; header: string }[],
  rows: Record<string, unknown>[],
): ColumnMeta[] {
  return columns.map((col) => {
    const samples: unknown[] = [];
    for (let i = 0; i < rows.length && samples.length < TYPE_SAMPLE_ROWS; i++) {
      const v = rows[i][col.key];
      if (v === null || v === undefined || v === "") {
        continue;
      }
      samples.push(v);
    }
    const displaySamples = samples.slice(0, SAMPLE_LIMIT);
    const inferredType = inferTypeFromValues(samples);
    return {
      key: col.key,
      header: col.header,
      inferredType,
      sampleValues: displaySamples,
    };
  });
}

function buildRows(
  headerKeys: string[],
  body: unknown[][],
): Record<string, unknown>[] {
  const rows: Record<string, unknown>[] = [];
  const norm = (v: unknown) => String(v ?? "").trim().toLowerCase();
  for (const line of body) {
    const row: Record<string, unknown> = {};
    let nonEmptyCount = 0;
    let headerEchoCount = 0;
    for (let i = 0; i < headerKeys.length; i++) {
      const value = line[i] ?? "";
      row[headerKeys[i]] = value;
      if (value !== "" && value !== null && value !== undefined) {
        nonEmptyCount++;
        if (norm(value) === norm(headerKeys[i])) {
          headerEchoCount++;
        }
      }
    }
    const hasValue = nonEmptyCount > 0;
    const isHeaderEchoNoise = nonEmptyCount > 0 && headerEchoCount === nonEmptyCount;
    if (hasValue && !isHeaderEchoNoise) {
      rows.push(row);
    }
  }
  return rows;
}

function normalizeExcelDateObjectsToCalendarDay(
  rows: Record<string, unknown>[],
): Record<string, unknown>[] {
  if (rows.length === 0) {
    return rows;
  }
  return rows.map((row) => {
    const next: Record<string, unknown> = { ...row };
    for (const [k, v] of Object.entries(next)) {
      if (!(v instanceof Date) || Number.isNaN(v.getTime())) {
        continue;
      }
      // Excel date-only ячейки из некоторых выгрузок приходят как 23:59:43
      // предыдущего календарного дня. Корректируем только этот артефакт,
      // не трогая реальные вечерние datetime.
      if (
        v.getHours() === 23 &&
        v.getMinutes() === 59 &&
        v.getSeconds() === 43
      ) {
        const shifted = new Date(v.getFullYear(), v.getMonth(), v.getDate() + 1);
        next[k] = new Date(
          shifted.getFullYear(),
          shifted.getMonth(),
          shifted.getDate(),
        );
      }
    }
    return next;
  });
}

export type ParseResult =
  | { ok: true; data: TabularData }
  | { ok: false; error: string };

const MAX_ROWS_WARN = 50_000;

export function parseWorkbookFromBuffer(
  buffer: ArrayBuffer,
  activeSheet?: string,
): ParseResult {
  let workbook: XLSX.WorkBook;
  try {
    workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  } catch {
    return { ok: false, error: "Не удалось прочитать файл. Проверьте формат (.xlsx / .xls)." };
  }

  const sheetNames = workbook.SheetNames.filter(Boolean);
  if (sheetNames.length === 0) {
    return { ok: false, error: "В книге нет листов." };
  }

  const sheetName =
    activeSheet && sheetNames.includes(activeSheet)
      ? activeSheet
      : sheetNames[0]!;
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) {
    return { ok: false, error: `Лист «${sheetName}» не найден.` };
  }

  const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
    header: 1,
    defval: "",
    raw: true,
  });

  if (!matrix.length) {
    return {
      ok: true,
      data: {
        sheetNames,
        activeSheet: sheetName,
        columns: [],
        rows: [],
      },
    };
  }

  const headerRow = matrix[0] as unknown[];
  const keys = dedupeHeaders(headerRow);
  const body = matrix.slice(1) as unknown[][];
  let rows = buildRows(keys, body);
  if (rows.length > MAX_ROWS_WARN) {
    rows = rows.slice(0, MAX_ROWS_WARN);
  }
  rows = normalizeExcelDateObjectsToCalendarDay(rows);

  const columnDefs = keys.map((key, i) => ({
    key,
    header: String(headerRow[i] ?? key),
  }));
  const columns = inferColumnTypes(columnDefs, rows);

  return {
    ok: true,
    data: {
      sheetNames,
      activeSheet: sheetName,
      columns,
      rows,
    },
  };
}
