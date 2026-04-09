import * as XLSX from "xlsx";
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

const RU_DATE = /^(\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/;
const ISO_DATE =
  /^(\d{4})-(\d{2})-(\d{2})(?:[T\s](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/;
const ISO_TIMESTAMP_TZ =
  /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:\d{2})$/i;

function expandYear(y: string): number {
  if (y.length === 4) {
    return Number(y);
  }
  const yy = Number(y);
  return yy >= 70 ? 1900 + yy : 2000 + yy;
}

function buildStrictLocalDate(
  year: number,
  month: number,
  day: number,
  hour: number,
  minute: number,
  second: number,
): Date | null {
  if (
    !Number.isInteger(year) ||
    !Number.isInteger(month) ||
    !Number.isInteger(day) ||
    !Number.isInteger(hour) ||
    !Number.isInteger(minute) ||
    !Number.isInteger(second)
  ) {
    return null;
  }
  if (
    month < 1 ||
    month > 12 ||
    day < 1 ||
    day > 31 ||
    hour < 0 ||
    hour > 23 ||
    minute < 0 ||
    minute > 59 ||
    second < 0 ||
    second > 59
  ) {
    return null;
  }
  const dt = new Date(year, month - 1, day, hour, minute, second);
  if (Number.isNaN(dt.getTime())) {
    return null;
  }
  if (
    dt.getFullYear() !== year ||
    dt.getMonth() !== month - 1 ||
    dt.getDate() !== day ||
    dt.getHours() !== hour ||
    dt.getMinutes() !== minute ||
    dt.getSeconds() !== second
  ) {
    return null;
  }
  return dt;
}

function excelSerialToDate(value: number): Date | null {
  if (!Number.isFinite(value)) {
    return null;
  }
  if (value < 20_000 || value > 80_000) {
    return null;
  }
  const ms = Math.round((value - 25569) * 86400 * 1000);
  const dt = new Date(ms);
  return Number.isNaN(dt.getTime()) ? null : dt;
}

function tryParseDate(value: unknown): Date | null {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value;
  }
  if (typeof value === "number") {
    return excelSerialToDate(value);
  }
  if (typeof value !== "string") {
    return null;
  }
  const s = value.trim();
  if (!s) {
    return null;
  }
  const ru = s.match(RU_DATE);
  if (ru) {
    const [, d, m, y, hh = "0", mm = "0", ss = "0"] = ru;
    return buildStrictLocalDate(
      expandYear(y),
      Number(m),
      Number(d),
      Number(hh),
      Number(mm),
      Number(ss),
    );
  }
  const iso = s.match(ISO_DATE);
  if (iso) {
    const [, y, m, d, hh = "0", mm = "0", ss = "0"] = iso;
    return buildStrictLocalDate(
      Number(y),
      Number(m),
      Number(d),
      Number(hh),
      Number(mm),
      Number(ss),
    );
  }
  if (ISO_TIMESTAMP_TZ.test(s)) {
    const dt = new Date(s);
    return Number.isNaN(dt.getTime()) ? null : dt;
  }
  return null;
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
  for (const line of body) {
    const row: Record<string, unknown> = {};
    for (let i = 0; i < headerKeys.length; i++) {
      row[headerKeys[i]] = line[i] ?? "";
    }
    const hasValue = Object.values(row).some(
      (v) => v !== "" && v !== null && v !== undefined,
    );
    if (hasValue) {
      rows.push(row);
    }
  }
  return rows;
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
