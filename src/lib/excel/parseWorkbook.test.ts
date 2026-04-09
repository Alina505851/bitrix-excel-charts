import * as XLSX from "xlsx";
import { describe, expect, it } from "vitest";
import {
  inferColumnTypes,
  parseWorkbookFromBuffer,
} from "@/lib/excel/parseWorkbook";

describe("parseWorkbookFromBuffer (как после загрузки файла)", () => {
  it("читает сгенерированный .xlsx из ArrayBuffer", () => {
    const wb = XLSX.utils.book_new();
    const aoa = [
      ["Стадия", "Сумма"],
      ["Новая", 1000],
      ["В работе", 2000],
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    XLSX.utils.book_append_sheet(wb, ws, "Сделки");
    const buffer = XLSX.write(wb, {
      type: "array",
      bookType: "xlsx",
    }) as ArrayBuffer;

    const result = parseWorkbookFromBuffer(buffer);

    expect(result.ok).toBe(true);
    if (!result.ok) {
      return;
    }
    expect(result.data.sheetNames).toContain("Сделки");
    expect(result.data.activeSheet).toBe("Сделки");
    expect(result.data.columns.length).toBe(2);
    expect(result.data.rows.length).toBe(2);
    const stageKey = result.data.columns.find((c) => c.header === "Стадия")?.key;
    const sumKey = result.data.columns.find((c) => c.header === "Сумма")?.key;
    expect(stageKey).toBeDefined();
    expect(sumKey).toBeDefined();
    expect(result.data.rows[0]?.[stageKey!]).toBe("Новая");
    expect(result.data.rows[0]?.[sumKey!]).toBe(1000);
    expect(result.data.rows[1]?.[sumKey!]).toBe(2000);
  });

  it("переключает лист по имени", () => {
    const wb = XLSX.utils.book_new();
    const ws1 = XLSX.utils.aoa_to_sheet([
      ["A"],
      [1],
    ]);
    XLSX.utils.book_append_sheet(wb, ws1, "Лист1");
    const ws2 = XLSX.utils.aoa_to_sheet([
      ["B"],
      [2],
    ]);
    XLSX.utils.book_append_sheet(wb, ws2, "Лист2");
    const buffer = XLSX.write(wb, {
      type: "array",
      bookType: "xlsx",
    }) as ArrayBuffer;

    const r2 = parseWorkbookFromBuffer(buffer, "Лист2");
    expect(r2.ok).toBe(true);
    if (!r2.ok) {
      return;
    }
    expect(r2.data.activeSheet).toBe("Лист2");
    expect(r2.data.rows[0]?.B).toBe(2);
  });

  it("при несуществующем имени листа выбирается первый лист книги", () => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
      ["Колонка"],
      [42],
    ]);
    XLSX.utils.book_append_sheet(wb, ws, "Данные");
    const buffer = XLSX.write(wb, {
      type: "array",
      bookType: "xlsx",
    }) as ArrayBuffer;
    const result = parseWorkbookFromBuffer(buffer, "НетТакогоЛиста");
    expect(result.ok).toBe(true);
    if (!result.ok) {
      return;
    }
    expect(result.data.activeSheet).toBe("Данные");
    expect(result.data.rows[0]?.["Колонка"]).toBe(42);
  });
});

describe("inferColumnTypes", () => {
  it("detects number-heavy column", () => {
    const cols = [{ key: "amt", header: "Сумма" }];
    const rows = Array.from({ length: 10 }, () => ({ amt: 100 }));
    const meta = inferColumnTypes(cols, rows);
    expect(meta[0]?.inferredType).toBe("number");
  });

  it("detects string column", () => {
    const cols = [{ key: "stage", header: "Стадия" }];
    const rows = [
      { stage: "Новая" },
      { stage: "В работе" },
      { stage: "Успех" },
    ];
    const meta = inferColumnTypes(cols, rows);
    expect(meta[0]?.inferredType).toBe("string");
  });
});
