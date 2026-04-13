import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import type { ColumnMeta, TabularData } from "@/lib/types";

function leadTabular(rows: Record<string, unknown>[]): TabularData {
  const columns: ColumnMeta[] = [
    {
      key: "st",
      header: "Стадия лида",
      inferredType: "string",
      sampleValues: [],
    },
    {
      key: "amt",
      header: "Сумма",
      inferredType: "number",
      sampleValues: [],
    },
  ];
  return {
    sheetNames: ["Лиды"],
    activeSheet: "Лиды",
    columns,
    rows,
  };
}

describe("exportChartsToExcelFile", () => {
  it("returns error and does not return buffer when no charts selected", async () => {
    const { exportChartsToExcelFile } = await import(
      "@/lib/excel/exportChartsExcel"
    );
    const res = await exportChartsToExcelFile(leadTabular([]), new Set(), [], {
      returnBuffer: true,
    });
    expect(res.ok).toBe(false);
    if (res.ok) {
      return;
    }
    expect(res.buffer).toBeUndefined();
  });

  it("writes workbook with overview sheet first, then chart worksheets", async () => {
    const { exportChartsToExcelFile } = await import(
      "@/lib/excel/exportChartsExcel"
    );
    const res = await exportChartsToExcelFile(
      leadTabular([
        { st: "Новая", amt: 100 },
        { st: "Новая", amt: 50 },
      ]),
      new Set(["leads_sum_by_stage"]),
      [],
      { fileNameBase: "report", sourceFileName: "source.xlsx", returnBuffer: true },
    );
    expect(res.ok).toBe(true);
    if (!res.ok) {
      return;
    }
    expect(res.fileName).toMatch(
      /^report_графики_\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}\.xlsx$/,
    );
    expect(res.sheetsWritten).toBe(3);
    expect(res.buffer).toBeDefined();
    expect(res.chartsEmbedded).toBe(false);

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(res.buffer!);
    expect(wb.worksheets.length).toBe(3);
    expect(wb.worksheets[0]!.name).toContain("Общая");
    expect(wb.worksheets[1]!.name).toContain("Исходн");
    expect(wb.worksheets[2]!.name).toMatch(/Лиды/i);
  });

  it("includes error sheet when chart cannot resolve", async () => {
    const { exportChartsToExcelFile } = await import(
      "@/lib/excel/exportChartsExcel"
    );
    const thin: TabularData = {
      sheetNames: ["S"],
      activeSheet: "S",
      columns: [
        {
          key: "n",
          header: "Только число",
          inferredType: "number",
          sampleValues: [],
        },
      ],
      rows: [{ n: 1 }],
    };
    const res = await exportChartsToExcelFile(
      thin,
      new Set(["companies_count_by_source"]),
      [],
      { returnBuffer: true },
    );
    expect(res.ok).toBe(true);
    if (!res.ok) {
      return;
    }
    expect(res.sheetsWritten).toBe(3);
    expect(res.chartsEmbedded).toBe(false);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(res.buffer!);
    expect(wb.worksheets.length).toBe(3);
  });

  it("writes empty-state sheet when filters exclude all rows", async () => {
    const { exportChartsToExcelFile } = await import(
      "@/lib/excel/exportChartsExcel"
    );
    const res = await exportChartsToExcelFile(
      leadTabular([{ st: "Новая", amt: 100 }]),
      new Set(["leads_sum_by_stage"]),
      [{ columnKey: "st", values: ["Другая"] }],
      { returnBuffer: true },
    );
    expect(res.ok).toBe(true);
    if (!res.ok) {
      return;
    }
    expect(res.chartsEmbedded).toBe(false);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(res.buffer!);
    expect(wb.worksheets.length).toBe(3);
    const ws = wb.worksheets[2]!;
    let flat = "";
    ws.eachRow((row) => {
      row.eachCell((cell) => {
        flat += `${cell.value ?? ""} `;
      });
    });
    expect(flat).toMatch(/Нет данных/i);
  });
});
