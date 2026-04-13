import { describe, expect, it } from "vitest";
import { buildLiveChartItems } from "@/lib/excel/buildLiveChartItems";
import type { ColumnMeta, TabularData } from "@/lib/types";

function tabular(cols: ColumnMeta[], rows: Record<string, unknown>[]): TabularData {
  return {
    sheetNames: ["S"],
    activeSheet: "S",
    columns: cols,
    rows,
  };
}

describe("buildLiveChartItems", () => {
  it("ставит linkToSourcePossible для текстовой оси без фильтров (формулы SUMIFS по точному совпадению X)", () => {
    const cols: ColumnMeta[] = [
      {
        key: "st",
        header: "Стадия",
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
    const data = tabular(cols, [
      { st: "Новая", amt: 10 },
      { st: "Успех", amt: 20 },
    ]);
    const items = buildLiveChartItems(
      data,
      new Set(["deals_sum_by_stage"]),
      [],
      null,
    );
    const item = items.find((x) => x.title === "Сумма по стадиям сделки");
    expect(item).toBeDefined();
    expect(item?.chartId).toBe("deals_sum_by_stage");
    expect(item?.linkToSourcePossible).toBe(true);
    expect(item?.xIsDate).toBe(false);
    expect(item?.reportThroughYmd).toBe("2026-04-12");
  });

  it("ставит linkToSourcePossible для оси-даты с шагом день без фильтров", () => {
    const cols: ColumnMeta[] = [
      {
        key: "d",
        header: "Дата создания сделки",
        inferredType: "date",
        sampleValues: [],
      },
      {
        key: "amt",
        header: "Сумма",
        inferredType: "number",
        sampleValues: [],
      },
    ];
    const data = tabular(cols, [{ d: "2024-01-05", amt: 1 }]);
    const items = buildLiveChartItems(
      data,
      new Set(["deals_line_avg_by_date"]),
      [],
      null,
    );
    const item = items.find((x) => x.title === "Средний чек по дням");
    expect(item?.chartDataEmpty).toBe(false);
    expect(item?.xIsDate).toBe(true);
    expect(item?.linkToSourcePossible).toBe(true);
    expect(item?.reportThroughYmd).toBe("2026-04-12");
  });

  it("для лидов с датой без фильтров оставляет linkToSourcePossible (нет границы отчёта по дате)", () => {
    const cols: ColumnMeta[] = [
      {
        key: "d",
        header: "Дата создания",
        inferredType: "date",
        sampleValues: [],
      },
      {
        key: "n",
        header: "N",
        inferredType: "number",
        sampleValues: [],
      },
    ];
    const data = tabular(cols, [{ d: "2024-01-05", n: 1 }]);
    const items = buildLiveChartItems(
      data,
      new Set(["leads_count_by_created_date"]),
      [],
      null,
    );
    const item = items.find((x) => x.title === "Приток лидов по датам");
    expect(item?.chartDataEmpty).toBe(false);
    expect(item?.linkToSourcePossible).toBe(true);
  });

  it("распознаёт строковую дату сделок как ось-даты; для графика с фильтрами формулы от источника отключены", () => {
    const cols: ColumnMeta[] = [
      {
        key: "d",
        header: "Дата отправки КП",
        inferredType: "string",
        sampleValues: ["15.01.2025", "28.02.2025"],
      },
      {
        key: "n",
        header: "N",
        inferredType: "number",
        sampleValues: [],
      },
    ];
    const data = tabular(cols, [{ d: "15.01.2025", n: 1 }]);
    const items = buildLiveChartItems(
      data,
      new Set(["deals_kp_sent_count_by_period"]),
      [],
      null,
    );
    const item = items.find((x) => x.chartId === "deals_kp_sent_count_by_period");
    expect(item?.xIsDate).toBe(true);
    expect(item?.linkToSourcePossible).toBe(false);
  });
});
