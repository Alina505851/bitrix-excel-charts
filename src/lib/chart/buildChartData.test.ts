import { describe, expect, it } from "vitest";
import {
  applyFilters,
  applyFiltersExcept,
  buildAggregatedChartRows,
  distinctStringValues,
} from "@/lib/chart/buildChartData";
import type { ChartConfig, ColumnMeta } from "@/lib/types";

const columns: ColumnMeta[] = [
  {
    key: "stage",
    header: "Стадия",
    inferredType: "string",
    sampleValues: [],
  },
  {
    key: "amount",
    header: "Сумма",
    inferredType: "number",
    sampleValues: [],
  },
];

describe("buildAggregatedChartRows", () => {
  it("sums by category", () => {
    const rows = [
      { stage: "A", amount: 10 },
      { stage: "A", amount: 20 },
      { stage: "B", amount: 5 },
    ];
    const cfg: ChartConfig = {
      xKey: "stage",
      yKeys: ["amount"],
      chartType: "bar",
      aggregation: "sum",
      filters: [],
    };
    const { data, warnings } = buildAggregatedChartRows(rows, cfg, columns);
    expect(warnings.length).toBe(0);
    expect(data.find((d) => d.name === "A")?.amount).toBe(30);
    expect(data.find((d) => d.name === "B")?.amount).toBe(5);
  });

  it("applies string filter", () => {
    const rows = [
      { stage: "A", amount: 10 },
      { stage: "B", amount: 99 },
    ];
    const cfg: ChartConfig = {
      xKey: "stage",
      yKeys: ["amount"],
      chartType: "bar",
      aggregation: "sum",
      filters: [{ columnKey: "stage", values: ["A"] }],
    };
    const { data } = buildAggregatedChartRows(rows, cfg, columns);
    expect(data.length).toBe(1);
    expect(data[0]?.name).toBe("A");
  });

  it("averages metric by category", () => {
    const rows = [
      { stage: "A", amount: 10 },
      { stage: "A", amount: 30 },
      { stage: "B", amount: 100 },
    ];
    const cfg: ChartConfig = {
      xKey: "stage",
      yKeys: ["amount"],
      chartType: "bar",
      aggregation: "avg",
      filters: [],
    };
    const { data, warnings } = buildAggregatedChartRows(rows, cfg, columns);
    expect(warnings.length).toBe(0);
    expect(data.find((d) => d.name === "A")?.amount).toBe(20);
    expect(data.find((d) => d.name === "B")?.amount).toBe(100);
  });

  it("cumulative sums period counts in order along X", () => {
    const dateCols: ColumnMeta[] = [
      {
        key: "d",
        header: "Дата",
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
    const rows = [
      { d: new Date(2024, 8, 5), n: 1 },
      { d: new Date(2024, 8, 15), n: 1 },
      { d: new Date(2024, 9, 2), n: 1 },
      { d: new Date(2024, 9, 10), n: 1 },
      { d: new Date(2024, 9, 20), n: 1 },
    ];
    const cfg: ChartConfig = {
      xKey: "d",
      yKeys: ["n"],
      chartType: "area",
      aggregation: "count",
      filters: [],
      dateGranularity: "month",
      cumulative: true,
    };
    const { data, warnings } = buildAggregatedChartRows(rows, cfg, dateCols);
    expect(warnings.length).toBe(0);
    const sep = data.find((x) => x.name === "сент. 2024");
    const oct = data.find((x) => x.name === "окт. 2024");
    expect(sep?.n).toBe(2);
    expect(oct?.n).toBe(5);
  });

  it("count_nonempty counts filled Y cells, not sum of ids", () => {
    const rows = [
      { ind: "IT", emp: 100500 },
      { ind: "IT", emp: 200600 },
      { ind: "IT", emp: null },
      { ind: "Trade", emp: 300 },
    ];
    const cols: ColumnMeta[] = [
      { key: "ind", header: "Отрасль", inferredType: "string", sampleValues: [] },
      { key: "emp", header: "ID сотрудника", inferredType: "number", sampleValues: [] },
    ];
    const cfg: ChartConfig = {
      xKey: "ind",
      yKeys: ["emp"],
      chartType: "bar",
      aggregation: "count_nonempty",
      filters: [],
    };
    const { data, warnings } = buildAggregatedChartRows(rows, cfg, cols);
    expect(warnings.length).toBe(0);
    expect(data.find((d) => d.name === "IT")?.emp).toBe(2);
    expect(data.find((d) => d.name === "Trade")?.emp).toBe(1);
  });

  it("groups date X-axis by calendar month", () => {
    const dateCols: ColumnMeta[] = [
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
    const rows = [
      { d: new Date(2024, 8, 5), n: 1 },
      { d: new Date(2024, 8, 20), n: 1 },
      { d: new Date(2024, 9, 2), n: 1 },
    ];
    const cfg: ChartConfig = {
      xKey: "d",
      yKeys: ["n"],
      chartType: "line",
      aggregation: "count",
      filters: [],
      dateGranularity: "month",
    };
    const { data, warnings } = buildAggregatedChartRows(rows, cfg, dateCols);
    expect(warnings.length).toBe(0);
    expect(data.find((x) => x.name === "сент. 2024")?.n).toBe(2);
    expect(data.find((x) => x.name === "окт. 2024")?.n).toBe(1);
  });

  it("groups date X-axis by calendar quarter", () => {
    const dateCols: ColumnMeta[] = [
      { key: "d", header: "Дата", inferredType: "date", sampleValues: [] },
      { key: "n", header: "N", inferredType: "number", sampleValues: [] },
    ];
    const rows = [
      { d: new Date(2024, 0, 10), n: 1 },
      { d: new Date(2024, 1, 5), n: 1 },
      { d: new Date(2024, 3, 1), n: 1 },
    ];
    const cfg: ChartConfig = {
      xKey: "d",
      yKeys: ["n"],
      chartType: "line",
      aggregation: "count",
      filters: [],
      dateGranularity: "quarter",
    };
    const { data, warnings } = buildAggregatedChartRows(rows, cfg, dateCols);
    expect(warnings.length).toBe(0);
    expect(data.find((x) => x.name === "1 кв. 2024")?.n).toBe(2);
    expect(data.find((x) => x.name === "2 кв. 2024")?.n).toBe(1);
  });

  it("groups RU datetime strings (DD.MM.YYYY HH:mm) into correct month", () => {
    const dateCols: ColumnMeta[] = [
      {
        key: "d",
        header: "Дата создания",
        inferredType: "date",
        sampleValues: [],
      },
      {
        key: "n",
        header: "ID",
        inferredType: "number",
        sampleValues: [],
      },
    ];
    const rows = [
      { d: "07.04.2026 16:32", n: 1 },
      { d: "07.04.2026 15:19", n: 2 },
      { d: "06.04.2026 13:00", n: 3 },
      { d: "06.04.2026 12:52", n: 4 },
      { d: "06.04.2026 12:48", n: 5 },
    ];
    const cfg: ChartConfig = {
      xKey: "d",
      yKeys: ["n"],
      chartType: "line",
      aggregation: "count",
      filters: [],
      dateGranularity: "month",
    };
    const { data, warnings } = buildAggregatedChartRows(rows, cfg, dateCols);
    expect(warnings.length).toBe(0);
    expect(data).toHaveLength(1);
    expect(data[0]?.name).toBe("апр. 2026");
    expect(data[0]?.n).toBe(5);
  });

  it("groups ISO Z timestamps into calendar months", () => {
    const dateCols: ColumnMeta[] = [
      {
        key: "d",
        header: "Дата создания",
        inferredType: "date",
        sampleValues: [],
      },
      {
        key: "n",
        header: "ID",
        inferredType: "number",
        sampleValues: [],
      },
    ];
    const rows = [
      { d: "2025-07-30T17:24:36.000Z", n: 1 },
      { d: "2025-08-01T12:00:00.000Z", n: 2 },
      { d: "2025-08-10T09:10:00.000Z", n: 3 },
    ];
    const cfg: ChartConfig = {
      xKey: "d",
      yKeys: ["n"],
      chartType: "line",
      aggregation: "count",
      filters: [],
      dateGranularity: "month",
    };
    const { data, warnings } = buildAggregatedChartRows(rows, cfg, dateCols);
    expect(warnings.length).toBe(0);
    expect(data.find((x) => x.name === "июл. 2025")?.n).toBe(1);
    expect(data.find((x) => x.name === "авг. 2025")?.n).toBe(2);
  });
});

describe("distinctStringValues", () => {
  it("collects unique labels with limit", () => {
    const rows = [
      { city: "Мск" },
      { city: "Спб" },
      { city: "Мск" },
      { city: "" },
    ];
    const v = distinctStringValues(rows, "city", 10);
    expect(v).toContain("Мск");
    expect(v).toContain("Спб");
    expect(v).toContain("(пусто)");
  });
});

describe("applyFilters", () => {
  it("filters by date range (ISO strings)", () => {
    const rows = [
      { d: "2024-01-15" },
      { d: "2024-02-01" },
      { d: "2024-03-10" },
    ];
    const f = [
      { columnKey: "d", dateFrom: "2024-02-01", dateTo: "2024-03-01" },
    ];
    const out = applyFilters(rows, f);
    expect(out.map((r) => r.d)).toEqual(["2024-02-01"]);
  });

  it("filters by calendar months (YYYY-MM)", () => {
    const rows = [
      { d: "2024-01-15" },
      { d: "2024-02-01" },
      { d: "2024-03-10" },
    ];
    const out = applyFilters(rows, [
      { columnKey: "d", calendarMonths: ["2024-01", "2024-03"] },
    ]);
    expect(out.map((r) => r.d)).toEqual(["2024-01-15", "2024-03-10"]);
  });

  it("filters by calendar quarters", () => {
    const rows = [
      { d: "2024-01-15" },
      { d: "2024-04-01" },
      { d: "2024-06-30" },
    ];
    const out = applyFilters(rows, [
      { columnKey: "d", calendarQuarters: ["2024-Q2"] },
    ]);
    expect(out.map((r) => r.d)).toEqual(["2024-04-01", "2024-06-30"]);
  });

  it("combines date range with calendar months (AND)", () => {
    const rows = [
      { d: "2024-01-20" },
      { d: "2024-02-15" },
      { d: "2024-03-05" },
    ];
    const out = applyFilters(rows, [
      {
        columnKey: "d",
        dateFrom: "2024-01-01",
        dateTo: "2024-02-28",
        calendarMonths: ["2024-02", "2024-03"],
      },
    ]);
    expect(out.map((r) => r.d)).toEqual(["2024-02-15"]);
  });

  it("combines several conditions", () => {
    const rows = [
      { region: "Мск", amount: 100 },
      { region: "Спб", amount: 50 },
      { region: "Мск", amount: 5 },
    ];
    const out = applyFilters(rows, [
      { columnKey: "region", values: ["Мск"] },
      { columnKey: "amount", min: 10, max: 200 },
    ]);
    expect(out).toEqual([{ region: "Мск", amount: 100 }]);
  });
});

describe("applyFiltersExcept", () => {
  it("ignores filter at given index", () => {
    const rows = [
      { a: "1", b: "x" },
      { a: "2", b: "y" },
    ];
    const filters = [
      { columnKey: "a", values: ["1"] },
      { columnKey: "b", values: ["y"] },
    ];
    const withoutSecond = applyFiltersExcept(rows, filters, 1);
    expect(withoutSecond.length).toBe(1);
    expect(withoutSecond[0]?.b).toBe("x");
  });
});
