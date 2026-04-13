import { describe, expect, it } from "vitest";
import {
  applyFilters,
  applyFiltersExcept,
  buildAggregatedChartRows,
  distinctStringValues,
} from "@/lib/chart/buildChartData";
import { parseDealStageOrderLines } from "@/lib/chart/dealStageOrder";
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
  it("literalPieComplement: two pie slices repeat vs non-repeat", () => {
    const cols: ColumnMeta[] = [
      { key: "t", header: "Тип", inferredType: "string", sampleValues: [] },
      { key: "n", header: "N", inferredType: "number", sampleValues: [] },
    ];
    const rows = [
      { t: "Повторная", n: 1 },
      { t: "Новая", n: 1 },
      { t: "Да", n: 1 },
      { t: "", n: 1 },
    ];
    const cfg: ChartConfig = {
      xKey: "t",
      yKeys: ["n"],
      chartType: "pie",
      aggregation: "count",
      filters: [],
      literalPieComplement: {
        repeatLabel: "Повторные",
        nonRepeatLabel: "Неповторные",
        repeatMatchAny: [
          { columnKey: "t", valuesContainAny: ["повтор"] },
          { columnKey: "t", values: ["Да"] },
        ],
      },
    };
    const { data } = buildAggregatedChartRows(rows, cfg, cols);
    expect(data).toHaveLength(2);
    expect(data.find((d) => d.name === "Повторные")?.n).toBe(2);
    expect(data.find((d) => d.name === "Неповторные")?.n).toBe(1);
  });

  it("omits empty X cells from categories (no «(пусто)» on charts)", () => {
    const rows = [
      { stage: "A", amount: 10 },
      { stage: "", amount: 99 },
      { stage: null, amount: 1 },
    ];
    const cfg: ChartConfig = {
      xKey: "stage",
      yKeys: ["amount"],
      chartType: "bar",
      aggregation: "sum",
      filters: [],
    };
    const { data } = buildAggregatedChartRows(rows, cfg, columns);
    expect(data.map((d) => d.name)).toEqual(["A"]);
    expect(data.find((d) => d.name === "A")?.amount).toBe(10);
  });

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

  it("sum is correct when partial rows cancel to zero before more values (no reset of accumulators)", () => {
    const rows = [
      { stage: "A", amount: 100 },
      { stage: "A", amount: -100 },
      { stage: "A", amount: 50 },
      { stage: "A", amount: 10 },
    ];
    const cfg: ChartConfig = {
      xKey: "stage",
      yKeys: ["amount"],
      chartType: "bar",
      aggregation: "sum",
      filters: [],
    };
    const { data } = buildAggregatedChartRows(rows, cfg, columns);
    expect(data.find((d) => d.name === "A")?.amount).toBe(60);
  });

  it("average uses all numeric rows even when running sum hits zero mid-group", () => {
    const rows = [
      { stage: "A", amount: 100 },
      { stage: "A", amount: -100 },
      { stage: "A", amount: 10 },
      { stage: "A", amount: 30 },
    ];
    const cfg: ChartConfig = {
      xKey: "stage",
      yKeys: ["amount"],
      chartType: "bar",
      aggregation: "avg",
      filters: [],
    };
    const { data } = buildAggregatedChartRows(rows, cfg, columns);
    expect(data.find((d) => d.name === "A")?.amount).toBe(10);
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

  it("orders X by calendar time for date columns (deal inflow by date)", () => {
    const dateCols: ColumnMeta[] = [
      { key: "d", header: "Дата", inferredType: "date", sampleValues: [] },
      { key: "n", header: "N", inferredType: "number", sampleValues: [] },
    ];
    const rows = [
      { d: new Date("2026-01-15"), n: 1 },
      { d: new Date("2026-01-01"), n: 1 },
      { d: new Date("2026-01-08"), n: 1 },
    ];
    const cfg: ChartConfig = {
      xKey: "d",
      yKeys: ["n"],
      chartType: "line",
      aggregation: "count",
      filters: [],
    };
    const { data } = buildAggregatedChartRows(rows, cfg, dateCols);
    expect(data.map((r) => r.name)).toEqual([
      "01.01.2026",
      "08.01.2026",
      "15.01.2026",
    ]);
  });

  it("maxXCellCalendarDayYmd drops rows strictly after the report boundary", () => {
    const dateCols: ColumnMeta[] = [
      { key: "d", header: "Дата", inferredType: "date", sampleValues: [] },
      { key: "n", header: "N", inferredType: "number", sampleValues: [] },
    ];
    const rows = [
      { d: new Date("2026-04-10"), n: 1 },
      { d: new Date("2026-04-12"), n: 1 },
      { d: new Date("2026-04-13"), n: 1 },
      { d: new Date("2026-05-01"), n: 1 },
    ];
    const cfg: ChartConfig = {
      xKey: "d",
      yKeys: ["n"],
      chartType: "line",
      aggregation: "count",
      filters: [],
    };
    const { data } = buildAggregatedChartRows(
      rows,
      cfg,
      dateCols,
      null,
      "2026-04-12",
    );
    expect(data.map((r) => r.name)).toEqual(["10.04.2026", "12.04.2026"]);
    expect(data.map((r) => r.n)).toEqual([1, 1]);
  });

  it("maxXCellCalendarDayYmd drops unparseable date labels on temporal axis", () => {
    const dateCols: ColumnMeta[] = [
      { key: "d", header: "Дата", inferredType: "date", sampleValues: [] },
      { key: "n", header: "N", inferredType: "number", sampleValues: [] },
    ];
    const rows = [
      { d: "12.04.2026", n: 1 },
      { d: "04/13/2026", n: 1 },
      { d: "2026-04-11", n: 1 },
    ];
    const cfg: ChartConfig = {
      xKey: "d",
      yKeys: ["n"],
      chartType: "line",
      aggregation: "count",
      filters: [],
    };
    const { data } = buildAggregatedChartRows(
      rows,
      cfg,
      dateCols,
      null,
      "2026-04-12",
    );
    expect(data.map((r) => r.name)).toEqual(["11.04.2026", "12.04.2026"]);
    expect(data.map((r) => r.n)).toEqual([1, 1]);
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

  it("countDistinctByKey counts unique deals instead of duplicate rows", () => {
    const dateCols: ColumnMeta[] = [
      { key: "d", header: "Дата создания", inferredType: "date", sampleValues: [] },
      { key: "id", header: "ID", inferredType: "number", sampleValues: [] },
    ];
    const rows = [
      { d: new Date(2026, 0, 5), id: 101 },
      { d: new Date(2026, 0, 5), id: 101 },
      { d: new Date(2026, 1, 7), id: 202 },
      { d: new Date(2026, 1, 10), id: 202 },
      { d: new Date(2026, 1, 11), id: 303 },
    ];
    const cfg: ChartConfig = {
      xKey: "d",
      yKeys: ["id"],
      chartType: "area",
      aggregation: "count",
      filters: [],
      dateGranularity: "month",
      cumulative: true,
      countDistinctByKey: "id",
    };
    const { data } = buildAggregatedChartRows(rows, cfg, dateCols);
    expect(data.find((x) => x.name === "янв. 2026")?.id).toBe(1);
    expect(data.find((x) => x.name === "февр. 2026")?.id).toBe(3);
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

  it("sorts categories by deal pipeline when X is стадия сделки", () => {
    const dealCols: ColumnMeta[] = [
      {
        key: "st",
        header: "Стадия сделки",
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
    const rows = [
      { st: "Оплата", amt: 1 },
      { st: "Назначенные заявки", amt: 2 },
      { st: "Договор", amt: 3 },
    ];
    const cfg: ChartConfig = {
      xKey: "st",
      yKeys: ["amt"],
      chartType: "bar",
      aggregation: "sum",
      filters: [],
    };
    const order = ["Назначенные заявки", "Договор", "Оплата"];
    const { data } = buildAggregatedChartRows(rows, cfg, dealCols, order);
    expect(data.map((d) => d.name)).toEqual([
      "Назначенные заявки",
      "Договор",
      "Оплата",
    ]);
  });

  it("orders deal stages by leading number when pipeline list not set", () => {
    const dealCols: ColumnMeta[] = [
      {
        key: "st",
        header: "Стадия сделки",
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
    const rows = [
      { st: "10. Сделка успешна", amt: 1 },
      { st: "2. В работе", amt: 2 },
      { st: "1. Новая", amt: 3 },
    ];
    const cfg: ChartConfig = {
      xKey: "st",
      yKeys: ["amt"],
      chartType: "bar",
      aggregation: "sum",
      filters: [],
    };
    const { data } = buildAggregatedChartRows(rows, cfg, dealCols, null);
    expect(data.map((d) => d.name)).toEqual([
      "1. Новая",
      "2. В работе",
      "10. Сделка успешна",
    ]);
  });

  it("orders terminal CRM stages after Оплата: успех, затем отказ", () => {
    const dealCols: ColumnMeta[] = [
      {
        key: "st",
        header: "Стадия сделки",
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
    const rows = [
      { st: "11. Отказ / Тендер проигран", amt: 1 },
      { st: "Оплата", amt: 2 },
      { st: "10. Сделка успешна", amt: 3 },
    ];
    const cfg: ChartConfig = {
      xKey: "st",
      yKeys: ["amt"],
      chartType: "bar",
      aggregation: "sum",
      filters: [],
    };
    const order = [
      "Оплата",
      "10. Сделка успешна",
      "11. Отказ / Тендер проигран",
    ];
    const { data } = buildAggregatedChartRows(rows, cfg, dealCols, order);
    expect(data.map((d) => d.name)).toEqual([
      "Оплата",
      "10. Сделка успешна",
      "11. Отказ / Тендер проигран",
    ]);
  });

  it("literalAvgBars: среднее по разным фильтрам и колонкам", () => {
    const literalCols: ColumnMeta[] = [
      {
        key: "stage",
        header: "Стадия",
        inferredType: "string",
        sampleValues: [],
      },
      {
        key: "kp",
        header: "Бюджет КП",
        inferredType: "number",
        sampleValues: [],
      },
      {
        key: "amt",
        header: "Сумма",
        inferredType: "number",
        sampleValues: [],
      },
    ];
    const literalRows = [
      { stage: "Отправка КП", kp: 100, amt: 1000 },
      { stage: "Оплата", kp: 200, amt: 500 },
      { stage: "Отправка КП", kp: 0, amt: 300 },
    ];
    const literalCfg: ChartConfig = {
      xKey: "stage",
      yKeys: ["kp", "amt"],
      chartType: "bar",
      aggregation: "avg",
      filters: [],
      literalAvgBars: [
        {
          label: "Средняя стоимость КП",
          filters: [{ columnKey: "kp", min: 0.01 }],
          valueKey: "kp",
          seriesKey: "kp",
        },
        {
          label: "Средний чек (стадия «Отправка КП»)",
          filters: [{ columnKey: "stage", values: ["Отправка КП"] }],
          valueKey: "amt",
          seriesKey: "amt",
        },
      ],
    };
    const { data, warnings } = buildAggregatedChartRows(
      literalRows,
      literalCfg,
      literalCols,
    );
    expect(warnings.length).toBe(0);
    expect(data).toHaveLength(2);
    const rowKp = data.find((d) => d.name === "Средняя стоимость КП");
    const rowDeal = data.find(
      (d) => d.name === "Средний чек (стадия «Отправка КП»)",
    );
    expect(rowKp?.kp).toBe(150);
    expect(rowKp?.amt).toBe(0);
    expect(rowDeal?.kp).toBe(0);
    expect(rowDeal?.amt).toBe(650);
  });
});

describe("parseDealStageOrderLines", () => {
  it("trims lines and skips empty rows", () => {
    expect(parseDealStageOrderLines(" A \n\n B \r\n C")).toEqual(["A", "B", "C"]);
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

  it("sorts deal stage values by pipeline when column meta is стадия сделки", () => {
    const meta: ColumnMeta = {
      key: "st",
      header: "Стадия сделки",
      inferredType: "string",
      sampleValues: [],
    };
    const rows = [{ st: "Оплата" }, { st: "Подготовка КП" }, { st: "Назначенные заявки" }];
    const order = ["Назначенные заявки", "Подготовка КП", "Оплата"];
    const v = distinctStringValues(rows, "st", 20, meta, order);
    expect(v.indexOf("Назначенные заявки")).toBeLessThan(v.indexOf("Подготовка КП"));
    expect(v.indexOf("Подготовка КП")).toBeLessThan(v.indexOf("Оплата"));
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

  it("filters by valuesContainAll / valuesContainAny (стадии с префиксом воронки)", () => {
    const rows = [
      { st: "7. Отправка КП", amt: 100 },
      { st: "Отправка КП", amt: 200 },
      { st: "Подготовка КП", amt: 1 },
      { st: "Оплата", amt: 50 },
    ];
    const f = [
      {
        columnKey: "st",
        valuesContainAll: ["отправка"],
        valuesContainAny: ["кп", "kp"],
      },
    ];
    const out = applyFilters(rows, f);
    expect(out.map((r) => r.st)).toEqual(["7. Отправка КП", "Отправка КП"]);
  });

  it("excludes rows when label matches excludeValues", () => {
    const rows = [
      { st: "Новая" },
      { st: "Успешно реализована" },
      { st: "Провал" },
    ];
    const out = applyFilters(rows, [
      {
        columnKey: "st",
        excludeValues: ["Успешно реализована", "Провал"],
      },
    ]);
    expect(out.map((r) => r.st)).toEqual(["Новая"]);
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
