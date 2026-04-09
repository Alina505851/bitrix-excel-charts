import { describe, expect, it } from "vitest";
import {
  PREDEFINED_CHARTS,
  chartsForEntity,
  getSpecById,
  normalizeConfigForTabular,
  suggestEntityAndCharts,
  suggestEntityFromFileName,
} from "@/lib/predefinedCharts";
import { buildAggregatedChartRows } from "@/lib/chart/buildChartData";
import type { ColumnMeta, TabularData } from "@/lib/types";

const leadColumnsMinimal: ColumnMeta[] = [
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

describe("PREDEFINED_CHARTS registry", () => {
  it("has unique ids and five entity sections", () => {
    const ids = PREDEFINED_CHARTS.map((c) => c.id);
    expect(new Set(ids).size).toBe(ids.length);
    expect(chartsForEntity("leads").length).toBeGreaterThan(0);
    expect(chartsForEntity("deals").length).toBeGreaterThan(0);
    expect(chartsForEntity("contacts").length).toBeGreaterThan(0);
    expect(chartsForEntity("companies").length).toBeGreaterThan(0);
    expect(chartsForEntity("quotes").length).toBeGreaterThan(0);
  });

  it("getSpecById returns spec or undefined", () => {
    expect(getSpecById("leads_sum_by_stage")?.entity).toBe("leads");
    expect(getSpecById("__missing__")).toBeUndefined();
  });
});

describe("PredefinedChartSpec.resolve", () => {
  it("leads_sum_by_stage succeeds on typical Bitrix-like headers", () => {
    const spec = getSpecById("leads_sum_by_stage")!;
    const r = spec.resolve(leadColumnsMinimal);
    expect(r.ok).toBe(true);
    if (r.ok) {
      expect(r.config.xKey).toBe("st");
      expect(r.config.yKeys).toEqual(["amt"]);
      expect(r.config.chartType).toBe("bar");
    }
  });

  it("leads_sum_by_stage fails without stage column", () => {
    const spec = getSpecById("leads_sum_by_stage")!;
    const r = spec.resolve([
      {
        key: "amt",
        header: "Сумма",
        inferredType: "number",
        sampleValues: [],
      },
    ]);
    expect(r.ok).toBe(false);
  });

  it("deals_count_by_company resolves with company header", () => {
    const spec = getSpecById("deals_count_by_company")!;
    const cols: ColumnMeta[] = [
      {
        key: "co",
        header: "Компания",
        inferredType: "string",
        sampleValues: [],
      },
      {
        key: "n",
        header: "Сумма",
        inferredType: "number",
        sampleValues: [],
      },
    ];
    const r = spec.resolve(cols);
    expect(r.ok).toBe(true);
  });

  it("deals_sum_by_stage does not use pure ID column as amount", () => {
    const spec = getSpecById("deals_sum_by_stage")!;
    const cols: ColumnMeta[] = [
      {
        key: "stage",
        header: "Стадия сделки",
        inferredType: "string",
        sampleValues: [],
      },
      {
        key: "deal_id",
        header: "ID сделки",
        inferredType: "number",
        sampleValues: [],
      },
    ];
    const r = spec.resolve(cols);
    expect(r.ok).toBe(false);
  });

  it("companies_count_by_source fails without source column", () => {
    const spec = getSpecById("companies_count_by_source")!;
    const r = spec.resolve([
      {
        key: "n",
        header: "Сумма",
        inferredType: "number",
        sampleValues: [],
      },
    ]);
    expect(r.ok).toBe(false);
  });

  it("companies_count_by_source resolves for explicit company source column", () => {
    const spec = getSpecById("companies_count_by_source")!;
    const cols: ColumnMeta[] = [
      {
        key: "company_source",
        header: "Источник компании",
        inferredType: "string",
        sampleValues: [],
      },
      {
        key: "n",
        header: "ID",
        inferredType: "number",
        sampleValues: [],
      },
    ];
    const r = spec.resolve(cols);
    expect(r.ok).toBe(true);
  });

  it("companies_count_by_source resolves for UF_CRM_SOURCEID", () => {
    const spec = getSpecById("companies_count_by_source")!;
    const cols: ColumnMeta[] = [
      {
        key: "UF_CRM_SOURCEID",
        header: "UF_CRM_SOURCEID",
        inferredType: "string",
        sampleValues: [],
      },
      {
        key: "n",
        header: "ID",
        inferredType: "number",
        sampleValues: [],
      },
    ];
    const r = spec.resolve(cols);
    expect(r.ok).toBe(true);
  });

  it("companies_count_by_industry requires explicit industry column", () => {
    const spec = getSpecById("companies_count_by_industry")!;
    const cols: ColumnMeta[] = [
      {
        key: "ctype",
        header: "Тип компании",
        inferredType: "string",
        sampleValues: [],
      },
      {
        key: "n",
        header: "ID",
        inferredType: "number",
        sampleValues: [],
      },
    ];
    const r = spec.resolve(cols);
    expect(r.ok).toBe(false);
  });

  it("leads_area_sum_by_date uses area chart and sum aggregation", () => {
    const spec = getSpecById("leads_area_sum_by_date")!;
    const cols: ColumnMeta[] = [
      {
        key: "dt",
        header: "Дата создания",
        inferredType: "date",
        sampleValues: [],
      },
      ...leadColumnsMinimal,
    ];
    const r = spec.resolve(cols);
    expect(r.ok).toBe(true);
    if (r.ok) {
      expect(r.config.chartType).toBe("area");
      expect(r.config.aggregation).toBe("sum");
    }
  });

  it("deals_won_q1_2026_sum_by_responsible includes Q1 2026 date bounds", () => {
    const spec = getSpecById("deals_won_q1_2026_sum_by_responsible")!;
    const cols: ColumnMeta[] = [
      {
        key: "st",
        header: "Стадия сделки",
        inferredType: "string",
        sampleValues: [],
      },
      {
        key: "dt",
        header: "Дата закрытия",
        inferredType: "date",
        sampleValues: [],
      },
      {
        key: "mgr",
        header: "Ответственный",
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
    const r = spec.resolve(cols);
    expect(r.ok).toBe(true);
    if (r.ok) {
      expect(r.config.filters).toHaveLength(2);
      const dateF = r.config.filters.find((f) => f.columnKey === "dt");
      expect(dateF?.dateFrom).toBe("2026-01-01");
      expect(dateF?.dateTo).toBe("2026-03-31");
    }
  });

  it("deals_won_sum_by_responsible filters by typical won stage labels", () => {
    const spec = getSpecById("deals_won_sum_by_responsible")!;
    const cols: ColumnMeta[] = [
      {
        key: "st",
        header: "Стадия сделки",
        inferredType: "string",
        sampleValues: [],
      },
      {
        key: "mgr",
        header: "Ответственный",
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
    const r = spec.resolve(cols);
    expect(r.ok).toBe(true);
    if (r.ok) {
      expect(r.config.filters).toHaveLength(1);
      expect(r.config.filters[0]?.columnKey).toBe("st");
      expect(r.config.filters[0]?.values?.includes("Успешно реализована")).toBe(
        true,
      );
    }
  });

  it("quotes_sent_with_amount_count_by_month resolves on typical КП export columns", () => {
    const spec = getSpecById("quotes_sent_with_amount_count_by_month")!;
    const cols: ColumnMeta[] = [
      {
        key: "dt",
        header: "Дата создания",
        inferredType: "date",
        sampleValues: [],
      },
      {
        key: "st",
        header: "Статус",
        inferredType: "string",
        sampleValues: [],
      },
      {
        key: "sum",
        header: "Сумма",
        inferredType: "number",
        sampleValues: [],
      },
    ];
    const r = spec.resolve(cols);
    expect(r.ok).toBe(true);
    if (r.ok) {
      expect(r.config.dateGranularity).toBe("month");
      expect(r.config.filters.length).toBe(2);
      expect(r.config.filters.some((f) => f.min === 0.01)).toBe(true);
    }
  });

  it("deals_line_avg_by_date uses line and avg", () => {
    const spec = getSpecById("deals_line_avg_by_date")!;
    const cols: ColumnMeta[] = [
      {
        key: "d",
        header: "Дата создания",
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
    const r = spec.resolve(cols);
    expect(r.ok).toBe(true);
    if (r.ok) {
      expect(r.config.chartType).toBe("line");
      expect(r.config.aggregation).toBe("avg");
    }
  });

  it("companies_count_by_responsible uses pie with count aggregation", () => {
    const spec = getSpecById("companies_count_by_responsible")!;
    const cols: ColumnMeta[] = [
      {
        key: "mgr",
        header: "Ответственный",
        inferredType: "string",
        sampleValues: [],
      },
      {
        key: "id",
        header: "ID",
        inferredType: "number",
        sampleValues: [],
      },
    ];
    const r = spec.resolve(cols);
    expect(r.ok).toBe(true);
    if (r.ok) {
      expect(r.config.chartType).toBe("pie");
      expect(r.config.aggregation).toBe("count");
    }
  });

  it("companies_sum_employees uses Кол-во сотрудников, not годовой оборот", () => {
    const spec = getSpecById("companies_sum_employees")!;
    const cols: ColumnMeta[] = [
      {
        key: "ind",
        header: "Отрасль",
        inferredType: "string",
        sampleValues: [],
      },
      {
        key: "staff",
        header: "Кол-во сотрудников",
        inferredType: "number",
        sampleValues: [],
      },
      {
        key: "rev",
        header: "Годовой оборот",
        inferredType: "number",
        sampleValues: [],
      },
    ];
    const r = spec.resolve(cols);
    expect(r.ok).toBe(true);
    if (r.ok) {
      expect(r.config.yKeys).toEqual(["staff"]);
      expect(r.config.aggregation).toBe("sum");
    }
  });

  it("companies_sum_employees requires explicit industry field", () => {
    const spec = getSpecById("companies_sum_employees")!;
    const cols: ColumnMeta[] = [
      {
        key: "ctype",
        header: "Тип компании",
        inferredType: "string",
        sampleValues: [],
      },
      {
        key: "staff",
        header: "Кол-во сотрудников",
        inferredType: "number",
        sampleValues: [],
      },
    ];
    const r = spec.resolve(cols);
    expect(r.ok).toBe(false);
  });

  it("companies_revenue_by_industry requires explicit revenue column", () => {
    const spec = getSpecById("companies_revenue_by_industry")!;
    const cols: ColumnMeta[] = [
      {
        key: "ind",
        header: "Отрасль",
        inferredType: "string",
        sampleValues: [],
      },
      {
        key: "n",
        header: "ID",
        inferredType: "number",
        sampleValues: [],
      },
    ];
    const r = spec.resolve(cols);
    expect(r.ok).toBe(false);
  });

  it("companies_sum_by_responsible requires explicit business metric (not ID)", () => {
    const spec = getSpecById("companies_sum_by_responsible")!;
    const cols: ColumnMeta[] = [
      {
        key: "mgr",
        header: "Ответственный",
        inferredType: "string",
        sampleValues: [],
      },
      {
        key: "id",
        header: "ID",
        inferredType: "number",
        sampleValues: [],
      },
    ];
    const r = spec.resolve(cols);
    expect(r.ok).toBe(false);
  });

  it("companies_count_by_company_type resolves with Тип компании", () => {
    const spec = getSpecById("companies_count_by_company_type")!;
    const cols: ColumnMeta[] = [
      {
        key: "ctype",
        header: "Тип компании",
        inferredType: "string",
        sampleValues: [],
      },
      {
        key: "v",
        header: "Годовой оборот",
        inferredType: "number",
        sampleValues: [],
      },
    ];
    const r = spec.resolve(cols);
    expect(r.ok).toBe(true);
    if (r.ok) {
      expect(r.config.xKey).toBe("ctype");
      expect(r.config.chartType).toBe("pie");
      expect(r.config.aggregation).toBe("count");
    }
  });

  it("companies_percent_with_deal resolves on 'Есть сделка' column", () => {
    const spec = getSpecById("companies_percent_with_deal")!;
    const cols: ColumnMeta[] = [
      {
        key: "has_deal",
        header: "Есть сделка?",
        inferredType: "string",
        sampleValues: [],
      },
      {
        key: "n",
        header: "Сумма",
        inferredType: "number",
        sampleValues: [],
      },
    ];
    const r = spec.resolve(cols);
    expect(r.ok).toBe(true);
    if (r.ok) {
      expect(r.config.chartType).toBe("pie");
      expect(r.config.aggregation).toBe("count");
    }
  });


  it("companies_avg_headcount_by_industry requires explicit industry (not only type company)", () => {
    const spec = getSpecById("companies_avg_headcount_by_industry")!;
    const cols: ColumnMeta[] = [
      {
        key: "ctype",
        header: "Тип компании",
        inferredType: "string",
        sampleValues: [],
      },
      {
        key: "staff",
        header: "Кол-во сотрудников",
        inferredType: "number",
        sampleValues: [],
      },
    ];
    const r = spec.resolve(cols);
    expect(r.ok).toBe(false);
  });

  it("companies_count_by_activity_sphere resolves with header containing suffix", () => {
    const spec = getSpecById("companies_count_by_activity_sphere")!;
    const cols: ColumnMeta[] = [
      {
        key: "sphere_clients",
        header: "Сфера деятельности (Клиенты)",
        inferredType: "string",
        sampleValues: [],
      },
      {
        key: "n",
        header: "ID",
        inferredType: "number",
        sampleValues: [],
      },
    ];
    const r = spec.resolve(cols);
    expect(r.ok).toBe(true);
  });

  it("companies_registered_by_month_line does not use update date as created date", () => {
    const spec = getSpecById("companies_registered_by_month_line")!;
    const cols: ColumnMeta[] = [
      {
        key: "updated_at",
        header: "Дата изменения",
        inferredType: "date",
        sampleValues: [],
      },
      {
        key: "id",
        header: "ID",
        inferredType: "number",
        sampleValues: [],
      },
    ];
    const r = spec.resolve(cols);
    expect(r.ok).toBe(false);
  });

  it("companies_registered_by_month_line prefers explicit created date when both exist", () => {
    const spec = getSpecById("companies_registered_by_month_line")!;
    const cols: ColumnMeta[] = [
      {
        key: "updated_at",
        header: "Дата изменения",
        inferredType: "date",
        sampleValues: [],
      },
      {
        key: "created_at",
        header: "Дата создания",
        inferredType: "date",
        sampleValues: [],
      },
      {
        key: "id",
        header: "ID",
        inferredType: "number",
        sampleValues: [],
      },
    ];
    const r = spec.resolve(cols);
    expect(r.ok).toBe(true);
    if (r.ok) {
      expect(r.config.xKey).toBe("created_at");
      expect(r.config.filters).toEqual([]);
    }
  });

  it("companies_activity_last_by_month_line resolves with last activity date field", () => {
    const spec = getSpecById("companies_activity_last_by_month_line")!;
    const cols: ColumnMeta[] = [
      {
        key: "last_activity_at",
        header: "Дата последней коммуникации",
        inferredType: "date",
        sampleValues: [],
      },
      {
        key: "id",
        header: "ID",
        inferredType: "number",
        sampleValues: [],
      },
    ];
    const r = spec.resolve(cols);
    expect(r.ok).toBe(true);
  });

  it("companies_activity_last_by_month_line does not resolve with next activity date field", () => {
    const spec = getSpecById("companies_activity_last_by_month_line")!;
    const cols: ColumnMeta[] = [
      {
        key: "next_activity_at",
        header: "Дата следующей активности",
        inferredType: "date",
        sampleValues: [],
      },
      {
        key: "id",
        header: "ID",
        inferredType: "number",
        sampleValues: [],
      },
    ];
    const r = spec.resolve(cols);
    expect(r.ok).toBe(false);
  });
});

describe("suggestEntityAndCharts", () => {
  it("returns leads with feasible chart ids for minimal lead export", () => {
    const s = suggestEntityAndCharts(leadColumnsMinimal);
    expect(s).not.toBeNull();
    expect(s!.entity).toBe("leads");
    expect(s!.chartIds).toContain("leads_sum_by_stage");
    expect(s!.chartIds.every((id) => id.startsWith("leads_"))).toBe(true);
  });

  it("prefers deals when колонка стадии явно про сделку", () => {
    const cols: ColumnMeta[] = [
      {
        key: "x",
        header: "Стадия сделки",
        inferredType: "string",
        sampleValues: [],
      },
      {
        key: "n",
        header: "Сумма",
        inferredType: "number",
        sampleValues: [],
      },
    ];
    const s = suggestEntityAndCharts(cols);
    expect(s).not.toBeNull();
    expect(s!.entity).toBe("deals");
  });

  it("returns null when no chart can be built", () => {
    const s = suggestEntityAndCharts([]);
    expect(s).toBeNull();
  });
});

describe("suggestEntityFromFileName", () => {
  it("detects companies by COMPANY marker", () => {
    expect(
      suggestEntityFromFileName("COMPANY_20260409_16f7faf1_69d753e232453.xls"),
    ).toBe("companies");
  });

  it("detects deals by DEALS marker", () => {
    expect(suggestEntityFromFileName("DEALS_EXPORT_2026.xlsx")).toBe("deals");
  });

  it("returns null for unknown file name", () => {
    expect(suggestEntityFromFileName("report_2026.xlsx")).toBeNull();
  });
});

describe("normalizeConfigForTabular", () => {
  it("drops missing keys from config", () => {
    const data: TabularData = {
      sheetNames: ["S"],
      activeSheet: "S",
      columns: leadColumnsMinimal,
      rows: [],
    };
    const normalized = normalizeConfigForTabular(data, {
      xKey: "ghost",
      yKeys: ["amt", "ghost_y"],
      chartType: "bar",
      aggregation: "sum",
      filters: [{ columnKey: "ghost_f", values: ["a"] }],
    });
    expect(normalized.xKey).toBeNull();
    expect(normalized.yKeys).toEqual(["amt"]);
    expect(normalized.filters).toEqual([]);
  });
});

describe("companies charts: field requirements and generation logic", () => {
  const companyMinimalColumns: ColumnMeta[] = [
    { key: "id", header: "ID", inferredType: "number", sampleValues: [] },
    {
      key: "created_at",
      header: "Дата создания",
      inferredType: "date",
      sampleValues: [],
    },
    {
      key: "responsible",
      header: "Ответственный",
      inferredType: "string",
      sampleValues: [],
    },
    {
      key: "company_name",
      header: "Название компании",
      inferredType: "string",
      sampleValues: [],
    },
    {
      key: "company_type",
      header: "Тип компании",
      inferredType: "string",
      sampleValues: [],
    },
    {
      key: "has_deal",
      header: "Есть сделка?",
      inferredType: "string",
      sampleValues: [],
    },
  ];

  const companyRows: Record<string, unknown>[] = [
    {
      id: 1,
      created_at: "2026-04-01",
      responsible: "Иванов",
      company_name: "ООО Альфа",
      company_type: "Партнер",
      has_deal: "Да",
    },
    {
      id: 2,
      created_at: "2026-04-02",
      responsible: "Петров",
      company_name: "ООО Бета",
      company_type: "Клиент",
      has_deal: "Нет",
    },
    {
      id: 3,
      created_at: "2026-04-03",
      responsible: "Иванов",
      company_name: "ООО Гамма",
      company_type: "Партнер",
      has_deal: "Да",
    },
  ];

  it("resolves and generates data for charts based on provided fields", () => {
    const ids = [
      "companies_count_by_company_type",
      "companies_percent_with_deal",
      "companies_count_by_responsible",
      "companies_registered_by_month_line",
    ] as const;

    for (const id of ids) {
      const spec = getSpecById(id)!;
      const resolved = spec.resolve(companyMinimalColumns);
      expect(resolved.ok).toBe(true);
      if (!resolved.ok) {
        continue;
      }
      const out = buildAggregatedChartRows(
        companyRows,
        resolved.config,
        companyMinimalColumns,
      );
      expect(out.warnings).toEqual([]);
      expect(out.data.length).toBeGreaterThan(0);
    }
  });

  it("keeps strict charts unavailable when required fields are absent", () => {
    const unavailable = [
      "companies_count_by_utm_source",
      "companies_count_by_utm_medium",
      "companies_revenue_by_industry",
      "companies_sum_employees",
      "companies_avg_headcount_by_industry",
    ] as const;

    for (const id of unavailable) {
      const spec = getSpecById(id)!;
      const resolved = spec.resolve(companyMinimalColumns);
      expect(resolved.ok).toBe(false);
    }
  });

  it("builds correct distribution for percent companies with deal", () => {
    const spec = getSpecById("companies_percent_with_deal")!;
    const resolved = spec.resolve(companyMinimalColumns);
    expect(resolved.ok).toBe(true);
    if (!resolved.ok) {
      return;
    }

    const out = buildAggregatedChartRows(
      companyRows,
      resolved.config,
      companyMinimalColumns,
    );
    expect(out.warnings).toEqual([]);
    expect(out.data.find((d) => d.name === "Да")?.id).toBe(2);
    expect(out.data.find((d) => d.name === "Нет")?.id).toBe(1);
  });
});
