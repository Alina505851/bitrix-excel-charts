import { describe, expect, it } from "vitest";
import { buildAggregatedChartRows } from "@/lib/chart/buildChartData";
import { PREDEFINED_CHARTS } from "@/lib/predefinedCharts";
import type { ColumnMeta } from "@/lib/types";

const smokeColumns: ColumnMeta[] = [
  { key: "lead_stage", header: "Стадия лида", inferredType: "string", sampleValues: [] },
  { key: "deal_stage", header: "Стадия сделки", inferredType: "string", sampleValues: [] },
  { key: "sum", header: "Сумма", inferredType: "number", sampleValues: [] },
  { key: "created_at", header: "Дата создания", inferredType: "date", sampleValues: [] },
  { key: "responsible", header: "Ответственный", inferredType: "string", sampleValues: [] },
  { key: "source", header: "Источник", inferredType: "string", sampleValues: [] },
  { key: "company", header: "Компания", inferredType: "string", sampleValues: [] },
  { key: "industry", header: "Отрасль", inferredType: "string", sampleValues: [] },
  { key: "headcount", header: "Кол-во сотрудников", inferredType: "number", sampleValues: [] },
  { key: "location", header: "Локация", inferredType: "string", sampleValues: [] },
  { key: "utm_source", header: "UTM Source", inferredType: "string", sampleValues: [] },
  { key: "utm_medium", header: "UTM Medium", inferredType: "string", sampleValues: [] },
  { key: "utm_campaign", header: "UTM Campaign", inferredType: "string", sampleValues: [] },
  { key: "quote_status", header: "Статус", inferredType: "string", sampleValues: [] },
];

const smokeRows: Record<string, unknown>[] = [
  {
    lead_stage: "Новая",
    deal_stage: "Успешно реализована",
    sum: 100,
    created_at: new Date("2026-02-10"),
    responsible: "Иванов",
    source: "Сайт",
    company: "ООО Альфа",
    industry: "IT",
    headcount: 10,
    location: "Москва",
    utm_source: "google",
    utm_medium: "cpc",
    utm_campaign: "brand",
    quote_status: "Отправлено",
  },
  {
    lead_stage: "В работе",
    deal_stage: "Успешно реализована",
    sum: 250,
    created_at: new Date("2026-03-12"),
    responsible: "Петров",
    source: "Реклама",
    company: "ООО Бета",
    industry: "Retail",
    headcount: 35,
    location: "СПб",
    utm_source: "yandex",
    utm_medium: "cpc",
    utm_campaign: "promo",
    quote_status: "Отправлено клиенту",
  },
];

describe("All predefined charts smoke", () => {
  it("every spec resolves without throw, and resolved charts build data", () => {
    for (const spec of PREDEFINED_CHARTS) {
      let resolved: ReturnType<typeof spec.resolve>;
      expect(() => {
        resolved = spec.resolve(smokeColumns);
      }).not.toThrow();

      if (!resolved!.ok) {
        continue;
      }

      expect(() =>
        buildAggregatedChartRows(smokeRows, resolved!.config, smokeColumns),
      ).not.toThrow();
    }
  });
});

