import { describe, expect, it } from "vitest";
import { seriesLegendLabel } from "@/lib/chart/seriesLegendLabel";

describe("seriesLegendLabel", () => {
  it("для count_nonempty по ID подставляет сущность раздела", () => {
    expect(
      seriesLegendLabel("ID компании", "count_nonempty", "companies"),
    ).toBe("Кол-во компаний · количество компаний");
    expect(seriesLegendLabel("ID сделки", "count_nonempty", "deals")).toBe(
      "Кол-во сделок · количество сделок",
    );
    expect(seriesLegendLabel("Company ID", "count_nonempty", "leads")).toBe(
      "Кол-во лидов · количество лидов",
    );
  });

  it("для count_nonempty без ID оставляет описание по колонке", () => {
    expect(
      seriesLegendLabel("Телефон", "count_nonempty", "deals"),
    ).toBe("Телефон · количество непустых значений");
  });
});
