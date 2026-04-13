import { describe, expect, it } from "vitest";
import type { ColumnMeta } from "@/lib/types";
import {
  collectParsedDatesForKpiColumn,
  formatKpiDateRu,
  formatKpiRecordDateRu,
  minMaxFromDateList,
} from "@/lib/report/kpiColumnDates";

describe("collectParsedDatesForKpiColumn", () => {
  it("учитывает и строки дд.мм.гггг, и числа-серии Excel в одной колонке", () => {
    const rows: Record<string, unknown>[] = [
      { d: "01.04.2026" },
      { d: 46024 },
      { d: "15.04.2026" },
    ];
    const dates = collectParsedDatesForKpiColumn(rows, "d");
    expect(dates).toHaveLength(3);
  });

  it("не трактует число как дату Excel, если колонка с типом string", () => {
    const meta: ColumnMeta = {
      key: "d",
      header: "Дата создания",
      inferredType: "string",
      sampleValues: [],
    };
    const rows: Record<string, unknown>[] = [{ d: 46024 }];
    expect(collectParsedDatesForKpiColumn(rows, "d", meta)).toHaveLength(0);
  });
});

describe("formatKpiDateRu / formatKpiRecordDateRu", () => {
  it("formatKpiDateRu: день.месяц в подписи", () => {
    const dotted = collectParsedDatesForKpiColumn(
      [{ d: "07.04.2026" }],
      "d",
    );
    const slash = collectParsedDatesForKpiColumn(
      [{ d: "07/04/2026" }],
      "d",
    );
    expect(formatKpiDateRu(dotted[0]!)).toBe("07.04.2026");
    expect(formatKpiDateRu(slash[0]!)).toBe("07.04.2026");
  });

  it("formatKpiRecordDateRu: дд.мм.гггг как formatKpiDateRu", () => {
    const d = new Date(2026, 3, 7);
    expect(formatKpiRecordDateRu(d)).toBe("07.04.2026");
    expect(formatKpiRecordDateRu(d)).toBe(formatKpiDateRu(d));
  });
});

describe("minMaxFromDateList", () => {
  it("сравнивает по календарному дню локально, не по времени суток", () => {
    const dates = [
      new Date(2026, 3, 7, 23, 0, 0),
      new Date(2026, 3, 8, 1, 0, 0),
      new Date(2026, 2, 1, 12, 0, 0),
    ];
    const mm = minMaxFromDateList(dates)!;
    expect(mm.min.getFullYear()).toBe(2026);
    expect(mm.min.getMonth()).toBe(2);
    expect(mm.min.getDate()).toBe(1);
    expect(mm.max.getMonth()).toBe(3);
    expect(mm.max.getDate()).toBe(8);
  });
});
