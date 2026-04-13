import { describe, expect, it } from "vitest";
import {
  formatDateDdMmYyyyRuLocal,
  formatDateTimeDdMmYyyyRuLocal,
  formatDateYmdLocal,
  groupLabel,
  tryParseDate,
  tryParseNumber,
} from "@/lib/chart/coerce";

describe("tryParseNumber", () => {
  it("parses finite number and normalized strings", () => {
    expect(tryParseNumber(42)).toBe(42);
    expect(tryParseNumber("1 000,5")).toBe(1000.5);
    expect(tryParseNumber("-")).toBe(null);
  });

  it("returns null for non-numeric input", () => {
    expect(tryParseNumber("abc")).toBe(null);
    expect(tryParseNumber(undefined)).toBe(null);
    expect(tryParseNumber(null)).toBe(null);
  });
});

describe("tryParseDate", () => {
  it("parses DD.MM.YYYY", () => {
    const d = tryParseDate("15.03.2024");
    expect(d).not.toBeNull();
    expect(d!.getFullYear()).toBe(2024);
    expect(d!.getMonth()).toBe(2);
    expect(d!.getDate()).toBe(15);
  });

  it("does not treat dotted date as US month.day (only дд.мм)", () => {
    expect(tryParseDate("03.15.2024")).toBe(null);
  });

  it("parses DD.MM.YYYY HH:mm:ss as day-first local datetime", () => {
    const d = tryParseDate("07.04.2026 16:32:33");
    expect(d).not.toBeNull();
    expect(d!.getFullYear()).toBe(2026);
    expect(d!.getMonth()).toBe(3);
    expect(d!.getDate()).toBe(7);
    expect(d!.getHours()).toBe(16);
    expect(d!.getMinutes()).toBe(32);
    expect(d!.getSeconds()).toBe(33);
  });

  it("parses DD.MM.YYYY HH:mm (without seconds) as day-first local datetime", () => {
    const d = tryParseDate("07.04.2026 16:32");
    expect(d).not.toBeNull();
    expect(d!.getFullYear()).toBe(2026);
    expect(d!.getMonth()).toBe(3);
    expect(d!.getDate()).toBe(7);
    expect(d!.getHours()).toBe(16);
    expect(d!.getMinutes()).toBe(32);
    expect(d!.getSeconds()).toBe(0);
  });

  it("parses ISO timestamp with timezone", () => {
    const d = tryParseDate("2025-07-30T17:24:36.000Z");
    expect(d).not.toBeNull();
    expect(Number.isNaN(d!.getTime())).toBe(false);
  });

  it("parses ISO date string", () => {
    const d = tryParseDate("2024-06-01");
    expect(d).not.toBeNull();
  });

  it("parses short RU year as day-first date", () => {
    const d = tryParseDate("09.04.26");
    expect(d).not.toBeNull();
    expect(d!.getFullYear()).toBe(2026);
    expect(d!.getMonth()).toBe(3);
    expect(d!.getDate()).toBe(9);
  });

  it("parses Excel serial date number", () => {
    const d = tryParseDate(46024);
    expect(d).not.toBeNull();
    expect(d!.getFullYear()).toBe(2026);
    expect(d!.getMonth()).toBe(0);
    expect(d!.getDate()).toBe(2);
  });

  it("does not parse small numeric ids as dates", () => {
    expect(tryParseDate(134)).toBe(null);
    expect(tryParseDate(522)).toBe(null);
  });

  it("parses DD/MM/YYYY and DD-MM-YYYY (день, месяц)", () => {
    const a = tryParseDate("07/04/2026");
    expect(a!.getMonth()).toBe(3);
    expect(a!.getDate()).toBe(7);
    const b = tryParseDate("07-04-2026");
    expect(b!.getMonth()).toBe(3);
    expect(b!.getDate()).toBe(7);
  });

  it("does not parse slash date with invalid month", () => {
    expect(tryParseDate("07/13/2026")).toBe(null);
  });

  it("returns null for invalid", () => {
    expect(tryParseDate("")).toBe(null);
    expect(tryParseDate("не дата")).toBe(null);
    expect(tryParseDate("31.06.2026 16:32:33")).toBe(null);
    // Невалидно и как дд.мм, и как мм.дд (месяц > 12 в обоих смыслах)
    expect(tryParseDate("15.13.2026 16:32:33")).toBe(null);
  });
});

describe("groupLabel", () => {
  it("normalizes empty and dates", () => {
    expect(groupLabel(null)).toBe("(пусто)");
    expect(groupLabel("")).toBe("(пусто)");
    const dt = new Date(2024, 0, 10, 12, 0, 0);
    expect(groupLabel(dt)).toBe("10.01.2024 12:00");
    const midnight = new Date(2024, 0, 10, 0, 0, 0);
    expect(groupLabel(midnight)).toBe("10.01.2024");
  });

  it("stringifies other values", () => {
    expect(groupLabel("Лид")).toBe("Лид");
    expect(groupLabel(0)).toBe("0");
  });
});

describe("formatDateYmdLocal", () => {
  it("formats date as local YYYY-MM-DD", () => {
    const dt = new Date(2026, 4, 31, 23, 30, 0);
    expect(formatDateYmdLocal(dt)).toBe("2026-05-31");
  });
});

describe("formatDateDdMmYyyyRuLocal", () => {
  it("formats as дд.мм.гггг from local calendar parts", () => {
    const dt = new Date(2026, 3, 7, 16, 32, 0);
    expect(formatDateDdMmYyyyRuLocal(dt)).toBe("07.04.2026");
  });
});

describe("formatDateTimeDdMmYyyyRuLocal", () => {
  it("formats date only at local midnight", () => {
    const dt = new Date(2026, 3, 7, 0, 0, 0, 0);
    expect(formatDateTimeDdMmYyyyRuLocal(dt)).toBe("07.04.2026");
  });

  it("adds time when non-midnight", () => {
    const dt = new Date(2026, 3, 7, 16, 32, 44, 0);
    expect(formatDateTimeDdMmYyyyRuLocal(dt)).toBe("07.04.2026 16:32:44");
  });
});
