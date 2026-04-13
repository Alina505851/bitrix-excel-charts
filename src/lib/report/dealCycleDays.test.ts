import { describe, expect, it } from "vitest";
import { calendarDaysBetweenLocalMidnight } from "@/lib/report/dealCycleDays";

describe("calendarDaysBetweenLocalMidnight", () => {
  it("returns 0 for the same calendar day (different times)", () => {
    const from = new Date(2026, 3, 10, 9, 30);
    const to = new Date(2026, 3, 10, 18, 0);
    expect(calendarDaysBetweenLocalMidnight(from, to)).toBe(0);
  });

  it("returns 1 for consecutive calendar days", () => {
    const from = new Date(2026, 3, 10, 23, 0);
    const to = new Date(2026, 3, 11, 1, 0);
    expect(calendarDaysBetweenLocalMidnight(from, to)).toBe(1);
  });

  it("returns negative when close is before creation (invalid row)", () => {
    const from = new Date(2026, 3, 12, 0, 0);
    const to = new Date(2026, 3, 10, 0, 0);
    expect(calendarDaysBetweenLocalMidnight(from, to)).toBeLessThan(0);
  });
});
