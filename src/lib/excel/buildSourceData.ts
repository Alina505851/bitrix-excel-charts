import { tryParseDate, tryParseNumber } from "@/lib/chart/coerce";
import type { TabularData } from "@/lib/types";

/** Таблица для листа «Исходные данные» в экспорте Excel (клиент и Python). */
export type SourceDataMatrix = {
  headers: string[];
  rows: Array<Array<string | number>>;
};

export function buildSourceData(tabular: TabularData): SourceDataMatrix {
  const headers = tabular.columns.map((c) => c.header || c.key);
  const keys = tabular.columns.map((c) => c.key);
  const isCreatedDateColumn = tabular.columns.map((c) =>
    /дата\s*создания|created[_\s-]*at/i.test(c.header || c.key),
  );
  const formatRuDate = (d: Date) => {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${day}.${m}.${y}`;
  };
  const rows = tabular.rows.map((row) =>
    keys.map((k, idx) => {
      const v = row[k];
      const meta = tabular.columns[idx];
      if (v === null || v === undefined) {
        return "";
      }
      if (isCreatedDateColumn[idx]) {
        const d = v instanceof Date ? v : tryParseDate(v);
        if (d) {
          return formatRuDate(d);
        }
      }
      if (meta?.inferredType === "number") {
        const n =
          typeof v === "number" && Number.isFinite(v)
            ? v
            : tryParseNumber(v);
        if (n !== null) {
          return n;
        }
      }
      if (v instanceof Date) {
        return formatRuDate(v);
      }
      return String(v);
    }),
  );
  return { headers, rows };
}
