export type InferredType = "string" | "number" | "date" | "unknown";

export type ColumnMeta = {
  key: string;
  header: string;
  inferredType: InferredType;
  sampleValues: unknown[];
};

export type TabularData = {
  sheetNames: string[];
  activeSheet: string;
  columns: ColumnMeta[];
  rows: Record<string, unknown>[];
};

export type ChartType = "bar" | "line" | "area" | "pie";

export type AggregationMode = "sum" | "count" | "avg" | "count_nonempty";

/** Группировка оси X для колонок с датой (иначе — по дням, YYYY-MM-DD). */
export type DateGranularity = "day" | "month" | "quarter" | "year";

export type ChartFilter = {
  columnKey: string;
  /** Для текста / даты как категорий — список допустимых значений (подпись ячейки) */
  values?: string[];
  min?: number;
  max?: number;
  /** Для колонок-дат: начало периода YYYY-MM-DD (включительно) */
  dateFrom?: string;
  /** Для колонок-дат: конец периода YYYY-MM-DD (включительно) */
  dateTo?: string;
  /** Для колонок-дат: календарные месяцы YYYY-MM (строка попадает в выборку, если дата в одном из месяцев) */
  calendarMonths?: string[];
  /** Для колонок-дат: кварталы вида YYYY-Q1 … YYYY-Q4 */
  calendarQuarters?: string[];
};

export type ChartConfig = {
  xKey: string | null;
  yKeys: string[];
  chartType: ChartType;
  aggregation: AggregationMode;
  filters: ChartFilter[];
  /** Для дат на оси X: месяц/год вместо отдельных дней. */
  dateGranularity?: DateGranularity;
  /**
   * Нарастающий итог по отсортированным точкам оси X: каждое значение Y = сумма всех предыдущих + текущее.
   * Имеет смысл для ряда «число компаний по периодам».
   */
  cumulative?: boolean;
};

export const defaultChartConfig = (): ChartConfig => ({
  xKey: null,
  yKeys: [],
  chartType: "bar",
  aggregation: "sum",
  filters: [],
});
