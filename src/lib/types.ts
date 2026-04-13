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
  /**
   * Если `values` не задан, подпись должна содержать каждую подстроку (без учёта регистра).
   * Удобно для стадий вида «7. Отправка КП», «C2: Отправка КП».
   */
  valuesContainAll?: string[];
  /**
   * В сочетании с `valuesContainAll` или отдельно: хотя бы одна подстрока (без учёта регистра).
   * Например `["кп", "kp"]` для латиницы в выгрузке.
   */
  valuesContainAny?: string[];
  /**
   * Подпись ячейки после `groupLabel` не должна совпадать ни с одним значением
   * (например исключить финальные стадии «успех» и «провал» для сделок «в работе»).
   */
  excludeValues?: string[];
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

/** Фиксированные категории на оси X: среднее по `valueKey` после `filters` (и общих `filters` конфига). */
export type LiteralAvgBarSpec = {
  label: string;
  filters: ChartFilter[];
  /** Колонка, по которой считается среднее. */
  valueKey: string;
  /** Ключ серии в `yKeys`, куда записать среднее (остальные серии в строке — 0). */
  seriesKey: string;
};

/**
 * Две доли круговой диаграммы: строки с непустым `xKey`, попавшие под repeatMatchAny,
 * и все остальные такие строки (взаимное исключение).
 */
export type LiteralPieComplementSpec = {
  repeatLabel: string;
  nonRepeatLabel: string;
  /** Строка в «повторные», если выполняется любой из фильтров (по колонке типа сделки и т.п.). */
  repeatMatchAny: ChartFilter[];
};

export type ChartConfig = {
  xKey: string | null;
  yKeys: string[];
  /**
   * Источник значений для каждой серии (длина = yKeys). Если не задано, берётся yKeys[i].
   * Нужно, когда ключ серии в точке данных отличается от ключа колонки (несколько метрик из одной колонки).
   */
  ySourceKeys?: string[];
  /**
   * Режим агрегации по сериям (длина = yKeys). Если не задано, для всех серий используется `aggregation`.
   */
  yAggregations?: AggregationMode[];
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
  /**
   * Для count/count_nonempty: считать уникальные записи по значению колонки (например, уникальные сделки по ID),
   * чтобы дубликаты строк из выгрузки не завышали результат.
   */
  countDistinctByKey?: string;
  /**
   * Если задано, агрегация не по `xKey`, а по списку «литералов»: по одной точке на элемент
   * (несколько метрик с разными фильтрами и колонками значения).
   */
  literalAvgBars?: LiteralAvgBarSpec[];
  /**
   * Круговая диаграмма из двух сегментов: «повторные» и «неповторные» по правилам в спецификации.
   * Требует `chartType: "pie"`, `aggregation: "count"`, непустой `xKey` (колонка признака).
   */
  literalPieComplement?: LiteralPieComplementSpec;
};

export const defaultChartConfig = (): ChartConfig => ({
  xKey: null,
  yKeys: [],
  chartType: "bar",
  aggregation: "sum",
  filters: [],
});
