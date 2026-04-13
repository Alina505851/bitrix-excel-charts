import type { ColumnMeta } from "@/lib/types";
import { dealCreatedDateColumnForKpi } from "@/lib/predefinedCharts";
import {
  collectParsedDatesForKpiColumn,
  formatKpiRecordDateRu,
  minMaxFromDateList,
} from "@/lib/report/kpiColumnDates";

/** Ключ колонки «Дата создания» для KPI и воронки (после выбора среди дубликатов по выборке строк). */
export function resolveDealKpiCreationDateColumnKey(
  columns: ColumnMeta[],
  sampleRows: Record<string, unknown>[],
): string | null {
  return dealCreatedDateColumnForKpi(columns, sampleRows)?.key ?? null;
}

export type DealFirstLastRecordLabels = {
  firstLabel: string;
  lastLabel: string;
  min: Date;
  max: Date;
};

/**
 * Колонка даты создания и распарсенные даты — для сводки и темпа роста без двойного обхода строк.
 */
export function computeDealCreationDateSeriesForKpi(
  columns: ColumnMeta[],
  rows: Record<string, unknown>[],
): { dateKey: string | null; dates: Date[] } {
  const dateKey = resolveDealKpiCreationDateColumnKey(columns, rows);
  if (!dateKey) {
    return { dateKey: null, dates: [] };
  }
  const meta = columns.find((c) => c.key === dateKey) ?? null;
  return {
    dateKey,
    dates: collectParsedDatesForKpiColumn(rows, dateKey, meta),
  };
}

/**
 * Первая и последняя дата записи сделки: min/max по колонке даты создания после фильтров.
 * Подписи — для отображения в UI (ru-RU).
 */
export function getDealFirstAndLastRecordLabels(
  columns: ColumnMeta[],
  rows: Record<string, unknown>[],
): DealFirstLastRecordLabels | null {
  const { dateKey, dates } = computeDealCreationDateSeriesForKpi(columns, rows);
  if (!dateKey) {
    return null;
  }
  const mm = minMaxFromDateList(dates);
  if (!mm) {
    return null;
  }
  return {
    firstLabel: formatKpiRecordDateRu(mm.min),
    lastLabel: formatKpiRecordDateRu(mm.max),
    min: mm.min,
    max: mm.max,
  };
}
