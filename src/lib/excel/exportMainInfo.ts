import type { EntityBlockId } from "@/lib/predefinedCharts";
import { formatDataDecimal, formatDataInt } from "@/lib/formatDataNumber";

/** Строки блока «Основные показатели» на первом листе экспорта. Числа — для формата ячеек (₽, %) как у компаний. */
export type ExcelMainInfoLine = { label: string; value: string | number };

/** Снимок KPI боковой панели для компаний/сделок (без React-типов). */
export type ExcelExportKpiSnapshot = {
  total: number;
  activeResponsiblesCount: number | null;
  hasResponsibleColumn: boolean;
  dateColumnHeader: string | null;
  firstRecordDate: string | null;
  lastRecordDate: string | null;
  avgPerMonth: number | null;
  avgPerMonthMeta: {
    monthSpan: number;
    minDateLabel: string;
    maxDateLabel: string;
  } | null;
  growthRateAvgPercent: number | null;
  hasDateColumn: boolean;
};

export type ExcelExportDealsFunnelSnapshot = {
  hasStageColumn: boolean;
  hasAmountColumn: boolean;
  total: number;
  inProgress: number;
  won: number;
  lost: number;
  conversionClosedPercent: number | null;
  potentialIncome: number | null;
  incomeInProgress: number | null;
  incomeWon: number | null;
  incomeLost: number | null;
  conversionMoneyPercent: number | null;
  avgCheckAll: number | null;
  avgCheckWon: number | null;
  avgDaysInWork: number | null;
};

/**
 * Основная информация для первого листа Excel — как в сводке боковой панели
 * (сущность выгрузки, KPI по датам; для сделок — воронка и деньги).
 */
export function buildExcelMainInfoLines(params: {
  entityLabel: string;
  importEntity: EntityBlockId;
  kpi: ExcelExportKpiSnapshot | null;
  dealsFunnel: ExcelExportDealsFunnelSnapshot | null;
}): ExcelMainInfoLine[] {
  const { entityLabel, importEntity, kpi, dealsFunnel } = params;
  const lines: ExcelMainInfoLine[] = [
    { label: "Сущность выгрузки", value: entityLabel },
  ];

  if (kpi) {
    if (kpi.hasResponsibleColumn && kpi.activeResponsiblesCount != null) {
      lines.push({
        label: "Активных ответственных",
        value: kpi.activeResponsiblesCount,
      });
    }
    if (kpi.firstRecordDate) {
      lines.push({
        label: "Дата первой записи",
        value: kpi.firstRecordDate,
      });
    }
    if (kpi.lastRecordDate) {
      lines.push({
        label: "Дата последней записи",
        value: kpi.lastRecordDate,
      });
    }
    if (kpi.avgPerMonth != null && kpi.avgPerMonthMeta) {
      lines.push({
        label: "Среднее в месяц",
        value: `${formatDataDecimal(kpi.avgPerMonth, 2, 0)} (${formatDataInt(kpi.total)} ÷ ${kpi.avgPerMonthMeta.monthSpan} мес., ${kpi.avgPerMonthMeta.minDateLabel} — ${kpi.avgPerMonthMeta.maxDateLabel})`,
      });
    }
    if (kpi.growthRateAvgPercent != null) {
      lines.push({
        label: "Темп роста (среднее помесячных %)",
        value: kpi.growthRateAvgPercent / 100,
      });
    }
  }

  if (importEntity === "deals" && dealsFunnel) {
    const f = dealsFunnel;
    if (f.hasStageColumn) {
      lines.push({
        label: "Всего сделок (воронка)",
        value: f.total,
      });
      lines.push({
        label: "В работе",
        value: f.inProgress,
      });
      lines.push({
        label: "Успешных",
        value: f.won,
      });
      lines.push({
        label: "Проигранных",
        value: f.lost,
      });
      lines.push({
        label: "Конверсия (успех к закрытым)",
        value:
          f.conversionClosedPercent != null
            ? f.conversionClosedPercent / 100
            : "—",
      });
    }
    if (f.hasAmountColumn) {
      if (f.potentialIncome != null) {
        lines.push({
          label: "Потенциальный доход",
          value: f.potentialIncome,
        });
      }
      if (f.incomeInProgress != null) {
        lines.push({
          label: "Доход в работе",
          value: f.incomeInProgress,
        });
      }
      if (f.incomeWon != null) {
        lines.push({
          label: "Доход от выигранных",
          value: f.incomeWon,
        });
      }
      if (f.incomeLost != null) {
        lines.push({
          label: "Упущенный доход",
          value: f.incomeLost,
        });
      }
      if (f.conversionMoneyPercent != null) {
        lines.push({
          label: "Конверсия по сумме закрытых",
          value: f.conversionMoneyPercent / 100,
        });
      }
      if (f.avgCheckAll != null) {
        lines.push({
          label: "Средний чек",
          value: f.avgCheckAll,
        });
      }
      if (f.avgCheckWon != null) {
        lines.push({
          label: "Средний чек (выигранные)",
          value: f.avgCheckWon,
        });
      }
    }
    if (f.avgDaysInWork != null) {
      lines.push({
        label: "Средний цикл в работе, дн.",
        value: f.avgDaysInWork,
      });
    }
  }

  return lines;
}
