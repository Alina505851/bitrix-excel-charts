import type { ColumnMeta } from "@/lib/types";
import {
  HARDCODED_DEAL_FUNNEL_STAGE_ORDER,
  USE_HARDCODED_DEAL_FUNNEL_STAGE_ORDER,
} from "@/lib/chart/dealStageOrderHardcoded";

/**
 * Порядок стадий для графиков: сначала список из UI; если пусто — опционально
 * {@link HARDCODED_DEAL_FUNNEL_STAGE_ORDER}; иначе `null` (используется
 * {@link sortDealStageKeysDefault}).
 */
export function getEffectiveDealStageOrder(
  userOrder: string[] | null | undefined,
): string[] | null {
  if (userOrder?.length) {
    return userOrder;
  }
  if (
    USE_HARDCODED_DEAL_FUNNEL_STAGE_ORDER &&
    HARDCODED_DEAL_FUNNEL_STAGE_ORDER.length > 0
  ) {
    return [...HARDCODED_DEAL_FUNNEL_STAGE_ORDER];
  }
  return null;
}

/** Нормализация подписи стадии для сопоставления с пользовательским списком. */
export function normStageLabel(s: string): string {
  return s
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/ё/g, "е");
}

/**
 * Разбор многострочного ввода из интерфейса: одна стадия — одна строка,
 * пустые строки отбрасываются.
 */
export function parseDealStageOrderLines(raw: string): string[] {
  return raw
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter((l) => l.length > 0);
}

function buildOrderIndex(order: string[]): Map<string, number> {
  const m = new Map<string, number>();
  for (let i = 0; i < order.length; i++) {
    const key = normStageLabel(order[i]!);
    if (!m.has(key)) {
      m.set(key, i);
    }
  }
  return m;
}

/**
 * Ведущий номер в подписи стадии (типичный экспорт CRM: «1. Новая», «10) КП»).
 * Числа в середине строки не считаются.
 */
export function leadingStageIndex(label: string): number | null {
  const s = label.trim();
  const dotted = /^\s*(\d+)\s*[\.\)]/.exec(s);
  if (dotted) {
    const n = Number(dotted[1]);
    return Number.isFinite(n) ? n : null;
  }
  return null;
}

/**
 * Порядок стадий без явного списка из UI: сначала по ведущему номеру (если есть),
 * подписи без номера — в конце, с сохранением исходного порядка (как в данных).
 */
export function sortDealStageKeysDefault(keys: string[]): string[] {
  if (keys.length <= 1) {
    return [...keys];
  }
  const anyNumbered = keys.some((k) => leadingStageIndex(k) !== null);
  if (!anyNumbered) {
    return [...keys];
  }
  const indexed = keys.map((k, i) => ({ k, i }));
  return [...indexed]
    .sort((a, b) => {
      const na = leadingStageIndex(a.k);
      const nb = leadingStageIndex(b.k);
      const va = na === null ? Number.POSITIVE_INFINITY : na;
      const vb = nb === null ? Number.POSITIVE_INFINITY : nb;
      if (va !== vb) {
        return va - vb;
      }
      return a.i - b.i;
    })
    .map((x) => x.k);
}

/**
 * Сортировка категорий оси X по пользовательскому порядку.
 * Если `order` пустой — алфавит (ru).
 */
export function sortDealPipelineStageKeys(
  keys: string[],
  order: string[] | null | undefined,
): string[] {
  if (!order?.length) {
    return [...keys].sort((a, b) => a.localeCompare(b, "ru", { numeric: true }));
  }
  const indexMap = buildOrderIndex(order);
  return [...keys].sort((a, b) => {
    const ia = indexMap.get(normStageLabel(a));
    const ib = indexMap.get(normStageLabel(b));
    const na = ia === undefined ? Number.POSITIVE_INFINITY : ia;
    const nb = ib === undefined ? Number.POSITIVE_INFINITY : ib;
    if (na !== nb) {
      return na - nb;
    }
    return a.localeCompare(b, "ru", { numeric: true });
  });
}

/** Колонки стадии сделки / предыдущей стадии — применять порядок из настроек. */
export function isDealStageLikeColumn(meta: ColumnMeta): boolean {
  if (meta.inferredType === "date") {
    return false;
  }
  const h = meta.header.trim().toLowerCase();
  const k = meta.key.toLowerCase();
  const t = `${h} ${k}`;
  if (t.includes("стадия") && t.includes("сделк")) {
    return true;
  }
  if (t.includes("предыдущ") && t.includes("стадия")) {
    return true;
  }
  // «Стадия» без явного «сделка» в заголовке (частые выгрузки). Не путать со стадией лида.
  if (t.includes("стадия") && !t.includes("лид")) {
    return true;
  }
  return false;
}
