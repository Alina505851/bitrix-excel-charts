import type { EntityBlockId } from "@/lib/predefinedCharts";
import type { AggregationMode } from "@/lib/types";

/** Подпись для count_nonempty по ID — как у компаний, но с сущностью раздела CRM. */
function legendCountNonEmptyById(entity: EntityBlockId | undefined): string {
  switch (entity ?? "generic") {
    case "companies":
      return "Кол-во компаний · количество компаний";
    case "deals":
      return "Кол-во сделок · количество сделок";
    case "leads":
      return "Кол-во лидов · количество лидов";
    case "contacts":
      return "Кол-во контактов · количество контактов";
    case "quotes":
      return "Кол-во коммерческих предложений · количество записей";
    default:
      return "Кол-во записей · количество записей";
  }
}

/**
 * Подпись серии для таблиц/графиков: «колонка · смысл агрегации».
 * Для суммы по колонке, в названии которой уже есть «сумма», не повторяем слово «сумма», а указываем рубли.
 * `entity` нужен для режима count_nonempty по колонке ID (раздел CRM).
 */
export function seriesLegendLabel(
  header: string,
  aggregation: AggregationMode,
  entity?: EntityBlockId,
): string {
  const h = header.trim();
  const lower = h.toLowerCase();

  if (aggregation === "count_nonempty") {
    if (lower.includes("id") || lower.includes("ид")) {
      return legendCountNonEmptyById(entity);
    }
    return `${h} · количество непустых значений`;
  }
  if (aggregation === "count") {
    return `${h} · количество записей`;
  }
  if (aggregation === "avg") {
    return `${h} · среднее значение`;
  }
  if (aggregation === "sum") {
    if (/\bсумм/i.test(h) || lower.includes("сумма")) {
      return `${h} · руб.`;
    }
    return `${h} · сумма, руб.`;
  }
  return `${h} · сумма, руб.`;
}
