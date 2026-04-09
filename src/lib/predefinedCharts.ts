import type {
  ChartConfig,
  ChartFilter,
  ColumnMeta,
  TabularData,
} from "@/lib/types";

export type EntityBlockId =
  | "leads"
  | "deals"
  | "contacts"
  | "companies"
  | "quotes";

export const ENTITY_BLOCKS: {
  id: EntityBlockId;
  label: string;
  description: string;
}[] = [
  {
    id: "leads",
    label: "Лиды",
    description:
      "Стадия, источник и сумма; воронка, UTM, причины отказа; динамика линией и областью, средние по дням.",
  },
  {
    id: "deals",
    label: "Сделки",
    description:
      "Стадии, суммы, контакты и менеджеры; выигранные сделки отдельными графиками; воронка, средний чек, динамика.",
  },
  {
    id: "contacts",
    label: "Контакты",
    description:
      "Компания, источник, ответственный; должность, география; приток по датам (линия и область).",
  },
  {
    id: "companies",
    label: "Компании",
    description:
      "Отрасль, тип, локация, UTM, источник, бренд, путь клиента, активность, КП и типовые поля выгрузки Битрикс24.",
  },
  {
    id: "quotes",
    label: "Комм. предложения (КП)",
    description:
      "Выгрузка КП из CRM: по месяцам только отправленные предложения с заполненной суммой.",
  },
];

type ResolveOk = { ok: true; config: ChartConfig };
export type ResolveFail = { ok: false; error: string };

export type PredefinedChartSpec = {
  id: string;
  entity: EntityBlockId;
  title: string;
  description: string;
  resolve: (columns: ColumnMeta[]) => ResolveOk | ResolveFail;
};

function normHeader(h: string): string {
  return h.trim().toLowerCase();
}

function headerMatches(meta: ColumnMeta, fragments: string[]): boolean {
  const h = normHeader(meta.header);
  const k = normHeader(meta.key);
  return fragments.some((f) => h.includes(f) || k.includes(f));
}

/**
 * Идентификаторы (ID) нельзя агрегировать суммой/средним:
 * это технические ключи, а не бизнес-метрика.
 */
function isLikelyIdentifierColumn(meta: ColumnMeta): boolean {
  const h = normHeader(meta.header);
  const k = normHeader(meta.key);
  const text = `${h} ${k}`;
  const hasLatinId =
    text.includes(" id") ||
    text.includes("id ") ||
    text.startsWith("id") ||
    text.endsWith("id") ||
    text.includes("_id") ||
    text.includes("guid") ||
    text.includes("uuid");
  const hasCyrId =
    text.includes(" ид") ||
    text.includes("ид ") ||
    text.startsWith("ид") ||
    text.endsWith("ид");
  return hasLatinId || hasCyrId;
}

function pickColumn(
  columns: ColumnMeta[],
  opts: {
    types: ColumnMeta["inferredType"][];
    anyHeader?: string[];
    preferHeaders?: string[];
  },
): ColumnMeta | null {
  const { types, anyHeader = [], preferHeaders = [] } = opts;
  const pool = columns.filter((c) => types.includes(c.inferredType));
  if (pool.length === 0) {
    return null;
  }
  for (const frag of preferHeaders) {
    const hit = pool.find((c) => headerMatches(c, [frag]));
    if (hit) {
      return hit;
    }
  }
  if (anyHeader.length > 0) {
    const hit = pool.find((c) => headerMatches(c, anyHeader));
    if (hit) {
      return hit;
    }
  }
  return pool[0] ?? null;
}

function pickStrictHeaderColumn(
  columns: ColumnMeta[],
  opts: {
    types: ColumnMeta["inferredType"][];
    exactHeaders: string[];
  },
): ColumnMeta | null {
  const { types, exactHeaders } = opts;
  const exact = new Set(exactHeaders.map((h) => normHeader(h)));
  const pool = columns.filter((c) => types.includes(c.inferredType));
  return (
    pool.find((c) => exact.has(normHeader(c.header)) || exact.has(normHeader(c.key))) ??
    null
  );
}

function firstNumberColumn(columns: ColumnMeta[]): ColumnMeta | null {
  const metricPool = columns.filter(
    (c) => c.inferredType === "number" && !isLikelyIdentifierColumn(c),
  );
  return metricPool[0] ?? null;
}

function stageColumnLeads(columns: ColumnMeta[]): ColumnMeta | null {
  return (
    pickColumn(columns, {
      types: ["string", "unknown"],
      preferHeaders: [
        "стадия лида",
        "статус лида",
        "стадия",
        "статус",
      ],
      anyHeader: ["стадия", "статус", "этап", "воронка"],
    }) ?? pickColumn(columns, { types: ["string", "unknown"], anyHeader: [] })
  );
}

function stageColumnDeals(columns: ColumnMeta[]): ColumnMeta | null {
  return (
    pickColumn(columns, {
      types: ["string", "unknown"],
      preferHeaders: [
        "стадия сделки",
        "стадия",
        "статус",
        "этап",
      ],
      anyHeader: ["стадия", "статус", "воронка"],
    }) ?? pickColumn(columns, { types: ["string", "unknown"], anyHeader: [] })
  );
}

function dateLikeColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return (
    pickColumn(columns, {
      types: ["date"],
      preferHeaders: ["дата", "date", "создан", "изменен"],
      anyHeader: ["дата", "date", "период"],
    }) ??
    pickColumn(columns, {
      types: ["string"],
      anyHeader: [
        "дата создания",
        "дата создания сделки",
        "дата создания лида",
        "дата",
      ],
    })
  );
}

function amountLikeColumn(columns: ColumnMeta[]): ColumnMeta | null {
  const metricPool = columns.filter(
    (c) => c.inferredType === "number" && !isLikelyIdentifierColumn(c),
  );
  return pickColumn(metricPool, {
    types: ["number"],
    preferHeaders: [
      "сумма",
      "опорт",
      "opportunity",
      "amount",
      "value",
      "стоимость",
      "годовой оборот",
    ],
    anyHeader: ["сумм", "цен", "руб", "оборот"],
  });
}

/** Сумма / плановая сумма лида */
function leadValueColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return (
    pickColumn(columns, {
      types: ["number"],
      preferHeaders: [
        "сумма",
        "опорт",
        "opportunity",
        "ожидаемая",
        "плановая",
      ],
      anyHeader: ["сумм", "цен"],
    }) ?? firstNumberColumn(columns)
  );
}

export function responsibleColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickColumn(columns, {
    types: ["string", "unknown"],
    anyHeader: [
      "ответственн",
      "ответств",
      "owner",
      "менеджер",
      "пользователь",
    ],
  });
}

function sourceColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickColumn(columns, {
    types: ["string", "unknown"],
    anyHeader: [
      "источник",
      "source",
      "канал",
      "рекламная система",
    ],
  });
}

/** Источник именно для карточки компании (строгий выбор без широких эвристик). */
function companySourceColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickStrictHeaderColumn(columns, {
    types: ["string", "unknown"],
    exactHeaders: [
      "источник",
      "источник компании",
      "uf_crm_sourceid",
      "source",
      "company source",
      "канал привлечения",
      "рекламная система",
    ],
  });
}

function companyColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickColumn(columns, {
    types: ["string", "unknown"],
    anyHeader: [
      "компания",
      "название компании",
      "связанная компания",
      "company",
      "работает в компании",
    ],
  });
}

function industryOrTypeColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickColumn(columns, {
    types: ["string", "unknown"],
    preferHeaders: [
      "отрасль",
      "сфера деятельности",
      "тип компании",
      "индустрия",
      "вид деятельности",
    ],
    anyHeader: [
      "отрасль",
      "индустрия",
      "сфера",
      "тип компании",
      "вид деятельности",
    ],
  });
}

/** Только «Сфера деятельности» (отдельно от отрасли / типа компании). */
function sphereOfActivityColumn(columns: ColumnMeta[]): ColumnMeta | null {
  const pool = columns.filter(
    (c) => c.inferredType === "string" || c.inferredType === "unknown",
  );
  return (
    pool.find((c) => {
      const h = normHeader(c.header);
      const k = normHeader(c.key);
      return h.includes("сфера деятельности") || k.includes("сфера деятельности");
    }) ?? null
  );
}

/** Поле «Тип компании» / тип клиента в выгрузке. */
function companyTypeColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickColumn(columns, {
    types: ["string", "unknown"],
    preferHeaders: ["тип компании"],
    anyHeader: [
      "тип компании",
      "тип клиента (клиенты)",
      "тип клиента",
      "тип клиента",
    ],
  });
}

/** Численность компании (не ID сотрудника): «Кол-во сотрудников» и т.п. */
function headcountColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickColumn(columns, {
    types: ["number"],
    preferHeaders: ["кол-во сотрудников", "численность"],
    anyHeader: [
      "кол-во сотруд",
      "количество сотруд",
      "численность",
      "штат",
      "employees",
      "staff",
    ],
  });
}

/** Колонка с привязкой к пользователю / сотруднику CRM, без «численности» и «кол-ва сотрудников». */
function employeeIdLikeColumn(columns: ColumnMeta[]): ColumnMeta | null {
  const exclude = [
    "кол-во",
    "количество",
    "численность",
    "штат",
    "годовой",
    "оборот",
    "бюджет",
    "валюта",
    "оплат",
    "inn",
    "инн",
  ];
  const preferFrags = [
    "связанный пользователь",
    "id сотруд",
    "ид сотруд",
    "employee id",
    "user id",
  ];
  const looseFrags = ["employee", "staff id"];
  const isExcluded = (c: ColumnMeta) => headerMatches(c, exclude);
  const tryPool = (pool: ColumnMeta[], frags: string[]) =>
    pool.find((c) => headerMatches(c, frags) && !isExcluded(c)) ?? null;
  const numPool = columns.filter((c) => c.inferredType === "number");
  const strPool = columns.filter(
    (c) => c.inferredType === "string" || c.inferredType === "unknown",
  );
  return (
    tryPool(numPool, preferFrags) ??
    tryPool(strPool, preferFrags) ??
    tryPool(numPool, looseFrags) ??
    tryPool(strPool, looseFrags) ??
    null
  );
}

/** Дата похожа на «дата создания компании» в выгрузке Битрикс24 */
export function companyCreatedDateColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return (
    pickStrictHeaderColumn(columns, {
      types: ["date"],
      exactHeaders: [
        "дата создания компании",
        "дата создания",
        "дата регистрации",
        "created at",
        "created_at",
        "date created",
        "registration date",
      ],
    }) ??
    pickStrictHeaderColumn(columns, {
      types: ["string"],
      exactHeaders: [
        "дата создания компании",
        "дата создания",
        "дата регистрации",
        "created at",
        "created_at",
        "date created",
        "registration date",
      ],
    })
  );
}

/** Город / регион / локация в карточке компании */
function locationColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickStrictHeaderColumn(columns, {
    types: ["string", "unknown"],
    exactHeaders: ["локация", "город", "регион", "страна", "city", "region"],
  });
}

function utmSourceColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickStrictHeaderColumn(columns, {
    types: ["string", "unknown"],
    exactHeaders: ["utm_source", "utm source", "uf_crm_sourceid"],
  });
}

function utmMediumColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickStrictHeaderColumn(columns, {
    types: ["string", "unknown"],
    exactHeaders: ["utm_medium", "utm medium"],
  });
}

function utmCampaignColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickStrictHeaderColumn(columns, {
    types: ["string", "unknown"],
    exactHeaders: ["utm_campaign", "utm campaign", "uf_crm_formname"],
  });
}

function messengerColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickStrictHeaderColumn(columns, {
    types: ["string", "unknown"],
    exactHeaders: ["мессенджер", "messenger"],
  });
}

function brandColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickStrictHeaderColumn(columns, {
    types: ["string", "unknown"],
    exactHeaders: ["бренд", "бренд (тест связь со смарт процессом)"],
  });
}

function clientJourneyColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickStrictHeaderColumn(columns, {
    types: ["string", "unknown"],
    exactHeaders: ["путь клиента", "customer journey"],
  });
}

function companyExportStatusColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickStrictHeaderColumn(columns, {
    types: ["string", "unknown"],
    exactHeaders: ["статус", "статус компании", "status"],
  });
}

function companyLifecycleColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickStrictHeaderColumn(columns, {
    types: ["string", "unknown"],
    exactHeaders: ["жизненный цикл", "жизненный цикл (клиенты)"],
  });
}

function companyHasDealColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickColumn(columns, {
    types: ["string", "unknown"],
    anyHeader: ["есть сделка"],
  });
}

function companyCreatedByColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickStrictHeaderColumn(columns, {
    types: ["string", "unknown"],
    exactHeaders: ["кем создана", "кем создано", "кем создан"],
  });
}

function companyNeedColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickStrictHeaderColumn(columns, {
    types: ["string", "unknown"],
    exactHeaders: ["потребность"],
  });
}

function parentCompanyColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickStrictHeaderColumn(columns, {
    types: ["string", "unknown"],
    exactHeaders: ["родительская компания", "родительская компания (клиенты)"],
  });
}

function companyBudgetKpColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickStrictHeaderColumn(columns, {
    types: ["number"],
    exactHeaders: ["бюджет кп", "бюджет планируемый"],
  });
}

/** Бизнес-метрика компании для сумм по менеджерам (не ID и не технические числа). */
function companyBusinessMetricColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickStrictHeaderColumn(columns, {
    types: ["number"],
    exactHeaders: [
      "выручка",
      "оборот",
      "годовой оборот",
      "доход",
      "revenue",
      "annual revenue",
      "бюджет кп",
      "бюджет планируемый",
      "сумма",
      "стоимость",
    ],
  });
}

/** Дата последней активности / коммуникации. */
function companyLastActivityColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return (
    pickStrictHeaderColumn(columns, {
      types: ["date"],
      exactHeaders: [
        "последняя активность",
        "дата последней активности",
        "дата последней коммуникации",
      ],
    }) ??
    pickStrictHeaderColumn(columns, {
      types: ["string", "unknown"],
      exactHeaders: [
        "последняя активность",
        "дата последней активности",
        "дата последней коммуникации",
      ],
    })
  );
}

function fallbackStringColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickColumn(columns, { types: ["string", "unknown"], anyHeader: [] });
}

/** Компания в карточке лида (в выгрузке часто отдельно от сделки/контакта). */
function leadCompanyColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return (
    pickColumn(columns, {
      types: ["string", "unknown"],
      preferHeaders: [
        "компания лида",
        "название компании лида",
        "компания (лид)",
      ],
      anyHeader: ["компания лида", "lead company"],
    }) ?? companyColumn(columns)
  );
}

/** Причина отказа / закрытия лида. */
function rejectReasonColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickColumn(columns, {
    types: ["string", "unknown"],
    preferHeaders: [
      "причина отказа",
      "причина закрытия",
      "причина отказа лида",
      "причина потери",
    ],
    anyHeader: [
      "причина отказа",
      "причина закрытия",
      "причина потер",
      "reason",
      "неуспех",
      "отказ",
    ],
  });
}

/** UTM / детализация источника (Битрикс24: реклама, вебформы). */
function utmOrDetailedSourceColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickColumn(columns, {
    types: ["string", "unknown"],
    preferHeaders: [
      "utm_source",
      "utm medium",
      "utm campaign",
      "детализированный источник",
    ],
    anyHeader: [
      "utm_source",
      "utm_medium",
      "utm_campaign",
      "utm content",
      "utm term",
      "рекламная кампания",
      "детализирован",
    ],
  });
}

/** Воронка / направление (лиды и сделки). */
function funnelOrDirectionColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickColumn(columns, {
    types: ["string", "unknown"],
    preferHeaders: [
      "воронка",
      "направление",
      "воронка продаж",
      "категория лида",
    ],
    anyHeader: ["воронка", "pipeline", "туннель", "направление продаж"],
  });
}

/** Контакт, привязанный к сделке. */
function dealContactColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickColumn(columns, {
    types: ["string", "unknown"],
    preferHeaders: [
      "контакт",
      "связанный контакт",
      "имя контакта",
    ],
    anyHeader: [
      "контакт",
      "contact id",
      "связанный контакт",
    ],
  });
}

/** Воронка/направление сделки (узче, чем общая «воронка»). */
function dealPipelineColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickColumn(columns, {
    types: ["string", "unknown"],
    preferHeaders: [
      "направление сделки",
      "воронка сделки",
      "воронка",
    ],
    anyHeader: ["направление", "воронка"],
  });
}

/** Новая / повторная и др. типизация сделки в выгрузке. */
function dealBusinessTypeColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickColumn(columns, {
    types: ["string", "unknown"],
    preferHeaders: ["тип сделки", "повторная сделка", "новая сделка"],
    anyHeader: [
      "повторн",
      "новая сделка",
      "repeat",
      "тип сделки",
    ],
  });
}

/**
 * Типичные подписи выигранной сделки в Битрикс24 / похожих выгрузках.
 * Если в вашей CRM другое название стадии — добавьте глобальный фильтр по стадии
 * или скорректируйте выгрузку.
 */
const WON_DEAL_STAGE_VALUES: string[] = [
  "Успешно реализована",
  "успешно реализована",
  "УСПЕШНО РЕАЛИЗОВАНА",
  "Успешная сделка",
  "успешная сделка",
  "WON",
  "won",
  "Won",
  "CLOSED WON",
  "Closed Won",
  "Закрыта и реализована",
  "Закрыта с выигрышем",
  "Сделка выиграна",
  "Выиграна",
  "выиграна",
  "Успех",
  "успех",
];

function wonDealStageFilters(stageKey: string): ChartFilter[] {
  return [{ columnKey: stageKey, values: [...WON_DEAL_STAGE_VALUES] }];
}

/** 1 квартал 2026 г. (январь–март включительно) по колонке даты в выгрузке. */
const DEALS_Q1_2026_FROM = "2026-01-01";
const DEALS_Q1_2026_TO = "2026-03-31";

function wonDealQ1_2026Filters(
  stage: ColumnMeta,
  dateCol: ColumnMeta,
): ChartFilter[] {
  return [
    ...wonDealStageFilters(stage.key),
    {
      columnKey: dateCol.key,
      dateFrom: DEALS_Q1_2026_FROM,
      dateTo: DEALS_Q1_2026_TO,
    },
  ];
}

/** Дата для динамики сделок: закрытие / изменение / создание. */
function dealClosedDateColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return (
    pickColumn(columns, {
      types: ["date"],
      preferHeaders: [
        "дата закрытия",
        "дата завершения",
        "дата успешного закрытия",
        "дата изменения",
        "дата создания",
      ],
      anyHeader: ["закрыт", "заверш"],
    }) ??
    pickColumn(columns, {
      types: ["string"],
      anyHeader: [
        "дата закрытия",
        "дата завершения",
        "дата изменения сделки",
      ],
    }) ??
    dateLikeColumn(columns)
  );
}

/** Должность контакта. */
function contactPositionColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickColumn(columns, {
    types: ["string", "unknown"],
    preferHeaders: ["должность", "должность контакта", "post"],
    anyHeader: [
      "должность",
      "position",
      "post",
      "job title",
    ],
  });
}

function barSum(
  x: ColumnMeta,
  y: ColumnMeta,
): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "bar",
    aggregation: "sum",
    filters: [],
  };
}

function barCount(x: ColumnMeta, y: ColumnMeta): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "bar",
    aggregation: "count",
    filters: [],
  };
}

function lineSum(x: ColumnMeta, y: ColumnMeta): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "line",
    aggregation: "sum",
    filters: [],
  };
}

function lineCount(x: ColumnMeta, y: ColumnMeta): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "line",
    aggregation: "count",
    filters: [],
  };
}

function barAvg(x: ColumnMeta, y: ColumnMeta): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "bar",
    aggregation: "avg",
    filters: [],
  };
}

function barCountNonempty(x: ColumnMeta, y: ColumnMeta): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "bar",
    aggregation: "count_nonempty",
    filters: [],
  };
}


function pieSum(x: ColumnMeta, y: ColumnMeta): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "pie",
    aggregation: "sum",
    filters: [],
  };
}

function pieCount(x: ColumnMeta, y: ColumnMeta): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "pie",
    aggregation: "count",
    filters: [],
  };
}

function areaSum(x: ColumnMeta, y: ColumnMeta): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "area",
    aggregation: "sum",
    filters: [],
  };
}

function areaCount(x: ColumnMeta, y: ColumnMeta): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "area",
    aggregation: "count",
    filters: [],
  };
}

function lineAvg(x: ColumnMeta, y: ColumnMeta): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "line",
    aggregation: "avg",
    filters: [],
  };
}

/** Число строк по месяцам (дата на оси X → ключи вида 2024-09, подписи «сент. 2024»). */
function lineCountByMonth(x: ColumnMeta, y: ColumnMeta): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "line",
    aggregation: "count",
    filters: [],
    dateGranularity: "month",
  };
}

/** Нарастающий итог числа строк по месяцам (кривая монотонно растёт). */
function areaCumulativeCountByMonth(x: ColumnMeta, y: ColumnMeta): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "area",
    aggregation: "count",
    filters: [],
    dateGranularity: "month",
    cumulative: true,
  };
}


function barSumWon(x: ColumnMeta, y: ColumnMeta, stage: ColumnMeta): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "bar",
    aggregation: "sum",
    filters: wonDealStageFilters(stage.key),
  };
}

function barCountWon(x: ColumnMeta, y: ColumnMeta, stage: ColumnMeta): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "bar",
    aggregation: "count",
    filters: wonDealStageFilters(stage.key),
  };
}

function lineSumByMonthWon(
  x: ColumnMeta,
  y: ColumnMeta,
  stage: ColumnMeta,
): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "line",
    aggregation: "sum",
    filters: wonDealStageFilters(stage.key),
    dateGranularity: "month",
  };
}

function lineCountByMonthWon(
  x: ColumnMeta,
  y: ColumnMeta,
  stage: ColumnMeta,
): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "line",
    aggregation: "count",
    filters: wonDealStageFilters(stage.key),
    dateGranularity: "month",
  };
}

function areaSumByMonthWon(
  x: ColumnMeta,
  y: ColumnMeta,
  stage: ColumnMeta,
): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "area",
    aggregation: "sum",
    filters: wonDealStageFilters(stage.key),
    dateGranularity: "month",
  };
}

function pieSumWon(x: ColumnMeta, y: ColumnMeta, stage: ColumnMeta): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "pie",
    aggregation: "sum",
    filters: wonDealStageFilters(stage.key),
  };
}

function barSumWonQ1_2026(
  x: ColumnMeta,
  y: ColumnMeta,
  stage: ColumnMeta,
  dateCol: ColumnMeta,
): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "bar",
    aggregation: "sum",
    filters: wonDealQ1_2026Filters(stage, dateCol),
  };
}

function barCountWonQ1_2026(
  x: ColumnMeta,
  y: ColumnMeta,
  stage: ColumnMeta,
  dateCol: ColumnMeta,
): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "bar",
    aggregation: "count",
    filters: wonDealQ1_2026Filters(stage, dateCol),
  };
}

function lineSumByMonthWonQ1_2026(
  x: ColumnMeta,
  y: ColumnMeta,
  stage: ColumnMeta,
  dateCol: ColumnMeta,
): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "line",
    aggregation: "sum",
    filters: wonDealQ1_2026Filters(stage, dateCol),
    dateGranularity: "month",
  };
}

function lineCountByMonthWonQ1_2026(
  x: ColumnMeta,
  y: ColumnMeta,
  stage: ColumnMeta,
  dateCol: ColumnMeta,
): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "line",
    aggregation: "count",
    filters: wonDealQ1_2026Filters(stage, dateCol),
    dateGranularity: "month",
  };
}

function areaSumByMonthWonQ1_2026(
  x: ColumnMeta,
  y: ColumnMeta,
  stage: ColumnMeta,
  dateCol: ColumnMeta,
): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "area",
    aggregation: "sum",
    filters: wonDealQ1_2026Filters(stage, dateCol),
    dateGranularity: "month",
  };
}

function pieSumWonQ1_2026(
  x: ColumnMeta,
  y: ColumnMeta,
  stage: ColumnMeta,
  dateCol: ColumnMeta,
): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "pie",
    aggregation: "sum",
    filters: wonDealQ1_2026Filters(stage, dateCol),
  };
}

// —— Коммерческие предложения (КП) ——

/** Типичные статусы «КП отправлено» в выгрузках Битрикс24 и похожих систем. */
const QUOTE_SENT_STATUS_VALUES: string[] = [
  "Отправлено",
  "отправлено",
  "Отправлено клиенту",
  "отправлено клиенту",
  "Отправлен клиенту",
  "отправлен клиенту",
  "Отослано",
  "отослано",
  "Sent",
  "sent",
];

function quoteDateColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return (
    pickColumn(columns, {
      types: ["date"],
      preferHeaders: [
        "дата создания",
        "дата создания коммерческого предложения",
        "дата отправки",
        "дата отправки коммерческого предложения",
        "дата отправления",
        "дата изменения",
      ],
      anyHeader: ["дата", "создан", "отправ"],
    }) ??
    pickColumn(columns, {
      types: ["string"],
      anyHeader: [
        "дата создания",
        "дата отправки",
        "дата изменения",
      ],
    }) ??
    dateLikeColumn(columns)
  );
}

function quoteAmountColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return (
    pickColumn(columns, {
      types: ["number"],
      preferHeaders: [
        "сумма",
        "стоимость",
        "сумма коммерческого предложения",
        "итого",
        "amount",
      ],
      anyHeader: ["сумм", "стоим", "руб", "цен"],
    }) ?? firstNumberColumn(columns)
  );
}

function quoteStatusColumn(columns: ColumnMeta[]): ColumnMeta | null {
  return pickColumn(columns, {
    types: ["string", "unknown"],
    preferHeaders: [
      "статус",
      "статус коммерческого предложения",
      "стадия",
    ],
    anyHeader: ["статус", "стадия", "состоян"],
  });
}

/** Отправленные КП с ненулевой суммой (минимум 0,01 — отсекаем пустые суммы). */
function quoteSentWithAmountFilters(
  status: ColumnMeta,
  amount: ColumnMeta,
): ChartFilter[] {
  return [
    { columnKey: status.key, values: [...QUOTE_SENT_STATUS_VALUES] },
    { columnKey: amount.key, min: 0.01 },
  ];
}

function lineCountByMonthQuotes(
  x: ColumnMeta,
  y: ColumnMeta,
  status: ColumnMeta,
  amount: ColumnMeta,
): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "line",
    aggregation: "count",
    filters: quoteSentWithAmountFilters(status, amount),
    dateGranularity: "month",
  };
}

function lineSumByMonthQuotes(
  x: ColumnMeta,
  y: ColumnMeta,
  status: ColumnMeta,
  amount: ColumnMeta,
): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "line",
    aggregation: "sum",
    filters: quoteSentWithAmountFilters(status, amount),
    dateGranularity: "month",
  };
}

function areaCountByMonthQuotes(
  x: ColumnMeta,
  y: ColumnMeta,
  status: ColumnMeta,
  amount: ColumnMeta,
): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "area",
    aggregation: "count",
    filters: quoteSentWithAmountFilters(status, amount),
    dateGranularity: "month",
  };
}

function areaSumByMonthQuotes(
  x: ColumnMeta,
  y: ColumnMeta,
  status: ColumnMeta,
  amount: ColumnMeta,
): ChartConfig {
  return {
    xKey: x.key,
    yKeys: [y.key],
    chartType: "area",
    aggregation: "sum",
    filters: quoteSentWithAmountFilters(status, amount),
    dateGranularity: "month",
  };
}

export const PREDEFINED_CHARTS: PredefinedChartSpec[] = [
  // —— Лиды ——
  {
    id: "leads_sum_by_stage",
    entity: "leads",
    title: "Сумма по стадиям лида",
    description: "Сумма (ожидаемая / опорт.) по стадии или статусу лида.",
    resolve(columns) {
      const x = stageColumnLeads(columns);
      const y = leadValueColumn(columns);
      if (!x) {
        return { ok: false, error: "Нет колонки стадии / статуса лида." };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки (сумма лида)." };
      }
      return { ok: true, config: barSum(x, y) };
    },
  },
  {
    id: "leads_count_by_stage",
    entity: "leads",
    title: "Количество лидов по стадиям",
    description: "Сколько лидов в каждой стадии.",
    resolve(columns) {
      const x = stageColumnLeads(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки стадии / статуса лида." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "leads_sum_by_date",
    entity: "leads",
    title: "Сумма лидов по датам",
    description: "Динамика суммы по дате создания / изменения.",
    resolve(columns) {
      const x = dateLikeColumn(columns);
      const y = leadValueColumn(columns);
      if (!x) {
        return { ok: false, error: "Нет колонки с датой." };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки для суммы." };
      }
      return { ok: true, config: lineSum(x, y) };
    },
  },
  {
    id: "leads_area_sum_by_date",
    entity: "leads",
    title: "Динамика суммы (область)",
    description:
      "Та же динамика суммы по датам, что и линия, но с мягкой заливкой под кривой — удобно для презентаций.",
    resolve(columns) {
      const x = dateLikeColumn(columns);
      const y = leadValueColumn(columns);
      if (!x) {
        return { ok: false, error: "Нет колонки с датой." };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки для суммы." };
      }
      return { ok: true, config: areaSum(x, y) };
    },
  },
  {
    id: "leads_line_avg_value_by_date",
    entity: "leads",
    title: "Средняя сумма лида по дням",
    description:
      "Тренд: средняя ожидаемая сумма в каждую дату — видно, растет ли «качество» потока.",
    resolve(columns) {
      const x = dateLikeColumn(columns);
      const y = leadValueColumn(columns);
      if (!x) {
        return { ok: false, error: "Нет колонки с датой." };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки суммы лида." };
      }
      return { ok: true, config: lineAvg(x, y) };
    },
  },
  {
    id: "leads_pie_by_stage",
    entity: "leads",
    title: "Доля суммы по стадиям",
    description: "Круговая диаграмма суммы по стадиям лида.",
    resolve(columns) {
      const x = stageColumnLeads(columns);
      const y = leadValueColumn(columns);
      if (!x) {
        return { ok: false, error: "Нет колонки стадии / статуса лида." };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки для суммы." };
      }
      return { ok: true, config: pieSum(x, y) };
    },
  },
  {
    id: "leads_sum_by_responsible",
    entity: "leads",
    title: "Сумма по ответственным",
    description: "Сумма по сотруднику (ответственный за лид).",
    resolve(columns) {
      const x = responsibleColumn(columns);
      const y = leadValueColumn(columns);
      if (!x) {
        return { ok: false, error: "Нет колон «Ответственный»." };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки для суммы." };
      }
      return { ok: true, config: barSum(x, y) };
    },
  },
  {
    id: "leads_count_by_source",
    entity: "leads",
    title: "Количество по источнику",
    description: "Сколько лидов с каждого источника.",
    resolve(columns) {
      const x = sourceColumn(columns) ?? fallbackStringColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки источника или текстового поля." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "leads_count_by_created_date",
    entity: "leads",
    title: "Приток лидов по датам",
    description:
      "Сколько лидов создано в каждую дату (по колонке даты в выгрузке Битрикс24).",
    resolve(columns) {
      const x = dateLikeColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки с датой создания / изменения." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: lineCount(x, y) };
    },
  },
  {
    id: "leads_area_inflow_by_date",
    entity: "leads",
    title: "Приток лидов по датам (область)",
    description:
      "Количество новых лидов по дням с заливкой — наглядный «объём» входящего потока.",
    resolve(columns) {
      const x = dateLikeColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки с датой создания / изменения." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: areaCount(x, y) };
    },
  },
  {
    id: "leads_pie_by_source",
    entity: "leads",
    title: "Доли лидов по источнику",
    description: "Круговая диаграмма: сколько лидов с каждого источника.",
    resolve(columns) {
      const x = sourceColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки «Источник»." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: pieCount(x, y) };
    },
  },
  {
    id: "leads_avg_value_by_stage",
    entity: "leads",
    title: "Средняя сумма лида по стадиям",
    description:
      "Средняя ожидаемая сумма (опортьюнити) в разрезе стадии — полезно для оценки качества воронки.",
    resolve(columns) {
      const x = stageColumnLeads(columns);
      const y = leadValueColumn(columns);
      if (!x) {
        return { ok: false, error: "Нет колонки стадии / статуса лида." };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки суммы лида." };
      }
      return { ok: true, config: barAvg(x, y) };
    },
  },
  {
    id: "leads_count_by_company",
    entity: "leads",
    title: "Лиды по компаниям",
    description:
      "Сколько лидов приходится на каждую компанию (поле компании в выгрузке лида).",
    resolve(columns) {
      const x = leadCompanyColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return {
          ok: false,
          error: "Нет колонки компании / названия компании лида.",
        };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "leads_count_by_reject_reason",
    entity: "leads",
    title: "Лиды по причине отказа",
    description:
      "Распределение закрытых / отклонённых лидов по причине (типичное поле Битрикс24).",
    resolve(columns) {
      const x = rejectReasonColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return {
          ok: false,
          error: "Нет колонки причины отказа или закрытия лида.",
        };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "leads_sum_by_reject_reason",
    entity: "leads",
    title: "Потери по сумме (причина отказа)",
    description:
      "Совокупная сумма лидов в разрезе причины закрытия — где «стоят» основные потери по воронке.",
    resolve(columns) {
      const x = rejectReasonColumn(columns);
      const y = leadValueColumn(columns);
      if (!x) {
        return {
          ok: false,
          error: "Нет колонки причины отказа или закрытия лида.",
        };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки суммы лида." };
      }
      return { ok: true, config: barSum(x, y) };
    },
  },
  {
    id: "leads_pie_by_funnel",
    entity: "leads",
    title: "Доли лидов по воронке",
    description:
      "Круговая диаграмма: распределение лидов между воронками или направлениями.",
    resolve(columns) {
      const x = funnelOrDirectionColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return {
          ok: false,
          error: "Нет колонки воронки, направления или категории лида.",
        };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: pieCount(x, y) };
    },
  },
  {
    id: "leads_count_by_utm",
    entity: "leads",
    title: "Лиды по UTM / детализации источника",
    description:
      "Количество лидов по UTM или детализированному источнику (маркировка рекламы).",
    resolve(columns) {
      const x = utmOrDetailedSourceColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return {
          ok: false,
          error: "Нет колонок UTM или детализированного источника.",
        };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "leads_sum_by_source",
    entity: "leads",
    title: "Сумма лидов по источнику",
    description:
      "Совокупная ожидаемая сумма лидов с каждого источника привлечения.",
    resolve(columns) {
      const x = sourceColumn(columns);
      const y = leadValueColumn(columns);
      if (!x) {
        return { ok: false, error: "Нет колонки «Источник»." };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки суммы лида." };
      }
      return { ok: true, config: barSum(x, y) };
    },
  },
  {
    id: "leads_count_by_funnel",
    entity: "leads",
    title: "Лиды по воронке / направлению",
    description:
      "Сколько лидов в каждой воронке или направлении (если поле есть в экспорте).",
    resolve(columns) {
      const x = funnelOrDirectionColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return {
          ok: false,
          error: "Нет колонки воронки, направления или категории лида.",
        };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },

  // —— Сделки ——
  {
    id: "deals_sum_by_stage",
    entity: "deals",
    title: "Сумма по стадиям сделки",
    description: "Сумма сделки по стадии или статусу.",
    resolve(columns) {
      const x = stageColumnDeals(columns);
      const y = amountLikeColumn(columns) ?? firstNumberColumn(columns);
      if (!x) {
        return { ok: false, error: "Нет колонки стадии сделки." };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки (сумма)." };
      }
      return { ok: true, config: barSum(x, y) };
    },
  },
  {
    id: "deals_count_by_stage",
    entity: "deals",
    title: "Количество сделок по стадиям",
    description: "Сколько сделок в каждой стадии.",
    resolve(columns) {
      const x = stageColumnDeals(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки стадии сделки." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "deals_sum_by_date",
    entity: "deals",
    title: "Сумма сделок по датам",
    description: "Динамика суммы по дате создания / закрытия.",
    resolve(columns) {
      const x = dateLikeColumn(columns);
      const y = amountLikeColumn(columns) ?? firstNumberColumn(columns);
      if (!x) {
        return { ok: false, error: "Нет колонки с датой." };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки для суммы." };
      }
      return { ok: true, config: lineSum(x, y) };
    },
  },
  {
    id: "deals_area_sum_by_date",
    entity: "deals",
    title: "Динамика суммы сделок (область)",
    description:
      "Сумма сделок по датам подчёркнута площадью под линией — удобно для отчётов и слайдов.",
    resolve(columns) {
      const x = dateLikeColumn(columns);
      const y = amountLikeColumn(columns) ?? firstNumberColumn(columns);
      if (!x) {
        return { ok: false, error: "Нет колонки с датой." };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки для суммы." };
      }
      return { ok: true, config: areaSum(x, y) };
    },
  },
  {
    id: "deals_pie_by_stage",
    entity: "deals",
    title: "Доля суммы по стадиям",
    description: "Круговая диаграмма суммы по стадиям сделки.",
    resolve(columns) {
      const x = stageColumnDeals(columns);
      const y = amountLikeColumn(columns) ?? firstNumberColumn(columns);
      if (!x) {
        return { ok: false, error: "Нет колонки стадии сделки." };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки для суммы." };
      }
      return { ok: true, config: pieSum(x, y) };
    },
  },
  {
    id: "deals_won_sum_by_responsible",
    entity: "deals",
    title: "Выигранные сделки: сумма по менеджерам",
    description:
      "Только успешные сделки (фильтр по типичным названиям стадии «Успешно реализована» и др.). Сумма по ответственным.",
    resolve(columns) {
      const stage = stageColumnDeals(columns);
      const x = responsibleColumn(columns);
      const y = amountLikeColumn(columns) ?? firstNumberColumn(columns);
      if (!stage) {
        return { ok: false, error: "Нет колонки стадии / статуса сделки." };
      }
      if (!x) {
        return { ok: false, error: "Нет колонки «Ответственный»." };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки суммы сделки." };
      }
      return { ok: true, config: barSumWon(x, y, stage) };
    },
  },
  {
    id: "deals_won_count_by_responsible",
    entity: "deals",
    title: "Выигранные сделки: количество по менеджерам",
    description:
      "Число выигранных сделок по каждому ответственному (по тем же названиям успешной стадии).",
    resolve(columns) {
      const stage = stageColumnDeals(columns);
      const x = responsibleColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!stage) {
        return { ok: false, error: "Нет колонки стадии / статуса сделки." };
      }
      if (!x) {
        return { ok: false, error: "Нет колонки «Ответственный»." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCountWon(x, y, stage) };
    },
  },
  {
    id: "deals_won_sum_by_month",
    entity: "deals",
    title: "Выигранные сделки: сумма по месяцам",
    description:
      "Динамика суммы выигранных сделок по календарным месяцам (дата в выгрузке — закрытия, изменения или создания).",
    resolve(columns) {
      const stage = stageColumnDeals(columns);
      const x = dealClosedDateColumn(columns);
      const y = amountLikeColumn(columns) ?? firstNumberColumn(columns);
      if (!stage) {
        return { ok: false, error: "Нет колонки стадии / статуса сделки." };
      }
      if (!x) {
        return { ok: false, error: "Нет колонки с датой." };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки суммы." };
      }
      return { ok: true, config: lineSumByMonthWon(x, y, stage) };
    },
  },
  {
    id: "deals_won_count_by_month",
    entity: "deals",
    title: "Выигранные сделки: количество по месяцам",
    description:
      "Сколько выигранных сделок закрыто в каждом месяце — тренд «скорости побед».",
    resolve(columns) {
      const stage = stageColumnDeals(columns);
      const x = dealClosedDateColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!stage) {
        return { ok: false, error: "Нет колонки стадии / статуса сделки." };
      }
      if (!x) {
        return { ok: false, error: "Нет колонки с датой." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: lineCountByMonthWon(x, y, stage) };
    },
  },
  {
    id: "deals_won_area_sum_by_month",
    entity: "deals",
    title: "Выигранные сделки: сумма по месяцам (область)",
    description:
      "То же, что сумма по месяцам для выигранных сделок, с заливкой под кривой.",
    resolve(columns) {
      const stage = stageColumnDeals(columns);
      const x = dealClosedDateColumn(columns);
      const y = amountLikeColumn(columns) ?? firstNumberColumn(columns);
      if (!stage) {
        return { ok: false, error: "Нет колонки стадии / статуса сделки." };
      }
      if (!x) {
        return { ok: false, error: "Нет колонки с датой." };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки суммы." };
      }
      return { ok: true, config: areaSumByMonthWon(x, y, stage) };
    },
  },
  {
    id: "deals_won_pie_by_company",
    entity: "deals",
    title: "Выигранные сделки: доли суммы по компаниям",
    description:
      "Круговая диаграмма: вклад клиентов в сумму выигранных сделок (только успешные стадии).",
    resolve(columns) {
      const stage = stageColumnDeals(columns);
      const x = companyColumn(columns);
      const y = amountLikeColumn(columns) ?? firstNumberColumn(columns);
      if (!stage) {
        return { ok: false, error: "Нет колонки стадии / статуса сделки." };
      }
      if (!x) {
        return { ok: false, error: "Нет колонки компании в сделках." };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки суммы." };
      }
      return { ok: true, config: pieSumWon(x, y, stage) };
    },
  },
  {
    id: "deals_won_q1_2026_sum_by_responsible",
    entity: "deals",
    title: "Выигранные, 1 кв. 2026: сумма по менеджерам",
    description:
      "Как обычные выигранные сделки по менеджерам, но только за период 01.01.2026–31.03.2026 по выбранной колонке даты (закрытие / изменение / создание).",
    resolve(columns) {
      const stage = stageColumnDeals(columns);
      const dateCol = dealClosedDateColumn(columns);
      const x = responsibleColumn(columns);
      const y = amountLikeColumn(columns) ?? firstNumberColumn(columns);
      if (!stage) {
        return { ok: false, error: "Нет колонки стадии / статуса сделки." };
      }
      if (!dateCol) {
        return { ok: false, error: "Нет колонки с датой." };
      }
      if (!x) {
        return { ok: false, error: "Нет колонки «Ответственный»." };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки суммы сделки." };
      }
      return {
        ok: true,
        config: barSumWonQ1_2026(x, y, stage, dateCol),
      };
    },
  },
  {
    id: "deals_won_q1_2026_count_by_responsible",
    entity: "deals",
    title: "Выигранные, 1 кв. 2026: количество по менеджерам",
    description:
      "Число выигранных сделок по ответственным только за 1 кв. 2026 (фильтр по дате + успешная стадия).",
    resolve(columns) {
      const stage = stageColumnDeals(columns);
      const dateCol = dealClosedDateColumn(columns);
      const x = responsibleColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!stage) {
        return { ok: false, error: "Нет колонки стадии / статуса сделки." };
      }
      if (!dateCol) {
        return { ok: false, error: "Нет колонки с датой." };
      }
      if (!x) {
        return { ok: false, error: "Нет колонки «Ответственный»." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return {
        ok: true,
        config: barCountWonQ1_2026(x, y, stage, dateCol),
      };
    },
  },
  {
    id: "deals_won_q1_2026_sum_by_month",
    entity: "deals",
    title: "Выигранные, 1 кв. 2026: сумма по месяцам",
    description:
      "Янв.–мар. 2026: динамика суммы выигранных сделок по месяцам (точки не выйдут за квартал благодаря фильтру даты).",
    resolve(columns) {
      const stage = stageColumnDeals(columns);
      const x = dealClosedDateColumn(columns);
      const y = amountLikeColumn(columns) ?? firstNumberColumn(columns);
      if (!stage) {
        return { ok: false, error: "Нет колонки стадии / статуса сделки." };
      }
      if (!x) {
        return { ok: false, error: "Нет колонки с датой." };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки суммы." };
      }
      return {
        ok: true,
        config: lineSumByMonthWonQ1_2026(x, y, stage, x),
      };
    },
  },
  {
    id: "deals_won_q1_2026_count_by_month",
    entity: "deals",
    title: "Выигранные, 1 кв. 2026: количество по месяцам",
    description:
      "Сколько выигранных сделок в каждом месяце 1 кв. 2026 — обычно три столбика/точки: янв., февр., март.",
    resolve(columns) {
      const stage = stageColumnDeals(columns);
      const x = dealClosedDateColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!stage) {
        return { ok: false, error: "Нет колонки стадии / статуса сделки." };
      }
      if (!x) {
        return { ok: false, error: "Нет колонки с датой." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return {
        ok: true,
        config: lineCountByMonthWonQ1_2026(x, y, stage, x),
      };
    },
  },
  {
    id: "deals_won_q1_2026_area_sum_by_month",
    entity: "deals",
    title: "Выигранные, 1 кв. 2026: сумма по месяцам (область)",
    description:
      "Сумма выигранных по месяцам в 1 кв. 2026 с заливкой под линией.",
    resolve(columns) {
      const stage = stageColumnDeals(columns);
      const x = dealClosedDateColumn(columns);
      const y = amountLikeColumn(columns) ?? firstNumberColumn(columns);
      if (!stage) {
        return { ok: false, error: "Нет колонки стадии / статуса сделки." };
      }
      if (!x) {
        return { ok: false, error: "Нет колонки с датой." };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки суммы." };
      }
      return {
        ok: true,
        config: areaSumByMonthWonQ1_2026(x, y, stage, x),
      };
    },
  },
  {
    id: "deals_won_q1_2026_pie_by_company",
    entity: "deals",
    title: "Выигранные, 1 кв. 2026: доли суммы по компаниям",
    description:
      "Круговая диаграмма по клиентам только для выигранных сделок и только за 01.01.2026–31.03.2026.",
    resolve(columns) {
      const stage = stageColumnDeals(columns);
      const dateCol = dealClosedDateColumn(columns);
      const x = companyColumn(columns);
      const y = amountLikeColumn(columns) ?? firstNumberColumn(columns);
      if (!stage) {
        return { ok: false, error: "Нет колонки стадии / статуса сделки." };
      }
      if (!dateCol) {
        return { ok: false, error: "Нет колонки с датой." };
      }
      if (!x) {
        return { ok: false, error: "Нет колонки компании в сделках." };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки суммы." };
      }
      return {
        ok: true,
        config: pieSumWonQ1_2026(x, y, stage, dateCol),
      };
    },
  },
  {
    id: "deals_sum_by_responsible",
    entity: "deals",
    title: "Сумма по ответственным",
    description: "Сумма сделок по менеджеру.",
    resolve(columns) {
      const x = responsibleColumn(columns);
      const y = amountLikeColumn(columns) ?? firstNumberColumn(columns);
      if (!x) {
        return { ok: false, error: "Нет колонки «Ответственный»." };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки для суммы." };
      }
      return { ok: true, config: barSum(x, y) };
    },
  },
  {
    id: "deals_count_by_company",
    entity: "deals",
    title: "Количество сделок по компаниям",
    description: "Сколько сделок у каждой компании.",
    resolve(columns) {
      const x = companyColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки компании в сделках." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "deals_count_by_responsible",
    entity: "deals",
    title: "Количество сделок по ответственным",
    description:
      "Сколько сделок закреплено за каждым менеджером (типовая выгрузка Битрикс24).",
    resolve(columns) {
      const x = responsibleColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки «Ответственный»." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "deals_sum_by_company",
    entity: "deals",
    title: "Сумма сделок по компаниям",
    description: "Совокупная сумма сделок по каждой компании-контрагенту.",
    resolve(columns) {
      const x = companyColumn(columns);
      const y = amountLikeColumn(columns) ?? firstNumberColumn(columns);
      if (!x) {
        return { ok: false, error: "Нет колонки компании в сделке." };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки суммы сделки." };
      }
      return { ok: true, config: barSum(x, y) };
    },
  },
  {
    id: "deals_count_by_created_date",
    entity: "deals",
    title: "Приток сделок по датам",
    description:
      "Число созданных сделок по дням (по дате из выгрузки: создания или изменения).",
    resolve(columns) {
      const x = dateLikeColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки с датой." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: lineCount(x, y) };
    },
  },
  {
    id: "deals_area_inflow_by_date",
    entity: "deals",
    title: "Приток сделок по датам (область)",
    description:
      "Число созданных сделок по дням с заливкой — насколько ровный или «пульсирующий» поток.",
    resolve(columns) {
      const x = dateLikeColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки с датой." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: areaCount(x, y) };
    },
  },
  {
    id: "deals_avg_amount_by_stage",
    entity: "deals",
    title: "Средний чек по стадиям",
    description:
      "Средняя сумма сделки в разрезе стадии — удобно сравнивать этапы воронки.",
    resolve(columns) {
      const x = stageColumnDeals(columns);
      const y = amountLikeColumn(columns) ?? firstNumberColumn(columns);
      if (!x) {
        return { ok: false, error: "Нет колонки стадии сделки." };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки суммы." };
      }
      return { ok: true, config: barAvg(x, y) };
    },
  },
  {
    id: "deals_pie_by_company",
    entity: "deals",
    title: "Доли суммы сделок по компаниям",
    description:
      "Круговая диаграмма: вклад каждой компании в общую сумму сделок в выборке.",
    resolve(columns) {
      const x = companyColumn(columns);
      const y = amountLikeColumn(columns) ?? firstNumberColumn(columns);
      if (!x) {
        return { ok: false, error: "Нет колонки компании." };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки суммы." };
      }
      return { ok: true, config: pieSum(x, y) };
    },
  },
  {
    id: "deals_count_by_contact",
    entity: "deals",
    title: "Сделки по контактам",
    description:
      "Сколько сделок приходится на каждый связанный контакт (если поле выгружено).",
    resolve(columns) {
      const x = dealContactColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return {
          ok: false,
          error: "Нет колонки контакта / связанного контакта.",
        };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "deals_sum_by_contact",
    entity: "deals",
    title: "Сумма сделок по контактам",
    description:
      "Кто из контактов «тянет» наибольший объём сделок (если контакт и сумма есть в выгрузке).",
    resolve(columns) {
      const x = dealContactColumn(columns);
      const y = amountLikeColumn(columns) ?? firstNumberColumn(columns);
      if (!x) {
        return {
          ok: false,
          error: "Нет колонки контакта / связанного контакта.",
        };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки суммы." };
      }
      return { ok: true, config: barSum(x, y) };
    },
  },
  {
    id: "deals_avg_by_responsible",
    entity: "deals",
    title: "Средний чек по менеджерам",
    description:
      "Средняя сумма сделки на одного ответственного — полезно сравнивать стиль продаж команды.",
    resolve(columns) {
      const x = responsibleColumn(columns);
      const y = amountLikeColumn(columns) ?? firstNumberColumn(columns);
      if (!x) {
        return { ok: false, error: "Нет колонки «Ответственный»." };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки суммы." };
      }
      return { ok: true, config: barAvg(x, y) };
    },
  },
  {
    id: "deals_pie_sum_by_responsible",
    entity: "deals",
    title: "Доли суммы по менеджерам",
    description:
      "Вклад каждого ответственного в общую сумму сделок в выборке — круговая диаграмма.",
    resolve(columns) {
      const x = responsibleColumn(columns);
      const y = amountLikeColumn(columns) ?? firstNumberColumn(columns);
      if (!x) {
        return { ok: false, error: "Нет колонки «Ответственный»." };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки суммы." };
      }
      return { ok: true, config: pieSum(x, y) };
    },
  },
  {
    id: "deals_line_avg_by_date",
    entity: "deals",
    title: "Средний чек по дням",
    description:
      "Как меняется средняя сумма сделки по датам в выгрузке — заметны ли провалы и всплески.",
    resolve(columns) {
      const x = dateLikeColumn(columns);
      const y = amountLikeColumn(columns) ?? firstNumberColumn(columns);
      if (!x) {
        return { ok: false, error: "Нет колонки с датой." };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки суммы." };
      }
      return { ok: true, config: lineAvg(x, y) };
    },
  },
  {
    id: "deals_count_by_pipeline",
    entity: "deals",
    title: "Сделки по воронке / направлению",
    description:
      "Распределение сделок по воронке или направлению из CRM.",
    resolve(columns) {
      const x = dealPipelineColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return {
          ok: false,
          error: "Нет колонки воронки или направления сделки.",
        };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "deals_count_by_business_type",
    entity: "deals",
    title: "Сделки по типу (новая / повторная)",
    description:
      "Сколько сделок каждого типа, если в выгрузке есть соответствующее поле.",
    resolve(columns) {
      const x = dealBusinessTypeColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return {
          ok: false,
          error: "Нет колонки типа сделки или признака повторности.",
        };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },

  // —— Контакты ——
  {
    id: "contacts_count_by_company",
    entity: "contacts",
    title: "Контакты по компаниям",
    description: "Сколько контактов приходится на каждую компанию.",
    resolve(columns) {
      const x = companyColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки компании." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "contacts_count_by_responsible",
    entity: "contacts",
    title: "Контакты по ответственным",
    description: "Распределение контактов по менеджерам.",
    resolve(columns) {
      const x = responsibleColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки «Ответственный»." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "contacts_count_by_source",
    entity: "contacts",
    title: "Контакты по источнику",
    description: "Количество контактов по источнику привлечения.",
    resolve(columns) {
      const x = sourceColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки источника." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "contacts_pie_by_company",
    entity: "contacts",
    title: "Доля контактов по компаниям (топ)",
    description: "Доли контактов по компаниям (по количеству строк).",
    resolve(columns) {
      const x = companyColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки компании." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return {
        ok: true,
        config: {
          xKey: x.key,
          yKeys: [y.key],
          chartType: "pie",
          aggregation: "count",
          filters: [],
        },
      };
    },
  },
  {
    id: "contacts_count_by_created_date",
    entity: "contacts",
    title: "Приток контактов по датам",
    description:
      "Сколько контактов создано в каждую дату (дата создания / изменения в экспорте).",
    resolve(columns) {
      const x = dateLikeColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки с датой." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: lineCount(x, y) };
    },
  },
  {
    id: "contacts_area_count_by_created_date",
    entity: "contacts",
    title: "Приток контактов (область)",
    description:
      "Новые контакты по датам с заливкой — визуальный объём роста базы.",
    resolve(columns) {
      const x = dateLikeColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки с датой." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: areaCount(x, y) };
    },
  },
  {
    id: "contacts_sum_by_responsible",
    entity: "contacts",
    title: "Сумма показателя по ответственным",
    description:
      "Если в выгрузке контактов есть числовое поле (балл, LTV и т.д.) — сумма по менеджерам.",
    resolve(columns) {
      const x = responsibleColumn(columns);
      const y = firstNumberColumn(columns);
      if (!x) {
        return { ok: false, error: "Нет колонки «Ответственный»." };
      }
      if (!y) {
        return {
          ok: false,
          error: "Нет числового столбца для суммирования.",
        };
      }
      return { ok: true, config: barSum(x, y) };
    },
  },
  {
    id: "contacts_avg_by_company",
    entity: "contacts",
    title: "Среднее число по компании",
    description:
      "По каждой компании — среднее по первому числовому полю (например средний «вес» контакта при наличии метрики).",
    resolve(columns) {
      const x = companyColumn(columns);
      const y = firstNumberColumn(columns);
      if (!x) {
        return { ok: false, error: "Нет колонки компании." };
      }
      if (!y) {
        return {
          ok: false,
          error: "Нет числового столбца для усреднения.",
        };
      }
      return { ok: true, config: barAvg(x, y) };
    },
  },
  {
    id: "contacts_count_by_position",
    entity: "contacts",
    title: "Контакты по должностям",
    description:
      "Распределение контактов по полю «Должность» (типичная карточка Битрикс24).",
    resolve(columns) {
      const x = contactPositionColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки должности контакта." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "contacts_pie_by_responsible",
    entity: "contacts",
    title: "Доли контактов по ответственным",
    description: "Круговая диаграмма: распределение базы контактов по менеджерам.",
    resolve(columns) {
      const x = responsibleColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки «Ответственный»." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: pieCount(x, y) };
    },
  },
  {
    id: "contacts_count_by_location",
    entity: "contacts",
    title: "Контакты по городам / регионам",
    description:
      "География контактов: город, регион или страна, если эти поля есть в выгрузке.",
    resolve(columns) {
      const x = locationColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return {
          ok: false,
          error: "Нет колонки города, региона или местоположения.",
        };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "contacts_pie_by_source",
    entity: "contacts",
    title: "Доли контактов по источнику",
    description: "Круговая диаграмма: откуда пришли контакты в разрезе источника.",
    resolve(columns) {
      const x = sourceColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки источника." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: pieCount(x, y) };
    },
  },

  // —— Компании ——
  {
    id: "companies_count_by_industry",
    entity: "companies",
    title: "Компании по отраслям",
    description:
      "Распределение компаний именно по отраслям (строится только при наличии явной колонки отрасли).",
    resolve(columns) {
      const x = pickStrictHeaderColumn(columns, {
        types: ["string", "unknown"],
        exactHeaders: ["отрасль", "индустрия", "industry"],
      });
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет явной колонки отрасли." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "companies_count_by_responsible",
    entity: "companies",
    title: "Распределение по ответственным",
    description:
      "Круговая диаграмма: сколько компаний ведёт каждый менеджер (поле «Ответственный»).",
    resolve(columns) {
      const x = responsibleColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки «Ответственный»." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: pieCount(x, y) };
    },
  },
  {
    id: "companies_registered_by_month_line",
    entity: "companies",
    title: "Динамика появления компаний — по месяцам (линия)",
    description:
      "Сколько компаний создано в каждом календарном месяце по дате создания. Узкий период (например сентябрь–декабрь) задаётся в «Фильтрах данных» по колонке с датой.",
    resolve(columns) {
      const x = companyCreatedDateColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return {
          ok: false,
          error:
            "Нет колонки с датой создания / изменения компании.",
        };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: lineCountByMonth(x, y) };
    },
  },
  {
    id: "companies_cumulative_count_by_month_area",
    entity: "companies",
    title: "Нарастающий итог компаний по месяцам",
    description:
      "Кумулятивная кривая: общее число компаний с учётом накопления от первого к последнему месяцу по дате создания (ось X — месяцы, значение не убывает).",
    resolve(columns) {
      const x = companyCreatedDateColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return {
          ok: false,
          error:
            "Нет колонки с датой создания / изменения компании.",
        };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: areaCumulativeCountByMonth(x, y) };
    },
  },
  {
    id: "companies_sum_by_responsible",
    entity: "companies",
    title: "Сумма бизнес-метрики по менеджерам",
    description:
      "Сумма выручки/оборота/бюджета по ответственным. Строится только при наличии явной числовой бизнес-метрики.",
    resolve(columns) {
      const x = responsibleColumn(columns);
      const y = companyBusinessMetricColumn(columns);
      if (!x) {
        return { ok: false, error: "Нет колонки «Ответственный»." };
      }
      if (!y) {
        return {
          ok: false,
          error:
            "Нет явной числовой бизнес-метрики (выручка/оборот/бюджет/сумма).",
        };
      }
      return { ok: true, config: barSum(x, y) };
    },
  },
  {
    id: "companies_count_by_source",
    entity: "companies",
    title: "Компании по источнику привлечения",
    description:
      "Сколько компаний пришло из каждого канала привлечения (строится только при наличии явной колонки источника компании).",
    resolve(columns) {
      const x = companySourceColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return {
          ok: false,
          error:
            "Нет явной колонки источника компании (например «Источник компании» / «Канал привлечения»).",
        };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "companies_revenue_by_industry",
    entity: "companies",
    title: "Выручка по отраслям",
    description:
      "Сумма выручки/оборота по отраслям (строится только при наличии явных колонок отрасли и выручки).",
    resolve(columns) {
      const x = pickStrictHeaderColumn(columns, {
        types: ["string", "unknown"],
        exactHeaders: ["отрасль", "индустрия", "сфера деятельности", "вид деятельности"],
      });
      const y = pickStrictHeaderColumn(columns, {
        types: ["number"],
        exactHeaders: [
          "выручка",
          "оборот",
          "годовой оборот",
          "доход",
          "revenue",
          "annual revenue",
        ],
      });
      if (!x) {
        return { ok: false, error: "Нет явной колонки отрасли." };
      }
      if (!y) {
        return {
          ok: false,
          error: "Нет явной числовой колонки выручки / оборота.",
        };
      }
      return { ok: true, config: barSum(x, y) };
    },
  },
  {
    id: "companies_count_employee_ids_by_industry",
    entity: "companies",
    title: "Кол-во компаний по отраслям (по заполненному ID)",
    description:
      "По отрасли считает количество компаний с заполненным ID/привязкой сотрудника, без суммирования значений идентификатора.",
    resolve(columns) {
      const x = industryOrTypeColumn(columns) ?? fallbackStringColumn(columns);
      const y = employeeIdLikeColumn(columns);
      if (!x) {
        return { ok: false, error: "Нет колонки отрасли / типа." };
      }
      if (!y) {
        return {
          ok: false,
          error:
            "Нет колонки «Связанный пользователь» / ID сотрудника (не путать с «Кол-во сотрудников»).",
        };
      }
      return { ok: true, config: barCountNonempty(x, y) };
    },
  },
  {
    id: "companies_sum_employees",
    entity: "companies",
    title: "Сумма числа сотрудников по отрасли",
    description:
      "Сумма численности сотрудников по отрасли (строится только при наличии явных колонок отрасли и численности).",
    resolve(columns) {
      const x = pickStrictHeaderColumn(columns, {
        types: ["string", "unknown"],
        exactHeaders: ["отрасль", "индустрия", "сфера деятельности", "вид деятельности"],
      });
      const y = pickStrictHeaderColumn(columns, {
        types: ["number"],
        exactHeaders: [
          "кол-во сотрудников",
          "количество сотрудников",
          "численность",
          "employees",
          "staff",
        ],
      });
      if (!x) {
        return { ok: false, error: "Нет явной колонки отрасли." };
      }
      if (!y) {
        return {
          ok: false,
          error:
            "Нет явной числовой колонки «Кол-во сотрудников» / численности.",
        };
      }
      return { ok: true, config: barSum(x, y) };
    },
  },
  {
    id: "companies_count_by_company_type",
    entity: "companies",
    title: "Типы компаний",
    description:
      "Круговая диаграмма распределения компаний по значениям поля «Тип компании».",
    resolve(columns) {
      const x = companyTypeColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return {
          ok: false,
          error: "Нет колонки «Тип компании» / типа клиента.",
        };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: pieCount(x, y) };
    },
  },
  {
    id: "companies_avg_headcount_by_industry",
    entity: "companies",
    title: "Средняя численность по отраслям",
    description:
      "Среднее значение поля «Кол-во сотрудников» по отрасли (строится только если есть явные колонки отрасли и численности сотрудников).",
    resolve(columns) {
      const x = pickStrictHeaderColumn(columns, {
        types: ["string", "unknown"],
        exactHeaders: ["отрасль", "индустрия", "сфера деятельности", "вид деятельности"],
      });
      const y = pickStrictHeaderColumn(columns, {
        types: ["number"],
        exactHeaders: [
          "кол-во сотрудников",
          "количество сотрудников",
          "численность",
          "employees",
          "staff",
        ],
      });
      if (!x) {
        return { ok: false, error: "Нет явной колонки отрасли." };
      }
      if (!y) {
        return {
          ok: false,
          error:
            "Нет явной числовой колонки «Кол-во сотрудников» / численности.",
        };
      }
      return { ok: true, config: barAvg(x, y) };
    },
  },
  {
    id: "companies_count_by_location",
    entity: "companies",
    title: "Компании по локации / региону",
    description:
      "Сколько компаний по колонке локации (только при наличии явного поля: «Локация»/«Город»/«Регион»/«Страна»).",
    resolve(columns) {
      const x = locationColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return {
          ok: false,
          error: "Нет явной колонки локации (город/регион/страна).",
        };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "companies_count_by_activity_sphere",
    entity: "companies",
    title: "Компании по сфере деятельности",
    description:
      "Распределение по колонке «Сфера деятельности» (если поле явно есть в выгрузке).",
    resolve(columns) {
      const x = sphereOfActivityColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return {
          ok: false,
          error: "Нет колонки «Сфера деятельности».",
        };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "companies_count_by_utm_source",
    entity: "companies",
    title: "Компании по UTM Source",
    description:
      "Сколько компаний по UTM Source (строится только при наличии колонки `utm_source`).",
    resolve(columns) {
      const x = utmSourceColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки UTM Source (`utm_source`)." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "companies_count_by_utm_medium",
    entity: "companies",
    title: "Компании по UTM Medium",
    description:
      "Распределение по UTM Medium (строится только при наличии колонки `utm_medium`).",
    resolve(columns) {
      const x = utmMediumColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки UTM Medium (`utm_medium`)." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "companies_count_by_utm_campaign",
    entity: "companies",
    title: "Компании по UTM Campaign",
    description:
      "Распределение по UTM Campaign (строится только при наличии колонки `utm_campaign`).",
    resolve(columns) {
      const x = utmCampaignColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки UTM Campaign (`utm_campaign`)." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "companies_count_by_messenger",
    entity: "companies",
    title: "Компании по мессенджеру",
    description:
      "Распределение по полю «Мессенджер» (строится только при наличии этой колонки).",
    resolve(columns) {
      const x = messengerColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки «Мессенджер»." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "companies_count_by_brand",
    entity: "companies",
    title: "Компании по бренду",
    description:
      "Сколько компаний по полю «Бренд» (если такое поле явно присутствует).",
    resolve(columns) {
      const x = brandColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки «Бренд»." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "companies_count_by_client_journey",
    entity: "companies",
    title: "Компании по пути клиента",
    description:
      "Распределение по полю «Путь клиента» (если оно есть в файле).",
    resolve(columns) {
      const x = clientJourneyColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки «Путь клиента»." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "companies_count_by_status_export",
    entity: "companies",
    title: "Компании по статусу",
    description:
      "Распределение по колонке «Статус» (если поле явно присутствует в выгрузке).",
    resolve(columns) {
      const x = companyExportStatusColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки «Статус»." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "companies_count_by_lifecycle",
    entity: "companies",
    title: "Компании по жизненному циклу",
    description:
      "Распределение по полю «Жизненный цикл» (если оно явно есть в файле).",
    resolve(columns) {
      const x = companyLifecycleColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return {
          ok: false,
          error: "Нет колонки «Жизненный цикл».",
        };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "companies_percent_with_deal",
    entity: "companies",
    title: "Процент компаний, у которых есть сделка",
    description:
      "Круговая диаграмма долей: какая часть компаний имеет сделку и какая — нет (по полю «Есть сделка?»).",
    resolve(columns) {
      const x = companyHasDealColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки «Есть сделка?»." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: pieCount(x, y) };
    },
  },
  {
    id: "companies_count_by_created_by",
    entity: "companies",
    title: "Компании по автору создания",
    description:
      "Сколько компаний создал каждый пользователь (строится только при наличии явной колонки «Кем создана»).",
    resolve(columns) {
      const x = companyCreatedByColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки «Кем создана»." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "companies_count_by_need",
    entity: "companies",
    title: "Компании по потребности",
    description:
      "Распределение по полю «Потребность» (если колонка явно присутствует).",
    resolve(columns) {
      const x = companyNeedColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return { ok: false, error: "Нет колонки «Потребность»." };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "companies_count_by_parent_company",
    entity: "companies",
    title: "Компании по родительской компании",
    description:
      "Сколько карточек привязано к каждой родительской компании (если поле есть).",
    resolve(columns) {
      const x = parentCompanyColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return {
          ok: false,
          error: "Нет колонки «Родительская компания».",
        };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: barCount(x, y) };
    },
  },
  {
    id: "companies_sum_budget_kp_by_responsible",
    entity: "companies",
    title: "Бюджет КП по ответственным",
    description:
      "Сумма поля «БЮДЖЕТ КП» по ответственным (только при наличии явной колонки бюджета).",
    resolve(columns) {
      const x = responsibleColumn(columns);
      const y = companyBudgetKpColumn(columns);
      if (!x) {
        return { ok: false, error: "Нет колонки «Ответственный»." };
      }
      if (!y) {
        return { ok: false, error: "Нет числовой колонки «БЮДЖЕТ КП»." };
      }
      return { ok: true, config: barSum(x, y) };
    },
  },
  {
    id: "companies_activity_last_by_month_line",
    entity: "companies",
    title: "Активность по месяцам (последняя)",
    description:
      "Сколько компаний попало в каждый месяц по дате «Последняя активность» или «Дата последней коммуникации».",
    resolve(columns) {
      const x = companyLastActivityColumn(columns);
      const y = firstNumberColumn(columns) ?? columns[0] ?? null;
      if (!x) {
        return {
          ok: false,
          error:
            "Нет колонки с датой последней активности / коммуникации.",
        };
      }
      if (!y) {
        return { ok: false, error: "Нет колонок в данных." };
      }
      return { ok: true, config: lineCountByMonth(x, y) };
    },
  },

  // —— Коммерческие предложения (КП) ——
  {
    id: "quotes_sent_with_amount_count_by_month",
    entity: "quotes",
    title: "КП отправленные с суммой: по месяцам (кол-во)",
    description:
      "Сколько коммерческих предложений появилось в каждом месяце: только статусы «Отправлено» и сходные + сумма в строке больше нуля. Ось времени — дата из выгрузки (создание, отправка или изменение).",
    resolve(columns) {
      const x = quoteDateColumn(columns);
      const y = quoteAmountColumn(columns);
      const status = quoteStatusColumn(columns);
      if (!x) {
        return { ok: false, error: "Нет колонки с датой." };
      }
      if (!y) {
        return {
          ok: false,
          error: "Нет числовой колонки суммы / стоимости КП.",
        };
      }
      if (!status) {
        return { ok: false, error: "Нет колонки статуса КП." };
      }
      return {
        ok: true,
        config: lineCountByMonthQuotes(x, y, status, y),
      };
    },
  },
  {
    id: "quotes_sent_with_amount_sum_by_month",
    entity: "quotes",
    title: "КП отправленные с суммой: по месяцам (сумма)",
    description:
      "Совокупная стоимость отправленных КП с заполненной суммой по календарным месяцам.",
    resolve(columns) {
      const x = quoteDateColumn(columns);
      const y = quoteAmountColumn(columns);
      const status = quoteStatusColumn(columns);
      if (!x) {
        return { ok: false, error: "Нет колонки с датой." };
      }
      if (!y) {
        return {
          ok: false,
          error: "Нет числовой колонки суммы / стоимости КП.",
        };
      }
      if (!status) {
        return { ok: false, error: "Нет колонки статуса КП." };
      }
      return {
        ok: true,
        config: lineSumByMonthQuotes(x, y, status, y),
      };
    },
  },
  {
    id: "quotes_sent_with_amount_area_count_by_month",
    entity: "quotes",
    title: "КП отправленные с суммой: по месяцам (кол-во, область)",
    description:
      "То же, что количество по месяцам, с заливкой — наглядный поток отправленных КП.",
    resolve(columns) {
      const x = quoteDateColumn(columns);
      const y = quoteAmountColumn(columns);
      const status = quoteStatusColumn(columns);
      if (!x) {
        return { ok: false, error: "Нет колонки с датой." };
      }
      if (!y) {
        return {
          ok: false,
          error: "Нет числовой колонки суммы / стоимости КП.",
        };
      }
      if (!status) {
        return { ok: false, error: "Нет колонки статуса КП." };
      }
      return {
        ok: true,
        config: areaCountByMonthQuotes(x, y, status, y),
      };
    },
  },
  {
    id: "quotes_sent_with_amount_area_sum_by_month",
    entity: "quotes",
    title: "КП отправленные с суммой: по месяцам (сумма, область)",
    description:
      "Динамика суммы отправленных КП по месяцам с мягкой заливкой под кривой.",
    resolve(columns) {
      const x = quoteDateColumn(columns);
      const y = quoteAmountColumn(columns);
      const status = quoteStatusColumn(columns);
      if (!x) {
        return { ok: false, error: "Нет колонки с датой." };
      }
      if (!y) {
        return {
          ok: false,
          error: "Нет числовой колонки суммы / стоимости КП.",
        };
      }
      if (!status) {
        return { ok: false, error: "Нет колонки статуса КП." };
      }
      return {
        ok: true,
        config: areaSumByMonthQuotes(x, y, status, y),
      };
    },
  },
];

export function getSpecById(id: string): PredefinedChartSpec | undefined {
  return PREDEFINED_CHARTS.find((c) => c.id === id);
}

export function chartsForEntity(entity: EntityBlockId): PredefinedChartSpec[] {
  return PREDEFINED_CHARTS.filter((c) => c.entity === entity);
}

const ENTITY_SUGGEST_PRIMARY: EntityBlockId[] = [
  "leads",
  "deals",
  "contacts",
  "quotes",
];

const FILE_ENTITY_MARKERS: {
  entity: EntityBlockId;
  markers: string[];
}[] = [
  { entity: "companies", markers: ["COMPANY", "COMPANIES"] },
  { entity: "deals", markers: ["DEAL", "DEALS"] },
  { entity: "leads", markers: ["LEAD", "LEADS"] },
  { entity: "contacts", markers: ["CONTACT", "CONTACTS"] },
  { entity: "quotes", markers: ["QUOTE", "QUOTES", "KP", "COMMERCIAL"] },
];

/** Смещение при выборе между лидами и сделками по подписи колонки стадии. */
const STAGE_HEADER_BIAS = 800;

function feasibleChartIdsForEntity(
  entity: EntityBlockId,
  columns: ColumnMeta[],
): string[] {
  const chartIds: string[] = [];
  for (const spec of chartsForEntity(entity)) {
    if (spec.resolve(columns).ok) {
      chartIds.push(spec.id);
    }
  }
  return chartIds;
}

function primarySuggestScore(
  entity: EntityBlockId,
  columns: ColumnMeta[],
  chartIds: string[],
): number {
  let score = chartIds.length;
  if (entity === "leads") {
    const st = stageColumnLeads(columns);
    const h = st ? normHeader(st.header) : "";
    if (h.includes("лид") && !h.includes("сделк")) {
      score += STAGE_HEADER_BIAS;
    }
  }
  if (entity === "deals") {
    const st = stageColumnDeals(columns);
    const h = st ? normHeader(st.header) : "";
    if (h.includes("сделк")) {
      score += STAGE_HEADER_BIAS;
    }
  }
  return score;
}

/** Определяет сущность CRM по имени файла, если в нём есть явный маркер. */
export function suggestEntityFromFileName(fileName: string): EntityBlockId | null {
  const upper = fileName.toUpperCase();
  for (const { entity, markers } of FILE_ENTITY_MARKERS) {
    if (markers.some((m) => upper.includes(m))) {
      return entity;
    }
  }
  return null;
}

/**
 * Подбор раздела CRM и графиков по фактическим колонкам файла.
 * Сначала сравниваются лиды, сделки, контакты и КП (типовые выгрузки); блок
 * «Компании» с множеством пресетов на общих колонках учитывается только если
 * в первых четырёх разделах нет ни одного построимого графика.
 * Между лидами и сделками учитывается подпись колонки стадии («лид» / «сделк»).
 * При равенстве итогового счёта — первый раздел в порядке ENTITY_SUGGEST_PRIMARY.
 */
export function suggestEntityAndCharts(columns: ColumnMeta[]): {
  entity: EntityBlockId;
  chartIds: string[];
} | null {
  const primary = ENTITY_SUGGEST_PRIMARY.map((entity) => {
    const chartIds = feasibleChartIdsForEntity(entity, columns);
    return {
      entity,
      chartIds,
      score: primarySuggestScore(entity, columns, chartIds),
    };
  });
  const maxScore = Math.max(0, ...primary.map((p) => p.score));
  if (maxScore > 0) {
    return primary.find((p) => p.score === maxScore)!;
  }
  const companyIds = feasibleChartIdsForEntity("companies", columns);
  if (companyIds.length > 0) {
    return { entity: "companies", chartIds: companyIds };
  }
  return null;
}

export function normalizeConfigForTabular(
  data: TabularData,
  cfg: ChartConfig,
): ChartConfig {
  const keys = new Set(data.columns.map((c) => c.key));
  const xKey = cfg.xKey && keys.has(cfg.xKey) ? cfg.xKey : null;
  const yKeys = cfg.yKeys.filter((k) => keys.has(k));
  const filters = cfg.filters
    .filter((f) => keys.has(f.columnKey))
    .map((f) => ({ ...f }));
  return { ...cfg, xKey, yKeys, filters };
}
