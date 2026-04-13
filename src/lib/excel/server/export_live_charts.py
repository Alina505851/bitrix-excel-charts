from __future__ import annotations

import json
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

import xlsxwriter

# Как в эталоне Excel: акцентный синий и заголовки таблиц на листе графика.
EXCEL_ACCENT_BLUE = "#0070C0"

# Единый вид столбчатых диаграмм: Excel по умолчанию gap=150 (узкие столбцы при многих точках).
COLUMN_BAR_GAP = 95

# Заголовок диаграммы как у основного графика без time_split (Calibri 13, жирный, синий).
COLUMN_CHART_TITLE_FONT: dict[str, object] = {
    "name": "Calibri",
    "size": 13,
    "bold": True,
    "color": EXCEL_ACCENT_BLUE,
}

# Подписи категорий на оси X: одинаковый наклон для месяцев/кварталов/лет.
COLUMN_CATEGORY_AXIS_NUM_FONT: dict[str, object] = {
    "size": 9,
    "color": EXCEL_ACCENT_BLUE,
    "rotation": -45,
}

# Как src/lib/chart/chartPalettes.ts — COMPANY_CHART_COLORS; «Сделки» и «Компании» делят одну палитру столбцов/линий Excel.
COMPANY_CHART_COLORS: list[str] = [
    EXCEL_ACCENT_BLUE,
    "#06b6d4",
    "#16a34a",
    "#f59e0b",
    "#ef4444",
    "#0891b2",
    "#db2777",
    "#0d9488",
    "#22c55e",
    "#0ea5e9",
    "#e11d48",
    "#eab308",
]

# Синхронно с ENTITY_CHART_PALETTES в chartPalettes.ts (лиды/контакты/прочие — свои ряды; сделки = компании).
LEADS_CHART_COLORS: list[str] = [
    "#22d3ee",
    "#06b6d4",
    "#0ea5e9",
    "#38bdf8",
    "#3b82f6",
    "#2563eb",
    "#60a5fa",
    "#6366f1",
    "#2dd4bf",
    "#0284c7",
    "#818cf8",
    "#7dd3fc",
]

CONTACTS_CHART_COLORS: list[str] = [
    "#10b981",
    "#22c55e",
    "#34d399",
    "#4ade80",
    "#14b8a6",
    "#2dd4bf",
    "#059669",
    "#65a30d",
    "#84cc16",
    "#0d9488",
    "#16a34a",
    "#047857",
]

QUOTES_CHART_COLORS: list[str] = [
    "#fb7185",
    "#f43f5e",
    "#ec4899",
    "#e11d48",
    "#f472b6",
    "#be123c",
    "#fda4af",
    "#db2777",
    "#f87171",
    "#ef4444",
    "#9f1239",
    "#d946ef",
]

GENERIC_CHART_COLORS: list[str] = [
    "#64748b",
    "#475569",
    "#0ea5e9",
    "#06b6d4",
    "#14b8a6",
    "#8b5cf6",
    "#a855f7",
    "#d946ef",
    "#ec4899",
    "#f43f5e",
    "#f59e0b",
    "#84cc16",
]

ENTITY_PALETTES: dict[str, list[str]] = {
    "leads": LEADS_CHART_COLORS,
    "deals": COMPANY_CHART_COLORS,
    "contacts": CONTACTS_CHART_COLORS,
    "companies": COMPANY_CHART_COLORS,
    "quotes": QUOTES_CHART_COLORS,
    "generic": GENERIC_CHART_COLORS,
}

# Фон/сетка/оси: единый «excel-синий» для всех сущностей (как макет выгрузки).
THEME_STYLE: dict[str, dict[str, object]] = {
    "classic": {
        "chart_style": 2,
        "pie_style": 10,
        "plot_bg": "#FFFFFF",
        "grid": "#D9E2F3",
        "axis": EXCEL_ACCENT_BLUE,
        "line_width": 2.0,
        "marker": 6,
    },
    "vivid": {
        "chart_style": 18,
        "pie_style": 26,
        "plot_bg": "#F2F8FC",
        "grid": "#B4C6E7",
        "axis": EXCEL_ACCENT_BLUE,
        "line_width": 2.6,
        "marker": 7,
    },
    "minimal": {
        "chart_style": 1,
        "pie_style": 2,
        "plot_bg": "#FFFFFF",
        "grid": "#E7EEF7",
        "axis": EXCEL_ACCENT_BLUE,
        "line_width": 1.6,
        "marker": 4,
    },
}

RU_MONTHS_MAP = {
    "янв.": 1,
    "февр.": 2,
    "мар.": 3,
    "апр.": 4,
    "мая": 5,
    "июн.": 6,
    "июл.": 7,
    "авг.": 8,
    "сент.": 9,
    "окт.": 10,
    "нояб.": 11,
    "дек.": 12,
}


def sanitize_sheet_name(raw: str, used: set[str]) -> str:
    name = (
        raw.replace(":", " ")
        .replace("\\", " ")
        .replace("/", " ")
        .replace("?", " ")
        .replace("*", " ")
        .replace("[", " ")
        .replace("]", " ")
        .strip()
    )
    if not name:
        name = "График"
    name = name[:31]
    base = name
    idx = 2
    while name in used:
        suffix = f" {idx}"
        name = (base[: 31 - len(suffix)] + suffix)[:31]
        idx += 1
    used.add(name)
    return name


def chart_type_for_xlsxwriter(chart_type: str) -> str:
    if chart_type in ("bar", "line", "area", "pie"):
        return chart_type
    return "column"


def time_split_chart_title(raw: object, block: str) -> str:
    """
    Заголовок блока М/К/Г: убираем «(линия)», не дублируем «— по месяцам», если уже есть в пресете.
    block: month | quarter | year
    """
    t = re.sub(r"\s*\(линия\)\s*$", "", str(raw or "График").strip())
    suffixes = {"month": "по месяцам", "quarter": "по кварталам", "year": "по годам"}
    needle = suffixes[block]
    if re.search(rf"{re.escape(needle)}\s*$", t, re.IGNORECASE):
        return t
    return f"{t} — {needle}"

def palette_for_entity(entity_id: object) -> list[str]:
    """Палитра серий как в веб-UI: сделки и компании — Accent Excel (#0070C0 и ряд)."""
    e = str(entity_id or "").strip().lower()
    return ENTITY_PALETTES.get(e, COMPANY_CHART_COLORS)

def theme_style(theme: str) -> dict[str, object]:
    return THEME_STYLE.get(theme, THEME_STYLE["classic"])

def clamp_int(v: int, lo: int, hi: int) -> int:
    return max(lo, min(hi, v))

def adaptive_chart_size(
    chart_kind: str,
    categories_count: int,
    series_count: int,
    is_time_split_block: bool = False,
) -> tuple[int, int]:
    """
    Подбор размера диаграммы под объём данных:
    - больше категорий/серий -> шире/выше;
    - для блоков М/К/Г ограничиваем высоту, чтобы лист оставался читаемым.
    """
    c = max(1, int(categories_count))
    s = max(1, int(series_count))

    if chart_kind == "pie":
        width = 760 + (c - 8) * 24 + (s - 1) * 12
        height = 420 + (c - 8) * 14
        return clamp_int(width, 760, 1200), clamp_int(height, 420, 760)

    if is_time_split_block:
        width = 760 + max(0, c - 6) * 18 + max(0, s - 1) * 16
        height = 300 + max(0, c - 8) * 8
        return clamp_int(width, 760, 1140), clamp_int(height, 300, 440)

    width = 760 + max(0, c - 10) * 22 + max(0, s - 1) * 18
    height = 420 + max(0, c - 12) * 10
    return clamp_int(width, 760, 1240), clamp_int(height, 420, 720)


def chart_height_to_rows(height_px: int) -> int:
    # В ExcelJS/xlsxwriter одна строка листа в среднем ~15 px.
    return max(18, int((height_px + 20) / 15))

def parse_category_to_year_month(label: str) -> tuple[int, int] | None:
    s = str(label).strip().lower()
    m_dmy = re.match(r"^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?$", s)
    if m_dmy:
        mo = int(m_dmy.group(2))
        y = int(m_dmy.group(3))
        if 1 <= mo <= 12:
            return y, mo
    m_day = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", s)
    if m_day:
        y = int(m_day.group(1))
        mo = int(m_day.group(2))
        if 1 <= mo <= 12:
            return y, mo
    m_month = re.match(r"^(\d{4})-(\d{2})$", s)
    if m_month:
        y = int(m_month.group(1))
        mo = int(m_month.group(2))
        if 1 <= mo <= 12:
            return y, mo
    m_ru = re.match(r"^(янв\.|февр\.|мар\.|апр\.|мая|июн\.|июл\.|авг\.|сент\.|окт\.|нояб\.|дек\.)\s+(\d{4})$", s)
    if m_ru:
        mo = RU_MONTHS_MAP.get(m_ru.group(1))
        y = int(m_ru.group(2))
        if mo:
            return y, mo
    return None

def month_label_ru(year: int, month: int) -> str:
    inv = {v: k for k, v in RU_MONTHS_MAP.items()}
    return f"{inv.get(month, str(month))} {year}"

def quarter_label(year: int, month: int) -> str:
    q = (month - 1) // 3 + 1
    return f"{year}-Q{q}"

def _pie_labels_category_value_pct(
    series_list: list | None,
) -> dict[str, object]:
    """Круговые: имя сегмента, значение и %, снаружи, выноски (как эталон Excel)."""
    s0: dict = series_list[0] if series_list and isinstance(series_list[0], dict) else {}
    out: dict[str, object] = {
        "category": True,
        "value": True,
        "percentage": True,
        "position": "outside_end",
        "leader_lines": True,
        "separator": "\n",
        "font": {"size": 9, "bold": True, "color": EXCEL_ACCENT_BLUE},
    }
    if aggregation_is_money(s0.get("aggregation")):
        out["num_format"] = '# ##0,00 "₽"'
    return out


def series_data_labels(
    chart_kind: str,
    chart_title: str = "",
    chart_id: str = "",
    series_list: list | None = None,
    categories_count: int = 0,
) -> dict[str, object]:
    """Подписи серий; для pie — адаптивный формат для читаемости."""
    if chart_kind == "pie":
        # Для долей по компаниям делаем «чистую» диаграмму:
        # подписи убираем, читаемость даёт расширенная легенда справа.
        if chart_id in {
            "deals_won_pie_by_company",
            "deals_pie_by_company",
            "contacts_pie_by_company",
        }:
            return {"value": False}
        # При большом числе сегментов full-label (категория + значение + %) визуально перегружает лист.
        # Делаем подписи компактнее; для очень плотных pie отключаем их совсем и оставляем легенду + таблицу.
        if categories_count > 10:
            return {"value": False}
        if categories_count > 6:
            return {
                "category": True,
                "percentage": True,
                "position": "best_fit",
                "separator": "\n",
                "font": {"size": 8, "bold": True, "color": EXCEL_ACCENT_BLUE},
            }
        return _pie_labels_category_value_pct(series_list)
    return {
        "value": True,
        "font": {"size": 9, "color": EXCEL_ACCENT_BLUE},
    }

def source_column_is_money(header: object) -> bool:
    """Колонка с суммами/деньгами в выгрузке — формат рублей."""
    h = norm_header(header)
    if not h:
        return False
    keys = (
        "сумма",
        "amount",
        "руб",
        "₽",
        "стоимость",
        "цена",
        "выруч",
        "оплат",
        "бюджет",
        "денег",
        "чек",
        "доход",
    )
    return any(k in h for k in keys)


def aggregation_is_money(agg: object) -> bool:
    a = str(agg or "").lower()
    return a in ("sum", "avg")


def normalize_excel_number(value: object) -> float | int:
    n = float(value or 0)
    if abs(n - round(n)) < 1e-3:
        return int(round(n))
    return n


def format_int_grouped(n: float) -> str:
    return f"{int(round(n)):,}".replace(",", " ")


def format_pie_legend_label(
    category: object,
    value: float,
    total: float,
    is_money: bool,
) -> str:
    pct = int(round((100.0 * value / total), 0)) if total > 0 else 0
    value_s = f"{format_int_grouped(value)} ₽" if is_money else format_int_grouped(value)
    return f"{str(category)} — {value_s} ({pct}%)"

def xl_col(col_index_zero_based: int) -> str:
    n = col_index_zero_based + 1
    out = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out

def norm_header(h: object) -> str:
    return str(h or "").strip().lower()


def report_through_exclusive_expr(ymd: object) -> str | None:
    s = str(ymd or "").strip()
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", s)
    if not m:
        return None
    y = int(m.group(1))
    mo = int(m.group(2))
    d = int(m.group(3))
    try:
        nxt = datetime(y, mo, d) + timedelta(days=1)
    except ValueError:
        return None
    return f"DATE({nxt.year},{nxt.month},{nxt.day})"


def excel_str_literal(s: object) -> str:
    return '"' + str(s).replace('"', '""') + '"'


def excel_criteria_for_value(v: object, negate: bool = False) -> str:
    sv = str(v or "").strip()
    if sv == "(пусто)":
        return excel_str_literal("<>") if negate else excel_str_literal("")
    if negate:
        return excel_str_literal(f"<>{sv}")
    return excel_str_literal(sv)


def excel_array_constant(values: list[str]) -> str:
    return "{" + ",".join(excel_str_literal(v) for v in values) + "}"


def parse_source_filter_criteria(
    chart_item: dict,
    src_map: dict[str, int],
) -> tuple[tuple[str, list[str]] | None, list[tuple[str, str]]]:
    """
    Возвращает:
    - include: (range_expr, values[]) — один OR-фильтр по значениям
    - excludes: [(range_expr, criterion_expr), ...]
    Поддерживаем только фильтры exact values / excludeValues (остальные игнорируются).
    """
    include: tuple[str, list[str]] | None = None
    excludes: list[tuple[str, str]] = []
    raw_filters = chart_item.get("filters", [])
    if not isinstance(raw_filters, list):
        return include, excludes
    for f in raw_filters:
        if not isinstance(f, dict):
            continue
        idx = src_map.get(norm_header(f.get("columnHeader", "")))
        if idx is None:
            continue
        col = xl_col(idx)
        # src_from/src_to добавляются позже, здесь только идентификатор колонки
        range_template = col
        vals = f.get("values")
        if isinstance(vals, list) and vals:
            safe_vals = [str(v) for v in vals]
            if include is None:
                include = (range_template, safe_vals)
            # Второй include-фильтр (OR по другой колонке) не поддерживаем формулой надёжно.
        ex_vals = f.get("excludeValues")
        if isinstance(ex_vals, list):
            for ex in ex_vals:
                excludes.append((range_template, excel_criteria_for_value(ex, negate=True)))
    return include, excludes


def range_by_col(src_sheet: str, col_letter: str, src_from: int, src_to: int) -> str:
    return f"'{src_sheet}'!${col_letter}${src_from}:${col_letter}${src_to}"


def countifs_formula(
    criteria_pairs: list[tuple[str, str]],
    include_or: tuple[str, list[str]] | None = None,
) -> str:
    crit = ",".join(f"{r},{c}" for r, c in criteria_pairs)
    if include_or and len(include_or[1]) > 0:
        ir, values = include_or
        arr = excel_array_constant([str(v) for v in values])
        return f"=SUM(COUNTIFS({crit},{ir},{arr}))"
    return f"=COUNTIFS({crit})"


def sumifs_formula(
    y_range: str,
    criteria_pairs: list[tuple[str, str]],
    include_or: tuple[str, list[str]] | None = None,
) -> str:
    crit = ",".join(f"{r},{c}" for r, c in criteria_pairs)
    if include_or and len(include_or[1]) > 0:
        ir, values = include_or
        arr = excel_array_constant([str(v) for v in values])
        return f"=SUM(SUMIFS({y_range},{crit},{ir},{arr}))"
    return f"=SUMIFS({y_range},{crit})"


WON_STAGE_VALUES = [
    "Успешно реализована",
    "успешно реализована",
    "УСПЕШНО РЕАЛИЗОВАНА",
    "Успешная сделка",
    "успешная сделка",
    "10. Сделка успешна",
    "Сделка успешна",
    "сделка успешна",
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
]

LOST_STAGE_VALUES = [
    "11. Отказ / Тендер проигран",
    "Отказ / Тендер проигран",
    "Отказ",
    "отказ",
    "Провал",
    "провал",
    "Неуспех",
    "неуспех",
    "LOST",
    "lost",
    "Lost",
    "CLOSED LOST",
    "Closed Lost",
    "Не реализована",
    "не реализована",
    "Сделка проиграна",
    "сделка проиграна",
    "Проиграна",
    "проиграна",
    "Закрыта без успеха",
    "Тендер проигран",
]


def normalize_stage_value_py(s: str) -> str:
    return " ".join(s.strip().lower().replace("ё", "е").split())


WON_STAGE_VALUES_NORM = sorted({normalize_stage_value_py(v) for v in WON_STAGE_VALUES})
LOST_STAGE_VALUES_NORM = sorted({normalize_stage_value_py(v) for v in LOST_STAGE_VALUES})


def first_present_col(src_map: dict[str, int], candidates: list[str]) -> int | None:
    for c in candidates:
        idx = src_map.get(norm_header(c))
        if idx is not None:
            return idx
    return None


def build_main_info_formula(
    label: str,
    src_sheet: str,
    src_map: dict[str, int],
    src_from: int,
    src_to: int,
) -> str | None:
    if src_to < src_from:
        return None
    id_idx = first_present_col(src_map, ["ID", "id"])
    stage_idx = first_present_col(src_map, ["Стадия сделки", "стадия сделки", "Стадия", "стадия"])
    sum_idx = first_present_col(src_map, ["Сумма", "сумма", "Amount", "amount"])
    resp_idx = first_present_col(src_map, ["Ответственный", "ответственный"])
    created_idx = first_present_col(
        src_map,
        [
            "Дата создания сделки",
            "дата создания сделки",
            "Дата создания",
            "дата создания",
            "Дата регистрации",
            "дата регистрации",
            "Дата и время создания",
            "дата и время создания",
            "Дата/время создания",
            "дата/время создания",
            "created at",
            "created_at",
            "date created",
            "registration date",
        ],
    )
    closed_idx = first_present_col(
        src_map,
        [
            "Дата закрытия, факт",
            "дата закрытия, факт",
            "Дата закрытия",
            "дата закрытия",
            "Дата завершения",
            "дата завершения",
            "Дата успешного закрытия",
            "дата успешного закрытия",
            "closed date",
            "closed_at",
            "date closed",
            "close date",
        ],
    )

    def rg(ci: int | None) -> str | None:
        return None if ci is None else range_by_col(src_sheet, xl_col(ci), src_from, src_to)

    id_r = rg(id_idx)
    stage_r = rg(stage_idx)
    sum_r = rg(sum_idx)
    resp_r = rg(resp_idx)
    created_r = rg(created_idx)
    closed_r = rg(closed_idx)

    if id_r is None and stage_r is None and sum_r is None:
        return None

    total_count = f"COUNTIFS({id_r},\"<>\")" if id_r else None
    stage_norm = (
        f'LOWER(SUBSTITUTE(TRIM({stage_r}),"ё","е"))'
        if stage_r
        else None
    )
    won_match = (
        f"ISNUMBER(MATCH({stage_norm},{excel_array_constant(WON_STAGE_VALUES_NORM)},0))"
        if stage_norm
        else None
    )
    lost_match = (
        f"ISNUMBER(MATCH({stage_norm},{excel_array_constant(LOST_STAGE_VALUES_NORM)},0))"
        if stage_norm
        else None
    )
    won_count = f"SUMPRODUCT(--({won_match}))" if won_match else None
    lost_count = f"SUMPRODUCT(--({lost_match}))" if lost_match else None
    total_sum = f"SUM({sum_r})" if sum_r else None
    won_sum = (
        f"SUMPRODUCT({sum_r}*--({won_match}))"
        if sum_r and won_match
        else None
    )
    lost_sum = (
        f"SUMPRODUCT({sum_r}*--({lost_match}))"
        if sum_r and lost_match
        else None
    )

    l = label.strip().lower()
    if "активных ответственных" in l and resp_r:
        return f"=IFERROR(SUMPRODUCT(({resp_r}<>\"\")/COUNTIF({resp_r},{resp_r}&\"\")),0)"
    if "дата первой записи" in l and created_r:
        return f"=IFERROR(AGGREGATE(15,6,{created_r}/({created_r}<>\"\"),1),\"\")"
    if "дата последней записи" in l and created_r:
        return f"=IFERROR(AGGREGATE(14,6,{created_r}/({created_r}<>\"\"),1),\"\")"
    if "темп роста (среднее помесячных %)" in l and created_r:
        # Для максимальной совместимости между версиями Excel
        # оставляем значение из payload (рассчитано в приложении).
        return None
    if "всего сделок" in l and total_count:
        return f"={total_count}"
    if l == "успешных" and won_count:
        return f"={won_count}"
    if l == "проигранных" and lost_count:
        return f"={lost_count}"
    if "в работе" in l and ("доход" not in l) and total_count and won_count and lost_count:
        return f"=MAX(({total_count})-({won_count})-({lost_count}),0)"
    if "конверсия (успех к закрытым)" in l and won_count and lost_count:
        return f"=IFERROR(({won_count})/(({won_count})+({lost_count})),0)"
    if "потенциальный доход" in l and total_sum:
        return f"={total_sum}"
    if "доход от выигранных" in l and won_sum:
        return f"={won_sum}"
    if "упущенный доход" in l and lost_sum:
        return f"={lost_sum}"
    if "доход в работе" in l and total_sum and won_sum and lost_sum:
        return f"=MAX(({total_sum})-({won_sum})-({lost_sum}),0)"
    if "конверсия по сумме закрытых" in l and won_sum and lost_sum:
        return f"=IFERROR(({won_sum})/(({won_sum})+({lost_sum})),0)"
    if l == "средний чек" and total_sum and total_count:
        return f"=IFERROR(({total_sum})/({total_count}),0)"
    if "средний чек (выигранные)" in l and won_sum and won_count:
        return f"=IFERROR(({won_sum})/({won_count}),0)"
    if "средний цикл в работе" in l and created_r and closed_r and stage_r:
        # Среднее по закрытым (won+lost): по каждой строке берём max(дата_закрытия - дата_создания, 0),
        # затем делим сумму на число закрытых строк с обеими датами (как в JS-логике KPI).
        mask = (
            f"({created_r}<>\"\")*({closed_r}<>\"\")*--((({won_match})+({lost_match}))>0)"
        )
        # Элементно max(x,0) = (x + ABS(x)) / 2
        pos_days = f"(({closed_r}-{created_r})+ABS({closed_r}-{created_r}))/2"
        return f"=IFERROR(SUMPRODUCT(({mask})*({pos_days}))/SUMPRODUCT({mask}),0)"
    return None

def count_formula_day_bucket(
    src_sheet: str,
    x_col: str,
    x_row: int,
    src_from: int,
    src_to: int,
    report_end_excl_expr: str | None = None,
) -> str:
    x_range = f"'{src_sheet}'!${x_col}${src_from}:${x_col}${src_to}"
    x_cell = f"$A${x_row}"
    # Счёт по календарному дню (без времени).
    formula = f'=COUNTIFS({x_range},">="&DATEVALUE({x_cell}),{x_range},"<"&DATEVALUE({x_cell})+1'
    if report_end_excl_expr:
        formula += f',{x_range},"<"&{report_end_excl_expr}'
    return formula + ")"

def count_formula_exact(src_sheet: str, x_col: str, x_row: int, src_from: int, src_to: int) -> str:
    x_range = f"'{src_sheet}'!${x_col}${src_from}:${x_col}${src_to}"
    x_cell = f"$A${x_row}"
    return f'=COUNTIFS({x_range},{x_cell})'

def count_nonempty_formula_day_bucket(
    src_sheet: str,
    x_col: str,
    y_col: str,
    x_row: int,
    src_from: int,
    src_to: int,
    report_end_excl_expr: str | None = None,
) -> str:
    x_range = f"'{src_sheet}'!${x_col}${src_from}:${x_col}${src_to}"
    y_range = f"'{src_sheet}'!${y_col}${src_from}:${y_col}${src_to}"
    x_cell = f"$A${x_row}"
    formula = (
        f'=COUNTIFS({x_range},">="&DATEVALUE({x_cell}),{x_range},"<"&DATEVALUE({x_cell})+1,'
        f'{y_range},"<>",{y_range},"<>-"'  # как в JS: "-" считаем пустым
    )
    if report_end_excl_expr:
        formula += f',{x_range},"<"&{report_end_excl_expr}'
    return formula + ")"


def count_nonempty_formula_month(
    src_sheet: str,
    x_col: str,
    y_col: str,
    year: int,
    month: int,
    src_from: int,
    src_to: int,
    report_end_excl_expr: str | None = None,
) -> str:
    x_range = f"'{src_sheet}'!${x_col}${src_from}:${x_col}${src_to}"
    y_range = f"'{src_sheet}'!${y_col}${src_from}:${y_col}${src_to}"
    start = f"DATE({year},{month},1)"
    end = f"EOMONTH({start},0)+1"
    formula = (
        f'=COUNTIFS({x_range},">="&{start},{x_range},"<"&{end},'
        f'{y_range},"<>",{y_range},"<>-"'
    )
    if report_end_excl_expr:
        formula += f',{x_range},"<"&{report_end_excl_expr}'
    return formula + ")"


def count_nonempty_formula_exact(
    src_sheet: str,
    x_col: str,
    y_col: str,
    x_row: int,
    src_from: int,
    src_to: int,
) -> str:
    x_range = f"'{src_sheet}'!${x_col}${src_from}:${x_col}${src_to}"
    y_range = f"'{src_sheet}'!${y_col}${src_from}:${y_col}${src_to}"
    x_cell = f"$A${x_row}"
    return f'=COUNTIFS({x_range},{x_cell},{y_range},"<>",{y_range},"<>-")'

def count_month_formula(
    src_sheet: str,
    x_col: str,
    year: int,
    month: int,
    src_from: int,
    src_to: int,
    report_end_excl_expr: str | None = None,
) -> str:
    x_range = f"'{src_sheet}'!${x_col}${src_from}:${x_col}${src_to}"
    start = f"DATE({year},{month},1)"
    end = f"EOMONTH({start},0)+1"
    formula = f'=COUNTIFS({x_range},">="&{start},{x_range},"<"&{end}'
    if report_end_excl_expr:
        formula += f',{x_range},"<"&{report_end_excl_expr}'
    return formula + ")"


def count_cumulative_through_month_from_chart_start_formula(
    src_sheet: str,
    x_col: str,
    start_year: int,
    start_month: int,
    end_year: int,
    end_month: int,
    src_from: int,
    src_to: int,
    report_end_excl_expr: str | None = None,
) -> str:
    """Нарастающий итог по месяцам: все записи с датой X от 1-го числа первого месяца оси до конца текущего месяца."""
    x_range = f"'{src_sheet}'!${x_col}${src_from}:${x_col}${src_to}"
    lower = f"DATE({start_year},{start_month},1)"
    end_excl = f"EOMONTH(DATE({end_year},{end_month},1),0)+1"
    formula = f'=COUNTIFS({x_range},">="&{lower},{x_range},"<"&{end_excl}'
    if report_end_excl_expr:
        formula += f',{x_range},"<"&{report_end_excl_expr}'
    return formula + ")"


def count_distinct_ids_between_dates_formula(
    src_sheet: str,
    x_col: str,
    id_col: str,
    start_expr: str,
    end_excl_expr: str,
    src_from: int,
    src_to: int,
    report_end_excl_expr: str | None = None,
) -> str:
    """Distinct count по ID в календарном интервале [start, end_excl)."""
    x_range = f"'{src_sheet}'!${x_col}${src_from}:${x_col}${src_to}"
    id_range = f"'{src_sheet}'!${id_col}${src_from}:${id_col}${src_to}"
    mask = f"({x_range}>={start_expr})*({x_range}<{end_excl_expr})*({id_range}<>\"\")"
    den = (
        f'COUNTIFS({id_range},{id_range},{x_range},">="&{start_expr},{x_range},"<"&{end_excl_expr}'
    )
    if report_end_excl_expr:
        mask += f"*({x_range}<{report_end_excl_expr})"
        den += f',{x_range},"<"&{report_end_excl_expr}'
    den += ")"
    return f"=IFERROR(SUMPRODUCT({mask}/({den})),0)"

def sum_formula(
    src_sheet: str,
    x_col: str,
    y_col: str,
    x_row: int,
    src_from: int,
    src_to: int,
    report_end_excl_expr: str | None = None,
) -> str:
    x_range = f"'{src_sheet}'!${x_col}${src_from}:${x_col}${src_to}"
    y_range = f"'{src_sheet}'!${y_col}${src_from}:${y_col}${src_to}"
    x_cell = f"$A${x_row}"
    formula = f'=SUMIFS({y_range},{x_range},">="&DATEVALUE({x_cell}),{x_range},"<"&DATEVALUE({x_cell})+1'
    if report_end_excl_expr:
        formula += f',{x_range},"<"&{report_end_excl_expr}'
    return formula + ")"


def sum_formula_exact(
    src_sheet: str,
    x_col: str,
    y_col: str,
    x_row: int,
    src_from: int,
    src_to: int,
) -> str:
    """Сумма Y по точному совпадению X с подписью категории (текстовая ось)."""
    x_range = f"'{src_sheet}'!${x_col}${src_from}:${x_col}${src_to}"
    y_range = f"'{src_sheet}'!${y_col}${src_from}:${y_col}${src_to}"
    x_cell = f"$A${x_row}"
    return f'=SUMIFS({y_range},{x_range},{x_cell})'


def sum_month_formula(
    src_sheet: str,
    x_col: str,
    y_col: str,
    year: int,
    month: int,
    src_from: int,
    src_to: int,
    report_end_excl_expr: str | None = None,
) -> str:
    x_range = f"'{src_sheet}'!${x_col}${src_from}:${x_col}${src_to}"
    y_range = f"'{src_sheet}'!${y_col}${src_from}:${y_col}${src_to}"
    start = f"DATE({year},{month},1)"
    end = f"EOMONTH({start},0)+1"
    formula = f'=SUMIFS({y_range},{x_range},">="&{start},{x_range},"<"&{end}'
    if report_end_excl_expr:
        formula += f',{x_range},"<"&{report_end_excl_expr}'
    return formula + ")"


def formula_literal_number(value: object) -> str:
    """Формула с тем же числовым результатом (через N), чтобы это была именно формула."""
    n = normalize_excel_number(value)
    return f"=N({n})"


def write_formula_literal_number(ws, row: int, col: int, value: object, fmt) -> None:
    n = normalize_excel_number(value)
    ws.write_formula(row, col, formula_literal_number(n), fmt, n)


def normalize_category_label(label: object) -> str:
    s = str(label or "").strip()
    if not s:
        return ""
    # ISO timestamps / datetimes: 2025-07-30T17:24:36.000Z -> 2025-07-30
    m_iso = re.match(r"^(\d{4})-(\d{2})-(\d{2})(?:[T\s].*)?$", s)
    if m_iso:
        return f"{m_iso.group(1)}-{m_iso.group(2)}-{m_iso.group(3)}"
    # RU datetime/date: 07.04.2026 16:32[:33] -> 2026-04-07
    m_ru = re.match(r"^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?$", s)
    if m_ru:
        d = int(m_ru.group(1))
        mo = int(m_ru.group(2))
        y = int(m_ru.group(3))
        return f"{y:04d}-{mo:02d}-{d:02d}"
    return s

def try_parse_source_date(value: object) -> datetime | None:
    if isinstance(value, str):
        s = value.strip()
        if not s:
            return None
        m_ru = re.match(r"^(\d{1,2})\.(\d{1,2})\.(\d{4})$", s)
        if m_ru:
            d = int(m_ru.group(1))
            mo = int(m_ru.group(2))
            y = int(m_ru.group(3))
            try:
                return datetime(y, mo, d)
            except ValueError:
                return None
        m_iso = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", s)
        if m_iso:
            y = int(m_iso.group(1))
            mo = int(m_iso.group(2))
            d = int(m_iso.group(3))
            try:
                return datetime(y, mo, d)
            except ValueError:
                return None
        m_iso_ts = re.match(r"^(\d{4})-(\d{2})-(\d{2})[T\s].*$", s)
        if m_iso_ts:
            y = int(m_iso_ts.group(1))
            mo = int(m_iso_ts.group(2))
            d = int(m_iso_ts.group(3))
            try:
                return datetime(y, mo, d)
            except ValueError:
                return None
    return None

def aggregate_time_buckets(
    categories: list[str],
    series: list[dict],
) -> tuple[dict | None, dict | None, dict | None]:
    parsed = [parse_category_to_year_month(c) for c in categories]
    if not any(parsed):
        return None, None, None
    month_keys = sorted({(y, m) for y, m in parsed if y and m})
    if not month_keys:
        return None, None, None

    month_index = {k: i for i, k in enumerate(month_keys)}
    month_labels = [month_label_ru(y, m) for y, m in month_keys]
    month_quarters = [quarter_label(y, m) for y, m in month_keys]

    month_series: list[dict] = []
    series_cumulative: list[bool] = []
    for s in series:
        cumulative = bool(s.get("cumulative"))
        source_header = s.get("sourceHeader", "")
        values = [0.0 for _ in month_keys]
        src = s.get("values", [])
        for i, key in enumerate(parsed):
            if key is None or i >= len(src):
                continue
            idx = month_index[key]
            v = float(src[i] or 0)
            if cumulative:
                # Нарастающий итог по месяцу — не суммируем с другими строками того же месяца.
                values[idx] = v
            else:
                values[idx] += v
        month_series.append(
            {
                "name": s.get("name", ""),
                "values": values,
                "aggregation": s.get("aggregation", "count"),
                "sourceHeader": source_header,
                "cumulative": cumulative,
            }
        )
        series_cumulative.append(cumulative)

    quarter_keys = sorted({q for q in month_quarters})
    quarter_index = {q: i for i, q in enumerate(quarter_keys)}
    quarter_series: list[dict] = []
    for si, s in enumerate(month_series):
        cumulative = series_cumulative[si]
        q_values = [0.0 for _ in quarter_keys]
        if cumulative:
            # Итог на конец квартала = значение последнего календарного месяца квартала,
            # а не сумма помесячных нарастающих значений.
            for qi, qk in enumerate(quarter_keys):
                last_v = 0.0
                for mi in range(len(month_keys)):
                    if month_quarters[mi] == qk:
                        last_v = float(s["values"][mi] or 0)
                q_values[qi] = last_v
        else:
            for i, v in enumerate(s["values"]):
                q_values[quarter_index[month_quarters[i]]] += float(v or 0)
        quarter_series.append(
            {
                "name": s.get("name", ""),
                "values": q_values,
                "aggregation": s.get("aggregation", "count"),
                "sourceHeader": s.get("sourceHeader", ""),
                "cumulative": cumulative,
            }
        )

    year_keys = sorted({y for y, _ in month_keys})
    year_index = {y: i for i, y in enumerate(year_keys)}
    year_series: list[dict] = []
    for si, s in enumerate(month_series):
        cumulative = series_cumulative[si]
        y_values = [0.0 for _ in year_keys]
        if cumulative:
            for yi, y in enumerate(year_keys):
                last_v = 0.0
                for mi in range(len(month_keys)):
                    if month_keys[mi][0] == y:
                        last_v = float(s["values"][mi] or 0)
                y_values[yi] = last_v
        else:
            for i, v in enumerate(s["values"]):
                y = month_keys[i][0]
                y_values[year_index[y]] += float(v or 0)
        year_series.append(
            {
                "name": s.get("name", ""),
                "values": y_values,
                "aggregation": s.get("aggregation", "count"),
                "sourceHeader": s.get("sourceHeader", ""),
                "cumulative": cumulative,
            }
        )

    month_payload = {
        "labels": month_labels,
        "quarters": month_quarters,
        "series": month_series,
    }
    quarter_payload = {
        "labels": quarter_keys,
        "series": quarter_series,
    }
    year_payload = {
        "labels": [str(y) for y in year_keys],
        "series": year_series,
    }
    return month_payload, quarter_payload, year_payload


def main() -> int:
    if len(sys.argv) != 3:
        return 2
    payload_path = Path(sys.argv[1])
    out_path = Path(sys.argv[2])
    payload = json.loads(payload_path.read_text(encoding="utf-8"))
    theme = str(payload.get("theme", "classic"))
    tstyle = theme_style(theme)

    charts = payload.get("charts", [])
    summary = payload.get("summary", {})
    source_data = payload.get("sourceData", {}) or {}
    workbook = xlsxwriter.Workbook(str(out_path))
    used_names: set[str] = set()

    title_fmt = workbook.add_format(
        {
            "bold": True,
            "font_size": 14,
            "bg_color": EXCEL_ACCENT_BLUE,
            "font_color": "#FFFFFF",
            "align": "left",
            "valign": "vcenter",
        }
    )
    head_fmt = workbook.add_format(
        {
            "bold": True,
            "bg_color": EXCEL_ACCENT_BLUE,
            "font_color": "#FFFFFF",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
        }
    )
    # Подписи столбцов под заголовком листа (как в эталоне Excel: светло-синий фон).
    table_col_head_fmt = workbook.add_format(
        {
            "bold": True,
            "bg_color": "#D9E2F3",
            "font_color": "#1F2937",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
        }
    )
    cell_fmt = workbook.add_format({"border": 1, "valign": "vcenter"})
    cell_date_fmt = workbook.add_format({"border": 1, "valign": "vcenter", "num_format": "dd.mm.yyyy"})
    # Пробел между разрядами; суммы/средние — с копейками и «₽» (как в ru-RU Excel).
    fmt_rub = workbook.add_format(
        {"border": 1, "valign": "vcenter", "num_format": '# ##0,00 "₽"'}
    )
    fmt_count = workbook.add_format({"border": 1, "valign": "vcenter", "num_format": "# ##0"})
    fmt_pct = workbook.add_format({"border": 1, "valign": "vcenter", "num_format": "0.0%"})
    thick_top_rub = workbook.add_format(
        {"border": 1, "top": 2, "valign": "vcenter", "num_format": '# ##0,00 "₽"'}
    )
    thick_top_count = workbook.add_format(
        {"border": 1, "top": 2, "valign": "vcenter", "num_format": "# ##0"}
    )
    axis_num_fmt = workbook.add_format({})

    # Source metadata (нужно заранее для формул KPI на Summary).
    src_headers = source_data.get("headers", []) or []
    src_rows = source_data.get("rows", []) or []
    src_from = 4
    src_to = 3 + len(src_rows)
    src_map: dict[str, int] = {norm_header(h): i for i, h in enumerate(src_headers)}
    src_sheet_formula_name = "Исходные данные"

    # Summary sheet.
    summary_name = sanitize_sheet_name("Общая информация", used_names)
    ws_sum = workbook.add_worksheet(summary_name)
    ws_sum.set_column("A:A", 28)
    ws_sum.set_column("B:B", 48)
    ws_sum.set_column("C:C", 64)
    ws_sum.merge_range("A1:C1", "Общая информация", title_fmt)

    main_info = summary.get("mainInfo") or []
    row = 2
    if main_info:
        section_fmt = workbook.add_format(
            {
                "bold": True,
                "font_size": 12,
                "bg_color": "#E0E7FF",
                "font_color": EXCEL_ACCENT_BLUE,
                "border": 1,
                "align": "left",
                "valign": "vcenter",
            }
        )
        ws_sum.merge_range(row, 0, row, 2, "Основные показатели", section_fmt)
        row += 1
        main_row_by_label: dict[str, int] = {}
        for item in main_info:
            if not isinstance(item, dict):
                continue
            label = str(item.get("label", ""))
            val = item.get("value", "")
            main_row_by_label[label.strip().lower()] = row + 1
            ws_sum.write(row, 0, label, cell_fmt)
            mi_formula = build_main_info_formula(
                label,
                src_sheet_formula_name,
                src_map,
                src_from,
                src_to,
            )
            if mi_formula:
                ll = label.lower()
                if any(k in ll for k in ("дата первой записи", "дата последней записи")):
                    mi_fmt = workbook.add_format({"border": 1, "valign": "vcenter", "num_format": "dd.mm.yyyy"})
                elif any(k in ll for k in ("конверсия", "темп роста")):
                    mi_fmt = fmt_pct
                elif "цикл" in ll:
                    mi_fmt = workbook.add_format({"border": 1, "valign": "vcenter", "num_format": "0.0"})
                elif any(k in ll for k in ("руб", "₽", "сумм", "выруч", "доход", "чек", "стоим", "оплат", "средн", "денег")):
                    mi_fmt = fmt_rub
                else:
                    mi_fmt = fmt_count
                ws_sum.write_formula(row, 1, mi_formula, mi_fmt)
                row += 1
                continue
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                ll = label.lower()
                if any(
                    k in ll
                    for k in (
                        "конверсия",
                        "темп роста",
                    )
                ):
                    mi_fmt = fmt_pct
                elif "цикл" in ll:
                    mi_fmt = workbook.add_format(
                        {"border": 1, "valign": "vcenter", "num_format": "0.0"}
                    )
                elif any(
                    k in ll
                    for k in (
                        "руб",
                        "₽",
                        "сумм",
                        "выруч",
                        "доход",
                        "чек",
                        "стоим",
                        "оплат",
                        "средн",
                        "денег",
                    )
                ):
                    mi_fmt = fmt_rub
                else:
                    mi_fmt = fmt_count
                write_formula_literal_number(ws_sum, row, 1, val, mi_fmt)
            else:
                ws_sum.write(row, 1, val, cell_fmt)
            row += 1
        # Конверсия в деньгах должна строго опираться на уже рассчитанные строки доходов.
        conv_money_row = main_row_by_label.get("конверсия по сумме закрытых")
        income_won_row = main_row_by_label.get("доход от выигранных")
        income_lost_row = main_row_by_label.get("упущенный доход")
        if conv_money_row and income_won_row and income_lost_row:
            conv_formula = (
                f"=IFERROR(B{income_won_row}/(B{income_won_row}+B{income_lost_row}),0)"
            )
            ws_sum.write_formula(conv_money_row - 1, 1, conv_formula, fmt_pct)
        row += 1

    meta_rows = [
        ("Дата формирования", summary.get("createdAt", "")),
        ("Исходный файл", summary.get("sourceFileName", "")),
        ("Графиков в файле", len(charts)),
    ]
    for k, v in meta_rows:
        ws_sum.write(row, 0, k, cell_fmt)
        if k == "Графиков в файле" and isinstance(v, int):
            write_formula_literal_number(ws_sum, row, 1, v, fmt_count)
        else:
            ws_sum.write(row, 1, v, cell_fmt)
        row += 1

    row += 1
    ws_sum.write_row(row, 0, ["Раздел", "График", "Статус"], head_fmt)
    row += 1
    for c in charts:
        status = "Готово"
        if c.get("resolveError") or c.get("resolve_error"):
            status = "Не построен"
        elif c.get("chartDataEmpty") or not (c.get("categories") or []):
            status = "Нет данных"
        ws_sum.write_row(
            row,
            0,
            [c.get("blockLabel", ""), c.get("title", ""), status],
            cell_fmt,
        )
        row += 1

    # Source data sheet.
    src_sheet_name = ""
    if src_headers:
        src_name = sanitize_sheet_name("Исходные данные", used_names)
        src_sheet_name = src_name
        ws_src = workbook.add_worksheet(src_name)
        ws_src.set_column(0, max(len(src_headers) - 1, 0), 20)
        ws_src.merge_range("A1:D1", "Исходные данные", title_fmt)
        ws_src.write_row(2, 0, src_headers, head_fmt)
        src_map = {norm_header(h): i for i, h in enumerate(src_headers)}
        for i, row_values in enumerate(src_rows):
            row_idx = 3 + i
            for j, raw in enumerate(row_values):
                if isinstance(raw, (int, float)):
                    hdr = src_headers[j] if j < len(src_headers) else ""
                    src_nf = fmt_rub if source_column_is_money(hdr) else fmt_count
                    write_formula_literal_number(ws_src, row_idx, j, raw, src_nf)
                elif isinstance(raw, str):
                    d = try_parse_source_date(raw)
                    if d is not None:
                        ws_src.write_datetime(row_idx, j, d, cell_date_fmt)
                    else:
                        ws_src.write(row_idx, j, raw, cell_fmt)
                else:
                    ws_src.write(row_idx, j, str(raw), cell_fmt)

    # Chart sheets.
    for c in charts:
        resolve_err = c.get("resolveError") or c.get("resolve_error")
        if resolve_err:
            base_sheet_name = c.get("sheetName") or c.get("title") or "График"
            sheet_name = sanitize_sheet_name(base_sheet_name, used_names)
            ws = workbook.add_worksheet(sheet_name)
            ws.set_column("A:A", 28)
            ws.set_column("B:Z", 16, axis_num_fmt)
            ws.merge_range("A1:D1", c.get("title", "График"), title_fmt)
            note_fmt = workbook.add_format(
                {
                    "italic": True,
                    "font_color": "#64748B",
                    "valign": "vcenter",
                    "text_wrap": True,
                }
            )
            ws.merge_range(
                3,
                0,
                3,
                25,
                str(resolve_err),
                note_fmt,
            )
            continue

        categories = [normalize_category_label(v) for v in (c.get("categories", []) or [])]
        series = c.get("series", [])
        month_payload, quarter_payload, year_payload = aggregate_time_buckets(categories, series)
        has_time_split = month_payload is not None and quarter_payload is not None and year_payload is not None

        base_sheet_name = c.get("sheetName") or c.get("title") or "График"
        sheet_name_raw = f"{base_sheet_name} [М+К+Г]" if has_time_split else base_sheet_name
        sheet_name = sanitize_sheet_name(sheet_name_raw, used_names)
        ws = workbook.add_worksheet(sheet_name)
        ws.set_column("A:A", 28)
        ws.set_column("B:Z", 16, axis_num_fmt)
        ws.merge_range("A1:D1", c.get("title", "График"), title_fmt)

        headers = [c.get("xHeader", "Категория"), *[s.get("name", "") for s in c.get("series", [])]]
        ws.write_row(1, 0, headers, table_col_head_fmt)

        chart_data_empty = bool(c.get("chartDataEmpty")) or len(categories) == 0
        if chart_data_empty:
            note_fmt = workbook.add_format(
                {
                    "italic": True,
                    "font_color": "#64748B",
                    "valign": "vcenter",
                    "text_wrap": True,
                }
            )
            ws.merge_range(
                2,
                0,
                2,
                25,
                "Нет данных для этого отчёта (проверьте фильтры и состав выгрузки).",
                note_fmt,
            )
            continue

        chart_month_start = (
            parse_category_to_year_month(str(categories[0])) if categories else None
        )

        for i, category in enumerate(categories):
            ws.write(2 + i, 0, category, cell_fmt)
            for j, s in enumerate(series):
                values = s.get("values", [])
                value = values[i] if i < len(values) else 0
                wrote_formula = False
                if c.get("linkToSourcePossible") and src_sheet_name:
                    x_idx_raw = c.get("xColumnIndex")
                    x_idx = (
                        int(x_idx_raw)
                        if isinstance(x_idx_raw, (int, float))
                        else src_map.get(norm_header(c.get("xHeader", "")))
                    )
                    y_idx_raw = s.get("sourceColumnIndex")
                    y_idx = (
                        int(y_idx_raw)
                        if isinstance(y_idx_raw, (int, float))
                        else src_map.get(norm_header(s.get("sourceHeader", "")))
                    )
                    distinct_idx_raw = s.get("countDistinctBySourceColumnIndex")
                    distinct_idx = (
                        int(distinct_idx_raw)
                        if isinstance(distinct_idx_raw, (int, float))
                        else None
                    )
                    gran = str(c.get("dateGranularity") or "")
                    x_is_date = bool(c.get("xIsDate"))
                    if x_is_date and not gran:
                        gran = "day"
                    report_end_excl_expr = (
                        report_through_exclusive_expr(c.get("reportThroughYmd"))
                        if x_is_date
                        else None
                    )
                    month_parsed: tuple[int, int] | None = None
                    if x_is_date and gran == "month":
                        month_parsed = parse_category_to_year_month(str(category))
                    if x_idx is not None:
                        src_from = 4
                        src_to = 3 + len(src_rows)
                        x_col = xl_col(x_idx)
                        out_row = 3 + i
                        agg = str(s.get("aggregation", "count"))
                        formula = None
                        if agg == "count":
                            if distinct_idx is not None and x_is_date and gran == "month" and month_parsed is not None:
                                id_col = xl_col(distinct_idx)
                                if s.get("cumulative") and chart_month_start is not None:
                                    start_expr = f"DATE({chart_month_start[0]},{chart_month_start[1]},1)"
                                else:
                                    start_expr = f"DATE({month_parsed[0]},{month_parsed[1]},1)"
                                end_excl_expr = (
                                    f"EOMONTH(DATE({month_parsed[0]},{month_parsed[1]},1),0)+1"
                                )
                                formula = count_distinct_ids_between_dates_formula(
                                    src_sheet_name,
                                    x_col,
                                    id_col,
                                    start_expr,
                                    end_excl_expr,
                                    src_from,
                                    src_to,
                                    report_end_excl_expr,
                                )
                            else:
                                if (
                                    s.get("cumulative")
                                    and x_is_date
                                    and gran == "month"
                                    and chart_month_start is not None
                                    and month_parsed is not None
                                ):
                                    formula = count_cumulative_through_month_from_chart_start_formula(
                                        src_sheet_name,
                                        x_col,
                                        chart_month_start[0],
                                        chart_month_start[1],
                                        month_parsed[0],
                                        month_parsed[1],
                                        src_from,
                                        src_to,
                                        report_end_excl_expr,
                                    )
                                else:
                                    if x_is_date and gran == "month" and month_parsed is not None:
                                        formula = count_month_formula(
                                            src_sheet_name,
                                            x_col,
                                            month_parsed[0],
                                            month_parsed[1],
                                            src_from,
                                            src_to,
                                            report_end_excl_expr,
                                        )
                                if formula is None:
                                    if x_is_date:
                                        formula = count_formula_day_bucket(
                                            src_sheet_name,
                                            x_col,
                                            out_row,
                                            src_from,
                                            src_to,
                                            report_end_excl_expr,
                                        )
                                    else:
                                        formula = count_formula_exact(
                                            src_sheet_name, x_col, out_row, src_from, src_to
                                        )
                                if s.get("cumulative") and i > 0:
                                    prev_excel_row = 2 + i
                                    val_col = xl_col(j + 1)
                                    if month_parsed is not None and gran == "month" and x_is_date:
                                        cur = count_month_formula(
                                            src_sheet_name,
                                            x_col,
                                            month_parsed[0],
                                            month_parsed[1],
                                            src_from,
                                            src_to,
                                            report_end_excl_expr,
                                        )[1:]
                                    elif x_is_date:
                                        cur = count_formula_day_bucket(
                                            src_sheet_name,
                                            x_col,
                                            out_row,
                                            src_from,
                                            src_to,
                                            report_end_excl_expr,
                                        )[1:]
                                    else:
                                        cur = count_formula_exact(
                                            src_sheet_name, x_col, out_row, src_from, src_to
                                        )[1:]
                                    formula = f"={val_col}{prev_excel_row}+{cur}"
                        elif agg == "sum" and y_idx is not None:
                            y_col = xl_col(y_idx)
                            if x_is_date and gran == "day":
                                formula = sum_formula(
                                    src_sheet_name,
                                    x_col,
                                    y_col,
                                    out_row,
                                    src_from,
                                    src_to,
                                    report_end_excl_expr,
                                )
                            elif x_is_date and gran == "month" and month_parsed is not None:
                                formula = sum_month_formula(
                                    src_sheet_name,
                                    x_col,
                                    y_col,
                                    month_parsed[0],
                                    month_parsed[1],
                                    src_from,
                                    src_to,
                                    report_end_excl_expr,
                                )
                            elif not x_is_date:
                                formula = sum_formula_exact(
                                    src_sheet_name, x_col, y_col, out_row, src_from, src_to
                                )
                        elif agg == "avg" and y_idx is not None:
                            y_col = xl_col(y_idx)
                            sum_f = None
                            cnt_f = None
                            chart_id = str(c.get("chartId") or "")
                            category_label = str(category)
                            if (
                                not x_is_date
                                and chart_id == "deals_kp_avg_and_check_at_send_stage"
                            ):
                                x_range = range_by_col(src_sheet_name, x_col, src_from, src_to)
                                y_range = range_by_col(src_sheet_name, y_col, src_from, src_to)
                                cl = category_label.strip().lower()
                                # Для строки «Средний чек (стадия «Отправка КП»): SUM(stage contains) / COUNT(stage contains)
                                if "отправк" in cl and ("кп" in cl or "kp" in cl):
                                    sum_f = (
                                        f"=SUMIFS({y_range},{x_range},\"*отправк*\",{x_range},\"*кп*\",{y_range},\">0\")"
                                    )
                                    cnt_f = (
                                        f"=COUNTIFS({x_range},\"*отправк*\",{x_range},\"*кп*\",{y_range},\">0\")"
                                    )
                                # Для «Средняя стоимость КП»: среднее по положительным значениям колонки показателя.
                                elif "стоим" in cl and "кп" in cl:
                                    sum_f = f"=SUMIFS({y_range},{y_range},\">0\")"
                                    cnt_f = f"=COUNTIFS({y_range},\">0\")"
                            if sum_f is None or cnt_f is None:
                                if x_is_date and gran == "day":
                                    sum_f = sum_formula(
                                        src_sheet_name,
                                        x_col,
                                        y_col,
                                        out_row,
                                        src_from,
                                        src_to,
                                        report_end_excl_expr,
                                    )
                                    cnt_f = count_formula_day_bucket(
                                        src_sheet_name,
                                        x_col,
                                        out_row,
                                        src_from,
                                        src_to,
                                        report_end_excl_expr,
                                    )
                                elif x_is_date and gran == "month" and month_parsed is not None:
                                    sum_f = sum_month_formula(
                                        src_sheet_name,
                                        x_col,
                                        y_col,
                                        month_parsed[0],
                                        month_parsed[1],
                                        src_from,
                                        src_to,
                                        report_end_excl_expr,
                                    )
                                    cnt_f = count_month_formula(
                                        src_sheet_name,
                                        x_col,
                                        month_parsed[0],
                                        month_parsed[1],
                                        src_from,
                                        src_to,
                                        report_end_excl_expr,
                                    )
                                elif not x_is_date:
                                    sum_f = sum_formula_exact(
                                        src_sheet_name, x_col, y_col, out_row, src_from, src_to
                                    )
                                    cnt_f = count_formula_exact(
                                        src_sheet_name, x_col, out_row, src_from, src_to
                                    )
                            if sum_f and cnt_f:
                                formula = f"=IFERROR(({sum_f[1:]})/({cnt_f[1:]}),0)"
                        elif agg == "count_nonempty" and y_idx is not None:
                            y_col = xl_col(y_idx)
                            if x_is_date and gran == "day":
                                formula = count_nonempty_formula_day_bucket(
                                    src_sheet_name,
                                    x_col,
                                    y_col,
                                    out_row,
                                    src_from,
                                    src_to,
                                    report_end_excl_expr,
                                )
                            elif x_is_date and gran == "month" and month_parsed is not None:
                                formula = count_nonempty_formula_month(
                                    src_sheet_name,
                                    x_col,
                                    y_col,
                                    month_parsed[0],
                                    month_parsed[1],
                                    src_from,
                                    src_to,
                                    report_end_excl_expr,
                                )
                            elif not x_is_date:
                                formula = count_nonempty_formula_exact(
                                    src_sheet_name,
                                    x_col,
                                    y_col,
                                    out_row,
                                    src_from,
                                    src_to,
                                )
                        if formula:
                            series_nf = (
                                fmt_rub
                                if aggregation_is_money(s.get("aggregation"))
                                else fmt_count
                            )
                            ws.write_formula(
                                2 + i,
                                j + 1,
                                formula,
                                series_nf,
                                normalize_excel_number(value),
                            )
                            wrote_formula = True
                if not wrote_formula:
                    series_nf = (
                        fmt_rub
                        if aggregation_is_money(s.get("aggregation"))
                        else fmt_count
                    )
                    write_formula_literal_number(ws, 2 + i, j + 1, value, series_nf)

        chart_kind = chart_type_for_xlsxwriter(c.get("chartType", "column"))
        palette = palette_for_entity(c.get("entityId") or c.get("entity_id") or "")
        if not has_time_split:
            chart = workbook.add_chart({"type": chart_kind})
            cat_col = 0
            if chart_kind == "pie":
                # Для pie делаем легенду информативной: категория + значение + доля.
                # Категории для таблицы остаются в колонке A, отдельные подписи для легенды — в служебной колонке.
                pie_legend_col = max(1, len(series) + 1)
                pie_total = 0.0
                if isinstance(series, list) and len(series) > 0 and isinstance(series[0], dict):
                    s0_vals = series[0].get("values", []) or []
                    pie_total = sum(float(v or 0) for v in s0_vals)
                    is_money = aggregation_is_money(series[0].get("aggregation"))
                    ws.write(1, pie_legend_col, "Легенда", table_col_head_fmt)
                    for i, category in enumerate(categories):
                        v = float(s0_vals[i] or 0) if i < len(s0_vals) else 0.0
                        ws.write(
                            2 + i,
                            pie_legend_col,
                            format_pie_legend_label(category, v, pie_total, is_money),
                            cell_fmt,
                        )
                    ws.set_column(pie_legend_col, pie_legend_col, 42)
                    cat_col = pie_legend_col
            cat_range = [sheet_name, 2, cat_col, 2 + max(len(categories) - 1, 0), cat_col]
            for j, s in enumerate(series):
                val_range = [sheet_name, 2, j + 1, 2 + max(len(categories) - 1, 0), j + 1]
                color = palette[j % len(palette)]
                series_opts: dict[str, object] = {
                    "name": s.get("name", f"Ряд {j + 1}"),
                    "categories": cat_range,
                    "values": val_range,
                    "fill": {"color": color},
                    "line": {"color": color, "width": float(tstyle["line_width"])},
                    "data_labels": series_data_labels(
                        chart_kind,
                        str(c.get("title", "")),
                        str(c.get("chartId", "")),
                        series if isinstance(series, list) else [],
                        len(categories),
                    ),
                }
                if chart_kind in ("column", "bar") and j == 0:
                    series_opts["gap"] = COLUMN_BAR_GAP
                if chart_kind != "pie":
                    series_opts["marker"] = {
                        "type": "circle",
                        "size": int(tstyle["marker"]),
                        "border": {"color": color},
                        "fill": {"color": color},
                    }
                if chart_kind == "pie":
                    series_opts["points"] = [
                        {
                            "fill": {"color": palette[i % len(palette)]},
                            "line": {"color": palette[i % len(palette)], "width": 1.25},
                        }
                        for i, _ in enumerate(categories)
                    ]
                chart.add_series(
                    series_opts
                )
            pie_title_font = (
                {"name": "Calibri", "size": 13, "bold": True, "color": "#000000"}
                if chart_kind == "pie"
                else COLUMN_CHART_TITLE_FONT
            )
            chart.set_title(
                {
                    "name": c.get("title", "График"),
                    "name_font": pie_title_font,
                }
            )
            pie_many_categories = chart_kind == "pie" and len(categories) > 14
            pie_medium_categories = chart_kind == "pie" and len(categories) > 9
            legend_opts: dict[str, object] = {
                "position": "right",
                "font": {"size": 10, "bold": True, "color": "#0070C0"},
            }
            if pie_many_categories:
                legend_opts = {
                    "position": "right",
                    "font": {"size": 8, "bold": False, "color": "#0070C0"},
                }
            elif pie_medium_categories:
                legend_opts = {
                    "position": "right",
                    "font": {"size": 9, "bold": True, "color": "#0070C0"},
                }
            chart.set_legend(legend_opts)
            chart.set_chartarea({"fill": {"color": "#FFFFFF"}, "border": {"none": True}})
            if chart_kind == "pie":
                chart.set_plotarea({"fill": {"color": "#FFFFFF"}, "border": {"none": True}})
            else:
                chart.set_plotarea(
                    {
                        "fill": {"color": str(tstyle["plot_bg"])},
                        "border": {"color": "#D9E2F3", "width": 1},
                    }
                )
            if chart_kind != "pie":
                x_axis_opts: dict[str, object] = {
                    "name": c.get("xHeader", ""),
                    "name_font": {"size": 10, "color": "#0070C0"},
                    "num_font": (
                        dict(COLUMN_CATEGORY_AXIS_NUM_FONT)
                        if chart_kind in ("column", "bar")
                        else {"size": 9, "color": "#0070C0"}
                    ),
                    "line": {"color": str(tstyle["axis"])},
                }
                chart.set_x_axis(x_axis_opts)
                chart.set_y_axis(
                    {
                        "name_font": {"size": 10, "color": "#0070C0"},
                        "num_font": {"size": 9, "color": "#0070C0"},
                        "major_gridlines": {
                            "visible": True,
                            "line": {"color": str(tstyle["grid"]), "dash_type": "dash"},
                        },
                        "line": {"color": str(tstyle["axis"])},
                    }
                )
            if chart_kind == "pie":
                chart.set_style(int(tstyle["pie_style"]))
            else:
                chart.set_style(int(tstyle["chart_style"]))
        if has_time_split:
            # Три графика на листе: месяц, квартал, год (без общего графика по исходной оси).
            # ---- Monthly block ----
            # Основная таблица: строка 1 — заголовок, 2 — подписи столбцов, с 3-й — данные.
            # Месячный блок начинается после неё.
            month_title_row = 5 + len(categories)
            ws.merge_range(
                f"A{month_title_row}:D{month_title_row}",
                "По месяцам (границы кварталов в таблице)",
                title_fmt,
            )
            month_head_row = month_title_row + 2
            month_headers = ["Квартал", "Месяц", *[s.get("name", "") for s in series]]
            ws.write_row(month_head_row - 1, 0, month_headers, table_col_head_fmt)
            month_start_ym = (
                parse_category_to_year_month(str(month_payload["labels"][0]))
                if month_payload.get("labels")
                else None
            )

            thick_top_fmt = workbook.add_format(
                {
                    "border": 1,
                    "top": 2,
                    "valign": "vcenter",
                }
            )
            for i, label in enumerate(month_payload["labels"]):
                q = month_payload["quarters"][i]
                prev_q = month_payload["quarters"][i - 1] if i > 0 else None
                row = month_head_row + i
                thick_border = prev_q is not None and q != prev_q
                row_fmt_text = thick_top_fmt if thick_border else cell_fmt
                ws.write(row, 0, q, row_fmt_text)
                ws.write(row, 1, label, row_fmt_text)
                for j, s in enumerate(month_payload["series"]):
                    is_rub = aggregation_is_money(s.get("aggregation"))
                    if thick_border:
                        num_fmt = thick_top_rub if is_rub else thick_top_count
                    else:
                        num_fmt = fmt_rub if is_rub else fmt_count
                    m_formula = None
                    if c.get("linkToSourcePossible") and src_sheet_name:
                        x_idx_raw = c.get("xColumnIndex")
                        x_idx = (
                            int(x_idx_raw)
                            if isinstance(x_idx_raw, (int, float))
                            else src_map.get(norm_header(c.get("xHeader", "")))
                        )
                        y_idx_raw = s.get("sourceColumnIndex")
                        y_idx = (
                            int(y_idx_raw)
                            if isinstance(y_idx_raw, (int, float))
                            else src_map.get(norm_header(s.get("sourceHeader", "")))
                        )
                        distinct_idx_raw = s.get("countDistinctBySourceColumnIndex")
                        distinct_idx = (
                            int(distinct_idx_raw)
                            if isinstance(distinct_idx_raw, (int, float))
                            else None
                        )
                        report_end_excl_expr = report_through_exclusive_expr(
                            c.get("reportThroughYmd")
                        )
                        month_parsed = parse_category_to_year_month(str(label))
                        if x_idx is not None and month_parsed is not None:
                            src_from = 4
                            src_to = 3 + len(src_rows)
                            x_col = xl_col(x_idx)
                            agg = str(s.get("aggregation", "count"))
                            if agg == "count":
                                if distinct_idx is not None:
                                    id_col = xl_col(distinct_idx)
                                    if bool(s.get("cumulative")) and month_start_ym is not None:
                                        start_expr = (
                                            f"DATE({month_start_ym[0]},{month_start_ym[1]},1)"
                                        )
                                    else:
                                        start_expr = (
                                            f"DATE({month_parsed[0]},{month_parsed[1]},1)"
                                        )
                                    end_excl_expr = (
                                        f"EOMONTH(DATE({month_parsed[0]},{month_parsed[1]},1),0)+1"
                                    )
                                    m_formula = count_distinct_ids_between_dates_formula(
                                        src_sheet_name,
                                        x_col,
                                        id_col,
                                        start_expr,
                                        end_excl_expr,
                                        src_from,
                                        src_to,
                                        report_end_excl_expr,
                                    )
                                else:
                                    if bool(s.get("cumulative")) and month_start_ym is not None:
                                        m_formula = count_cumulative_through_month_from_chart_start_formula(
                                            src_sheet_name,
                                            x_col,
                                            month_start_ym[0],
                                            month_start_ym[1],
                                            month_parsed[0],
                                            month_parsed[1],
                                            src_from,
                                            src_to,
                                            report_end_excl_expr,
                                        )
                                    else:
                                        m_formula = count_month_formula(
                                            src_sheet_name,
                                            x_col,
                                            month_parsed[0],
                                            month_parsed[1],
                                            src_from,
                                            src_to,
                                            report_end_excl_expr,
                                        )
                            elif agg == "sum" and y_idx is not None:
                                y_col = xl_col(y_idx)
                                m_formula = sum_month_formula(
                                    src_sheet_name,
                                    x_col,
                                    y_col,
                                    month_parsed[0],
                                    month_parsed[1],
                                    src_from,
                                    src_to,
                                    report_end_excl_expr,
                                )
                    if m_formula:
                        ws.write_formula(
                            row,
                            j + 2,
                            m_formula,
                            num_fmt,
                            normalize_excel_number(s["values"][i]),
                        )
                    else:
                        write_formula_literal_number(
                            ws, row, j + 2, s["values"][i], num_fmt
                        )

            m_chart = workbook.add_chart({"type": "column"})
            m_cat = [
                sheet_name,
                month_head_row,
                1,
                month_head_row + len(month_payload["labels"]) - 1,
                1,
            ]
            for j, s in enumerate(month_payload["series"]):
                color = palette[j % len(palette)]
                m_series: dict[str, object] = {
                    "name": s.get("name", f"Ряд {j + 1}"),
                    "categories": m_cat,
                    "values": [
                        sheet_name,
                        month_head_row,
                        j + 2,
                        month_head_row + len(month_payload["labels"]) - 1,
                        j + 2,
                    ],
                    "fill": {"color": color},
                    "line": {"color": color},
                    "data_labels": {"value": True, "font": {"size": 9, "color": "#0070C0"}},
                }
                if j == 0:
                    m_series["gap"] = COLUMN_BAR_GAP
                m_chart.add_series(m_series)
            m_chart.set_title(
                {
                    "name": time_split_chart_title(c.get("title", "График"), "month"),
                    "name_font": COLUMN_CHART_TITLE_FONT,
                }
            )
            m_chart.set_legend({"position": "right", "font": {"size": 10, "bold": True, "color": "#0070C0"}})
            m_chart.set_x_axis(
                {
                    "num_font": dict(COLUMN_CATEGORY_AXIS_NUM_FONT),
                    "line": {"color": str(tstyle["axis"])},
                }
            )
            m_chart.set_y_axis(
                {
                    "num_font": {"size": 9, "color": "#0070C0"},
                    "major_gridlines": {"visible": True, "line": {"color": str(tstyle["grid"]), "dash_type": "dash"}},
                    "line": {"color": str(tstyle["axis"])},
                }
            )
            m_chart.set_style(int(tstyle["chart_style"]))
            m_chart.set_chartarea({"fill": {"color": "#FFFFFF"}, "border": {"none": True}})
            m_chart.set_plotarea(
                {
                    "fill": {"color": str(tstyle["plot_bg"])},
                    "border": {"color": "#D9E2F3", "width": 1},
                }
            )
            m_w, m_h = adaptive_chart_size(
                "column",
                len(month_payload["labels"]),
                len(month_payload["series"]),
                is_time_split_block=True,
            )
            m_chart.set_size({"width": m_w, "height": m_h})
            ws.insert_chart("F3", m_chart)

            # ---- Quarterly block ----
            q_title_row = month_head_row + len(month_payload["labels"]) + 4
            ws.merge_range(
                f"A{q_title_row}:D{q_title_row}",
                "Разбивка по кварталам",
                title_fmt,
            )
            q_head_row = q_title_row + 2
            q_headers = ["Квартал", *[s.get("name", "") for s in series]]
            ws.write_row(q_head_row - 1, 0, q_headers, table_col_head_fmt)
            quarter_to_month_rows: dict[str, list[int]] = {}
            for mi, qk in enumerate(month_payload["quarters"]):
                quarter_to_month_rows.setdefault(str(qk), []).append(month_head_row + mi)
            for i, q_label in enumerate(quarter_payload["labels"]):
                row = q_head_row + i
                ws.write(row, 0, q_label, cell_fmt)
                for j, s in enumerate(quarter_payload["series"]):
                    q_nf = fmt_rub if aggregation_is_money(s.get("aggregation")) else fmt_count
                    month_rows = quarter_to_month_rows.get(str(q_label), [])
                    q_formula = None
                    month_col = xl_col(j + 2)
                    if month_rows:
                        if bool(s.get("cumulative")):
                            last_row = month_rows[-1] + 1
                            q_formula = f"={month_col}{last_row}"
                        else:
                            first_row = month_rows[0] + 1
                            last_row = month_rows[-1] + 1
                            q_formula = f"=SUM({month_col}{first_row}:{month_col}{last_row})"
                    if q_formula:
                        ws.write_formula(
                            row,
                            j + 1,
                            q_formula,
                            q_nf,
                            normalize_excel_number(s["values"][i]),
                        )
                    else:
                        write_formula_literal_number(
                            ws, row, j + 1, s["values"][i], q_nf
                        )

            q_chart = workbook.add_chart({"type": "column"})
            q_cat = [
                sheet_name,
                q_head_row,
                0,
                q_head_row + len(quarter_payload["labels"]) - 1,
                0,
            ]
            for j, s in enumerate(quarter_payload["series"]):
                color = palette[j % len(palette)]
                q_series: dict[str, object] = {
                    "name": s.get("name", f"Ряд {j + 1}"),
                    "categories": q_cat,
                    "values": [
                        sheet_name,
                        q_head_row,
                        j + 1,
                        q_head_row + len(quarter_payload["labels"]) - 1,
                        j + 1,
                    ],
                    "fill": {"color": color},
                    "line": {"color": color},
                    "data_labels": {"value": True, "font": {"size": 9, "color": "#0070C0"}},
                }
                if j == 0:
                    q_series["gap"] = COLUMN_BAR_GAP
                q_chart.add_series(q_series)
            q_chart.set_title(
                {
                    "name": time_split_chart_title(c.get("title", "График"), "quarter"),
                    "name_font": COLUMN_CHART_TITLE_FONT,
                }
            )
            q_chart.set_legend({"position": "right", "font": {"size": 10, "bold": True, "color": "#0070C0"}})
            q_chart.set_x_axis(
                {
                    "num_font": dict(COLUMN_CATEGORY_AXIS_NUM_FONT),
                    "line": {"color": str(tstyle["axis"])},
                }
            )
            q_chart.set_y_axis(
                {
                    "num_font": {"size": 9, "color": "#0070C0"},
                    "major_gridlines": {"visible": True, "line": {"color": str(tstyle["grid"]), "dash_type": "dash"}},
                    "line": {"color": str(tstyle["axis"])},
                }
            )
            q_chart.set_style(int(tstyle["chart_style"]))
            q_chart.set_chartarea({"fill": {"color": "#FFFFFF"}, "border": {"none": True}})
            q_chart.set_plotarea(
                {
                    "fill": {"color": str(tstyle["plot_bg"])},
                    "border": {"color": "#D9E2F3", "width": 1},
                }
            )
            q_w, q_h = adaptive_chart_size(
                "column",
                len(quarter_payload["labels"]),
                len(quarter_payload["series"]),
                is_time_split_block=True,
            )
            q_chart.set_size({"width": q_w, "height": q_h})
            q_top_row = 3 + chart_height_to_rows(m_h) + 1
            ws.insert_chart(f"F{q_top_row}", q_chart)

            # ---- Yearly block ----
            y_title_row = q_head_row + len(quarter_payload["labels"]) + 4
            ws.merge_range(
                f"A{y_title_row}:D{y_title_row}",
                "По годам",
                title_fmt,
            )
            y_head_row = y_title_row + 2
            y_headers = ["Год", *[s.get("name", "") for s in series]]
            ws.write_row(y_head_row - 1, 0, y_headers, table_col_head_fmt)
            year_to_month_rows: dict[int, list[int]] = {}
            for mi, m_label in enumerate(month_payload["labels"]):
                ym = parse_category_to_year_month(str(m_label))
                if ym is None:
                    continue
                year_to_month_rows.setdefault(int(ym[0]), []).append(month_head_row + mi)
            for i, y_label in enumerate(year_payload["labels"]):
                row = y_head_row + i
                ws.write(row, 0, y_label, cell_fmt)
                for j, s in enumerate(year_payload["series"]):
                    y_nf = fmt_rub if aggregation_is_money(s.get("aggregation")) else fmt_count
                    year_rows = year_to_month_rows.get(int(str(y_label)), [])
                    y_formula = None
                    month_col = xl_col(j + 2)
                    if year_rows:
                        if bool(s.get("cumulative")):
                            last_row = year_rows[-1] + 1
                            y_formula = f"={month_col}{last_row}"
                        else:
                            first_row = year_rows[0] + 1
                            last_row = year_rows[-1] + 1
                            y_formula = f"=SUM({month_col}{first_row}:{month_col}{last_row})"
                    if y_formula:
                        ws.write_formula(
                            row,
                            j + 1,
                            y_formula,
                            y_nf,
                            normalize_excel_number(s["values"][i]),
                        )
                    else:
                        write_formula_literal_number(
                            ws, row, j + 1, s["values"][i], y_nf
                        )

            y_chart = workbook.add_chart({"type": "column"})
            y_cat = [
                sheet_name,
                y_head_row,
                0,
                y_head_row + len(year_payload["labels"]) - 1,
                0,
            ]
            for j, s in enumerate(year_payload["series"]):
                color = palette[j % len(palette)]
                y_series: dict[str, object] = {
                    "name": s.get("name", f"Ряд {j + 1}"),
                    "categories": y_cat,
                    "values": [
                        sheet_name,
                        y_head_row,
                        j + 1,
                        y_head_row + len(year_payload["labels"]) - 1,
                        j + 1,
                    ],
                    "fill": {"color": color},
                    "line": {"color": color},
                    "data_labels": {"value": True, "font": {"size": 9, "color": "#0070C0"}},
                }
                if j == 0:
                    y_series["gap"] = COLUMN_BAR_GAP
                y_chart.add_series(y_series)
            y_chart.set_title(
                {
                    "name": time_split_chart_title(c.get("title", "График"), "year"),
                    "name_font": COLUMN_CHART_TITLE_FONT,
                }
            )
            y_chart.set_legend({"position": "right", "font": {"size": 10, "bold": True, "color": "#0070C0"}})
            y_chart.set_x_axis(
                {
                    "num_font": dict(COLUMN_CATEGORY_AXIS_NUM_FONT),
                    "line": {"color": str(tstyle["axis"])},
                }
            )
            y_chart.set_y_axis(
                {
                    "num_font": {"size": 9, "color": "#0070C0"},
                    "major_gridlines": {"visible": True, "line": {"color": str(tstyle["grid"]), "dash_type": "dash"}},
                    "line": {"color": str(tstyle["axis"])},
                }
            )
            y_chart.set_style(int(tstyle["chart_style"]))
            y_chart.set_chartarea({"fill": {"color": "#FFFFFF"}, "border": {"none": True}})
            y_chart.set_plotarea(
                {
                    "fill": {"color": str(tstyle["plot_bg"])},
                    "border": {"color": "#D9E2F3", "width": 1},
                }
            )
            y_w, y_h = adaptive_chart_size(
                "column",
                len(year_payload["labels"]),
                len(year_payload["series"]),
                is_time_split_block=True,
            )
            y_chart.set_size({"width": y_w, "height": y_h})
            y_top_row = q_top_row + chart_height_to_rows(q_h) + 1
            ws.insert_chart(f"F{y_top_row}", y_chart)
        elif not has_time_split:
            w, h = adaptive_chart_size(chart_kind, len(categories), len(series))
            chart.set_size({"width": w, "height": h})
            ws.insert_chart("F3", chart)

    workbook.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
