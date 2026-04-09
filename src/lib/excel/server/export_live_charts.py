from __future__ import annotations

import json
import re
import sys
from datetime import datetime
from pathlib import Path

import xlsxwriter

ENTITY_PALETTES: dict[str, list[str]] = {
    "leads": ["#0EA5E9", "#38BDF8", "#0284C7", "#7DD3FC", "#0369A1"],
    "deals": ["#8B5CF6", "#A78BFA", "#7C3AED", "#C4B5FD", "#6D28D9"],
    "contacts": ["#14B8A6", "#2DD4BF", "#0F766E", "#5EEAD4", "#0D9488"],
    "companies": [
        "#2563EB",
        "#06B6D4",
        "#16A34A",
        "#C026D3",
        "#7C3AED",
        "#0891B2",
        "#DC2626",
        "#DB2777",
        "#0D9488",
        "#4F46E5",
        "#E11D48",
        "#84CC16",
    ],
    "quotes": ["#E11D48", "#F43F5E", "#BE123C", "#FB7185", "#9F1239"],
}

THEME_STYLE: dict[str, dict[str, object]] = {
    "classic": {
        "chart_style": 2,
        "pie_style": 10,
        "plot_bg": "#F8FAFC",
        "grid": "#E2E8F0",
        "axis": "#CBD5E1",
        "line_width": 2.0,
        "marker": 6,
    },
    "vivid": {
        "chart_style": 18,
        "pie_style": 26,
        "plot_bg": "#F1F5F9",
        "grid": "#CBD5E1",
        "axis": "#94A3B8",
        "line_width": 2.6,
        "marker": 7,
    },
    "minimal": {
        "chart_style": 1,
        "pie_style": 2,
        "plot_bg": "#FFFFFF",
        "grid": "#EEF2F7",
        "axis": "#D7DEE8",
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

def palette_for_entity(entity_id: str) -> list[str]:
    return ENTITY_PALETTES.get(entity_id, ENTITY_PALETTES["leads"])

def theme_style(theme: str) -> dict[str, object]:
    return THEME_STYLE.get(theme, THEME_STYLE["classic"])

def parse_category_to_year_month(label: str) -> tuple[int, int] | None:
    s = str(label).strip().lower()
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

def series_data_labels(chart_kind: str, chart_title: str = "") -> dict[str, object]:
    if chart_kind == "pie":
        title = str(chart_title).lower()
        if "типы компаний" in title:
            return {
                "value": True,
                "percentage": True,
                "position": "outside_end",
                "leader_lines": True,
                "separator": "\n",
                "font": {"size": 9, "bold": True, "color": "#0F172A"},
            }
        if "распределение по ответственным" in title:
            return {
                "category": True,
                "value": True,
                "percentage": True,
                "position": "outside_end",
                "leader_lines": True,
                "separator": "\n",
                "font": {"size": 9, "bold": True, "color": "#0F172A"},
            }
        return {
            "percentage": True,
            "position": "outside_end",
            "leader_lines": True,
            "separator": "\n",
            "num_format": "0%",
            "font": {"size": 9, "bold": True, "color": "#0F172A"},
        }
    return {
        "value": True,
        "font": {"size": 8, "color": "#334155"},
    }

def normalize_excel_number(value: object) -> float | int:
    n = float(value or 0)
    if abs(n - round(n)) < 1e-3:
        return int(round(n))
    return n

def xl_col(col_index_zero_based: int) -> str:
    n = col_index_zero_based + 1
    out = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out

def norm_header(h: object) -> str:
    return str(h or "").strip().lower()

def count_formula_day_bucket(src_sheet: str, x_col: str, x_row: int, src_from: int, src_to: int) -> str:
    x_range = f"'{src_sheet}'!${x_col}${src_from}:${x_col}${src_to}"
    x_cell = f"$A${x_row}"
    # Счёт по календарному дню (без времени).
    return f'=COUNTIFS({x_range},">="&DATEVALUE({x_cell}),{x_range},"<"&DATEVALUE({x_cell})+1)'

def count_formula_exact(src_sheet: str, x_col: str, x_row: int, src_from: int, src_to: int) -> str:
    x_range = f"'{src_sheet}'!${x_col}${src_from}:${x_col}${src_to}"
    x_cell = f"$A${x_row}"
    return f'=COUNTIFS({x_range},{x_cell})'

def count_month_formula(src_sheet: str, x_col: str, year: int, month: int, src_from: int, src_to: int) -> str:
    x_range = f"'{src_sheet}'!${x_col}${src_from}:${x_col}${src_to}"
    start = f"DATE({year},{month},1)"
    end = f"EOMONTH({start},0)+1"
    return f'=COUNTIFS({x_range},">="&{start},{x_range},"<"&{end})'


def count_cumulative_through_month_from_chart_start_formula(
    src_sheet: str,
    x_col: str,
    start_year: int,
    start_month: int,
    end_year: int,
    end_month: int,
    src_from: int,
    src_to: int,
) -> str:
    """Нарастающий итог по месяцам: все записи с датой X от 1-го числа первого месяца оси до конца текущего месяца."""
    x_range = f"'{src_sheet}'!${x_col}${src_from}:${x_col}${src_to}"
    lower = f"DATE({start_year},{start_month},1)"
    end_excl = f"EOMONTH(DATE({end_year},{end_month},1),0)+1"
    return f'=COUNTIFS({x_range},">="&{lower},{x_range},"<"&{end_excl})'

def sum_formula(
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
    return f'=SUMIFS({y_range},{x_range},">="&DATEVALUE({x_cell}),{x_range},"<"&DATEVALUE({x_cell})+1)'

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
) -> tuple[list[dict], list[dict]] | tuple[None, None]:
    parsed = [parse_category_to_year_month(c) for c in categories]
    if not any(parsed):
        return None, None
    month_keys = sorted({(y, m) for y, m in parsed if y and m})
    if not month_keys:
        return None, None

    month_index = {k: i for i, k in enumerate(month_keys)}
    month_labels = [month_label_ru(y, m) for y, m in month_keys]
    month_quarters = [quarter_label(y, m) for y, m in month_keys]

    month_series: list[dict] = []
    series_cumulative: list[bool] = []
    for s in series:
        cumulative = bool(s.get("cumulative"))
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
        month_series.append({"name": s.get("name", ""), "values": values})
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
        quarter_series.append({"name": s.get("name", ""), "values": q_values})

    month_payload = {
        "labels": month_labels,
        "quarters": month_quarters,
        "series": month_series,
    }
    quarter_payload = {
        "labels": quarter_keys,
        "series": quarter_series,
    }
    return month_payload, quarter_payload


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
            "bg_color": "#0EA5E9",
            "font_color": "#FFFFFF",
            "align": "left",
            "valign": "vcenter",
        }
    )
    head_fmt = workbook.add_format(
        {
            "bold": True,
            "bg_color": "#38BDF8",
            "font_color": "#FFFFFF",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
        }
    )
    cell_fmt = workbook.add_format({"border": 1, "valign": "vcenter"})
    cell_date_fmt = workbook.add_format({"border": 1, "valign": "vcenter", "num_format": "dd.mm.yyyy"})
    cell_num_fmt = workbook.add_format({"border": 1, "valign": "vcenter", "num_format": "0"})
    axis_num_fmt = workbook.add_format({"num_format": "0"})

    # Summary sheet.
    summary_name = sanitize_sheet_name("Общая информация", used_names)
    ws_sum = workbook.add_worksheet(summary_name)
    ws_sum.set_column("A:A", 28)
    ws_sum.set_column("B:B", 48)
    ws_sum.set_column("C:C", 64)
    ws_sum.merge_range("A1:C1", "Общая информация", title_fmt)

    meta_rows = [
        ("Дата формирования", summary.get("createdAt", "")),
        ("Исходный файл", summary.get("sourceFileName", "")),
        ("Строк на листе", summary.get("rowsTotal", 0)),
        ("Строк после фильтров", summary.get("rowsFiltered", 0)),
        ("Графиков в файле", len(charts)),
    ]
    row = 2
    for k, v in meta_rows:
        ws_sum.write(row, 0, k, cell_fmt)
        ws_sum.write(row, 1, v, cell_fmt)
        row += 1

    row += 1
    ws_sum.write_row(row, 0, ["Раздел", "График", "Статус"], head_fmt)
    row += 1
    for c in charts:
        ws_sum.write_row(
            row,
            0,
            [c.get("blockLabel", ""), c.get("title", ""), "Готово"],
            cell_fmt,
        )
        row += 1

    # Source data sheet.
    src_headers = source_data.get("headers", []) or []
    src_rows = source_data.get("rows", []) or []
    src_sheet_name = ""
    src_map: dict[str, int] = {}
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
                    ws_src.write_number(row_idx, j, normalize_excel_number(raw), cell_num_fmt)
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
        categories = [normalize_category_label(v) for v in (c.get("categories", []) or [])]
        series = c.get("series", [])
        month_payload, quarter_payload = aggregate_time_buckets(categories, series)
        has_time_split = month_payload is not None and quarter_payload is not None

        base_sheet_name = c.get("sheetName") or c.get("title") or "График"
        sheet_name_raw = f"{base_sheet_name} [M+Q]" if has_time_split else base_sheet_name
        sheet_name = sanitize_sheet_name(sheet_name_raw, used_names)
        ws = workbook.add_worksheet(sheet_name)
        ws.set_column("A:A", 28)
        ws.set_column("B:Z", 16, axis_num_fmt)
        ws.merge_range("A1:D1", c.get("title", "График"), title_fmt)

        headers = [c.get("xHeader", "Категория"), *[s.get("name", "") for s in c.get("series", [])]]
        ws.write_row(2, 0, headers, head_fmt)

        chart_month_start = (
            parse_category_to_year_month(str(categories[0])) if categories else None
        )

        for i, category in enumerate(categories):
            ws.write(3 + i, 0, category, cell_fmt)
            for j, s in enumerate(series):
                values = s.get("values", [])
                value = values[i] if i < len(values) else 0
                wrote_formula = False
                if c.get("linkToSourcePossible") and src_sheet_name:
                    x_idx = src_map.get(norm_header(c.get("xHeader", "")))
                    y_idx = src_map.get(norm_header(s.get("sourceHeader", "")))
                    if x_idx is not None:
                        src_from = 4
                        src_to = 3 + len(src_rows)
                        x_col = xl_col(x_idx)
                        out_row = 4 + i
                        agg = str(s.get("aggregation", "count"))
                        gran = str(c.get("dateGranularity") or "")
                        x_is_date = bool(c.get("xIsDate"))
                        formula = None
                        month_parsed: tuple[int, int] | None = None
                        if agg == "count":
                            if x_is_date and gran == "month":
                                month_parsed = parse_category_to_year_month(str(category))
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
                                    )
                                if formula is None:
                                    if x_is_date:
                                        formula = count_formula_day_bucket(
                                            src_sheet_name, x_col, out_row, src_from, src_to
                                        )
                                    else:
                                        formula = count_formula_exact(
                                            src_sheet_name, x_col, out_row, src_from, src_to
                                        )
                                if s.get("cumulative") and i > 0:
                                    prev_excel_row = 3 + i
                                    val_col = xl_col(j + 1)
                                    if month_parsed is not None and gran == "month" and x_is_date:
                                        cur = count_month_formula(
                                            src_sheet_name,
                                            x_col,
                                            month_parsed[0],
                                            month_parsed[1],
                                            src_from,
                                            src_to,
                                        )[1:]
                                    elif x_is_date:
                                        cur = count_formula_day_bucket(
                                            src_sheet_name, x_col, out_row, src_from, src_to
                                        )[1:]
                                    else:
                                        cur = count_formula_exact(
                                            src_sheet_name, x_col, out_row, src_from, src_to
                                        )[1:]
                                    formula = f"={val_col}{prev_excel_row}+{cur}"
                        elif agg == "sum" and y_idx is not None:
                            y_col = xl_col(y_idx)
                            formula = sum_formula(src_sheet_name, x_col, y_col, out_row, src_from, src_to)
                        if formula:
                            ws.write_formula(
                                3 + i,
                                j + 1,
                                formula,
                                cell_num_fmt,
                                normalize_excel_number(value),
                            )
                            wrote_formula = True
                if not wrote_formula:
                    ws.write_number(3 + i, j + 1, normalize_excel_number(value), cell_num_fmt)

        chart_kind = chart_type_for_xlsxwriter(c.get("chartType", "column"))
        palette = palette_for_entity(str(c.get("entityId", "")))
        chart = workbook.add_chart({"type": chart_kind})
        cat_range = [sheet_name, 3, 0, 3 + max(len(categories) - 1, 0), 0]
        for j, s in enumerate(series):
            val_range = [sheet_name, 3, j + 1, 3 + max(len(categories) - 1, 0), j + 1]
            color = palette[j % len(palette)]
            series_opts = {
                "name": s.get("name", f"Ряд {j + 1}"),
                "categories": cat_range,
                "values": val_range,
                "fill": {"color": color},
                "line": {"color": color, "width": float(tstyle["line_width"])},
                "data_labels": series_data_labels(chart_kind, str(c.get("title", ""))),
                "marker": {
                    "type": "circle",
                    "size": int(tstyle["marker"]),
                    "border": {"color": color},
                    "fill": {"color": color},
                },
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
        chart.set_title(
            {
                "name": c.get("title", "График"),
                "name_font": {"name": "Calibri", "size": 13, "bold": True, "color": "#0F172A"},
            }
        )
        chart.set_legend({"position": "right", "font": {"size": 10, "bold": True, "color": "#334155"}})
        chart.set_chartarea({"fill": {"color": "#FFFFFF"}, "border": {"none": True}})
        chart.set_plotarea(
            {
                "fill": {"color": str(tstyle["plot_bg"])},
                "border": {"color": "#E2E8F0", "width": 1},
            }
        )
        if chart_kind != "pie":
            chart.set_x_axis(
                {
                    "name": c.get("xHeader", ""),
                    "name_font": {"size": 10, "color": "#334155"},
                    "num_font": {"size": 9, "color": "#475569"},
                    "line": {"color": str(tstyle["axis"])},
                }
            )
            chart.set_y_axis(
                {
                    "name_font": {"size": 10, "color": "#334155"},
                    "num_font": {"size": 9, "color": "#475569"},
                    "major_gridlines": {
                        "visible": True,
                        "line": {"color": str(tstyle["grid"]), "dash_type": "dash"},
                    },
                    "line": {"color": str(tstyle["axis"])},
                }
            )
        if chart_kind == "pie":
            chart.set_style(int(tstyle["pie_style"]))
            chart.set_legend({"position": "right", "font": {"size": 10, "bold": True, "color": "#334155"}})
        else:
            chart.set_style(int(tstyle["chart_style"]))
        if month_payload and quarter_payload:
            chart.set_size({"width": 760, "height": 340})
            ws.insert_chart("F3", chart)
            # ---- Monthly block ----
            # Основная таблица занимает строки: заголовок в 3-й и данные с 4-й.
            # Чтобы не перезаписывать исходные категории/значения, месячный блок
            # начинаем строго после основной таблицы.
            month_title_row = 6 + len(categories)
            ws.merge_range(
                f"A{month_title_row}:D{month_title_row}",
                "Разбивка по месяцам (границы кварталов выделены)",
                title_fmt,
            )
            month_head_row = month_title_row + 2
            month_headers = ["Квартал", "Месяц", *[s.get("name", "") for s in series]]
            ws.write_row(month_head_row - 1, 0, month_headers, head_fmt)

            thick_top_fmt = workbook.add_format(
                {
                    "border": 1,
                    "top": 2,
                    "valign": "vcenter",
                    "num_format": "0",
                }
            )
            for i, label in enumerate(month_payload["labels"]):
                q = month_payload["quarters"][i]
                prev_q = month_payload["quarters"][i - 1] if i > 0 else None
                row = month_head_row + i
                row_fmt_text = thick_top_fmt if prev_q is not None and q != prev_q else cell_fmt
                row_fmt_num = thick_top_fmt if prev_q is not None and q != prev_q else cell_num_fmt
                ws.write(row, 0, q, row_fmt_text)
                ws.write(row, 1, label, row_fmt_text)
                for j, s in enumerate(month_payload["series"]):
                    ws.write_number(row, j + 2, normalize_excel_number(s["values"][i]), row_fmt_num)

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
                m_chart.add_series(
                    {
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
                        "data_labels": {"value": True, "font": {"size": 8, "color": "#334155"}},
                    }
                )
            m_chart.set_title({"name": f"{c.get('title', 'График')} — по месяцам"})
            m_chart.set_legend({"position": "right", "font": {"size": 10, "bold": True, "color": "#334155"}})
            m_chart.set_x_axis({"num_font": {"size": 9, "color": "#475569"}, "line": {"color": str(tstyle["axis"])}})
            m_chart.set_y_axis(
                {
                    "num_font": {"size": 9, "color": "#475569"},
                    "major_gridlines": {"visible": True, "line": {"color": str(tstyle["grid"]), "dash_type": "dash"}},
                    "line": {"color": str(tstyle["axis"])},
                }
            )
            m_chart.set_style(int(tstyle["chart_style"]))
            m_chart.set_size({"width": 760, "height": 320})
            ws.insert_chart("F22", m_chart)

            # ---- Quarterly block ----
            q_title_row = month_head_row + len(month_payload["labels"]) + 4
            ws.merge_range(
                f"A{q_title_row}:D{q_title_row}",
                "Разбивка по кварталам",
                title_fmt,
            )
            q_head_row = q_title_row + 2
            q_headers = ["Квартал", *[s.get("name", "") for s in series]]
            ws.write_row(q_head_row - 1, 0, q_headers, head_fmt)
            for i, q_label in enumerate(quarter_payload["labels"]):
                row = q_head_row + i
                ws.write(row, 0, q_label, cell_fmt)
                for j, s in enumerate(quarter_payload["series"]):
                    ws.write_number(row, j + 1, normalize_excel_number(s["values"][i]), cell_num_fmt)

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
                q_chart.add_series(
                    {
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
                        "data_labels": {"value": True, "font": {"size": 8, "color": "#334155"}},
                    }
                )
            q_chart.set_title({"name": f"{c.get('title', 'График')} — по кварталам"})
            q_chart.set_legend({"position": "right", "font": {"size": 10, "bold": True, "color": "#334155"}})
            q_chart.set_x_axis({"num_font": {"size": 9, "color": "#475569"}, "line": {"color": str(tstyle["axis"])}})
            q_chart.set_y_axis(
                {
                    "num_font": {"size": 9, "color": "#475569"},
                    "major_gridlines": {"visible": True, "line": {"color": str(tstyle["grid"]), "dash_type": "dash"}},
                    "line": {"color": str(tstyle["axis"])},
                }
            )
            q_chart.set_style(int(tstyle["chart_style"]))
            q_chart.set_size({"width": 760, "height": 300})
            ws.insert_chart("F40", q_chart)
        else:
            chart.set_size({"width": 760, "height": 420})
            ws.insert_chart("F3", chart)

    workbook.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
