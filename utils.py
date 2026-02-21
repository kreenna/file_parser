import re

import pandas as pd


def parse_price(value) -> float | None:
    """Нормализация цены в корректный формат."""

    if value is None:
        return None

    # убираем все пробелы
    s = str(value).replace(" ", "").replace(u"\u00a0", "")

    # заменяем запятые на точки
    s = s.replace(",", ".")

    # убираем точки-разделители тысяч (все кроме последней)
    parts = s.split(".")
    if len(parts) > 1:
        # последняя часть — дробная, остальное — целое
        integer_part = "".join(parts[:-1])
        decimal_part = parts[-1]
        s = integer_part + "." + decimal_part if decimal_part else integer_part

    # regex находим число без лишних данных
    m = re.search(r"[-+]?\d+(?:\.\d{0,2})?", s)
    if not m:
        return None

    try:
        return float(m.group(0))
    except ValueError:
        return None


SIMPLE_HEADER_KEYWORDS = {
    "name": ["наименован", "назван", "товар", "позици", "description", "item", "product"],
    "sku": ["артикул", "арт", "код", "sku", "article"],
    "unit": ["ед", "ед. изм", "единиц", "unit", "шт", "кг", "м"],
    "price": ["цена", "стоим", "руб", "₽", "price"],
    "manufacturer": ["производ", "бренд", "vendor", "manufacturer", "поставщик"],
    "notes": ["примечан", "коммент", "notes", "описан", "description"],
}

def simple_match_header(row) -> dict[str, int]:
    """
    Очень простой матчинг заголовка по подстрокам.
    Возвращает col_map: {\"name\": idx, ...}.
    """
    col_map: dict[str, int] = {}
    for idx, val in enumerate(row):
        if val is None:
            continue
        text = str(val).strip().lower()
        if not text or len(text) > 100:
            continue
        for col_type, keywords in SIMPLE_HEADER_KEYWORDS.items():
            if col_type in col_map:
                continue
            for kw in keywords:
                if kw in text:
                    col_map[col_type] = idx
                    break
    return col_map


def normalize_unit(unit: str) -> str:
    """Нормализация единицы измерения."""

    unit = unit.lower().strip()
    if not unit:
        return ""
    mapping = {
        "шт": ["шт", "штука", "штук"],
        "м": ["м", "метр", "метра", "метров"],
        "м2": ["м2", "м^2", "кв.м", "квм"],
        "кг": ["кг", "килограмм", "килограмма", "килограммoв"],
        "л": ["л", "литр", "литра", "литров"],
    }
    for norm, variants in mapping.items():
        if any(variant in unit for variant in variants):
            return norm
    return unit


def is_junk_row(text: str) -> bool:
    """Проверка строки на валидность."""

    t = text.lower().strip()
    if not t:
        return True
    junk_words = ["итого", "всего", "номер", "№", "сумма", "огрн", "инн", "кпп", "окпо", "ооо", "г.", "ул.", "д.",
                  "стр."]
    if len(t) < 3:
        return True
    if sum(ch.isalpha() for ch in t) < 2 and sum(ch.isdigit() for ch in t) > 3:
        return True
    return any(w in t for w in junk_words)


def extract_vendor(name: str) -> str:
    """Получение поставщика."""

    # простейшая эвристика: заглавные латинские слова 2–15 символов
    tokens = re.split(r"[,\s;]+", name)
    candidates = [
        t for t in tokens
        if 2 <= len(t) <= 15 and t.isupper() and any("A" <= ch <= "Z" for ch in t)
    ]
    return candidates[0] if candidates else ""


def read_csv_iter(path: str, sep: str):
    """Итератор для CSV, возвращает df (для совместимости с multi-sheet логикой)."""
    yield pd.read_csv(path, sep=sep, dtype=str)


def split_into_tables(all_rows: list[list], min_non_empty_cells: int = 2) -> list[tuple[int, int]]:
    """
    Возвращает список (start_row, end_row) для блоков таблиц.
    Разделитель: 2+ подряд строк, где < min_non_empty_cells непустых ячеек.
    Индексы 0-based по all_rows.
    """
    tables = []
    cur_start = None
    empty_count = 0

    for row_id, row in enumerate(all_rows):
        non_empty = sum(1 for value in row if value not in (None, "") and str(value).strip())
        if non_empty >= min_non_empty_cells:
            if cur_start is None:
                cur_start = row_id
            empty_count = 0
        else:
            empty_count += 1
            if empty_count >= 2 and cur_start is not None:
                end = row_id - empty_count
                if end - cur_start >= 2:
                    tables.append((cur_start, end))
                cur_start = None

    if cur_start is not None and len(all_rows) - 1 - cur_start >= 2:
        tables.append((cur_start, len(all_rows) - 1))

    if not tables:
        tables = [(0, len(all_rows) - 1)]

    return tables