import re
from pathlib import Path
from typing import Optional, List, Tuple, Any

import numpy as np
import pandas as pd
from rapidfuzz import fuzz, process

COLUMN_KEYWORDS = {
    "name": ["наименование ", "название ", "product", "item", " имя", "название товара",
             "номенклатура "],
    "sku": ["артикул ", " арт.", "код ", "sku", "article", "тип ", "марка ", "марки ", "обозначение ",
            " модель"],
    "unit": ["ед. измерения", "единица измерения", "unit", "ед. изм"],
    "quantity": ["количество", "кол-во", "кол.  "],
    # unit price (generic)
    "price_unit": [" цена", "price", "расценка ", "РРЦ "],
    # explicit “без НДС”
    "price_base": [" цена без ндс", "без ндс", "без НДС", "цена за ед без ндс"],
    # explicit “с НДС”
    "price_vat": [" цена с ндс", "с НДС", "вкл. НДС", "цена с учетом ндс", "базовая цена с ндс"],
    # totals
    "total_no_vat": ["стоимость без ндс", "сумма без ндс", "итого без ндс"],
    "total_vat": ["стоимость с ндс", "сумма с ндс", "итого с ндс"],
    "total": ["стоимость ", "сумма ", "итого ", "всего ", "total ", "amount "],
    "manufacturer": ["производитель ", "бренд", "vendor", "manufacturer", "поставщик "],
    "notes": ["примечани", "notes", "описание", "комментари", "характеристик"]
}

SIMPLE_HEADER_KEYWORDS = {
    "name": ["наименован", "назван", "товар", "позици", "description", "item", "product"],
    "sku": ["артикул", "арт", "код", "sku", "article"],
    "unit": ["ед", "ед. изм", "ед.изм", "единиц", "unit", "изм"],
    "quantity": ["количество", "кол-во", "кол."],
    "price_unit": ["цена", "стоимость", "price", "расценк", "РРЦ", "МРЦ"],
    # explicit “без НДС”
    "price_base": ["цена без ндс", "без ндс", "без НДС", "цена за ед без ндс"],
    # explicit “с НДС”
    "price_vat": ["цена с ндс", "с НДС", "вкл. НДС", "цена с учетом ндс"],
    # totals
    "total_no_vat": ["стоимость без ндс", "сумма без ндс", "итого без ндс"],
    "total_vat": ["стоимость с ндс", "сумма с ндс", "итого с ндс"],
    "total": ["стоимость", "сумма", "итого", "всего", "total", "amount"],
    "manufacturer": ["производ", "бренд", "vendor", "manufacturer", "поставщик"],
    "notes": ["примечан", "коммент", "notes", "описан", "description"],
}


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

    # c regex находим число без лишних данных
    m = re.search(r"[-+]?\d+(?:\.\d{0,2})?", s)
    if not m:
        return None

    try:
        return float(m.group(0))
    except ValueError:
        return None


def simple_match_header(row) -> dict[str, int]:
    """Простой матчинг заголовка по подстрокам."""

    col_map: dict[str, int] = {}

    for idx, value in enumerate(row):
        # проходимся по всем ячейкам строки

        if value is None:
            continue

        text = str(value).strip().lower()
        if not text or len(text) > 100:
            continue

        for col_type, keywords in SIMPLE_HEADER_KEYWORDS.items():
            if col_type in col_map:
                continue

            for keyword in keywords:
                if keyword in text:
                    col_map[col_type] = idx
                    break

    return col_map


def normalize_unit(unit: str) -> str:
    """Нормализация единицы измерения."""

    unit = unit.lower().strip()
    if not unit:
        return "шт"

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
    return ""


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


def read_csv_iter(path: str, sep: str):
    """Итератор для CSV, возвращает df (для совместимости с multi-sheet логикой)."""
    yield pd.read_csv(path, sep=sep, dtype=str)


def collect_prices_for_row(row, col_map: dict) -> dict:
    """
    Сбор всех найденных ценовых значений по типам:
    price_base, price_unit, price_vat, total_no_vat, total_vat, total.
    """

    price_cols = ["price_base", "price_unit", "price_vat", "total_no_vat", "total_vat", "total"]
    prices = {}

    for key in price_cols:
        idx: int = col_map.get(key)

        if idx is None or idx >= len(row):
            continue

        value = row[idx] if isinstance(row, tuple) else row.iloc[idx]
        price = parse_price(value) if value else 0.0

        if price is not None and price > 0:
            prices[key] = price

    return prices


def pick_best_price(prices: dict, quantity: float | None = None, vat_rate: Optional[float] = 0.2) -> float:
    """
    Выбирает лучшую ЦЕНУ ЗА ЕДИНИЦУ без НДС.
    Приоритет:
      1) price_base
      2) price_unit
      3) price_vat / 1.2
      4) total_no_vat / quantity
      5) total / quantity
      6) total_vat / quantity / 1.2
    """

    if "price_base" in prices:
        return prices["price_base"]

    if "price_unit" in prices:
        return prices["price_unit"]

    if "price_vat" in prices:
        return prices["price_vat"] / (1 + vat_rate)

    if quantity and quantity > 0:
        if "total_no_vat" in prices:
            return prices["total_no_vat"] / quantity
        if "total" in prices:
            return prices["total"] / quantity
        if "total_vat" in prices:
            return prices["total_vat"] / quantity / (1 + vat_rate)

    # fallback: any total if quantity unknown
    for key in ["total_no_vat", "total", "total_vat"]:
        if key in prices:
            return prices[key]

    return 0.0


def get_output_json_path(input_file: str) -> Path:
    """Меняет формат файла на JSON."""
    in_path = Path(input_file)
    return in_path.with_suffix(".json")


def detect_table_regions(df: pd.DataFrame,
                         min_rows: int = 3,
                         min_cols: int = 2,
                         max_empty_row_gap: int = 2) -> List[Tuple[int, int, int, int]]:
    """
    Поиск прямоугольных регионов таблиц на всем листе.

    Возвращает список (row_start, row_end, col_start, col_end) углов региона
    в df для каждой таблицы. Использует пустые строки и столбцы как разделители.
    """

    # True where cell has non‑empty, non‑whitespace content
    mask = df.map(lambda x: isinstance(x, str) and x.strip() != "")
    mask |= df.map(lambda x: not isinstance(x, str) and pd.notna(x))
    has_value = mask.values.copy()  # shape [n_rows, n_cols]

    n_rows, n_cols = has_value.shape

    # "склеиваем" короткие вертикальные разрывы
    for col in range(n_cols):
        row = 0
        while row < n_rows:
            # пропускаем пустые снизу/сверху
            if has_value[row, col]:
                row += 1
                continue

            start = row
            while row < n_rows and not has_value[row, col]:
                row += 1
            end = row  # [start, end) — подряд пустые

            gap_len = end - start
            # если по обе стороны есть данные и разрыв маленький — считаем его частью таблицы
            if 0 < gap_len <= max_empty_row_gap:
                if start > 0 and end < n_rows:
                    if has_value[start - 1, col] and has_value[end, col]:
                        has_value[start:end, col] = True

    visited = np.zeros_like(has_value, dtype=bool)
    regions: List[Tuple[int, int, int, int]] = []

    def bfs(r0: int, c0: int):
        """Flood‑fill для получения всех блоков с данными."""
        stack = [(r0, c0)]
        visited[r0, c0] = True
        row_min = row_max = r0
        col_min = col_max = c0

        while stack:
            r, c = stack.pop()

            # обновляем границы
            row_min = min(row_min, r)
            row_max = max(row_max, r)
            col_min = min(col_min, c)
            col_max = max(col_max, c)

            # 4‑соседство
            for dr, dc in ((1, 0), (-1, 0), (0, 1), (0, -1)):
                rr, cc = r + dr, c + dc
                if 0 <= rr < n_rows and 0 <= cc < n_cols:
                    if has_value[rr, cc] and not visited[rr, cc]:
                        visited[rr, cc] = True
                        stack.append((rr, cc))

        return row_min, row_max, col_min, col_max

    for r in range(n_rows):
        for c in range(n_cols):
            if has_value[r, c] and not visited[r, c]:
                r0, r1, c0, c1 = bfs(r, c)

                # "сырой" bounding box
                sub = has_value[r0:r1 + 1, c0:c1 + 1]

                # drop empty outer rows
                while sub.shape[0] > 0 and not sub[0].any():
                    r0 += 1
                    sub = has_value[r0:r1 + 1, c0:c1 + 1]
                while sub.shape[0] > 0 and not sub[-1].any():
                    r1 -= 1
                    sub = has_value[r0:r1 + 1, c0:c1 + 1]
                # drop empty outer cols
                while sub.shape[1] > 0 and not sub[:, 0].any():
                    c0 += 1
                    sub = has_value[r0:r1 + 1, c0:c1 + 1]
                while sub.shape[1] > 0 and not sub[:, -1].any():
                    c1 -= 1
                    sub = has_value[r0:r1 + 1, c0:c1 + 1]

                # фильтр по минимальному размеру
                if (r1 - r0 + 1) >= min_rows and (c1 - c0 + 1) >= min_cols:
                    regions.append((r0, r1, c0, c1))

    return regions


def score_header_row(col_map: dict[str, int]) -> int:
    """
    Оценка строки в соответствии с headers, которые были найдены.
    """
    score = 0

    if "name" in col_map:
        score += 30
    if "sku" in col_map:
        score += 30

    for key in ["unit", "quantity", "manufacturer", "notes"]:
        if key in col_map:
            score += 10

    for key in ["price_unit", "price_base", "price_vat", "total_no_vat", "total_vat", "total"]:
        if key in col_map:
            score += 6

    score += len(col_map) * 2
    return score


def merge_missing(base_map: dict[str, int], candidate_map: dict[str, int]) -> dict[str, int]:
    """
    Включение отсутствующих headers из строк после основной строки с headers (multi-headers processing).
    """
    merged_map = dict(base_map)
    for key, value in candidate_map.items():
        if key not in merged_map:
            merged_map[key] = value
    return merged_map


def merge_keep_old(base: dict[str, int], candidate_map: dict[str, int]) -> dict[str, int]:
    """
    Переписывание старых headers на более актуальные без удаления оставшихся.
    """
    merged_map = dict(base)
    for key, idx in candidate_map.items():
        merged_map[key] = idx
    return merged_map


def safe_lower_processor(value):
    """Обработчик для приведения содержимого сроки к нижнему регистру без ошибок."""
    return value.lower() if isinstance(value, str) else value


def match_keyword_headers(headers_list: list, col_map: dict[str, int] = None) -> dict[str, int]:
    """
    Проверка названий из строки на соответствие с headers столбцов.
    """

    col_map = dict(col_map or {})

    for col_type, keywords in COLUMN_KEYWORDS.items():
        # проходимся по каждому типу и ключевому слову

        best_score = 0
        best_idx = None

        for keyword in keywords:
            # проверяем совпадения по каждому слову и определяем лучшее
            match = process.extractOne(keyword, headers_list, processor=safe_lower_processor,
                                       scorer=fuzz.partial_ratio)

            if match and match[1] > best_score:
                best_score = match[1]
                best_idx = match[2]

        if best_idx is not None and best_score >= 80:
            col_map[col_type] = best_idx

    return col_map


CORE_KEYS = {"name", "sku"}
OPTIONAL_KEYS = {"unit", "quantity", "manufacturer", "notes", "price_unit", "price_base", "price_vat", "total_no_vat",
                 "total_vat", "total"}


def find_col_map_header_row(df: pd.DataFrame, lookahead: int = 3) -> tuple[dict[str, int], int | Any]:
    """
    Поиск строки с headers.
    """
    # ищем строку шапки: первая строка, где есть >=2 ячейки с буквами
    header_row: int = 0

    # non-empty rows containing text
    while header_row < len(df) and df.iloc[header_row].astype(str).str.contains(r"[а-яa-z]", case=False,
                                                                                regex=True,
                                                                                na=False).sum() < 2:
        header_row += 1

    if header_row >= len(df):  # если строка за пределами таблицы, завершаем работу
        return {}, 0

    base_map: dict[str, int] = {}
    best_score = -1
    last_header_row = header_row

    for offset in range(lookahead):
        row_idx = header_row + offset

        if row_idx >= len(df):  # если следующих строк нет, завершаем цикл
            break

        candidate_map = match_keyword_headers(df.iloc[row_idx].astype(str).tolist(), {})
        candidate_score = score_header_row(candidate_map)

        if not candidate_score:  # если нет основных полей, продолжаем цикл
            continue

        if not base_map:
            base_map = dict(candidate_map)
            best_score = candidate_score
            last_header_row = row_idx
            continue

        if candidate_score > best_score:
            base_map = merge_keep_old(base_map, candidate_map)
            best_score = candidate_score
            last_header_row = row_idx

        else:
            base_map = merge_missing(base_map, candidate_map)

    return base_map, last_header_row
