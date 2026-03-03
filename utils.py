import re
from typing import Optional

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
    "unit": ["ед", "ед. изм", "ед.изм", "единиц", "unit", "изм"],
    "quantity": ["количество", "кол-во", "кол."],
    "price_unit": ["цена", "стоимость", "price", "расценк"],
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


def simple_match_header(row) -> dict[str, int]:
    """
    Очень простой матчинг заголовка по подстрокам.
    Возвращает col_map: {\"name\": idx, ...}.
    """

    col_map: dict[str, int] = {}

    for idx, value in enumerate(row):
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


def collect_prices_for_row(row, col_map: dict) -> dict:
    """
    Собирает все найденные ценовые значения по типам:
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
