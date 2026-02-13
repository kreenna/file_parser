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


def match_header(row) -> dict | None:
    """Проверка на соответствие названиям шапки."""

    COLUMN_PATTERNS = {
        "name": [r"наименован", r"названи", r"товар", r"материал",
                 r"номенклатур", r"продукци", r"работ", r"^name$", r"product", r"item", r"имя"],
        "sku": [r"артикул", r"арт[\.\s]", r"код\s*товар", r"sku", r"^код$", r"article"],
        "unit": [r"ед[\.\s]", r"единиц", r"unit", r"шт[\.\s", r"кг", r"м"],
        "price": [r"цена", r"стоимость", r"руб", r"₽", r"price", r"расценк"],
        "manufacturer": [r"производител", r"бренд", r"vendor", r"manufactur", r"поставщик"],
        "notes": [r"примечани", r"notes", r"описание", r"комментари"]
    }

    if not row:
        return None

    found: dict[str, int] = {}

    for i, value in enumerate(row):
        if value is None:
            continue

        stripped = str(value).lower().strip()

        if not stripped or len(stripped) > 200:
            continue
        for ftype, pats in COLUMN_PATTERNS.items():
            if ftype in found:
                continue
            if any(re.search(pat, stripped) for pat in pats):
                found[ftype] = i
                break
    return found if "name" in found and len(found) >= 2 else None


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
