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
    "unit": ["ед", "ед. изм", "ед.изм", "единиц", "unit", "изм"],
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


HEADER_HINT_KEYWORDS = [
    "наименован", "назван", "товар", "позици", "материал",
    "description", "item", "product",
    "артикул", "арт", "код", "sku", "article",
    "ед", "изм", "unit",
    "цена", "стоим", "руб", "₽", "price",
    "производ", "бренд", "vendor", "manufacturer", "поставщик",
    "примечан", "коммент", "notes", "описан",
]

_header_hint_re = re.compile("|".join(HEADER_HINT_KEYWORDS), re.IGNORECASE)


def is_probable_header_row(row: pd.Series,
                           min_label_cells: int = 2,
                           max_numeric_fraction: float = 0.5) -> bool:
    """Heuristic: is this row likely to be a header row of a table?"""

    cells = row.astype(str)
    non_empty = cells.str.strip().ne("").sum()
    if non_empty == 0:
        return False

    # Count cells with letters
    has_letters = cells.str.contains(r"[A-Za-zА-Яа-я]", regex=True, na=False)
    n_letters = has_letters.sum()
    print(f"n_letters: {n_letters}")

    # Count mostly numeric cells
    numeric_like = cells.str.fullmatch(r"[-+]?\d+([.,]\d+)?", na=False)
    n_numeric = numeric_like.sum()

    # Basic filters
    if n_letters < min_label_cells:
        return False

    # Bonus: any header-like keyword
    if cells.str.contains(_header_hint_re, na=False).any():
        return True

    if n_numeric / max(non_empty, 1) > max_numeric_fraction:
        # Too many numeric cells → likely data row
        return False

    # Fallback: if there are enough text cells and not too many numbers, accept
    return True


def detect_header_block(df: pd.DataFrame,
                        max_scan_rows: int = 30,
                        max_header_block_height: int = 5
                        ) -> tuple[list[int], int]:
    """
    Detect header rows indices and first data row index.
    Returns ([header_row_indices], data_start_index).
    If nothing found, returns ([], 0) → assume header at 0.
    """
    n_rows = len(df)
    if n_rows == 0:
        return [], 0

    max_scan_rows = min(max_scan_rows, n_rows)

    first_header_idx = None
    header_rows: list[int] = []

    # 1) find first header row
    for i in range(max_scan_rows):
        row = df.iloc[i]
        if is_probable_header_row(row):
            first_header_idx = i
            header_rows.append(i)
            break

    if first_header_idx is None:
        # No good candidate; fallback to row 0
        return [0], 1 if n_rows > 1 else n_rows

    # 2) include immediately following header-like rows (multi-row header)
    for j in range(first_header_idx + 1,
                   min(first_header_idx + 1 + max_header_block_height, max_scan_rows)):
        row = df.iloc[j]
        if is_probable_header_row(row):
            header_rows.append(j)
        else:
            break

    data_start = max(header_rows) + 1
    return header_rows, data_start


def build_multirow_header(df: pd.DataFrame, header_row_indices: list[int], data_start_index: int, sep: str = " | ") -> \
        tuple[pd.DataFrame, list[str]]:
    """
    Take first `n_header_rows` rows as header rows, combine them into one
    string header per column, and return (df_data, headers).
    """
    if df.empty or not header_row_indices:
        return df, df.columns.astype(str).tolist()

    header_block = df.iloc[header_row_indices].copy()
    headers: list[str] = []

    n_cols = header_block.shape[1]

    # Fill upward merged top-cells (you already handle this earlier if you use SmartSheet).
    # If df came from SmartSheet, merged cells are already expanded.

    # Column-wise, gather non-empty header fragments
    for col_idx in range(n_cols):
        fragments = []
        col_vals = header_block.iloc[:, col_idx]
        for val in col_vals:
            if pd.isna(val):
                continue
            text = str(val).strip()
            if not text:
                continue
            fragments.append(text)

        if not fragments:
            headers.append(f"col_{col_idx}")
        else:
            uniq = []
            for f in fragments:
                if f not in uniq:
                    uniq.append(f)
            headers.append(sep.join(uniq))

        data = df.iloc[data_start_index:].reset_index(drop=True)

        if data.shape[1] != len(headers):
            raise ValueError(
                f"Header/data width mismatch: df has {data.shape[1]} cols, headers has {len(headers)}"
            )

        data.columns = headers
        return data, headers
