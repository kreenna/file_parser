import os
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
import structlog
from rapidfuzz import process, fuzz

from metrics import quality, hash_file
from utils import parse_price, normalize_unit, is_junk_row, extract_vendor, read_csv_iter, match_header

logger = structlog.get_logger()


@dataclass
class ParseResult:
    file_path: str
    file_hash: str
    items: list[dict] = field(default_factory=list)
    confidence: float = 0.0
    needs_ai: bool = False
    errors: list[str] = field(default_factory=list)
    method: str = "none"
    col_map: dict = field(default_factory=dict)


class ExcelParser:
    """
    XLSX/XLS/CSV парсер:
    - основной метод: pandas + fuzzy по шапке
    - fallback: простой regex/эвристический парсер по строкам
    - метрики качества _quality + needs_ai
    - PDF/DOC/DOCX пока игнорируем
    """

    # Публичный метод

    def parse_file(self, file_path: str) -> ParseResult:
        """Метод фактического парсинга файла, возвращает объект ParseResult с данными о файле и содержимом."""

        p = Path(file_path)
        res = ParseResult(file_path=file_path, file_hash=hash_file(p))

        # пробуем парсить файл в зависимости от его формата
        try:
            print("Trying to parse")
            file_type = p.suffix.lower()

            if file_type in (".xlsx", ".xlsm", ".xls", ".csv"):
                self._parse_spreadsheet(p, res, [pd.ExcelFile])
            elif file_type in ['.csv', '.tsv']:
                self._parse_spreadsheet(p, res, [], self._standard_csv_fallback)
            else:
                res.errors.append(f"Unsupported type: {file_type}.")
                return res

            AI_FALLBACK_THRESHOLD = 0.55
            res.needs_ai = (res.confidence < AI_FALLBACK_THRESHOLD or not res.items)

        except Exception as e:
            print("Failed to parse")
            logger.error("parse_error", file=str(p), error=str(e))
            res.errors.append(str(e))
            res.needs_ai = True

        return res

    # XLSX / XLS

    def _parse_spreadsheet(self, file_path: Path, res: ParseResult, readers: list, fallback_method: callable = None):
        """Универсальный парсер для spreadsheet-форматов (xlsx, xls, csv)."""

        file_type = file_path.suffix.lower()

        try:
            if file_type == ".csv":

                # CSV требует попыток разных разделителей
                for sep in [";", ",", "\t", "|"]:

                    try:
                        df_iterator = read_csv_iter(str(file_path), sep)
                        sheet_items = []

                        for df in df_iterator:
                            # если колонок больше, чем 2, парсим с использованием fuzzy
                            if df.shape[1] < 2:
                                continue

                            items, col_map = self._parse_dataframe_with_fuzzy(df)
                            if items:  # если данные после парсинга есть, сохраняем
                                sheet_items.extend(items)

                        if sheet_items:
                            self._handle_success(res, sheet_items, col_map, f"{file_type}+fuzzy", sep)
                            logger.info(f"{file_type}_fuzzy_ok", items=len(sheet_items), conf=f"{res.confidence:.2f}",
                                        sep=sep)
                            return

                    except Exception as e:
                        res.errors.append(f"{file_type}({sep}): {e}")
                        continue
            else:
                # Excel-форматы: один reader, multi-sheet
                print("Trying to parse excel")
                if file_type == ".xlsx":
                    fallback_method = self._standard_xlsx_fallback
                elif file_type == ".xls":
                    fallback_method = self._standard_xls_fallback

                xls = readers[0](str(file_path))  # pd.ExcelFile для xlsx/xls
                sheet_items = []

                # проходимся по каждому листу
                for sheet_name in xls.sheet_names:
                    df = xls.parse(sheet_name=sheet_name, dtype=str)
                    items, col_map = self._parse_dataframe_with_fuzzy(df)

                    if items:
                        sheet_items.extend(items)
                        if not res.col_map:
                            res.col_map = col_map

                if sheet_items:
                    self._handle_success(res, sheet_items, None, "fuzzy")
                    logger.info(f"fuzzy_ok_{file_type}", items=len(sheet_items), conf=f"{res.confidence:.2f}")
                    return

        except Exception as e:
            res.errors.append(f"fuzzy_{file_type}: {e}")
            logger.warning(f"fuzzy_{file_type}_failed", error=str(e))

        fallback_method(file_path, res)  # _standard_xlsx_fallback или _standard_csv_fallback

    # основной путь: pandas + fuzzy

    @staticmethod
    def _parse_dataframe_with_fuzzy(df: pd.DataFrame) -> tuple[list[dict], dict]:
        """Парсинг файла с использованием Fuzzy для поиска данных."""

        if df.empty:
            return [], {}

        # ищем строку шапки: первая строка, где есть >=2 ячейки с буквами
        header_row = 0
        while header_row < len(df) and df.iloc[header_row].astype(str).str.contains(r"[а-яa-z]", case=False, regex=True,
                                                                                    na=False).sum() < 2:
            header_row += 1

        headers = df.iloc[header_row].astype(str)
        data = df.iloc[header_row + 1:].reset_index(drop=True)

        COLUMN_KEYWORDS = {
            "name": ["наименовани", "названи", "товар", "позици", "материал", "product", "item", "имя",
                     "название товара", "материал", "номенклатур"],
            "sku": ["артикул", "арт", "код", "sku", "article"],
            "unit": ["ед. измерения", "единица измерения", "unit", "шт.", "кг", "м."],
            "price": ["цена", "стоимость", "руб", "₽", "price", "расценк"],
            "manufacturer": ["производител", "бренд", "vendor", "manufacturer", "поставщик"],
            "notes": ["примечани", "notes", "описание", "комментари"]
        }

        col_map: dict[str, int] = {}

        for col_type, keywords in COLUMN_KEYWORDS.items():
            best_score = 0
            best_idx = None

            # проверяем совпадения по каждому слову и определяем лучшее
            for keyword in keywords:
                match = process.extractOne(keyword, headers.astype(str).tolist(), scorer=fuzz.partial_ratio)

                if match and match[1] > best_score:
                    best_score = match[1]
                best_idx = match[2]
                if best_score >= 70:
                    col_map[col_type] = best_idx

        if "name" not in col_map:
            return [], col_map

        name_idx = col_map["name"]
        items: list[dict] = []

        for _, row in data.iterrows():
            name_value = str(row.iloc[name_idx]).strip()
            if not name_value or len(name_value) < 3 or is_junk_row(name_value):
                continue

            item = {
                "name": name_value,
                "sku": "",
                "unit": "шт",
                "price": 0.0,
                "manufacturer": "",
                "notes": ""
            }

            for col_type, column_id in col_map.items():
                if col_type == "name" or column_id >= len(row):
                    continue

                value = row.iloc[column_id]
                if pd.isna(value):
                    continue

                if col_type == "sku":
                    item["sku"] = str(value).strip()
                elif col_type == "unit":
                    item["unit"] = normalize_unit(str(value).strip())


                elif col_type == "price":

                    parsed_price = parse_price(value)

                    # защита от артикула вместо цены

                    sku_val = ""

                    if "sku" in col_map:
                        sku_raw = row.iloc[col_map["sku"]]
                        sku_val = str(sku_raw).strip() if not pd.isna(sku_raw) else ""

                    sku_clean = sku_val.replace("-", "").replace(" ", "")

                    if sku_clean and parsed_price is not None:
                        try:
                            if str(int(parsed_price)) in sku_clean:
                                item["price"] = 0.0
                            else:
                                item["price"] = parsed_price

                        except ValueError:
                            item["price"] = parsed_price or 0.0

                    else:

                        item["price"] = parsed_price or 0.0

                elif col_type == "manufacturer":
                    item["manufacturer"] = str(value).strip()

                elif col_type == "notes":
                    item["notes"] = str(value).strip()

                if not item["manufacturer"]:
                    item["manufacturer"] = extract_vendor(name_value)

                items.append(item)

        return items, col_map

    @staticmethod
    def _handle_success(res: ParseResult, items: list, col_map: dict, method_suffix: str, sep: str = None):
        """Общий обработчик успеха fuzzy-парсинга."""
        res.items.extend(items)
        if not res.col_map and col_map:
            res.col_map = col_map
        res.method = f"{method_suffix}"
        res.confidence = quality(res.items)

    # Fallback: стандартный парсер

    def _standard_xlsx_fallback(self, p: Path, res: ParseResult):
        """Логика обработки XLSX-файлов с openpyxl, если Fuzzy не сработал."""

        try:
            print("triying excel fallback")
            import openpyxl
            wb = openpyxl.load_workbook(str(p), read_only=True, data_only=True)
            for sn in wb.sheetnames:
                rows = list(wb[sn].iter_rows(values_only=True))
                items = self._standard_parse(rows)
                if items:
                    res.items.extend(items)
                    res.method = "standard"
            wb.close()
        except Exception as e:
            res.errors.append(f"standard_xlsx: {e}")

        if res.items:
            res.confidence = quality(res.items)

    def _standard_xls_fallback(self, p: Path, res: ParseResult):
        """Логика обработки XLS-файлов с xlrd, если Fuzzy не сработал."""

        try:
            import xlrd
            wb = xlrd.open_workbook(str(p))
            for i in range(wb.nsheets):
                ws = wb.sheet_by_index(i)
                rows = [tuple(ws.cell_value(r, c) for c in range(ws.ncols))
                        for r in range(ws.nrows)]
                items = self._standard_parse(rows)
                if items:
                    res.items.extend(items)
                    res.method = "standard"
        except Exception as e:
            res.errors.append(f"standard_xls: {e}")

        if res.items:
            res.confidence = quality(res.items)

    def _standard_csv_fallback(self, p: Path, res: ParseResult):
        """Логика обработки CSV-файлов с pandas, если Fuzzy не сработал."""

        for sep in [";", ",", "\t", "|"]:
            try:
                df = pd.read_csv(str(p), sep=sep, header=None, dtype=str)
                rows = [tuple(r) for _, r in df.iterrows()]
                items = self._standard_parse(rows)
                if items:
                    res.items.extend(items)
                    res.method = "csv_standard"
                    res.confidence = quality(items)
                    return
            except Exception as e:
                res.errors.append(f"csv_std({sep}): {e}")
                continue

    @staticmethod
    def _standard_parse(rows) -> list[dict] | None:
        """Стандартный парсинг."""

        # поиск заголовка
        for row_id, row in enumerate(rows[:30]):
            cols = match_header(row)
            if not cols or "name" not in cols:
                continue

            col_name = cols["name"]
            items: list[dict] = []

            for dr in rows[row_id + 1:]:
                if col_name >= len(dr) or dr[col_name] is None:
                    continue
                name = str(dr[col_name]).strip()
                if not name or len(name) < 3 or is_junk_row(name):
                    continue

                item = {
                    "name": name,
                    "sku": "",
                    "unit": "шт",
                    "price": 0.0,
                    "manufacturer": "",
                    "notes": "",
                }

                for col_type, col_id in cols.items():
                    if col_type == "name" or col_id >= len(dr):
                        continue
                    value = dr[col_id]

                    # находим и заполняем информацию о позиции
                    if col_type == "sku":
                        item["sku"] = str(value).strip() if value else ""

                    elif col_type == "unit":
                        item["unit"] = normalize_unit(str(value).strip()) if value else "шт"

                    elif col_type == "price":
                        item["price"] = parse_price(value) if value else 0.0

                    elif col_type == "manufacturer":
                        item["manufacturer"] = str(value).strip() if value else ""

                    elif col_type == "notes":
                        item["notes"] = str(value).strip() if value else ""

                if not item["manufacturer"]:
                    item["manufacturer"] = extract_vendor(name)
                items.append(item)

            if items:
                return items
        return None


parser = ExcelParser()

result = parser.parse_file(os.path.join("test-files", "1.Спецификация (Гибрид)  1АПС д.1 — копия.xlsx"))
print(result)
