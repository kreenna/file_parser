import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import pandas as pd
import structlog
from rapidfuzz import process, fuzz
import openpyxl

from metrics import quality, hash_file
from smart_reader import SmartSheet
from utils import parse_price, normalize_unit, is_junk_row, extract_vendor, read_csv_iter, match_header, \
    split_into_tables

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

            if file_type in (".xlsx", ".xlsm", ".xls"):
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

        file_type: str = file_path.suffix.lower()
        col_map: dict = {}

        try:
            if file_type == ".csv":

                # CSV требует попыток разных разделителей
                for sep in [";", ",", "\t", "|"]:

                    try:
                        df_iterator = read_csv_iter(str(file_path), sep)
                        sheet_items: list[dict] = []

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
                # Excel-форматы: multi-sheet + merged cells
                print("Trying to parse excel (smart)")

                if file_type == ".xlsx":
                    fallback_method = self._standard_xlsx_fallback
                elif file_type == ".xls":
                    fallback_method = self._standard_xls_fallback

                wb = openpyxl.load_workbook(str(file_path), read_only=True, data_only=True, keep_links=False)
                sheet_items: list[dict] = []

                for sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    smart = SmartSheet(ws)
                    all_rows = smart.all_rows()
                    if len(all_rows) < 2:
                        continue

                    tables = split_into_tables(all_rows)

                    # превращаем в df и используем fuzzy
                    for start, end in tables:
                        sub_rows = all_rows[start:end + 1]
                        df = pd.DataFrame(sub_rows)
                        items, col_map = self._parse_dataframe_with_fuzzy(df)
                        if items:
                            sheet_items.extend(items)
                            if not res.col_map:
                                res.col_map = col_map

                wb.close()

                if sheet_items:
                    self._handle_success(res, sheet_items, col_map, "fuzzy+merged")
                    logger.info(f"fuzzy_merged_ok_{file_type}", items=len(sheet_items), conf=f"{res.confidence:.2f}")
                    return

        except Exception as e:
            res.errors.append(f"fuzzy_{file_type}: {e}")
            logger.warning(f"fuzzy_{file_type}_failed", error=str(e))

        fallback_method(file_path, res)   # _standard_xlsx_fallback или _standard_csv_fallback

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

            if best_idx and best_score >= 70:
                col_map[col_type] = best_idx

        if "name" not in col_map and "sku" not in col_map:
            return [], col_map

        items: list[dict] = []
        current_section = ""

        for _, row in data.iterrows():
            name_value = None
            sku_value = None
            unit_value = None
            price_value = None
            manufacturer_value = None
            notes_value = None

            if "name" in col_map:
                name_value = row.iloc[col_map["name"]]
            if "sku" in col_map:
                sku_value = row.iloc[col_map["sku"]]
            if "unit" in col_map:
                unit_value = row.iloc[col_map["unit"]]
            if "price" in col_map:
                price_value = row.iloc[col_map["price"]]
            if "manufacturer" in col_map:
                manufacturer_value = row.iloc[col_map["manufacturer"]]
            if "notes" in col_map:
                notes_value = row.iloc[col_map["notes"]]

            item_values: dict = {
                "name": str(name_value) if name_value and not pd.isna(name_value) else "",
                "sku": str(sku_value) if sku_value and not pd.isna(sku_value) else "",
                "unit": str(unit_value) if unit_value and not pd.isna(unit_value) else None,
                "price": price_value,
                "manufacturer": manufacturer_value,
                "notes": notes_value if notes_value != name_value else ""
            }

            item = ExcelParser._build_item(item_values, raw_row=row, sku_col_idx=col_map.get("sku"))

            has_price = item["price"] > 0
            has_sku = bool(item["sku"])

            if not has_sku and not has_price:
                # treat as section header
                if item["name"]:
                    current_section = item["name"]
                continue

            # normal item
            item["section"] = current_section
            items.append(item)

            if item:
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

    # fallback: стандартный парсер

    def _standard_xlsx_fallback(self, p: Path, res: ParseResult):
        """Логика обработки XLSX-файлов с openpyxl, если Fuzzy не сработал."""

        try:
            print("trying excel fallback")
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
        for row_id, row in enumerate(rows):
            cols = match_header(row)
            if not cols:
                continue

            col_name = cols.get("name")
            col_sku = cols.get("sku")
            col_unit = cols.get("unit")
            col_price = cols.get("price")
            col_manufacturer = cols.get("manufacturer")
            col_notes = cols.get("notes")

            items: list[dict] = []

            for dr in rows[row_id + 1:]:
                # обозначаем переменные
                name_value = None
                sku_value = None
                unit_value = None
                price_value = None
                manufacturer_value = None
                notes_value = None

                if col_name is not None and col_name < len(dr):
                    name_value = dr[col_name]

                if col_sku is not None and col_sku < len(dr):
                    sku_value = dr[col_sku]

                if col_unit is not None and col_unit < len(dr):
                    unit_value = dr[col_unit]

                if col_price is not None and col_price < len(dr):
                    price_value = dr[col_price]

                if col_manufacturer is not None and col_manufacturer < len(dr):
                    manufacturer_value = dr[col_manufacturer]

                if col_notes is not None and col_notes < len(dr):
                    notes_value = dr[col_notes]

                item_values: dict = {
                    "name": str(name_value) if name_value and not pd.isna(name_value) else "",
                    "sku": str(sku_value) if sku_value and not pd.isna(sku_value) else "",
                    "unit": str(unit_value) if unit_value and not pd.isna(unit_value) else None,
                    "price": price_value,
                    "manufacturer": manufacturer_value,
                    "notes": notes_value
                }

                item = ExcelParser._build_item(item_values)

                if item:
                    items.append(item)

            if items:
                return items

        return None

    @staticmethod
    def _build_item(item: dict, raw_row=None, sku_col_idx: Optional[int] = None) -> Optional[dict]:
        """
        Универсальная сборка item из сырых значений.
        name и sku могут быть пустыми; если оба пустые -> вернуть None.
        """

        name = item.get("name") or ""
        sku = item.get("sku") or ""

        # вызов ИИ для поиска данных
        if not name and not sku:
            # TODO: call AI here later to guess name/sku from raw_row
            # ai_name, ai_sku = self._ai_guess_item(raw_row)
            # if ai_name or ai_sku: reuse builder
            return None
        if name and (len(name) < 3 or is_junk_row(name)):
            # если name есть, но мусорный — считаем, что позиции нет
            return None

        # eд.измерения
        unit: str = ""
        unit_value: str = item.get("unit")
        if unit_value:
            unit = normalize_unit(str(unit_value).strip())

        # цена (с защитой от спутывания с SKU
        price_value = item.get("price")
        parsed_price = parse_price(price_value) if price_value not in (None, "") else None

        if all(x is not None for x in (parsed_price, raw_row, sku_col_idx)):
            sku_raw = raw_row.iloc[sku_col_idx]
            sku_value = str(sku_raw) if not pd.isna(sku_raw) else ""

            sku_clean = sku_value.replace("-", "").replace(" ", "")
            try:
                if sku_clean and str(int(parsed_price)) in sku_clean:
                    price = 0.0
                else:
                    price = parsed_price
            except ValueError:
                price = parsed_price or 0.0
        else:
            price = parsed_price or 0.0

        # производитель
        manufacturer_value: str = item.get("manufacturer")
        manufacturer = (manufacturer_value or "").strip()
        if not manufacturer:
            # если производитель пустой, пробуем определить по name
            if name:
                manufacturer = extract_vendor(name)

        # примечания
        notes = item.get("notes") if item.get("notes") and not pd.isna(item.get("notes")) else ""

        if not any([unit, price, manufacturer, notes]):
            return None

        item = {
            "name": name,
            "sku": sku,
            "unit": unit,
            "price": price,
            "manufacturer": manufacturer,
            "notes": notes,
        }

        return item


parser = ExcelParser()

result = parser.parse_file(os.path.join("test-files", "V3.7 Шаблон ТКП апрель.xlsx"))
print(result)
