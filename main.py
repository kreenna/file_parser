import os

from file_parser import ExcelParser
from utils import get_output_json_path

file_name = "Токовый расчет.xlsx"

if __name__ == "__main__":
    parser = ExcelParser()
    result = parser.parse_file(os.path.join("test-files", file_name))

    output_file = get_output_json_path(os.path.join("result-files", file_name))
    result.write_parse_result_to_json(output_file)
