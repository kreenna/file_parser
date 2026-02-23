class SmartSheet:
    def __init__(self, ws):
        self.ws = ws
        self.data = [list(row) for row in ws.values]

        if not self.data:
            self.max_row = 0
            self.max_col = 0
            return

        self.max_row = len(self.data)
        self.max_col = len(self.data[0])

        if hasattr(ws, "merged_cells"):
            for merged_cell in ws.merged_cells.ranges:
                min_r, min_c = merged_cell.min_row - 1, merged_cell.min_col - 1
                max_r, max_c = merged_cell.max_row, merged_cell.max_col

                try:
                    value = self.data[min_r][min_c]
                except IndexError:
                    continue

                for row in range(min_r, max_r):
                    for col in range(min_c, max_c):
                        if row == min_r and col == min_c:
                            continue
                        if row < self.max_row and col < self.max_col:
                            self.data[row][col] = value

    def cell(self, row: int, col: int):
        if 1 <= row <= self.max_row and 1 <= col <= self.max_col:
            return self.data[row - 1][col - 1]
        return None

    def all_rows(self) -> list[list]:
        return self.data
