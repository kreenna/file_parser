class SmartSheet:
    def __init__(self, ws):
        self.ws = ws
        self.merged_map = {}
        if hasattr(ws, "merged_cells"):
            for mg in ws.merged_cells.ranges:
                try:
                    val = ws.cell(mg.min_row, mg.min_col).value
                    for r in range(mg.min_row, mg.max_row + 1):
                        for c in range(mg.min_col, mg.max_col + 1):
                            self.merged_map[(r, c)] = val
                except Exception:
                    continue
        self.max_row = ws.max_row or 0
        self.max_col = ws.max_column or 0

    def cell(self, row, col):
        return self.merged_map.get((row, col), self.ws.cell(row, col).value)

    def all_rows(self) -> list[list]:
        return [
            [self.cell(r, c) for c in range(1, self.max_col + 1)]
            for r in range(1, self.max_row + 1)
        ]
