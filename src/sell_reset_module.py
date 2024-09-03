from base_module import BaseModule

class ResetModule(BaseModule):
    def process_row(self, row: int):
        _id = self.sheet[f"A{row}"].value
        if self.sheet[f"P{row}"].value == 100:
            self.market_operation(row, "SELL", "N", "S", "Sell-Reset", _id)

def run(workbook):
    ResetModule(workbook).run()
