from base_module import BaseModule

class SellModule(BaseModule):
    def process_row(self, row: int):
        _id = self.sheet[f"A{row}"].value
        self.market_operation(row, "SELL", "N", "M", "Sell-Profit", _id)

def run(workbook):
    SellModule(workbook).run()
