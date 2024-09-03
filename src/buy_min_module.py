from base_module import BaseModule
from utils import send_email, CLIENT

class MinModule(BaseModule):
    def process_row(self, row: int):
        statsH2 = self.workbook.sheets["Stats"]["H2"].value
        columnH = self.sheet[f"H{row}"].value
        sym = self.sheet[f"C{row}"].value
        prefix = "Buy-Min"
        _id = self.sheet[f"A{row}"].value
        if statsH2 >= columnH:
            self.market_operation(row, "BUY", "J", "H", prefix, _id)
        else:
            send_email(f"Crypto-Binance-{prefix}-Insufficient", CLIENT.generate_min_insufficient_email(sym, columnH, statsH2, id))

def run(workbook):
    MinModule(workbook).run()
