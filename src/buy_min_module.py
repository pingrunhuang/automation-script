from base_module import BaseModule
from utils import send_email, CLIENT

class MinModule(BaseModule):
    def process_row(self, row: int):
        statsH2 = float(self.workbook.sheets["Stats"]["H2"].value)
        try:
            columnH = float(self.sheet[f"H{row}"].value)
        except:
            print(f"column H{row} is not a number, continue...")
            return
        
        sym = self.sheet[f"C{row}"].value
        prefix = "Buy-Min"
        _id = self.sheet[f"A{row}"].value
        if statsH2 >= columnH:
            self.market_operation(row, "BUY", "J", "H", prefix, _id)
        else:
            send_email(f"Crypto-Binance-{prefix}-Insufficient", CLIENT.generate_min_insufficient_email(sym, columnH, statsH2, id))

def run(workbook):
    MinModule(workbook).run()
