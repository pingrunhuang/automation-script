from base_module import BaseModule
from utils import send_email, CLIENT

class BuyMoreModule(BaseModule):
    def process_row(self, row: int):
        statsH2 = float(self.workbook.sheets["Stats"]["H2"].value)
        try:
            columnI = float(self.sheet[f"I{row}"].value)
        except:
            print(f"column I{row} is not a number, continue...")
            return
        sym = self.sheet[f"C{row}"].value
        prefix = "Buy-More"
        _id = self.sheet[f"A{row}"].value
        if statsH2 >= columnI:
            self.market_operation(row, "BUY", "J", "I", prefix, _id)
        else:
            send_email(f"Crypto-Binance-{prefix}-Insufficient", CLIENT.generate_buy_insufficient_email(sym, columnI, statsH2, _id))

def run(workbook):
    BuyMoreModule(workbook).run()
