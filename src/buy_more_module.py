from base_module import BaseModule
from utils import send_email, CLIENT

class BuyMoreModule(BaseModule):
    def process_row(self, row: int):
        statsH2 = self.workbook.sheets["Stats"]["H2"].value
        columnI = self.sheet[f"I{row}"].value
        sym = self.sheet[f"C{row}"].value
        prefix = "Buy-More"
        _id = self.sheet[f"A{row}"].value
        if statsH2 >= columnI:
            self.market_operation(row, "BUY", "J", "I", prefix, _id)
        else:
            send_email(f"Crypto-Binance-{prefix}-Insufficient", CLIENT.generate_buy_insufficient_email(sym, columnI, statsH2, _id))

def run(workbook):
    BuyMoreModule(workbook).run()
