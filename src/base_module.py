from abc import ABC, abstractmethod
from utils import CLIENT, send_email, call_vb, fetch_market_price, generate_qty
from xlwings import Book
from datetime import datetime
from binance.exceptions import BinanceAPIException, BinanceOrderException
import time

class BaseModule(ABC):
    def __init__(self, workbook: Book):
        self.workbook = workbook
        self.sheet = workbook.sheets["Overview"]

    @abstractmethod
    def process_row(self, row: int):
        pass

    def run(self):
        print(f"Running {self.__class__.__name__}")
        row = 3
        while self.sheet[f"C{row}"].value is not None:
            self.process_row(row)
            row += 1

    def create_order(self, sym: str, quote_qty: float, side: str):
        pair = f"{sym}USDT"
        print(f"Creating {side.lower()} order of {pair} with {quote_qty} and wait for 5 seconds...")
        time.sleep(5)
        if side == "BUY":
            return CLIENT.order_market_buy(symbol=pair, quoteOrderQty=quote_qty)
        elif side == "SELL":
            return CLIENT.order_market_sell(symbol=pair, quoteOrderQty=quote_qty)
        else:
            raise ValueError(f"Invalid side: {side}")

    def process_binance_sheet(self, _id: str, dt: datetime, quote_qty: str, executed_qty: float, side: str, order_id:str):
        sheet = self.workbook.sheets["Binance"]
        str_dt = dt.strftime("%d-%b-%y")
        row = sheet.range("A3").end("down").row + 1
        sheet.cells(row, 1).value = _id  # column A
        sheet.cells(row, 3 if side == "BUY" else 4).value = "USDT"  # column C or D
        sheet.cells(row, 5).value = str_dt  # column E
        sheet.cells(row, 8 if self.__class__.__name__ == "SellModule" else 7).value = -float(quote_qty) if side == "BUY" else quote_qty  # column G or H
        sheet.cells(row, 10).value = executed_qty if side == "BUY" else -executed_qty  # column J
        sheet.cells(row, 14).value = order_id # column N set to order_id
        print("check binance sheet?")
        return row

    def market_operation(self, row: int, side: str, price_column: str, qty_column: str, email_prefix: str, _id:str):
        exch = self.sheet[f"HA{row}"].value
        qty = self.sheet[f"{qty_column}{row}"].value
        sym = self.sheet[f"C{row}"].value
        market_price = self.sheet[f"{price_column}{row}"].value

        if qty != "-" and exch == "Binance":
            call_vb(self.workbook)
            data, url = fetch_market_price(sym, _id, email_prefix)
            if not data:
                return

            prx = float(data["price"])
            price_condition = prx < market_price if side == "SELL" else prx > market_price
            if price_condition:
                send_email(f"Crypto-Binance-{email_prefix}-Reject", CLIENT.generate_reject_email(sym, data["price"], market_price, url, data, _id))
            else:
                try:
                    order_detail = self.create_order(sym, qty, side)
                    if order_detail:
                        quote_qty, executed_qty  = generate_qty(order_detail)
                        order_id = order_detail["orderId"]
                        self.process_binance_sheet(
                            _id=_id, 
                            dt=datetime.today(), 
                            quote_qty=quote_qty, 
                            executed_qty=executed_qty, 
                            side=side,
                            order_id=order_id
                        )
                        xcell = self.sheet[f"X{row}"].value
                        str_xcell = "{:.8f}".format(float(xcell))
                        table = CLIENT.generate_done_email(sym, order_detail, qty, market_price, email_prefix, _id, str_xcell)
                        send_email(f"Crypto-Binance-{email_prefix}-Done", table)
                except (BinanceOrderException, BinanceAPIException) as e:
                    send_email(f"Crypto-Binance-{email_prefix}-OrderError", CLIENT.generate_order_error_mail(sym, e.message, _id, side))
