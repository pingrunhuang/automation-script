from utils import CLIENT, send_email, format_usd_ntl, call_vb, fetch_market_price
import time
from pprint import pprint
from datetime import datetime
from xlwings import Book, Sheet
from binance.exceptions import BinanceAPIException, BinanceOrderException

def run(workbook):
    print("Running min module")
    sheet = workbook.sheets["Overview"]
    row = 3
    while sheet[f"C{row}"].value is not None:
        print(f"checking row {row}")
        _min(row, workbook)
        row+=1


def _min(row, workbook):
    print(f"checking row {row}")
    sheet = workbook.sheets["Overview"]
    exch = sheet[f"HA{row}"].value
    statsH2 = workbook.sheets["Stats"]["H2"].value
    columnH = sheet[f"H{row}"].value
    _id = sheet[f"A{row}"].value
    sym = sheet[f"C{row}"].value
    marketBuy = sheet[f"J{row}"].value

    if columnH!="-" and exch=="Binance":
        call_vb(workbook)
        if statsH2 >= columnH:
            print("statsH2 greater than columnI")
            data, url = fetch_market_price(sym, "min")
            if not data:
                return
            prx = float(data["price"])
            if prx > marketBuy:
                print(f"price from api:{prx} > marketBuy:{marketBuy}")
                send_email("Crypto-Binance-MinReject", CLIENT.generate_reject_email(sym, data["price"], marketBuy, url, data))
            else:
                print(f"price from api:{prx} <= marketBuy:{marketBuy}")
                try:
                    order_detail = create_buy_order(sym, columnH)
                    if order_detail:
                        send_email("Crypto-Binance-MinDone", CLIENT.generate_min_email(sym, order_detail, columnH, marketBuy))
                        process_binance_sheet(workbook, _id, datetime.today(), order_detail["cummulativeQuoteQty"], order_detail["executedQty"])
                except (BinanceOrderException, BinanceAPIException) as e:
                    send_email("Crypto-Binance-MinOrderError", CLIENT.generate_order_error_mail(sym, e.message, "BUY"))
        else:
            print(f"statsH2:{statsH2} less then columnH: {columnH}")
            send_email("Crypto-Binance-MinInsufficient", CLIENT.generate_min_insufficient_email(sym, columnH, statsH2))
    print("######################################################################")
    


def create_buy_order(sym, quote_qty):
    pair = f"{sym}USDT"
    print(f"Creating buy order of {pair} with {quote_qty} and wait for 5 seconds...")
    time.sleep(5)
    resp = CLIENT.order_market_buy(symbol=pair, quoteOrderQty=format_usd_ntl(quote_qty))
    print("Order confirmation:")
    pprint(resp, indent=4)
    return resp


def process_binance_sheet(workbook:Book, _id:str, dt:datetime, cummulativeQuoteQty:str, executedQty:float, to="USDT"):
    sheet:Sheet = workbook.sheets["Binance"]
    str_dt = dt.strftime("%d-%b-%y")
    row = sheet.range("A3").end("down").row + 1
    sheet.cells(row, 1).value= _id # column A
    sheet.cells(row, 3).value= to # column C
    sheet.cells(row, 5).value= str_dt # column E
    sheet.cells(row, 7).value= -float(cummulativeQuoteQty) # column G
    sheet.cells(row, 10).value= executedQty # column J
    print("check binance sheet?")
    return row
