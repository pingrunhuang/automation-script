import requests
from utils import CLIENT, send_email, format_usd_ntl, call_vb, read_and_save_workbook
import time
from pprint import pprint
from datetime import datetime
import json
from xlwings import Book, Sheet

def process_overview_sheet():
    with read_and_save_workbook() as workbook:
        sheet = workbook.sheets["Overview"]
        row = 2
        while sheet[f"C{row}"].value is not None:
            sheet = workbook.sheets["Overview"]
            _id = sheet[f"A{row}"].value
            sym = sheet[f"C{row}"].value
            marketsell = sheet[f"N{row}"].value
            str_usdt_ntl = sheet[f"M{row}"].value
            call_vb(workbook)
            if str_usdt_ntl=="-":
                print("Illegal number of quote qty")
            else:
                pair = f"{sym}USDT"
                print(f"processing symbol={sym}")
                url = f"https://api.binance.com/api/v1/ticker/price?symbol={pair}"
                res = requests.get(url, proxies={"https":"127.0.0.1:7890"})
                data = res.json()
                # show info
                print(f"Response: {data}")
                prx = float(data["price"])

                if prx < marketsell:
                    send_email("Crypto-Binance-Reject", f"{sym}\t{marketsell}")
                else:
                    order_detail = create_sell_order(pair, str_usdt_ntl)
                    if order_detail:
                        # send_email("Crypto-Binance-Sell", json.dumps(order_detail))
                        # checked that executedQty is the base qty
                        process_binance_sheet(workbook, _id, datetime.today(), str_usdt_ntl, -float(order_detail["executedQty"]))
            workbook.save()
            row+=1
            print("######################################################################")

def create_sell_order(pair, quote_qty):
    print(f"Creating selling order of {pair} with {quote_qty} and wait for 5 seconds...")
    time.sleep(5)
    resp = CLIENT.order_market_sell(symbol=pair, quoteOrderQty=format_usd_ntl(quote_qty))
    """
    {
        "symbol": "BTCUSDT",
        "orderId": 28,
        "orderListId": -1, //Unless an order list, value will be -1
        "clientOrderId": "6gCrw2kRUAF9CvJDGP16IP",
        "transactTime": 1507725176595,
        "price": "0.00000000",
        "origQty": "10.00000000",
        "executedQty": "10.00000000",
        "cummulativeQuoteQty": "10.00000000",
        "status": "FILLED",
        "timeInForce": "GTC",
        "type": "MARKET",
        "side": "SELL",
        "workingTime": 1507725176595,
        "selfTradePreventionMode": "NONE"
    }
    """
    print("Order confirmation:")
    pprint(resp, indent=4)
    return resp


def process_binance_sheet(workbook:Book, _id:str, dt:datetime, usdt_profit:str, qty:float, to="USDT"):
    sheet:Sheet = workbook.sheets["Binance"]
    row = sheet.used_range.last_cell.row+1
    str_dt = dt.strftime("%d-%b-%y")
    sheet.cells(row, 1).value = _id # column A
    sheet.cells(row, 4).value=to # column D 
    sheet.cells(row, 5).value=str_dt # column E
    sheet.cells(row, 8).value=usdt_profit # column H
    sheet.cells(row, 10).value=qty # column J


process_overview_sheet()