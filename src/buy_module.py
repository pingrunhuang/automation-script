from utils import CLIENT, send_email, format_usd_ntl, call_vb, fetch_market_price
import time
from pprint import pprint
from datetime import datetime
from xlwings import Book, Sheet
from binance.exceptions import BinanceAPIException, BinanceOrderException

def buy(row, workbook):
    print(f"checking row {row}")
    sheet = workbook.sheets["Overview"]
    exch = sheet[f"AG{row}"].value
    AN23 = sheet[f"AN23"].value
    columnI = sheet[f"I{row}"].value
    _id = sheet[f"A{row}"].value
    sym = sheet[f"C{row}"].value
    marketBuy = sheet[f"J{row}"].value
    pair = f"{sym}USDT"
    if columnI!="-" and exch=="Binance":
        if AN23 >= columnI:
            print("AN23 greater than columnI")
            data, url = fetch_market_price(pair, workbook)
            if not data:
                return
            prx = float(data["price"])
            if prx > marketBuy:
                print(f"price from api:{prx} > marketBuy:{marketBuy}")
                send_email("Crypto-Binance-BuyReject", CLIENT.generate_reject_email(sym, data["price"], marketBuy, url, data), workbook)
            else:
                print(f"price from api:{prx} <= marketBuy:{marketBuy}")
                try:
                    order_detail = create_buy_order(pair, columnI)
                    if order_detail:
                        print("proceeding step 6.2.3.2.2.1?")
                        send_email("Crypto-Binance-BuyDone", CLIENT.generate_buy_email(sym, order_detail, columnI))
                        print("processing step 6.2.3.2.2.2?")
                        process_binance_sheet(workbook, _id, datetime.today(), order_detail["cummulativeQuoteQty"], order_detail["executedQty"])
                except (BinanceOrderException, BinanceAPIException) as e:
                    print("Binance buy order error step 6.2.3.2.1, continue?")
                    send_email("Crypto-Binance-BuyOrderError", CLIENT.generate_order_error_mail(e.message), workbook)
        else:
            print("AN23 less then columnI")
            send_email("Crypto-Binance-BuyInsufficient", CLIENT.generate_buy_insufficient_email(sym, columnI, AN23))
    print("######################################################################")
    

def run(workbook):
    print("Running buy module")
    call_vb(workbook)
    sheet = workbook.sheets["Overview"]
    row = 2
    while sheet[f"C{row}"].value is not None:

        buy(row, workbook)
        row+=1



def create_buy_order(pair, quote_qty):
    print(f"Creating buy order of {pair} with {quote_qty} and wait for 5 seconds...")
    time.sleep(5)
    resp = CLIENT.order_market_buy(symbol=pair, quoteOrderQty=format_usd_ntl(quote_qty))
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
        "side": "BUY",
        "workingTime": 1507725176595,
        "selfTradePreventionMode": "NONE"
    }
    """
    print("Order confirmation:")
    pprint(resp, indent=4)
    return resp


def process_binance_sheet(workbook:Book, _id:str, dt:datetime, cummulativeQuoteQty:str, executedQty:float, to="USDT"):
    sheet:Sheet = workbook.sheets["Binance"]
    str_dt = dt.strftime("%d-%b-%y")
    # find last row of a certain column
    # row = sheet.used_range.last_cell.row: this is used for the whole sheet
    row = sheet.range("A1").end("down").row + 1
    sheet.cells(row, 1).value = _id # column A
    sheet.cells(row, 3).value=to # column C
    sheet.cells(row, 5).value=str_dt # column E
    sheet.cells(row, 7).value=-float(cummulativeQuoteQty) # column G
    sheet.cells(row, 10).value=executedQty # column J
    print("check binance sheet?")
    return row
