from utils import CLIENT, send_email, format_usd_ntl, call_vb, fetch_market_price
import time
from pprint import pprint
from datetime import datetime
from xlwings import Book, Sheet
from binance.exceptions import BinanceAPIException, BinanceOrderException

def process_overview_sheet(row, workbook):
    sheet = workbook.sheets["Overview"]
    exch = sheet[f"AG{row}"].value
    str_usdt_ntl = sheet[f"M{row}"].value
    if exch != "Binance":
        print("not from binance")
    elif str_usdt_ntl=="-":
        print("Illegal number of quote qty")
    else:
        print("called vb on step 3, continue?")
        call_vb(workbook)
        _id = sheet[f"A{row}"].value
        sym = sheet[f"C{row}"].value
        print("Proceeding step 4?")
        marketsell = sheet[f"N{row}"].value
        pair = f"{sym}USDT"
        print("proceeding step 5?")
        print(f"processing symbol={sym}")
        data, url = fetch_market_price(pair)
        if not data:
            return
        prx = float(data["price"])
        if prx < marketsell:
            print(f"price from api:{prx} < marketsell:{marketsell}, proceeding step 6.2.2?")
            send_email("Crypto-Binance-SellReject", CLIENT.generate_reject_email(sym, data["price"], marketsell, url, data))
        else:
            print(f"price from api:{prx} >= marketsell:{marketsell}, proceeding step 6.2.3?")
            try:
                order_detail = create_sell_order(pair, str_usdt_ntl)
                if order_detail:
                    print("proceeding step 6.2.3.2.2?")
                    # checked that executedQty is the base qty
                    process_binance_sheet(workbook, _id, datetime.today(), order_detail["cummulativeQuoteQty"], order_detail["executedQty"])
                    print("proceeding step 6.2.3.2.2.7?")
                    send_email("Crypto-Binance-SellDone", CLIENT.generate_sell_email(sym, order_detail, str_usdt_ntl))
            except (BinanceOrderException, BinanceAPIException) as e:
                print("Binance sell order error step 5.2.3.2.1, continue?")
                send_email("Crypto-Binance-SellOrderError", CLIENT.generate_order_error_mail(e.message))

def run(workbook):
    print("Running sell module")
    sheet = workbook.sheets["Overview"]
    row = 2
    while sheet[f"C{row}"].value is not None:
        call_vb(workbook)
        process_overview_sheet(row, workbook)
        print("######################################################################")
        row+=1

def create_sell_order(pair, quote_qty):
    print("create sell order?")
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
    str_dt = dt.strftime("%d-%b-%y")
    row = sheet.range("A1").end("down").row + 1
    sheet.cells(row, 1).value = _id # column A
    sheet.cells(row, 4).value=to # column D 
    sheet.cells(row, 5).value=str_dt # column E
    sheet.cells(row, 8).value=usdt_profit # column G
    sheet.cells(row, 10).value=-float(qty) # column J
    print("check binance sheet?")
    return row
