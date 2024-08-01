from utils import CLIENT, send_email, format_usd_ntl, call_vb, fetch_market_price
import buy_module
import time
from pprint import pprint
from datetime import datetime
from xlwings import Book, Sheet
from binance.exceptions import BinanceAPIException, BinanceOrderException

def run(workbook):
    print("running reset module")
    call_vb(workbook)
    sheet = workbook.sheets["Overview"]
    AS27 = sheet["AS27"].value
    if AS27 == 0 or AS27 is None:
        buy_module.run(workbook)
    else:
        row = 2
        while sheet[f"C{row}"].value is not None:
            if sheet[f"P{row}"].value==100:
                reset(row, workbook)
            row+=1

def reset(row, workbook):
    sheet = workbook.sheets["Overview"]
    exch = sheet[f"AG{row}"].value
    columnS = sheet[f"S{row}"].value
    if exch != "Binance":
        print("not from binance")
    else:
        print("called vb on step 5, continue?")
        sheet = workbook.sheets["Overview"]
        _id = sheet[f"A{row}"].value
        sym = sheet[f"C{row}"].value
        marketsell = sheet[f"N{row}"].value
        pair = f"{sym}USDT"
        data, url = fetch_market_price(pair, workbook)
        if not data:
            return
        prx = float(data["price"])
        if prx < marketsell:
            print(f"price from api:{prx} < marketsell:{marketsell}, proceeding step 6.2.2?")
            send_email("Crypto-Binance-ResetReject", CLIENT.generate_reject_email(sym, data["price"], marketsell, url, data), workbook)
        else:
            print(f"price from api:{prx} >= marketsell:{marketsell}, proceeding step 6.2.3.1?")
            try:
                order_detail = create_sell_order(pair, columnS)
                if order_detail:
                    print("proceeding step 6.2.3.2.2.1?")
                    send_email("Crypto-Binance-ResetDone", CLIENT.generate_reset_email(sym, order_detail, columnS))
                    # checked that executedQty is the base qty
                    print("proceeding step 6.2.3.2.2.2?")
                    process_binance_sheet(workbook, _id, datetime.today(), order_detail["cummulativeQuoteQty"], order_detail["executedQty"])
                    
            except (BinanceOrderException, BinanceAPIException) as e:
                print("Binance sell order error step 6.2.3.2.1, continue?")
                send_email("Crypto-Binance-ResetOrderError", CLIENT.generate_order_error_mail(e.message), workbook)

    print("######################################################################")


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
    sheet.cells(row, 8).value=usdt_profit # column H
    sheet.cells(row, 10).value=-float(qty) # column J
    print("check binance sheet?")
    return row
