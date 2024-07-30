import requests
from utils import CLIENT, send_email, format_usd_ntl, call_vb, read_and_save_workbook, PROXY
import time
from pprint import pprint
from datetime import datetime
from xlwings import Book, Sheet
from binance.exceptions import BinanceAPIException, BinanceOrderException

def process_overview_sheet():
    with read_and_save_workbook() as workbook:
        sheet = workbook.sheets["Overview"]
        row = 2
        while sheet[f"C{row}"].value is not None:
            exch = sheet[f"AG{row}"].value
            str_usdt_ntl = sheet[f"M{row}"].value
            if exch != "Binance":
                print("not from binance")
            elif str_usdt_ntl=="-":
                print("Illegal number of quote qty")
            else:
                print("called vb on step 3, continue?")
                call_vb(workbook)
                sheet = workbook.sheets["Overview"]
                _id = sheet[f"A{row}"].value
                sym = sheet[f"C{row}"].value
                print("Proceeding step 4.1?")
                marketsell = sheet[f"N{row}"].value
                pair = f"{sym}USDT"
                print("proceeding step 5.1?")
                print(f"processing symbol={sym}")
                url = f"https://api.binance.com/api/v1/ticker/price?symbol={pair}"
                try:
                    if PROXY:
                        print(f"using proxy: {PROXY}")
                        res = requests.get(url, proxies=PROXY)
                    else:
                        res = requests.get(url)
                    data = res.json()
                except Exception as e:
                    data = {"msg": str(e)}

                if data.get("code") or data.get("price") in (0, "0"):
                    print(f"error fetching {pair} price: {data}, continue?")
                    send_email("Crypto-Binance-SellQueryError", CLIENT.generate_error_email(url, data), workbook)
                else:
                    # show info
                    print(f"Response: {data}")
                    prx = float(data["price"])

                    if prx < marketsell:
                        print(f"price from api:{prx} < marketsell:{marketsell}, proceeding step 5.2.2?")
                        send_email("Crypto-Binance-SellReject", CLIENT.generate_reject_email(sym, data["price"], marketsell, url, data), workbook)
                    else:
                        print(f"price from api:{prx} >= marketsell:{marketsell}, proceeding step 5.2.3.1?")
                        try:
                            order_detail = create_sell_order(pair, str_usdt_ntl)
                            if order_detail:
                                print("proceeding step 5.2.3.2.2?")
                                # checked that executedQty is the base qty
                                process_binance_sheet(workbook, _id, datetime.today(), float(order_detail["cummulativeQuoteQty"]), -float(order_detail["executedQty"]))
                                print("proceeding step 5.2.3.2.2.7?")
                                send_email("Crypto-Binance-SellDone", CLIENT.generate_sell_email(sym, order_detail, str_usdt_ntl))
                        except (BinanceOrderException, BinanceAPIException) as e:
                            print("Binance sell order error step 5.2.3.2.1, continue?")
                            send_email("Crypto-Binance-SellOrderError", CLIENT.generate_sell_error_mail(e.message), workbook) 
                workbook.save()
            row+=1
            print("######################################################################")
        total_usd = sheet["AQ2"].value
        total_aed = sheet["AS2"].value
        pl_pct = sheet["AQ4"].value
        aj_usdt = sheet["AN23"].value
        body = CLIENT.generate_balance_email(int(total_usd), int(total_aed), pl_pct, aj_usdt)
        send_email("Crypto-Binance-SellEnd", body)

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


def process_binance_sheet(workbook:Book, _id:str, dt:datetime, usdt_profit:str, qty:float, to="USDT", row=1):
    sheet:Sheet = workbook.sheets["Binance"]
    str_dt = dt.strftime("%d-%b-%y")
    sheet.cells(row, 1).value = _id # column A
    sheet.cells(row, 4).value=to # column D 
    sheet.cells(row, 5).value=str_dt # column E
    sheet.cells(row, 8).value=usdt_profit # column H
    sheet.cells(row, 10).value=qty # column J
    print("check binance sheet?")
    return row+1

process_overview_sheet()