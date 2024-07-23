import requests
from utils import read_sheet, color_print, CLIENT, send_email, format_usd_ntl, format_currency
import time
from pprint import pprint
from datetime import datetime


def update_market_price():
    url = "https://api.binance.com/api/v1/ticker/price"
    res = requests.get(url)
    data = res.json()
    if type(data) is dict:
        msg = data["msg"]
        raise ValueError(msg)
    result = {}
    for entry in data:
        pair = entry["symbol"]
        price = entry["price"]
        if pair.endswith("USDT"):
            symbol = pair[:-len("USDT")]
            result[symbol]=float(price)
    with read_sheet("Analysis") as sheet:
        row = 2
        rows = len(sheet["B"])
        while row <= rows:
            sym = sheet[f"B{row}"].value
            # TODO: what if no price get from here?
            # print(f"{sym} price:", result.get(sym))
            sheet[f"E{row}"] = result.get(sym)
            row+=1
    return result


def process_overview_sheet(market_price_map:dict):
    email_body = []
    with read_sheet("Overview", False) as sheet:
        row = 2
        rows = len(sheet["C"])
        while row <= rows:
            _id = sheet[f"A{row}"].value
            sym = sheet[f"C{row}"].value
            pair = f"{sym}USDT"
            print(f"processing symbol={sym}")
            url = f"https://api.binance.com/api/v1/ticker/price?symbol={pair}"
            res = requests.get(url)
            data = res.json()
            # show info
            color_print(f"Response: {data}")
            prx = float(data["price"])
            exp_prx = market_price_map[sym]
            color_print(f"expected price:{exp_prx}")
            exp_diff = float(sheet[f"EX{row}"].value)
            color_print(f"expected diff: {exp_diff}")
            diff = (prx-exp_prx)/exp_prx
            if diff > exp_diff:
                color_print(f"calculated difference: {diff}", "red")
            else:
                color_print(f"calculated difference: {diff}", "green")

            str_usdt_ntl = sheet[f"M{row}"].value
            order_detail = create_sell_order(pair, str_usdt_ntl)
            if order_detail:
                # checked that executedQty is the base qty
                process_binance_sheet(_id, datetime.today(), str_usdt_ntl, -float(order_detail["executedQty"]))
                email_body.append(order_detail)
            row+=1
            print("######################################################################")
    
    #TODO: send out email here


def create_sell_order(pair, quote_qty):
    if quote_qty=="-":
        color_print("Illegal number of quote qty")
        return None
    
    is_sell = input("wanna sell? (Yes/No)")
    if is_sell=="Yes":
        color_print(f"Creating selling order of {pair} with {quote_qty} and wait for 5 seconds...")
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
        color_print("Order confirmation:")
        pprint(resp, indent=4)
        return resp
    
    elif is_sell=="No":
        color_print("Wait for 5 seconds...")
        time.sleep(5)
        return None
    else:
        print(f"Illegal input: {input}")
        return None


def process_binance_sheet(_id:str, dt:datetime, usdt_profit:str, qty:float, to="USDT"):
    with read_sheet("Binance") as sheet:
        row = len(sheet["A"])+1
        str_dt = dt.strftime("%d-%b-%y")
        sheet.cell(row=row, column=1, value=_id) # column A
        sheet.cell(row=row, column=4, value=to) # column D 
        sheet.cell(row=row, column=5, value=str_dt) # column E
        sheet.cell(row=row, column=8, value=format_currency(usdt_profit)) # column H
        sheet.cell(row=row, column=10, value=qty) # column J


result = update_market_price()
process_overview_sheet(result)