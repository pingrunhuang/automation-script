from pathlib import Path
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import Dict
from binance.client import BaseClient
from colorama import Fore, Style
from requests.sessions import Session as Session
import yaml
from binance.client import Client
import xlwings as xw
import json
import datetime
import requests
import time

with open("config.yaml") as f:
    CONFIGS = yaml.safe_load(f)

# excel_path = "D:\Google Drive\AJ\Excel"
excel_path = CONFIGS["EXCEL_PATH"]

PROXY = CONFIGS.get("PROXY", {})

def generate_table(lines):
    html = """
    <!DOCTYPE html>
    <html>
        <body>
            <table>
                <tbody>
    """
    for line in lines:
        html += "<tr>"
        for column in line:
            html += f'<td style="border: 1px solid #ddd; padding: 8px; text-align: left;">{column}</td>'
        html += "</tr>"
    html += """
            </tbody>
        </table>
      </body>
    </html>
    """
    return html

if PROXY:
    def send_email(subject, body="", html=""):
        print(subject)
        print(body)
        print(html)
else:
    SMTP_SERVER = smtplib.SMTP('smtp.gmail.com', 587)  # Specify your SMTP server and port
    SMTP_SERVER.starttls()  # Secure the connection
    SMTP_SERVER.login(CONFIGS["EMAIL_FROM"], CONFIGS["EMAIL_PASS"])
    def send_email(subject , html="", body=""):
        # Create the email message
        to_emails = CONFIGS["EMAIL_TO"]
        msg = MIMEMultipart("alternative")
        msg['From'] = CONFIGS["EMAIL_FROM"]
        msg['To'] = ', '.join(to_emails)
        msg['Subject'] = subject

        # Attach the body with the msg instance
        if body:
            msg.attach(MIMEText(body, 'plain'))
        if html:
            msg.attach(MIMEText(html, 'html'))
        # Set up the server and send the email
        try:
            
            text = msg.as_string()
            SMTP_SERVER.sendmail(CONFIGS["EMAIL_FROM"], to_emails, text)
            print(f"Email sent successfully:{text}")
        except Exception as e:
            print(f"Failed to send email: {str(e)}")


def format_numbers(num)->str:
    # if type(num) is str:
    #     num = float(num)
    # return "%.2f" % num
    return num


def color_print(msg, color=None):
    if color=="green":
        print(Fore.GREEN+msg)
    elif color=="red":
        print(Fore.RED+msg)
    else:
        print(Style.RESET_ALL+msg)


def timestamp2date(ts:float, format="%d/%m/%y %H:%M:%S"):
    dt = datetime.datetime.fromtimestamp(ts//1000)
    return dt.strftime(format)


def format_usd_ntl(ntl:str)->float:
    # return float(ntl.replace("US$", ""))
    return float(ntl)

def format_currency(ntl:float)->str:
    return "US${:10,.8f}".format(ntl)


def call_vb(wb:xw.Book, macro_name="Sheet9.BinanceRate"):
    print(f"Calling vba {macro_name}......")
    macro = wb.macro(macro_name)
    try:
        macro()
    except Exception as e:
        print(f"some error happen when calling vba: {e}")
        return
    finally:
        print("Done vba")

if CONFIGS["TESTNET"] is True:
    color_print("Running on testnet", "green")


class MyBNCClient(Client):
    sell_params = None
    buy_params = None

    def _get_request_kwargs(self, method, signed: bool, force_params: bool = False, **kwargs):
        kwargs = super()._get_request_kwargs(method, signed, force_params, **kwargs)
        if PROXY:
            kwargs.update({"proxies": PROXY})
        return kwargs

    def create_order(self, **params):
        if params["side"]=="SELL":
            print(f"Creating sell order with parameters: {params}")
            self.sell_params = params
        elif params["side"]=="BUY":
            print(f"Creating sell order with parameters: {params}")
            self.buy_params = params
        else:
            raise ValueError(f"side is illegal: {params['side']}")
        uri = self._create_api_uri("order", True, BaseClient.PUBLIC_API_VERSION)
        print(f"Endpoint: {uri}")    
        return super().create_order(**params)
    
    def generate_reject_email(self, symbol:str, price:str, ntl:str, endpoint:str, response:dict, _id):
        """
        if this is for buy: ntl is columnJ
        if this is for sell: ntl is columnN
        """
        lines = [
            ("ID", _id),
            ("Symbol", symbol),
            ("Price", format_numbers(price)),
            ("Expd Price",format_numbers(ntl)),
            ("",""),
            ("Sent", endpoint),
            ("",""),
            ("Received", json.dumps(response))
        ]
        return generate_table(lines)

    def generate_order_error_mail(self, sym:str, msg:str, _id:str, side:str="SELL"):
        """
        send email with subject: “Crypto-Binance-SellOrderError”, Body:
        - line break
        - full details of 5.2.3.2
        """
        assert side in ("BUY", "SELL")
        lines = [
            ("ID", _id),
            ("Symbol", sym),
            ("Sent", self._create_api_uri('order', True, BaseClient.PUBLIC_API_VERSION)), 
            ("params", json.dumps(self.sell_params) if side=="SELL" else json.dumps(self.buy_params)),
            ("",""),
            ("Received", msg)
        ]
        return generate_table(lines)


    def generate_balance_email(self, AQ2:str, AS2:str, AQ4:str, AN23:str, statsC15:str, start_time:float):
        url = self._create_api_uri("account", signed=False, version=BaseClient.PUBLIC_API_VERSION)
        resp = self.get_asset_balance("USDT")
        print(f"Posting api: {url}, resp: {resp}")
        assert type(resp)==dict
        duration = time.time()-start_time
        dur = duration_formating(duration)
        lines = [
            ("Total-USD", "USD.%.f"%int(AQ2)),
            ("Total-AED", "AED.%.f"%int(AS2)),
            ("P/L%", "{:.1%}".format(float(AQ4))), # 100 percent format
            ("Profit", "USD.%.f"%int(statsC15)),
            ("AJ-USDT", "USDT.%.f"%int(AN23)),
            ("Binance-USDT", "USDT.%.f"%float(resp['free'])),
            ("",""),
            ("Sent", url),
            ("",""),
            ("Received", json.dumps(resp)),
            ("",""),
            ("Runtime", dur)
        ]
        return generate_table(lines)

    def generate_buy_insufficient_email(self, columnC, ColumnI, AN23, _id):
        lines =[
            ("ID", _id),
            ("Symbol", columnC),
            ("Buy", ColumnI),
            ("Cash", AN23)
        ]
        return generate_table(lines)

    def generate_min_insufficient_email(self, columnC, columnH, _min, _id):
        lines = [
            ("ID", _id),
            ("Symbol", columnC),
            ("Min", columnH),
            ("Cash", _min)
        ]
        return generate_table(lines)

    def generate_error_email(self, symbol, url, resp):
        lines = [
            ("Symbol", symbol),
            ("",""),
            ("Sent", url),
            ("",""),
            ("Received", json.dumps(resp))
        ]
        return generate_table(lines)
    
    def _generate_buy_email(self, sym, resp:dict, columnIcell:str, columnJcell:str):
        ts_date = timestamp2date(float(resp['transactTime']))
        prx = resp["fills"][0]["price"]
        lines = [
            ("Symbol", sym),
            ("Buy", format_numbers(resp['cummulativeQuoteQty'])),
            ("Expd Buy", format_numbers(columnIcell)),
            ("Price", prx),
            ("Expd Price", columnJcell),
            ("Qty", format_numbers(resp['executedQty'])),
            ("Date/Time", ts_date),
            ("",""),
            ("Sent", self._create_api_uri('order', True, BaseClient.PUBLIC_API_VERSION)),
            ("Parameters", json.dumps(self.buy_params)),
            ("",""),
            ("Received", json.dumps(resp))
        ]
        return lines
    
    def _generate_min_email(self, sym, resp:dict, columnHcell:str, columnJcell:str):
        ts_date = timestamp2date(float(resp['transactTime'])) 
        prx = resp["fills"][0]["price"]
        lines = [
            ("Symbol", sym),
            ("Min", format_numbers(resp['cummulativeQuoteQty'])),
            ("Expd Min", format_numbers(columnHcell)),
            ("Price", prx),
            ("Expd Price", columnJcell),
            ("Qty", format_numbers(resp['executedQty'])),
            ("Date/Time", ts_date),
            ("",""),
            ("Sent", self._create_api_uri('order', True, BaseClient.PUBLIC_API_VERSION)),
            ("Parameters", json.dumps(self.buy_params)),
            ("",""),
            ("Received", json.dumps(resp))
        ]
        return lines
    
    def _generate_sell_email(self, sym, resp:dict, columnMcell:str, columnNcell:str):
        ts_date = timestamp2date(float(resp['transactTime']))
        prx = resp["fills"][0]["price"]
        lines = [
            ("Symbol", sym),
            ("Sell-Profit", format_numbers(resp['cummulativeQuoteQty'])), 
            ("Expd Sell-Profit", format(columnMcell)), 
            ("Price", prx),
            ("Expd Price", columnNcell),
            ("Qty", format_numbers(resp['executedQty'])),
            ("Date/Time", ts_date),
            ("",""),
            ("Sent", self._create_api_uri('order', True, BaseClient.PUBLIC_API_VERSION)),
            ("Parameters", json.dumps(self.sell_params)),
            ("",""),
            ("Received", json.dumps(resp))
        ]
        return lines

    def _generate_reset_email(self, sym, resp:dict, columnScell:str, columnNcell:str):
        ts_date = timestamp2date(float(resp['transactTime']))
        prx = resp["fills"][0]["price"]
        lines = [
            ("Symbol", sym),
            ("Sell-Reset", format_numbers(resp['cummulativeQuoteQty'])), 
            ("Expd Sell-Reset", columnScell), 
            ("Price", prx),
            ("Expd Price", columnNcell),
            ("Qty", format_numbers(resp['executedQty'])),
            ("Date/Time", ts_date),
            ("",""),
            ("Sent", self._create_api_uri('order', True, BaseClient.PUBLIC_API_VERSION)),
            ("Parameters", json.dumps(self.sell_params)),
            ("",""),
            ("Received", json.dumps(resp))
        ]
        return lines

    def generate_done_email(self, sym, resp, qty, market_price, email_prefix, _id):
        lines = [("ID", _id)]
        match email_prefix:
            case "Buy-More":
                func = self._generate_buy_email
            case "Buy-Min":
                func = self._generate_min_email
            case "Sell-Profit":
                func = self._generate_sell_email
            case "Sell-Reset":
                func = self._generate_reset_email
        lines += func(sym, resp, qty, market_price)
        return generate_table(lines)


def fetch_market_price(sym, email_prefix:str):
    pair = f"{sym}USDT"
    url = f"https://api.binance.com/api/v1/ticker/price?symbol={pair}"
    print("proceeding step 6.1?")
    print(f"processing pair={pair}: url={url}")
    try:
        if PROXY:
            print(f"using proxy: {PROXY}")
            res = requests.get(url, proxies=PROXY)
        else:
            res = requests.get(url)
        data = res.json()
        print(f"Response: {data}")
        if data.get("msg") or data.get("price") in (0, "0"):
            raise TimeoutError(data["msg"])
        return data, url
    except Exception as e:
        data = {"msg": str(e)}
        print(f"error fetching {pair} price: {data}, continue?")
        send_email(f"Crypto-Binance-{email_prefix}-QueryError", CLIENT.generate_error_email(sym, url, data))
        return {}, ""
    

CLIENT = MyBNCClient(CONFIGS["API_KEY"], CONFIGS["SECRET_KEY"], testnet=CONFIGS["TESTNET"])

def duration_formating(duration:int):
    dur = datetime.timedelta(seconds=duration)
    seconds = dur.seconds
    minutes = (seconds % 3600) // 60
    seconds = seconds % 60
    if minutes<10:
        s_minutes = f"0{minutes}"
    else:
        s_minutes = str(minutes)
    if seconds<10:
        s_seconds = f"0{seconds}"
    else:
        s_seconds = str(seconds)
    return f"{s_minutes}:{s_seconds}"