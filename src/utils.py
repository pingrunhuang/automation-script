from pathlib import Path
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from contextlib import contextmanager
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

if PROXY:
    def send_email(subject:str, body:str):
        print(subject)
        print(body)
else:
    SMTP_SERVER = smtplib.SMTP('smtp.gmail.com', 587)  # Specify your SMTP server and port
    SMTP_SERVER.starttls()  # Secure the connection
    SMTP_SERVER.login(CONFIGS["EMAIL_FROM"], CONFIGS["EMAIL_PASS"])
    def send_email(subject:str, body:str):
        # Create the email message
        to_emails = CONFIGS["EMAIL_TO"]
        msg = MIMEMultipart()
        msg['From'] = CONFIGS["EMAIL_FROM"]
        msg['To'] = ', '.join(to_emails)
        msg['Subject'] = subject

        # Attach the body with the msg instance
        msg.attach(MIMEText(body, 'plain'))

        # Set up the server and send the email
        try:
            
            text = msg.as_string()
            SMTP_SERVER.sendmail(CONFIGS["EMAIL_FROM"], to_emails, text)
            print(f"Email sent successfully:{text}")
        except Exception as e:
            print(f"Failed to send email: {str(e)}")


@contextmanager
def read_and_save_workbook():
    workbook = xw.Book(excel_path)
    app = workbook.app
    yield workbook
    workbook.save()
    workbook.close()
    app.kill()
    if not PROXY:
        SMTP_SERVER.close()



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


def call_vb(wb:xw.Book):
    print("Calling vba......")
    macro = wb.macro("Sheet9.BinanceRate")
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
        print(f"Endpoing: {uri}")    
        return super().create_order(**params)
    
    def generate_reject_email(self, symbol:str, price:str, ntl:str, endpoint:str, response:dict):
        """
        if this is for buy: ntl is columnJ
        if this is for sell: ntl is columnN
        """
        lines = [
            f"Symbol:       {symbol}",
            f"Price:        {price}",
            f"Expd Price:   {ntl}\n",
            f"Sent:         {endpoint}\n",
            f"Received:     {json.dumps(response)}"
        ]
        return "\n".join(lines)

    def generate_sell_email(self, sym, resp:dict, columnMcell:str):
        ts_date = timestamp2date(float(resp['transactTime']))
        lines = [
            f"Symbol:           {sym}",
            f"Sell-Profit:      {resp['cummulativeQuoteQty']}", 
            f"Expd Sell-Profit: {columnMcell}", 
            f"Qty:              {resp['executedQty']}",
            f"Date/Time:        {ts_date}\n",
            f"Sent:             {self._create_api_uri('order', True, BaseClient.PUBLIC_API_VERSION)}"
            f"Parameters:       {json.dumps(self.sell_params)}\n",
            f"Received:         {json.dumps(resp)}"
        ]
        return "\n".join(lines)

    def generate_reset_email(self, sym, resp:dict, columnScell:str):
        ts_date = timestamp2date(float(resp['transactTime']))
        lines = [
            f"Symbol:           {sym}",
            f"Sell-Reset:       {resp['cummulativeQuoteQty']}", 
            f"Expd Sell-Reset:  {columnScell}", 
            f"Qty:              {resp['executedQty']}",
            f"Date/Time:        {ts_date}\n",
            f"Sent:             {self._create_api_uri('order', True, BaseClient.PUBLIC_API_VERSION)}"
            f"Parameters:       {json.dumps(self.sell_params)}\n",
            f"Received:         {json.dumps(resp)}"
        ]
        return "\n".join(lines)

    def generate_order_error_mail(self, msg:str):
        """
        send email with subject: “Crypto-Binance-SellOrderError”, Body:
        - line break
        - full details of 5.2.3.2
        """
        lines = [
            f"Sent:     {self._create_api_uri('order', True, BaseClient.PUBLIC_API_VERSION)}\n"
            f"params:   {json.dumps(self.sell_params)}\n",
            f"Received: {msg}"
        ]
        return "\n".join(lines)


    def generate_balance_email(self, AQ2, AS2, AQ4, AN23, start_time):
        url = self._create_api_uri("account", signed=False, version=BaseClient.PUBLIC_API_VERSION)
        resp = self.get_asset_balance("USDT")
        print(f"Posting api: {url}, resp: {resp}")
        assert type(resp)==dict
        duration = time.time()-start_time
        dur = duration_formating(duration)
        lines = [
            f"Total-USD:        {int(AQ2)}",
            f"Total-AED:        {int(AS2)}",
            f"P/L%:             {int(AQ4*100)}%",
            f"AJ-USDT:          {int(AN23)}",
            f"Binance-USDT:     {int(float(resp['free']))}\n",
            f"Sent:             {url}\n",
            f"Received:         {json.dumps(resp)}\n",
            f"Runtime:          {dur}"
        ]
        return "\n".join(lines)

    def generate_buy_insufficient_email(self, columnC, ColumnI, AN23):
        lines =[
            f"Symbol: {columnC}",
            f"Buy:    {ColumnI}",
            f"Cash:   {AN23}"
        ]
        return "\n".join(lines)

    def generate_error_email(self, url, resp):
        lines = [
            f"- Sent:     {url}\n",
            f"- Received: {json.dumps(resp)}"
        ]
        return "\n".join(lines)
    
    def generate_buy_email(self, sym, resp:dict, columnIcell:str):
        ts_date = timestamp2date(float(resp['transactTime']))
        lines = [
            f"Symbol:     {sym}",
            f"Buy         {resp['cummulativeQuoteQty']}",
            f"Expd Buy:   {columnIcell}",
            f"Qty:        {resp['executedQty']}",
            f"Date/Time:  {ts_date}\n"
            f"\nSent:     {self._create_api_uri('order', True, BaseClient.PUBLIC_API_VERSION)}",
            f"Parameters: {json.dumps(self.buy_params)}\n",
            f"\nReceived: {json.dumps(resp)}\n"
        ]
        return "\n".join(lines)


def fetch_market_price(pair, module:str="sell"):
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
        if module=="sell":
            send_email("Crypto-Binance-SellQueryError", CLIENT.generate_error_email(url, data))
        elif module=="reset":
            send_email("Crypto-Binance-ResetQueryError", CLIENT.generate_error_email(url, data))
        elif module=="buy":
            send_email("Crypto-Binance-BuyQueryError", CLIENT.generate_error_email(url, data))
        return {}, ""
    

CLIENT = MyBNCClient(CONFIGS["API_KEY"], CONFIGS["SECRET_KEY"], testnet=CONFIGS["TESTNET"])

def duration_formating(duration:int):
    dur = datetime.timedelta(seconds=duration)
    days, seconds = dur.days, dur.seconds
    hours = days * 24 + seconds // 3600
    minutes = (seconds % 3600) // 60
    seconds = seconds % 60
    return f"{hours}:{minutes}:{seconds}"