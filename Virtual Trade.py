import os
import sys
import time
import json
import ctypes
import threading
import webbrowser
import requests
import pyotp
import keyboard
import pandas as pd
from datetime import datetime
from pathlib import Path
from threading import Thread
import upstox_client
from upstox_client.rest import ApiException
import xlwings as xw

import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

dict_lock = threading.Lock()
excel_lock = threading.Lock()
access = None
no_login_data_counter=10
##############################################################################
def enable_ansi_support():
    if os.name == 'nt':  # Check if the OS is Windows
        kernel32 = ctypes.windll.kernel32
        hStdOut = kernel32.GetStdHandle(-11)  # Get handle to standard output
        mode = ctypes.c_uint32()
        kernel32.GetConsoleMode(hStdOut, ctypes.byref(mode))
        mode.value |= 0x0004  # Enable virtual terminal processing
        kernel32.SetConsoleMode(hStdOut, mode)

enable_ansi_support()

tdate = datetime.now().date()
code = None

base_dir = Path(__file__).resolve().parent
while base_dir.name != "Live Virtual Trade - Scalping":
    if base_dir.parent == base_dir:
        raise FileNotFoundError("'Live Virtual-Scalping Trade - Upstox' folder not found in path hierarchy.")
    base_dir = base_dir.parent


def show_totp(secret):
    totp = pyotp.TOTP(secret)
    otp = totp.now()
    return otp


if not os.path.exists('Credentials/login_details.json'):
    print("User Details not found. First Create a User Base & Retry. Exiting program in 10 Seconds:", end=" ")
    while no_login_data_counter >= 0:
        print(no_login_data_counter, end=" ", flush=True)
        no_login_data_counter -= 1
        time.sleep(1)
    print()  # just to move to a new line after countdown
    sys.exit()

with open('Credentials/login_details.json', 'r') as file_read:
    users_data = json.load(file_read)

allowed_namess = users_data.keys()
allowed_names = [name.lower() for name in allowed_namess]

name_dict = {}

for i in range(len(allowed_names)):
    name_dict[f'{allowed_names[i]}'] = f'{tdate}_access_code_{allowed_names[i]}.json'

name_list = name_dict.values()

os.makedirs("Credentials/Data", exist_ok=True)
os.makedirs("Credentials/Trade_Log", exist_ok=True)

file_list = os.listdir(f'Credentials/Data')

for name in name_list:
    if name in file_list:
        with open(f'Credentials/Data/{name}', 'r') as file_read:
            access = json.load(file_read)
            acc_name = name[23:][:-5]

if not access:

    while True:
        acc_name = input(f'\nEnter Name of Account Holder to Login From {list(allowed_namess)} : ').lower()
        if acc_name in allowed_names:
            break
        else:
            print(f"\nInvalid User. Please Enter Registered User Name {list(allowed_namess)}'.")

    try:
        with open(f'Credentials/Data/{tdate}_access_code_{acc_name}.json', 'r') as file_read:
            access = json.load(file_read)

    except:

        with open('Credentials/login_details.json', 'r') as file_read:
            login_details = json.load(file_read)

        api_key = login_details[f'{acc_name.capitalize()}']['api_key']
        api_secret = login_details[f'{acc_name.capitalize()}']['api_secret']
        api_auth = login_details[f'{acc_name.capitalize()}']['api_auth']
        api_pin = login_details[f'{acc_name.capitalize()}']['pin']
        mobile_no = login_details[f'{acc_name.capitalize()}']['Mob No.']
        hold_name = login_details[f'{acc_name.capitalize()}']['full_name']

        print(f'\nTrying to Login from Account Holder: {hold_name}')

        uri = 'https://www.google.com/'
        url1 = f'https://api.upstox.com/v2/login/authorization/dialog?response_type=code&client_id={api_key}&redirect_uri={uri}\n'

        options = uc.ChromeOptions()
        options.headless = True
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        driver = uc.Chrome(options=options)

        # driver = uc.Chrome() # Use this line instead to run Chrome in normal (visible) mode, (In that case, comment out the 5 lines above that set headless options)

        driver.get(url1)
        wait = WebDriverWait(driver, 20)
        phone_input = wait.until(EC.presence_of_element_located((By.ID, "mobileNum")))
        phone_input.send_keys(mobile_no)
        otp_button = wait.until(EC.element_to_be_clickable((By.ID, "getOtp")))
        otp_button.click()
        # print("✅ Phone number entered, now captcha should appear normally")

        totp_value = show_totp(api_auth)
        totp_input = wait.until(EC.presence_of_element_located((By.ID, "otpNum")))
        totp_input.send_keys(totp_value)
        proceed_button = wait.until(EC.element_to_be_clickable((By.ID, "continueBtn")))
        proceed_button.click()
        # print("✅ TOTP entered and Continue clicked!")

        pin_input = wait.until(EC.presence_of_element_located((By.ID, "pinCode")))
        pin_input.send_keys(api_pin)
        proceed_button = wait.until(EC.element_to_be_clickable((By.ID, "pinContinueBtn")))
        proceed_button.click()

        # print("✅ PIN entered and proceed button clicked!")
        time.sleep(3)
        code_url = driver.current_url

        driver.quit()

        start = code_url.find('code=')
        if start != -1:
            start =start + 5  # move past 'code='
            code = code_url[start:start+6]
        else:
            print("No code found in the URL")

        url = 'https://api.upstox.com/v2/login/authorization/token'
        headers = {
            'accept': 'application/json',
            'Content-Type': 'application/x-www-form-urlencoded',
        }

        data = {
            'code': code,
            'client_id': api_key,
            'client_secret': api_secret,
            'redirect_uri': uri,
            'grant_type': 'authorization_code',
        }

        response = requests.post(url, headers=headers, data=data)
        access = response.json()['access_token']
        print(f'\nLogin Successful, Status Code : {response.status_code}')
        print(f"User Name : {response.json()['user_name']}\nEmail ID : {response.json()['email']}")

        with open(f'Credentials/Data/{tdate}_access_code_{acc_name}.json', 'w') as file_write:
            json.dump(access, file_write)

print(f'\nLogin Successful from Account : {acc_name.capitalize()}')
hold_name = users_data[f'{acc_name.capitalize()}']['full_name']

streamer = None
live_data = {}
configuration = upstox_client.Configuration()
configuration.access_token = access
def on_message(message):
    # print(message)
    global live_data
    dict_data = message
    if 'feeds' in dict_data:
        data = dict_data['feeds']
        for key, value in data.items():
            ltp = value['ltpc']['ltp']

            with dict_lock:
                live_data[key] = ltp

def start_stream():
    global streamer, final_list, configuration

    streamer = upstox_client.MarketDataStreamerV3(
        upstox_client.ApiClient(configuration), final_list, "ltpc")

    streamer.on("message", on_message)
    streamer.connect()

def main():
    thread = threading.Thread(target=start_stream)
    thread.start()

def lot_size():
    df = pd.read_csv('Credentials/instrument.csv')

    df_nifty = df[(df['exchange'] == 'NSE_FO') & (df['instrument_type'] == 'OPTIDX') & (df['name'] == 'NIFTY')]
    expiry_list_nifty = df_nifty['expiry'].unique().tolist()
    expiry_list_nifty.sort()
    nifty_lot_size = df_nifty[df_nifty['expiry'] == expiry_list_nifty[0]].reset_index(drop=True).loc[0, 'lot_size']

    df_bnf = df[(df['exchange'] == 'NSE_FO') & (df['instrument_type'] == 'OPTIDX') & (df['name'] == 'BANKNIFTY')]
    expiry_list_bnf = df_bnf['expiry'].unique().tolist()
    expiry_list_bnf.sort()
    bnf_lot_size = df_bnf[df_bnf['expiry'] == expiry_list_bnf[0]].reset_index(drop=True).loc[0, 'lot_size']

    df_sensex = df[(df['exchange'] == 'BSE_FO') & (df['instrument_type'] == 'OPTIDX') & (df['name'] == 'SENSEX')]
    expiry_list_sensex = df_sensex['expiry'].unique().tolist()
    expiry_list_sensex.sort()
    sensex_lot_size = df_sensex[df_sensex['expiry'] == expiry_list_sensex[0]].reset_index(drop=True).loc[0, 'lot_size']

    return {'nifty': nifty_lot_size, 'bnf': bnf_lot_size, 'sensex': sensex_lot_size}

def option_chain(instrument_key,expiry_date,inst,ocs):
    global access
    url = 'https://api.upstox.com/v2/option/chain'
    params = {
            'instrument_key': instrument_key,
            'expiry_date': expiry_date
    }
    headers = {
        'Accept': 'application/json',
        'Authorization': f'Bearer {access}'
    }

    response = requests.get(url, params=params, headers=headers)
    time.sleep(1)
    time_stamp = datetime.now().strftime("%H:%M:%S")
    option = response.json()
    option_df = pd.json_normalize(option['data'])
    option_df = option_df[['expiry', 'strike_price', 'underlying_spot_price', 'call_options.instrument_key', 'call_options.market_data.ltp',  'put_options.instrument_key', 'put_options.market_data.ltp', ]]
    option_df = option_df.rename(columns={'call_options.instrument_key' : 'CE_instrument_key', 'call_options.market_data.ltp' : 'CE_ltp', 'put_options.instrument_key' : 'PE_instrument_key', 'put_options.market_data.ltp' : 'PE_ltp', 'underlying_spot_price' : 'spot_price'})
    option_df[['signal_ce', 'signal_pe']] = None

    lot_index = lot_size()

    if instrument_key == 'NSE_INDEX|Nifty 50':
        option_df[['lotsize', 'Index']] = [int(lot_index['nifty']), 'Nifty 50']
    elif instrument_key == 'NSE_INDEX|Nifty Bank':
        option_df[['lotsize', 'Index']] = [int(lot_index['bnf']), 'Bank Nifty']
    else:
        option_df[['lotsize', 'Index']] = [int(lot_index['sensex']), 'Sensex']

    option_df['symbol_ce'] = option_df['strike_price'].astype(str) + '_CE'
    option_df['symbol_pe'] = option_df['strike_price'].astype(str) + '_PE'
    
    option_df = option_df[['Index','expiry','lotsize','CE_instrument_key' ,'symbol_ce','CE_ltp','signal_ce','strike_price','signal_pe','PE_ltp','symbol_pe','PE_instrument_key','spot_price']]

    option_df['diff'] = abs(option_df['spot_price'] - option_df['strike_price'])
    ce = option_df.loc[option_df['diff'].idxmin(),'CE_ltp']
    strike = option_df.loc[option_df['diff'].idxmin(),'strike_price']
    pe = option_df.loc[option_df['diff'].idxmin(),'PE_ltp']

    fut_spot_price = ce-pe+strike

    option_df['spot_price'] = fut_spot_price
    option_df['diff'] = abs(option_df['spot_price'] - option_df['strike_price'])
    atm_strike = option_df.loc[option_df['diff'].idxmin(), 'strike_price']

    ce_atm_ltp = option_df[option_df['strike_price'] == atm_strike].iloc[0]['CE_ltp']
    pe_atm_ltp = option_df[option_df['strike_price'] == atm_strike].iloc[0]['PE_ltp']

    x = option_df['strike_price'].diff().mode()[0]
    upper_limit = atm_strike + inst*x
    lower_limit = atm_strike - inst*x
    option_df = option_df[(option_df['strike_price'] >= lower_limit) & (option_df['strike_price'] <= upper_limit)]
    strike_list = option_df['strike_price'].tolist()
    list1 = option_df['CE_instrument_key'].tolist()
    list2 = option_df['PE_instrument_key'].tolist()
    t_list = list1 + list2
    return t_list, strike_list

def instrument():
    inst_url = 'https://assets.upstox.com/market-quote/instruments/exchange/complete.csv.gz'
    instrument = pd.read_csv(inst_url)
    instrument.to_csv('Credentials/instrument.csv')


def update_subscription_list(inst):
    global expiry_list_nifty, expiry_list_bnf, expiry_list_sensex
    instrument_key_nifty = 'NSE_INDEX|Nifty 50'
    instrument_key_bnf = 'NSE_INDEX|Nifty Bank'
    instrument_key_sensex = 'BSE_INDEX|SENSEX'
    index_ltp = [instrument_key_nifty, instrument_key_bnf, instrument_key_sensex]

    nifty_0_list, nifty_strikes = option_chain(instrument_key_nifty,expiry_list_nifty[0],inst,ocs=0)
    bnf_0_list, bnf_strikes = option_chain(instrument_key_bnf,expiry_list_bnf[0],inst,ocs=0)
    sensex_0_list, sensex_strikes = option_chain(instrument_key_sensex,expiry_list_sensex[0],inst,ocs=0)

    final_list = nifty_0_list + bnf_0_list + sensex_0_list + index_ltp
    return final_list, nifty_strikes, bnf_strikes, sensex_strikes

##############################################################################
try:
    with open(f'Credentials/Data/{tdate}_inputs.json', 'r') as file_read:
        inputs = json.load(file_read)
        ref_inst = inputs['instrument']
        sub_list = inputs['subscription']

except:
    while True:
        # ref_inst = input('Do you want to refresh Instrument Data : 1 / 0 : ')
        ref_inst = '1'
        if ref_inst == '1' or ref_inst == '0':
            break
        else:
            print('Invalid Input, Enter either 1 or 0')

    while True:
        # sub_list = input('\nDo you want to Update Subscription List : 1 / 0 : ')
        sub_list = '1'
        if sub_list == '1' or sub_list == '0':
            break
        else:
            print("\nInvalid Selection. Please enter either '0' or '1'.")

    inputs = {'instrument': 0, 'subscription':0 }

    with open(f'Credentials/Data/{tdate}_inputs.json', 'w') as file_write:
        json.dump(inputs, file_write)

##############################################################################
if ref_inst == '1' :
        instrument()
        print('################---->| Instrument Data Updated |<----################')
else:
    pass

df = pd.read_csv('Credentials/instrument.csv')

df_nifty = df[(df['exchange'] == 'NSE_FO') & (df['instrument_type'] == 'OPTIDX') & (df['name'] == 'NIFTY')]
expiry_list_nifty = df_nifty['expiry'].unique().tolist()
expiry_list_nifty.sort()
nifty_lot_size = df_nifty[df_nifty['expiry'] == expiry_list_nifty[0]].reset_index(drop=True).loc[0, 'lot_size']

df_bnf = df[(df['exchange'] == 'NSE_FO') & (df['instrument_type'] == 'OPTIDX') & (df['name'] == 'BANKNIFTY')]
expiry_list_bnf = df_bnf['expiry'].unique().tolist()
expiry_list_bnf.sort()
bnf_lot_size = df_bnf[df_bnf['expiry'] == expiry_list_bnf[0]].reset_index(drop=True).loc[0, 'lot_size']

df_sensex = df[(df['exchange'] == 'BSE_FO') & (df['instrument_type'] == 'OPTIDX') & (df['name'] == 'SENSEX')]
expiry_list_sensex = df_sensex['expiry'].unique().tolist()
expiry_list_sensex.sort()
sensex_lot_size = df_sensex[df_sensex['expiry'] == expiry_list_sensex[0]].reset_index(drop=True).loc[0, 'lot_size']

##############################################################################

inst = 5 # ATMs +- OTMs you require for each expiry to Subscribe in Websocket

##############################################################################
if sub_list == '1':
    final_list, nifty_strikes, bnf_strikes, sensex_strikes = update_subscription_list(inst)
    sub_list_strikes = {'sub_list':final_list, 'nifty_strikes':nifty_strikes, 'bnf_strikes':bnf_strikes, 'sensex_strikes':sensex_strikes}
    print('##########---->| Websocket Subscription List Updated |<----##########')
    with open('Credentials/final_list.json', 'w') as file_write:
        json.dump(sub_list_strikes, file_write, indent=4)
else:
    try:
        with open('Credentials/final_list.json', 'r') as file_read:
            sub_list_strikes = json.load(file_read)
            final_list = sub_list_strikes['sub_list']
            nifty_strikes = sub_list_strikes['nifty_strikes']
            bnf_strikes = sub_list_strikes['bnf_strikes']
            sensex_strikes = sub_list_strikes['sensex_strikes']
    except:
        final_list, nifty_strikes, bnf_strikes, sensex_strikes = update_subscription_list(inst)
        sub_list_strikes = {'sub_list':final_list, 'nifty_strikes':nifty_strikes, 'bnf_strikes':bnf_strikes, 'sensex_strikes':sensex_strikes}
        with open('Credentials/final_list.json', 'w') as file_write:
            json.dump(sub_list_strikes, file_write, indent=4)
            print('Subscription List File Not Found, but now Created & Updated')
###############################################################################


if __name__ == "__main__":
    main()

while not live_data:
    print("Waiting for live data to populate...")
    time.sleep(1)

# app = xw.App(visible=True, add_book=False)
# app.display_alerts = False
# wb = app.books.open(f'Scapling.xlsx')

wb = xw.Book('Scalping - Virtual.xlsx')
nifty = wb.sheets('Nifty')
bnf = wb.sheets('Bank-Nifty')
sensex = wb.sheets('Sensex')

condition = True
def hello():
    # time.sleep(2)
    global condition
    condition = False

run_ce = False
run_pe = False
run_both = False
sell_sig = False
set_target_2 = False
set_target_0 = False

def set_flag(option):
    time.sleep(0.5)
    global run_ce, run_pe, run_both, sell_sig, set_target_2, set_target_0
    if option == 'ce':
        run_ce = True
    elif option == 'pe':
        run_pe = True
    elif option == 'both' :
        run_both = True

    elif option == 'ce_s':
        run_ce = True
        sell_sig = True
    elif option == 'pe_s':
        run_pe = True
        sell_sig = True
    elif option == 'both_s' :
        run_both = True
        sell_sig = True

    elif option == 'ce_t2':
        run_ce = True
        set_target_2 = True
    elif option == 'pe_t2':
        run_pe = True
        set_target_2 = True
    elif option == 'both_t2' :
        run_both = True
        set_target_2 = True

    elif option == 'ce_t0':
        run_ce = True
        set_target_0 = True
    elif option == 'pe_t0':
        run_pe = True
        set_target_0 = True
    elif option == 'both_t0' :
        run_both = True
        set_target_0 = True



def buy(tradevalue, lot, index):
    
    multiplier = ((lot + 23) // 24) if index=='Nifty' else ((lot + 16) // 17) if index=='Bank-Nifty' else ((lot + 49) // 50)

    brokerage = 10*multiplier
    transaction_charge = 0.0003503 * tradevalue
    sebi_charge = 0.000001 * tradevalue
    gst = 0.18*(brokerage + transaction_charge + sebi_charge)
    stamp_charge = 0.00003*tradevalue
    total = brokerage + transaction_charge + sebi_charge + gst + stamp_charge
    return total

def sell(tradevalue, lot, index):

    if index=='Nifty':
        multiplier = ((lot + 23) // 24)
    elif index=='Bank-Nifty':
        multiplier = ((lot + 16) // 17)
    else:
        multiplier = ((lot + 49) // 50)

    brokerage = 10*multiplier
    transaction_charge = 0.0003503 * tradevalue
    sebi_charge = 0.000001 * tradevalue
    gst = 0.18*(brokerage + transaction_charge + sebi_charge)
    stt = 0.001 * tradevalue
    total = brokerage + stt + transaction_charge + sebi_charge + gst
    return total


def Brokerage_cal(entrydf, lot, index):
    summ = []
    for index_df, row in entrydf.iterrows():
        if index_df == 'BUY':
            xy = buy(row['value'], lot, index)
        elif index_df == 'SELL':
            xy = sell(row['value'], lot, index)
        summ.append(xy)

    return int(sum(summ))

trade_cancel = False

def cancel_trade():
    global trade_cancel
    trade_cancel = True

keyboard.add_hotkey('z+right', lambda: set_flag('ce_t2'))
keyboard.add_hotkey('x+right', lambda: set_flag('pe_t2'))
keyboard.add_hotkey('b+right', lambda: set_flag('both_t2'))

keyboard.add_hotkey('z+left', lambda: set_flag('ce_t0'))
keyboard.add_hotkey('x+left', lambda: set_flag('pe_t0'))
keyboard.add_hotkey('b+left', lambda: set_flag('both_t0'))

keyboard.add_hotkey('z+up', lambda: set_flag('ce'))
keyboard.add_hotkey('x+up', lambda: set_flag('pe'))
keyboard.add_hotkey('b+up', lambda: set_flag('both'))

keyboard.add_hotkey('z+down', lambda: set_flag('ce_s'))
keyboard.add_hotkey('x+down', lambda: set_flag('pe_s'))
keyboard.add_hotkey('b+down', lambda: set_flag('both_s'))

keyboard.add_hotkey('r', lambda: reset_margin_check())

keyboard.add_hotkey('ctrl+delete', lambda: cancel_trade())

last_del_time = 0
double_press_interval = 0.3  # seconds allowed between presses

def on_delete_press(event):
    global last_del_time
    now = time.time()
    
    if run_ce or run_pe or run_both:
        if now - last_del_time <= double_press_interval:
            hello()  # Call your custom function here
            last_del_time = 0
        else:
            last_del_time = now

keyboard.on_press_key("delete", on_delete_press)

def tradelog(df, timing):
    now = datetime.now()
    now_date = now.strftime('%d-%m-%Y')
    now_time = now.time().replace(microsecond=0)
    now_day  = now.strftime('%A')
    time_df = pd.DataFrame([{'Date':now_date, 'Day':now_day, 'Entry Time':timing[0], 'Exit Time':timing[1], 'Duration':timing[2]}])
    add_df = pd.concat([time_df, df], axis=1).reset_index(drop=True)

    try:
        old_df = pd.read_excel(f'Credentials/Trade_Log/{now_date}_log.xlsx', index_col=0)
    except:
        old_df = pd.DataFrame()

    new_df = pd.concat([old_df, add_df], axis=0).reset_index(drop=True)
    new_df['Cumm'] = new_df['net_profit'].cumsum()
    new_df = new_df[[ "Date", "Day", "Entry Time", "Exit Time", "Duration", "Index", "strike", "type", "qty/lot", "lot", "total_qty", "Points", "min_pts", "max_pts", "profit", "brokerage", "net_profit", "Cumm", "gain", 'margin']]



    new_df.to_excel(f'Credentials/Trade_Log/{now_date}_log.xlsx')

def margin_pnl():
    global check_margin
    profit, brokerage, net_profit, total_profits, total_losses, total_trades, total_pts, pos_sum, neg_sum, trade_acc = (None,) * 10
    now = datetime.now()
    now_date = now.strftime('%d-%m-%Y')
    profit_nos = []
    loss_nos = []
    try:
        df_log = pd.read_excel(f'Credentials/Trade_Log/{now_date}_log.xlsx', index_col=0)
        profit = df_log['profit'].sum()
        brokerage = df_log['brokerage'].sum()
        net_profit = df_log['net_profit'].sum()
        total_trades = len(df_log)
        total_profits = (df_log['net_profit'] > 0).sum()
        trade_acc = round((total_profits/total_trades)*100,2)
        total_losses = (df_log['net_profit'] <= 0).sum()
        total_pts = df_log['Points'].sum()
        pos_sum = df_log.loc[df_log['net_profit']>0, 'net_profit'].sum()
        neg_sum = df_log.loc[df_log['net_profit']<0, 'net_profit'].sum()
    except:
        pass

    nifty.range('I2').value = profit
    nifty.range('I3').value = brokerage
    nifty.range('I4').value = net_profit
    nifty.range('I5').value = total_pts
    
    nifty.range('K2').value = total_profits
    nifty.range('K3').value = total_losses
    nifty.range('K4').value = total_trades

    nifty.range('J6').value = pos_sum
    nifty.range('K6').value = neg_sum

    nifty.range('K5').value = f'{trade_acc} %'


    check_margin = False

def margin_trade(ce,pe,qty):
    global configuration, margin_check

    margin_instance = upstox_client.ChargeApi(upstox_client.ApiClient(configuration))

    # Sell Margin
    inst_cepe_s = [upstox_client.Instrument(instrument_key=ce,quantity=qty,product="D",transaction_type="SELL"),
                 upstox_client.Instrument(instrument_key=pe,quantity=qty,product="D",transaction_type="SELL")]

    inst_ce_s = [upstox_client.Instrument(instrument_key=ce,quantity=qty,product="D",transaction_type="SELL")]

    inst_pe_s = [upstox_client.Instrument(instrument_key=pe,quantity=qty,product="D",transaction_type="SELL")]

    margin_body_cepe_s = upstox_client.MarginRequest(inst_cepe_s)
    margin_body_ce_s = upstox_client.MarginRequest(inst_ce_s)
    margin_body_pe_s = upstox_client.MarginRequest(inst_pe_s)

    margin_cepe_s = margin_instance.post_margin(margin_body_cepe_s).to_dict()['data']['required_margin']
    margin_ce_s = margin_instance.post_margin(margin_body_ce_s).to_dict()['data']['required_margin']
    margin_pe_s = margin_instance.post_margin(margin_body_pe_s).to_dict()['data']['required_margin']

    sheets[index].range('M9').value = margin_ce_s
    sheets[index].range('N9').value = margin_pe_s
    sheets[index].range('O9').value = margin_cepe_s

    # Buy Margin
    inst_cepe_b = [upstox_client.Instrument(instrument_key=ce,quantity=qty,product="D",transaction_type="BUY"),
                 upstox_client.Instrument(instrument_key=pe,quantity=qty,product="D",transaction_type="BUY")]

    inst_ce_b = [upstox_client.Instrument(instrument_key=ce,quantity=qty,product="D",transaction_type="BUY")]

    inst_pe_b = [upstox_client.Instrument(instrument_key=pe,quantity=qty,product="D",transaction_type="BUY")]

    margin_body_cepe_b = upstox_client.MarginRequest(inst_cepe_b)
    margin_body_ce_b = upstox_client.MarginRequest(inst_ce_b)
    margin_body_pe_b = upstox_client.MarginRequest(inst_pe_b)

    margin_cepe_b = margin_instance.post_margin(margin_body_cepe_b).to_dict()['data']['required_margin']
    margin_ce_b = margin_instance.post_margin(margin_body_ce_b).to_dict()['data']['required_margin']
    margin_pe_b = margin_instance.post_margin(margin_body_pe_b).to_dict()['data']['required_margin']

    sheets[index].range('M4').value = margin_ce_b
    sheets[index].range('N4').value = margin_pe_b
    sheets[index].range('O4').value = margin_cepe_b

    margin_sell = [margin_ce_s, margin_pe_s, margin_cepe_s]
    margin_check = False

    return margin_sell


def reset_margin_check():
    global margin_check
    margin_check = True

def normal_brok(qty, lot, ce_ltp, pe_ltp, index_pass):
    ce_buy_value = qty*lot*ce_ltp
    ce_brok_dff = pd.DataFrame({'signal':['BUY', 'SELL'], 'value':[ce_buy_value, ce_buy_value]})
    ce_brok_dff = ce_brok_dff.set_index('signal')
    ce_brok = Brokerage_cal(ce_brok_dff, lot, index_pass)

    pe_buy_value = qty*lot*pe_ltp
    pe_brok_dff = pd.DataFrame({'signal':['BUY', 'SELL'], 'value':[pe_buy_value, pe_buy_value]})
    pe_brok_dff = pe_brok_dff.set_index('signal')
    pe_brok = Brokerage_cal(pe_brok_dff, lot, index_pass)

    cepe_brok_dff = pd.DataFrame({'signal':['BUY', 'BUY', 'SELL', 'SELL'], 'value':[ce_buy_value, pe_buy_value, ce_buy_value, pe_buy_value]})
    cepe_brok_dff = cepe_brok_dff.set_index('signal')
    cepe_brok = Brokerage_cal(cepe_brok_dff, lot, index_pass)

    sheets[index].range('M6').value = -ce_brok
    sheets[index].range('N6').value = -pe_brok
    sheets[index].range('O6').value = -cepe_brok

m=1
flip = True
check_margin = True
margin_check = True
sheets = {'nifty' : nifty, 'bnf':bnf, 'sensex':sensex}
expiry = {'nifty':expiry_list_nifty, 'bnf':expiry_list_bnf, 'sensex':expiry_list_sensex}
segment = {'nifty':'NSE_FO', 'bnf':'NSE_FO', 'sensex':'BSE_FO'}
symbol = {'nifty':'NIFTY', 'bnf':'BANKNIFTY', 'sensex':'SENSEX'}
index_name = {'nifty':'NSE_INDEX|Nifty 50', 'bnf':'NSE_INDEX|Nifty Bank', 'sensex':'BSE_INDEX|SENSEX'}
step_size = {'nifty':50, 'bnf':100, 'sensex':100}
print('#################----------->| Monitoring Started, Ready to Place Trade |<-----------#################')
one_time = True

# Auto Spot ATM Entry in Cell C6 - First Thing
#################################################################################
with dict_lock:
    nifty_spot = live_data[index_name['nifty']]
    bnf_spot = live_data[index_name['bnf']]
    sensex_spot = live_data[index_name['sensex']]

spot_atm_nifty = round(nifty_spot/step_size['nifty'])*step_size['nifty']
spot_atm_bnf = round(bnf_spot/step_size['bnf'])*step_size['bnf']
spot_atm_sensex = round(sensex_spot/step_size['sensex'])*step_size['sensex']

nifty.range('C6').value = spot_atm_nifty
bnf.range('C6').value = spot_atm_bnf
sensex.range('C6').value = spot_atm_sensex

nifty.range('L3').value = [[i] for i in nifty_strikes]
bnf.range('L3').value = [[i] for i in bnf_strikes]
sensex.range('L3').value = [[i] for i in sensex_strikes]

lot_index = lot_size()

nifty.range('C2').value = int(lot_index['nifty'])
bnf.range('C2').value = int(lot_index['bnf'])
sensex.range('C2').value = int(lot_index['sensex'])
#################################################################################

while True :
    active_sheet = wb.sheets.active
    var = {}
    exit_cond=0

    if active_sheet.name in ('Nifty','Bank-Nifty','Sensex'):

        if active_sheet.name == 'Nifty':
            index = 'nifty'
        elif active_sheet.name == 'Bank-Nifty':
            index = 'bnf'
        elif active_sheet.name == 'Sensex':
            index = 'sensex'

        if flip:
            sheets[index].range('H10').color = (0, 255, 0)
            flip = False
        else:
            sheets[index].range('H10').color = None
            flip = True

        if check_margin:
            time.sleep(1)
            margin_pnl()

        curr_time = datetime.now().strftime("%I:%M:%S %p")
        curr_date = datetime.today().date()
        curr_date_str = curr_date.strftime("%d-%m-%Y")
        today_day = curr_date.strftime("%A")
        sheets[index].range('H1').value = f'{active_sheet.name} : {hold_name.title()} (Live Virtual Trade) | Today : {curr_date_str} / {today_day} | {curr_time}'

        try :

            if one_time == True:
                with dict_lock:
                    var[f'{index}_spot'] = live_data[index_name[f'{index}']]
                    spot_atm = round(var[f'{index}_spot']/step_size[index])*step_size[index]
                    sheets[index].range('C6').value = spot_atm
                one_time = False

            var[f'strike_{index}'] = sheets[index].range('C6').value
            var[f'qty_{index}'] = sheets[index].range('C2').value
            var[f'lot_{index}'] = sheets[index].range('C3').value
            var[f'target_{index}'] = sheets[index].range('I8').value
            var[f'sl_{index}'] = sheets[index].range('I9').value

            var[f'ceent_{index}'] = sheets[index].range('B7').value
            var[f'peent_{index}'] = sheets[index].range('D7').value
            var[f'signal_{index}'] = str(sheets[index].range('C7').value or "").lower()

            var[f'target_{index}_leg1'] = sheets[index].range('I12').value
            var[f'sl_{index}_leg1'] = sheets[index].range('J12').value
            var[f'target_{index}_leg2'] = sheets[index].range('I13').value
            var[f'sl_{index}_leg2'] = sheets[index].range('J13').value
            var[f'activate_{index}'] = sheets[index].range('H11').value

            reset_tar = sheets[index].range('F9').value

            var[f'trade_ce_{index}'] = sheets[index].range('E2').value
            var[f'trade_pe_{index}'] = sheets[index].range('F2').value

            var[f'{index}_expiry'] = expiry[f'{index}'][0]

            var[f'df_{index}'] = df[(df['exchange'] == segment[f'{index}']) & (df['instrument_type'] == 'OPTIDX') & (df['name'] == symbol[f'{index}']) & (df['expiry'] == var[f'{index}_expiry']) & (df['strike'] == var[f'strike_{index}'])]

            var[f'{index}_ce'] = var[f'df_{index}'][var[f'df_{index}']['option_type'] == 'CE'].iloc[0]['instrument_key']
            with dict_lock:
                var[f'{index}_ce_ltp'] = live_data[var[f'{index}_ce']]
            sheets[index].range('B6').value = var[f'{index}_ce_ltp']

            var[f'{index}_pe'] = var[f'df_{index}'][var[f'df_{index}']['option_type'] == 'PE'].iloc[0]['instrument_key']
            with dict_lock:
                var[f'{index}_pe_ltp'] = live_data[var[f'{index}_pe']]
                var[f'{index}_spot'] = live_data[index_name[f'{index}']]

            synthetic_spot = var[f'{index}_ce_ltp'] - var[f'{index}_pe_ltp'] + var[f'strike_{index}']
            synthetic_atm = round(synthetic_spot/step_size[index])*step_size[index]

            sheets[index].range('D1').value = f'Synthetic ATM Strike : {synthetic_atm}'
            sheets[index].range('D3').value = f"Order Quantity : {int(var[f'qty_{index}']*var[f'lot_{index}'])}"

            sheets[index].range('D6').value = var[f'{index}_pe_ltp']
            
            sheets[index].range('B1').value = f"{active_sheet.name} SPOT : {var[f'{index}_spot']}"
            sheets[index].range('C4').value = var[f'{index}_expiry']
            curr_exp = datetime.strptime(var[f'{index}_expiry'], "%Y-%m-%d").date()
            exp_day = curr_exp.strftime("%A")            
            dte = (curr_exp - curr_date).days
            sheets[index].range('D4').value = f'Days to Expiry : {dte} ({exp_day})'

        except Exception as e:
            print(f'Error | Strike Not Subscribed:{e}')
            pass

        if margin_check:
            margin_sell = margin_trade(var[f'{index}_ce'], var[f'{index}_pe'], var[f'qty_{index}'])
            normal_brok(var[f'qty_{index}'], var[f'lot_{index}'], var[f'{index}_ce_ltp'], var[f'{index}_pe_ltp'], active_sheet.name)

        if (var[f'signal_{index}'] == 'bbo'): # Buy Break-Out
            if (var[f'ceent_{index}'] != 0) and (var[f'{index}_ce_ltp'] >= var[f'ceent_{index}']):
                set_flag('ce')
            elif (var[f'peent_{index}'] != 0) and (var[f'{index}_pe_ltp'] >= var[f'peent_{index}']):
                set_flag('pe')

        elif var[f'signal_{index}'] == 'sbd': # Sell Break-Down
            if (var[f'ceent_{index}'] != 0) and (var[f'{index}_ce_ltp'] <= var[f'ceent_{index}']):
                set_flag('ce_s')
            elif (var[f'peent_{index}'] != 0) and (var[f'{index}_pe_ltp'] <= var[f'peent_{index}']):
                set_flag('pe_s')

        elif (var[f'signal_{index}'] == 'blo'): # Buy Limit Order
            if (var[f'ceent_{index}'] != 0) and (var[f'{index}_ce_ltp'] <= var[f'ceent_{index}']):
                set_flag('ce')
            elif (var[f'peent_{index}'] != 0) and (var[f'{index}_pe_ltp'] <= var[f'peent_{index}']):
                set_flag('pe')

        elif var[f'signal_{index}'] == 'slo': # Sell Limit Order
            if (var[f'ceent_{index}'] != 0) and (var[f'{index}_ce_ltp'] >= var[f'ceent_{index}']):
                set_flag('ce_s')
            elif (var[f'peent_{index}'] != 0) and (var[f'{index}_pe_ltp'] >= var[f'peent_{index}']):
                set_flag('pe_s')

        entry_time_sec = datetime.now().time().second
        # sheets[index].range('D2').value = entry_time_sec

        if entry_time_sec == 59:
            if var[f'trade_ce_{index}']:
                set_flag('ce')
            elif var[f'trade_pe_{index}']:
                set_flag('pe')

        if run_ce or run_pe:

            if set_target_2:
                sheets[index].range('I8').value = 2
                var[f'target_{index}'] = 2
                set_target_2 = False

            elif set_target_0:
                sheets[index].range('I8').value = 0
                var[f'target_{index}'] = 0
                set_target_0 = False



            dt1 = datetime.now().replace(microsecond=0)
            entry_time = dt1.time()

            sheets[index].range('B7').value = 0
            sheets[index].range('D7').value = 0

            # sheets[index].range('E2').value = 0
            # sheets[index].range('F2').value = 0
            
            print("Executing CE logic inline...") if run_ce else print("Executing PE logic inline...")
            structure={}
            sheets[index].range("A15:O23").clear_contents()

            sheets[index].range('I12:J13').color = None
            sheets[index].range('I8:I9').color = None

            sheets[index].range('M12').color = (0, 255, 0)
            sheets[index].range('M12').value = 'Position Live...'

            if run_ce and not sell_sig:
                sheets[index].range('M3').color = (0, 255, 0)
                sheets[index].range('M2').color = (0, 255, 0)
            elif run_ce and sell_sig:
                sheets[index].range('M7').color = (0, 255, 0)
                sheets[index].range('M8').color = (0, 255, 0)
            elif run_pe and not sell_sig:
                sheets[index].range('M2').color = (0, 255, 0)
                sheets[index].range('N3').color = (0, 255, 0)
            elif run_pe and sell_sig:
                sheets[index].range('M7').color = (0, 255, 0)
                sheets[index].range('N8').color = (0, 255, 0)
            
            if m==1:
                structure['Instrument Token'] = var[f'{index}_ce'] if run_ce else var[f'{index}_pe']
                structure['Signal'] = 'BUY' if not sell_sig else 'SELL'
                structure['strike'] = var[f'strike_{index}']
                structure['type'] = 'CE' if run_ce else 'PE'
                structure['Lot Size'] = var[f'qty_{index}']
                structure['lot'] = var[f'lot_{index}']
                structure['total_qty'] = var[f'qty_{index}']*var[f'lot_{index}']
                structure['entry_ltp'] = live_data[var[f'{index}_ce']] if run_ce else live_data[var[f'{index}_pe']]
                m=2

            lowest_pts = float('inf')
            highest_pts = float('-inf')
            while True:

                with dict_lock:
                    structure['current_ltp'] = live_data[var[f'{index}_ce']] if run_ce else live_data[var[f'{index}_pe']]

                if not sell_sig:
                    points = (structure['current_ltp'] - structure['entry_ltp'])
                    profit = (structure['current_ltp'] - structure['entry_ltp'])*structure['total_qty']
                else:
                    points = (structure['entry_ltp'] - structure['current_ltp'])
                    profit = (structure['entry_ltp'] - structure['current_ltp'])*structure['total_qty']

                structure['Points'] = points
                structure['profit'] = profit

                brok_df = pd.DataFrame({'signal': [structure['Signal'], 'BUY' if structure['Signal'] == 'SELL' else 'SELL'], 'value': [structure['total_qty'] * structure['entry_ltp'], structure['total_qty'] * structure['current_ltp']]})
                brok_df = brok_df.set_index('signal')

                structure['brokerage'] = -Brokerage_cal(brok_df, structure['lot'], active_sheet.name)
                structure['net_profit'] = structure['profit'] + structure['brokerage']
                dff = pd.DataFrame(structure, index=[0])

                net_profit = dff.iloc[0]['net_profit']

                if (run_ce and not sell_sig) or (run_pe and not sell_sig):
                    margin = dff.iloc[0]['total_qty']*dff.iloc[0]['entry_ltp']
                else:
                    if run_ce and sell_sig:
                        margin = margin_sell[0]*var[f'lot_{index}']
                    elif run_pe and sell_sig:
                        margin = margin_sell[1]*var[f'lot_{index}']


                new_rows = pd.DataFrame([{'brokerage': 'Gain %', 'net_profit': f'{round(net_profit/margin*100,2)} %'},
                                         {'brokerage': 'Margin', 'net_profit': margin}], columns=dff.columns)

                dff = pd.concat([dff, new_rows], ignore_index=True)
                dff_excel = dff.copy()
                dff_excel.columns = ['Instrument Token', 'Signal', 'Strike', 'Type', 'Lot Size', 'Lot', 'Total Qty','Entry LTP', 'Current LTP', 'Points','Profit', 'Brokerage', 'Net Profit']

                points = dff.at[0, 'Points']
                if points < lowest_pts:
                    lowest_pts = points 
                if points > highest_pts:
                    highest_pts = points
                min_max  = [round(lowest_pts,2), round(highest_pts,2)]

                try:
                    sheets[index].range('A15').value = dff_excel
                except Exception as e:
                    print(f" Error writing to Excel: {e}")

                if (var[f'target_{index}'] != 0 and points >= var[f'target_{index}']) or (var[f'sl_{index}'] != 0 and points <= var[f'sl_{index}']):
                    if points >= var[f'target_{index}']:
                        sheets[index].range('I8').color = (0, 255, 0)
                    else:
                        sheets[index].range('I9').color = (255, 0, 0)
                    hello()


                if not condition :
                    exit_cond = exit_cond + 1

                iteration = 10

                if exit_cond==iteration or trade_cancel:

                    dt2 = datetime.now().replace(microsecond=0)
                    exit_time = dt2.time()

                    trade_dur = (dt2-dt1).total_seconds()
                    timing = [entry_time, exit_time, trade_dur]

                    sheets[index].range('M2:O3').color = None
                    sheets[index].range('M7:O8').color = None

                    sheets[index].range('M12').color = None
                    sheets[index].range('M12').value = 'Position Closed...' if exit_cond==iteration else 'Position Cancelled Abruptly - Not Allowed...'

                    condition = True if exit_cond == iteration else condition
                    m=1
                    if exit_cond==iteration:
                        dff = pd.DataFrame([{'Index':symbol[index], 'strike':dff.at[0, 'strike'], 'type':dff.at[0, 'type'], 'qty/lot':dff.at[0, 'Lot Size'], 'lot':dff.at[0, 'lot'], 'total_qty':dff.at[0, 'total_qty'], 'Points':dff.at[0, 'Points'], 'min_pts':min_max[0], 'max_pts':min_max[1], 'profit':dff.at[0, 'profit'], 'brokerage': dff.at[0, 'brokerage'], 'net_profit':dff.at[0, 'net_profit'], 'gain':dff.at[1, 'net_profit'], 'margin':dff.at[2,'net_profit']}])
                        tradelog(dff, timing)
                    trade_cancel = False
                    check_margin = True
                    margin_check = True
                    run_ce = False
                    run_pe = False
                    sell_sig = False

                    if reset_tar:
                        sheets[index].range('I8').value = reset_tar

                    print('Position Closed') if exit_cond==iteration else print('Position Cancelled Abruptly - Not Allowed')
                    break
##############################################################

        ce_last_ltp = None
        pe_last_ltp = None
        leg1 = True
        leg2 = True

        if run_both:

            if set_target_2:
                sheets[index].range('I8').value = 2
                var[f'target_{index}'] = 2
                set_target_2 = False

            if set_target_0:
                sheets[index].range('I8').value = 0
                var[f'target_{index}'] = 0
                set_target_0 = False

            dt1 = datetime.now().replace(microsecond=0)
            entry_time = dt1.time()
            
            print("Executing PE-CE logic inline...")
            structure={}
            sheets[index].range("A15:O23").clear_contents()

            sheets[index].range('I12:J13').color = None
            sheets[index].range('I8:I9').color = None

            sheets[index].range('M12').color = (0, 255, 0)
            sheets[index].range('M12').value = 'Position Live...'

            if run_both and not sell_sig:
                sheets[index].range('M2').color = (0, 255, 0)
                sheets[index].range('O3').color = (0, 255, 0)
            elif run_both and sell_sig:
                sheets[index].range('M7').color = (0, 255, 0)
                sheets[index].range('O8').color = (0, 255, 0)

            if m==1:
                structure['Instrument Token'] = [var[f'{index}_ce'],var[f'{index}_pe']]
                structure['signal'] = ['BUY','BUY'] if not sell_sig else ['SELL','SELL']
                structure['strike'] = [var[f'strike_{index}'], var[f'strike_{index}']]
                structure['type'] = ['CE','PE']
                structure['Lot Size'] = [var[f'qty_{index}'], var[f'qty_{index}']]
                structure['lot'] = [var[f'lot_{index}'], var[f'lot_{index}']]
                structure['total_qty'] = [var[f'qty_{index}']*var[f'lot_{index}'], var[f'qty_{index}']*var[f'lot_{index}']]
                structure['entry_ltp'] = [live_data[var[f'{index}_ce']], live_data[var[f'{index}_pe']]]
                m=2

            lowest_pts = float('inf')
            highest_pts = float('-inf')
            while True:
                if leg1:
                    ce_last_ltp = live_data[var[f'{index}_ce']]
                if leg2:
                    pe_last_ltp = live_data[var[f'{index}_pe']]

                with dict_lock:
                    structure['current_ltp'] = [ce_last_ltp, pe_last_ltp]


                if not sell_sig:
                    points = [(structure['current_ltp'][0] - structure['entry_ltp'][0]), (structure['current_ltp'][1] - structure['entry_ltp'][1])]
                    profit = [(structure['current_ltp'][0] - structure['entry_ltp'][0])*structure['total_qty'][0], (structure['current_ltp'][1] - structure['entry_ltp'][1])*structure['total_qty'][1]]
                else:
                    points = [(structure['entry_ltp'][0] - structure['current_ltp'][0]), (structure['entry_ltp'][1] - structure['current_ltp'][1])]
                    profit = [(structure['entry_ltp'][0] - structure['current_ltp'][0])*structure['total_qty'][0], (structure['entry_ltp'][1] - structure['current_ltp'][1])*structure['total_qty'][1]]

                structure['Points'] = points
                structure['profit'] = profit

                both_df = pd.DataFrame(structure)

                total_profit = both_df['profit'].sum()
                total_points = both_df['Points'].sum()

                brok_df = both_df[['signal', 'total_qty', 'entry_ltp', 'current_ltp']].copy()

                brok_df_1 = brok_df[['signal', 'total_qty', 'entry_ltp']]
                brok_df_1['value'] = brok_df_1['total_qty']*brok_df_1['entry_ltp']
                brok_df_1 = brok_df_1[['signal', 'value']]

                brok_df_2 = brok_df[['signal', 'total_qty', 'current_ltp']]
                brok_df_2['value'] = brok_df_2['total_qty'] * brok_df_2['current_ltp']
                brok_df_2['signal'] = ['SELL', 'SELL'] if not sell_sig else ['BUY', 'BUY']
                brok_df_2 = brok_df_2[['signal', 'value']]

                brok_df = pd.concat([brok_df_1, brok_df_2], ignore_index=True).set_index('signal')


                brokerage = -Brokerage_cal(brok_df, structure['lot'][0], active_sheet.name)

                net_profit = total_profit + brokerage

                if (run_both and not sell_sig):
                    margin = (both_df.iloc[0]['total_qty']*both_df.iloc[0]['entry_ltp'])+(both_df.iloc[1]['total_qty']*both_df.iloc[1]['entry_ltp'])
                else:
                    margin = margin_sell[2]*var[f'lot_{index}']

                
                gain = f'{round(net_profit/margin*100,2)} %'

                new_rows = pd.DataFrame([{'current_ltp':both_df['Points'].sum(),  'Points': 'Total Profit', 'profit': total_profit},
                                         {'Points': 'Brokerage', 'profit': brokerage},
                                         {'Points': 'Net Profit', 'profit': net_profit}, 
                                         {'entry_ltp': 'Margin', 'current_ltp':margin, 'Points': 'Gain %', 'profit': gain}], columns=both_df.columns)
                                        
                both_df = pd.concat([both_df, new_rows], ignore_index=True)
                both_df_excel = both_df.copy()
                both_df_excel.columns = ['Instrument Token', 'Signal', 'Strike', 'Type', 'Lot Size', 'Lot', 'Total Qty', 'Entry LTP', 'Current LTP', 'Points', 'Profit']

                points = both_df.loc[2, 'current_ltp']
                if points < lowest_pts:
                    lowest_pts = points
                if points > highest_pts:
                    highest_pts = points
                min_max  = [round(lowest_pts,2), round(highest_pts,2)]

                try:
                    sheets[index].range('A15').value = both_df_excel
                except Exception as e:
                    print(f" Error writing to Excel: {e}")

                if var[f'activate_{index}']:

                    if (var[f'target_{index}_leg1'] != 0 and both_df.at[0,'Points'] >= var[f'target_{index}_leg1']) or (var[f'sl_{index}_leg1'] != 0 and both_df.at[0,'Points'] <= -var[f'sl_{index}_leg1']):
                        if both_df.at[0,'Points'] >= var[f'target_{index}_leg1']:
                            sheets[index].range('I12').color = (0, 255, 0)
                        if both_df.at[0,'Points'] <= -var[f'sl_{index}_leg1']:
                            sheets[index].range('J12').color = (255, 0, 0)
                        leg1=False

                    if (var[f'target_{index}_leg2'] != 0 and both_df.at[1,'Points'] >= var[f'target_{index}_leg2']) or (var[f'sl_{index}_leg2'] != 0 and both_df.at[1,'Points'] <= -var[f'sl_{index}_leg2']):
                        if both_df.at[1,'Points'] >= var[f'target_{index}_leg2']:
                            sheets[index].range('I13').color = (0, 255, 0)
                        if both_df.at[1,'Points'] <= -var[f'sl_{index}_leg2']:
                            sheets[index].range('J13').color = (255, 0, 0)
                        leg2=False

                    if not leg1 and not leg2:
                        hello()
                
                if condition:
                    
                    if (var[f'target_{index}'] != 0 and points >= var[f'target_{index}']) or (var[f'sl_{index}'] != 0 and points <= -var[f'sl_{index}']):
                        if points >= var[f'target_{index}']:
                            sheets[index].range('I8').color = (0, 255, 0)
                        else:
                            sheets[index].range('I9').color = (255, 0, 0)
                        hello()

                if not condition :
                    exit_cond = exit_cond + 1

                iteration = 10

                if exit_cond==iteration or trade_cancel:

                    dt2 = datetime.now().replace(microsecond=0)
                    exit_time = dt2.time()

                    trade_dur = (dt2-dt1).total_seconds()
                    timing = [entry_time, exit_time, trade_dur]

                    sheets[index].range('M2:O3').color = None
                    sheets[index].range('M7:O8').color = None

                    sheets[index].range('M12').color = None
                    sheets[index].range('M12').value = 'Position Closed' if exit_cond==iteration else 'Position Cancelled Abruptly - Not Allowed...'

                    condition = True if exit_cond==iteration else condition
                    m=1
                    if exit_cond==iteration:
                        both_df = pd.DataFrame([{'Index':symbol[index], 'strike':both_df.at[0, 'strike'], 'type':'CE-PE', 'qty/lot':both_df.at[0, 'Lot Size'], 'lot':both_df.at[0, 'lot'], 'total_qty':both_df.at[2, 'total_qty'], 'Points':both_df.at[2, 'current_ltp'], 'min_pts':min_max[0], 'max_pts':min_max[1],  'profit':both_df.at[2, 'profit'], 'brokerage': both_df.at[3, 'profit'], 'net_profit':both_df.at[4, 'profit'], 'gain':both_df.at[5, 'profit'], 'margin':both_df.at[5,'current_ltp']}])
                        tradelog(both_df, timing)
                    trade_cancel = False
                    check_margin = True
                    margin_check = True
                    run_both = False
                    sell_sig = False

                    if reset_tar:
                        sheets[index].range('I8').value = reset_tar
                    print('Position Closed') if exit_cond==iteration else print('Position Cancelled Abruptly - Not Allowed')
                    break