from numpy.core.fromnumeric import transpose
import yfinance as yf
import xlwings as xw

from datetime import datetime 
from IPython.display import clear_output
import pandas as pd
from pandas import json_normalize  
import functools
import threading
import time
import schedule
import requests
import json
from fii import get_fiis
from company import Company
import fundamentus

WORKBOOK = "Ações_Escolha-Online.xlsx"
criptos = ['PAXG-USD','PAXG-BRL','ETH-USD','ETH-BRL','BRLUSD=X','BTC-USD','BTC-BRL', 'DOGE-USD']
def job():
    return f"{threading.current_thread()}"


def run_threaded(job_func):
    job_thread = threading.Thread(target=job_func)
    job_thread.start()

def catch_exceptions(cancel_on_failure=False):
    def catch_exceptions_decorator(job_func):
        @functools.wraps(job_func)
        def wrapper(*args, **kwargs):
            try:
                return job_func(*args, **kwargs)
            except:
                import traceback
                print(traceback.format_exc())
                if cancel_on_failure:
                    return schedule.CancelJob
        return wrapper
    return catch_exceptions_decorator

assets = []

@catch_exceptions(cancel_on_failure=False)
def connect_workbook(workbook=WORKBOOK):
    current = xw.apps.keys()
    wb = xw.apps[current[0]].books[workbook] 
    print(f"pid:{current} - workbook:{wb}")
    return wb

@catch_exceptions(cancel_on_failure=True)
def refresh_wb_and_save(workbook=WORKBOOK):
    wb = connect_workbook(workbook)
    wb.api.refresh_all(wb.api)
    wb.api.save(timeout=3000)
    set_time()

@catch_exceptions(cancel_on_failure=True)
def get_sheet(sheet, at=WORKBOOK ):
    try:
        wb = connect_workbook(at)
        if wb.sheets[sheet].api.show_all_data:
            # wb.sheets[sheet].api.autofilter.ShowAllData()
            return wb.sheets[sheet]
    except Exception as e:
        print(f"Ooops Crash! {e}" )
        return []


@catch_exceptions(cancel_on_failure=True)
def get_assets(sheet='ATIVOS'):
    print(f"Getting Assets - {job()}")
    
    assets = []
    sht_assets = get_sheet(sheet)
    active = True
    i = 2
    print(f"Starts assets:{datetime.now()}- {job()}")
    # Enquanto houver assets na coluna A
    while active:
        asset = sht_assets.range(f"A{i}").value
        assets.append(asset)
        if asset is None: 
            active = False
            break
        i = i + 1
    print(f"Ends assets:{datetime.now()}- {job()}")  
    return assets



@catch_exceptions(cancel_on_failure=True)
def set_time():
    global failured_before
    try:
        sht_assets = get_sheet('INDICADOR')
        sht_assets.range(f"D3").value = datetime.now()
        if failured_before == True: 
            failured_before = False
            print("\nClient is back!!", end="\n")
            stop_server_time()
            schedule_server_time(sec=1)
            schedule_at_bovespa_time()
    except Exception as e:
        # print("Problema ao se conectar com o cliente Excel")
        # print(f"{e} ")
        # print("Cancelando Time Server...")
        print(".", end="", flush=True)
        if not failured_before:
            print("Excel Client is Down!!")
            stop_server_time()
            # print("Tentando novamente em 15 seg")
            schedule_server_time(sec=15)

            stops_bovespas_jobs()
        
            failured_before = True
            print("Waiting Excel Client", end="", flush=True)
        

@catch_exceptions(cancel_on_failure=False)    
def populate_prices(assets_to_update = None):
    # 1 - Get prices for the assets_to_update, store in a dict.
    # 2 - Read the spreadsheet and insert the price in the right spot.
    sht_assets = get_sheet('ATIVOS')
    all_assets = get_assets()
    all_prices = {}
    assets_to_update = assets_to_update or all_assets
    for asset in assets_to_update:
        if asset == None: continue
        try: # Local Market
            all_prices[asset] = yf.Ticker(f"{asset}.SA").history().tail(1)['Close'].iloc[0]
            if asset == "SELIC": all_prices[asset] = 1
            print(f"Price ({all_prices[asset]}) for { asset } loaded.")
        except Exception as e:
            try: 
                if asset!='PAXG-USD':
                    all_prices[asset] = yf.Ticker(f"{asset}").history().tail(1)['Close'].iloc[0]
                    if asset == "SELIC": all_prices[asset] = 1
                    # print(f"Price ({all_prices[asset]}) for { asset } loaded.")
                else:
                    r = requests.get("https://pyinvesting.com/invest/instrument/data/paxg-usdcc")
                    all_prices[asset] = json.loads(r.content)['price_latest']
                print(f"Price ({all_prices[asset]}) for { asset } loaded.")
            except:
                print(f"Error {e} with asset {asset}")
                pass

    print("Trying to insert price into table:")
    for i, asset in enumerate(all_assets, start = 2):
        if sht_assets.range(f"A{i}").value in assets_to_update:
            if asset in all_prices.keys():
                sht_assets.range(f"D{i}").value = all_prices[asset]
                print(f"\tAtivo: {asset} -> R$ {all_prices[asset]}")
    
    sht_assets = get_sheet('INDICADOR')
    sht_assets.range(f"F3").value = datetime.now()
    # wb.screen_updating = True
    print(f"Ends Prices:{datetime.now()}- {job()}")
    print("Updating Pivots Tables and External Sources...")
    refresh_wb_and_save(WORKBOOK)
    print(f"Finished...\n\tTotal PL: R$ {get_sheet('MOBILE')['PL_TOTAL'].value: ,.2f}")

@catch_exceptions(cancel_on_failure=False)    
def populate_cripto():
    populate_prices(criptos)

@catch_exceptions(cancel_on_failure=False)
def get_selic(amount, date):
        str_data = f"{date.year}-{date.month}-{date.day}"
        r = requests.get(f"https://www.felipemaion.com.br/selic?amount={amount}&date={str_data}")
        # print(r.text)
        return f"{r.text}"

@catch_exceptions(cancel_on_failure=False)    
def populate_selic():
    now = datetime.now()
    print(f"Starts SELIC:{now}- {job()}")
    try:
        sht = get_sheet('APORTES')
        last_update = sht.range('K1').value
        index_update = float(get_selic('1000', last_update)) / 1000.

        aportes = sht.range('A1').options(pd.DataFrame,
                            index=True, expand='table').value
        print(aportes.head(10))
        
        for i, aporte in aportes.iterrows():
            data, valor, old_selic, this_last_update = aporte['DATA'], aporte['VALOR'], aporte['SELIC'], aporte[last_update]

            if data and valor:
                if this_last_update == last_update:
                    selic_value = old_selic * index_update
                else:    
                    selic_value = get_selic(valor, data)#.replace(".",",")
                aportes.at[i,'SELIC'] = selic_value
                aportes.at[i,last_update] = now
                # print(selic_value, aportes.at[i,'SELIC'])
                print(f"\t #{ int(i) }: R$ {aporte['VALOR']}: {aporte['DATA']}= R$ {selic_value}")
        # print(aportes['SELIC']) # this should be a test? I don't know how to make.
        sht = get_sheet('APORTES')
        print("Setting Values into Excel...")
        sht.range('A1').value = aportes
        sht.range('K1').value = now
        print("Updating Pivots Tables and External Sources...")

        refresh_wb_and_save(WORKBOOK)
        print("ENDING SELIC.")
        return aportes
    except Exception as e:
        print(f"SELIC: erro... {e}")



@catch_exceptions(cancel_on_failure=True)
def populate_fiis():
    df_fii = get_fiis()
    sht_fundamentos = get_sheet('FIIs')
    sht_fundamentos.range("A1").value = df_fii
    print(df_fii.head())
    return df_fii

def populate_radar():
    data = fundamentus.get_dataframe()
    sht_fundamentos = get_sheet('RADAR')
    # print(data.head())
    sht_fundamentos.range('A1').options(index=False).value = data
    # print(sht_fundamentos.range('A1').options(expand='table').value)
    return data
    


@catch_exceptions(cancel_on_failure=True)
def populate_fundamentos(assets=None):
    if assets==None: assets = get_assets('RADAR')
    if type(assets) == type('string'):
        assets = [assets]
    
    print(f"Starts Fundamentos:{datetime.now()}")
    load_Company = Company(assets[0])
    total_DataFrame = json_normalize(json.loads(load_Company.json_data)['data'])
    if len(assets) > 1:
        for asset in assets[1:]:
            load_Company = Company(asset)
            if load_Company.data:
                load_DataFrame = json_normalize(json.loads(load_Company.json_data)['data'])
                total_DataFrame = pd.concat([total_DataFrame,load_DataFrame],sort=True, ignore_index=True).dropna(thresh=5)
    print(f"Ends Fundamentos:{datetime.now()}")
    sht_fundamentos = get_sheet('FUNDAMENTOS')
    sht_fundamentos.range("A1").value = total_DataFrame
    refresh_wb_and_save()
    return total_DataFrame
    
def starts_bovespas_jobs():
    print(f"Starting Bovespas's Jobs")
    schedule.every(15).minutes.do(run_threaded, populate_prices).tag("bovespa")
    populate_selic() #When service starts
#     schedule.every(1).week.do(run_threaded, populate_fundamentos).tag("fundamentos")
def stops_bovespas_jobs():
    populate_prices()
    print(f"ENDING BOVESPA'S JOBS.")
    schedule.clear("bovespa")
    
def schedule_at_bovespa_time():
    print(f"SCHEDULLING BOVESPA'S JOBS: 9:45 - 18:30, every 15 min.")
    now = datetime.now()
    today9am = now.replace(hour=9, minute=45, second=0, microsecond=0)
    today19pm = now.replace(hour=18, minute=30, second=0, microsecond=0)
    if datetime.now() > today9am and datetime.now() < today19pm:
        print("It is after 9:45, starting...\n Let me catch up...")
        populate_prices()
        starts_bovespas_jobs()
    schedule.every(1).day.at("09:45").do(starts_bovespas_jobs)
    schedule.every(1).day.at("18:30").do(stops_bovespas_jobs)

def schedule_fundamentos(days=15):
    print(f"Schedulling Fundamentos to every {days} days")
    schedule.every(days).days.do(run_threaded, populate_fundamentos)

def schedule_SELIC(minute=1):
    print(f"Schedulling SELIC to every {minute} day")
    schedule.every(minute).days.do(run_threaded, populate_selic)

def schedule_server_time(sec=1):
    print(f"Schedulling Server Time to every {sec} seconds")
    schedule.every(sec).seconds.do(run_threaded, set_time).tag("horario")

def stop_server_time():
    print(f"Stopping Server Time.")
    schedule.clear("horario")

def run_all():
    populate_selic();
    populate_radar();
    populate_fundamentos();
    populate_fiis();
    populate_prices(); 
    set_time();
# schedule_SELIC()
failured_before = False
def start():
    # populate_prices()    
    # total_DataFrame = populate_fundamentos()
    # total_DataFrame.to_excel("output20201019.xlsx")
    print(f"Starts schedulling:{datetime.now()}")

    schedule_server_time()
    schedule_at_bovespa_time()
    schedule_fundamentos()
    populate_prices()
    while 1:
    #     clear_output()
        schedule.run_pending()
        time.sleep(1)
print("Run 'start()' to run the schedulled jobs")
print("Run 'run_all()' to populate: selic, radar, fundamentos, fiis, prices, and set_time in Excel")
# start()
