import yfinance as yf
import xlwings as xw
import schedule
import time
from datetime import datetime 
from IPython.display import clear_output
import pandas as pd
from pandas import json_normalize  
import functools
import threading
import time
import schedule


def job():
    return f"{threading.current_thread()}\n"


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

@catch_exceptions(cancel_on_failure=True)
def get_assets():
    print("Getting Assets")
    job()
    assets = []
    current = xw.apps.keys()
    print(current)
    wb = xw.apps[current[0]].books[0] 
    sht_assets = wb.sheets['ATIVOS']
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
    
@catch_exceptions(cancel_on_failure=False)    
def populate_prices():
    assets = get_assets()
#     print(assets)
    print(f"Starts Prices:{datetime.now()}- {job()}")
    job()
    current = xw.apps.keys()
    wb = xw.apps[current[0]].books[0] 
    sht_assets = wb.sheets['ATIVOS']
        # Pega último preço do Ativo
    for i, asset in enumerate(assets, start = 2):
        try:
            yf_asset = yf.Ticker(f"{asset}.SA").history().tail(1)['Close'].iloc[0]
            if asset == "SELIC": yf_asset = 1
            sht_assets.range(f"D{i}").value = yf_asset
#             print(f"\tAtivo: {asset} -> R$ {yf_asset}")
            
        except Exception as e:
            print(f"Erro: { asset } - {e}")
            pass
    sht_assets = wb.sheets['INDICADOR']   
    sht_assets.range(f"F3").value = datetime.now()
    print(f"Ends Prices:{datetime.now()}- {job()}")    
#     return schedule.CancelJob

@catch_exceptions(cancel_on_failure=True)
def set_time():
    job()
    current = xw.apps.keys()
    wb = xw.apps[current[0]].books[0] 
    sht_assets = wb.sheets['INDICADOR']
    sht_assets.range(f"D3").value = datetime.now()

@catch_exceptions(cancel_on_failure=True)
def populate_fundamentos():
    assets = get_assets()
    print(f"Starts Fundamentos:{datetime.now()}")
    load_Company = Company(assets[0])
    total_DataFrame = json_normalize(json.loads(load_Company.json_data)['data'])

    for asset in assets[1:]:
        load_Company = Company(asset)
        if load_Company.data:
            load_DataFrame = json_normalize(json.loads(load_Company.json_data)['data'])
            total_DataFrame = pd.concat([total_DataFrame,load_DataFrame],sort=True, ignore_index=True).dropna(thresh=2)
    print(f"Ends Fundamentos:{datetime.now()}")
    current = xw.apps.keys()
    wb = xw.apps[current[0]].books[0] 
    sht_fundamentos = wb.sheets['FUNDAMENTOS']
    sht_fundamentos.range(f"A1").value = total_DataFrame
    
    return total_DataFrame
    
def starts_bovespas_jobs():
    print(f"Starting Jobs")
    schedule.every(15).minutes.do(run_threaded, populate_prices).tag("bovespa")
    
#     schedule.every(1).week.do(run_threaded, populate_fundamentos).tag("fundamentos")
def stops_bovespas_jobs():
    print(f"ENDING BOVESPA'S JOBS.")
    schedule.clear("bovespa")
    
def schedule_at_bovespa_time():
    print(f"SCHEDULLING BOVESPA'S JOBS: 9:45 - 18:30")
    schedule.every(1).day.at("09:45").do(starts_bovespas_jobs)
    schedule.every(1).day.at("18:30").do(stops_bovespas_jobs)

def schedule_fundamentos(minute=1):
    print(f"Schedulling Fundamentos to every {minute} minutes")
    schedule.every(minute).to(5).minutes.do(run_threaded, populate_fundamentos)

def schedule_server_time(sec=1):
    print(f"Schedulling Server Time to every {sec} seconds")
    schedule.every(sec).seconds.do(run_threaded, set_time).tag("horario")
populate_prices()    
# total_DataFrame = populate_fundamentos()
# total_DataFrame.to_excel("output20201019.xlsx")
print(f"Starts schedulling:{datetime.now()}")

schedule_server_time()
schedule_at_bovespa_time()
schedule_fundamentos()



while 1:
#     clear_output()
    schedule.run_pending()
    time.sleep(1)