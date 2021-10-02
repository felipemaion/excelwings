#!/usr/bin/env python3

import urllib.request, json

def cotacaoBRL():
    """
    Retorna a última cotação do Bitcoin em BRL - Mercado Bitcoin via API BitValor
    """

    with urllib.request.urlopen("https://api.bitvalor.com/v1/ticker.json") as url:

        data = json.loads(url.read().decode())
        last = data['ticker_24h']['exchanges']['MBT']['last']

        return last

def cotacaoUSD():
    """
    Retorna a última cotação do Bitcoin em USD - API Bitfinex
    """

    with urllib.request.urlopen("https://api.bitfinex.com/v1/pubticker/btcusd") as url:

        data = json.loads(url.read().decode())
        last = data['last_price']

        return last


def btc2brl(btc):
    """
    Recebe uma quantidade de BTC como parâmetro e retorna seu valor aproximado em Real
    """

    price = float(cotacaoBRL())
    brl = price * btc
    brl = "{0:.2f}".format(brl)

    return(brl)

def btc2usd(btc):
    """
    Recebe uma quantidade de BTC como parâmetro e retorna seu valor aproximado em USD
    """

    price = float(cotacaoUSD())
    usd = price * btc
    usd = "{0:.2f}".format(usd)

    return(usd)


#leitura da localidade (parâmetro via linha de comando)
import sys

btc=''
if len(sys.argv)>0:
    sys.argv.pop(0)
    btc = ' '.join(sys.argv).replace(',','.')

if not btc:
    print("\nModo de uso:\n\n./bitcoin [unidade]\n\nExemplos:\n\n./bitcoin 0.35 (para Bitcoin)\n./bitcoin R100 (para Real)\n./bitcoin U100 (para Dólar)\n./bitcoin C (para consultar a cotação atual)\n\n")

elif (btc[0].upper()=='U'):
    usd = float(btc[1:])
    value = usd/float(cotacaoUSD())
    print(value)

elif (btc[0].upper()=='R'):
    brl = float(btc[1:])
    value = brl/float(cotacaoBRL())
    print(value)

elif (btc[0].upper()=='C'):
    print('USD: %s | BRL: %s' %(cotacaoUSD(), cotacaoBRL()))

elif float(btc):
    btc = float(btc)
    print("R$\t", btc2brl(btc)) 
    print("USD\t", btc2usd(btc)) 