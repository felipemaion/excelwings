#!/usr/bin/python
# -*- coding: utf-8 -*-

import requests
from bs4 import BeautifulSoup
import pandas as pd


def get_fiis():
    url = 'https://www.fundsexplorer.com.br/ranking'
    response = requests.get(url)
    soup = BeautifulSoup(response.text, "html.parser")

    data = []
    table = soup.find(id="table-ranking")
    table_head = table.find('thead')

    rows = table_head.find_all('tr')
    for row in rows:
        cols = row.find_all('th')
        # Get the Header
        header = [element.get_text(separator=" ").strip() for element in cols]
        data.append([element for element in header])


    table_body = table.find('tbody')

    rows = table_body.find_all('tr')
    for row in rows:
        cols = row.find_all('td')
        # Clean data in collumns:
        clean_col = [element.text.replace('R$','').replace('%','').replace('.0','').replace('.','').replace('N/A','').strip() for element in cols]
        data.append([element for element in clean_col])

    df = pd.DataFrame(data=data[1:],columns=data[0])
    # df.to_excel("fii.xlsx")
    
    return df