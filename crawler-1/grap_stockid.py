import requests
import pandas as pd


res = requests.get("https://isin.twse.com.tw/isin/C_public.jsp?strMode=4")
df = pd.read_html(res.text)[0]
df.columns = df.iloc[0]
df = df.iloc[1:]
df = df.dropna(thresh=3, axis=0).dropna(thresh=3, axis=1)
df = df.set_index('有價證券代號及名稱')
df.to_excel('data/taiwan_stock_id_list.xlsx', sheet_name='stock_id')