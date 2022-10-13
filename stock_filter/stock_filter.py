import pip._vendor.requests
from bs4 import BeautifulSoup
import pandas as pd
import time

##修改threshold直接改占比資料
threshold=5
##修改delay_time會是更動從新像網頁送出請求的時間，目前是0.2秒
delay_time=0.2

def process_1(stock_number):
    target_url="https://fubon-ebrokerdj.fbs.com.tw/z/zc/zco/zco_"+stock_number+"_1.djhtm"
    r = pip._vendor.requests.get(target_url) 
    soup = BeautifulSoup(r.text,"html.parser")
    target_broker = soup.select("TD.t4t1 a")[0].get_text()
    roll_broker_data = soup.select("TD.t3n1")[:4]
    target_broker_data={"amount":int(roll_broker_data[2].get_text().replace(',', '')),
                        "percentage":float(roll_broker_data[3].get_text().strip('%')),}
    if target_broker_data['percentage']>threshold and (target_broker_data['amount']*100/target_broker_data['percentage'])>1000:
        return [True,target_broker,target_broker_data]
    else:
        return [False]

def process_2(data_1):
    target_url="https://fubon-ebrokerdj.fbs.com.tw/z/zc/zco/zco_"+stock_number+"_2.djhtm"
    r = pip._vendor.requests.get(target_url) 
    soup = BeautifulSoup(r.text,"html.parser") 
    target_broker_list = [stock_tag.get_text() for stock_tag in soup.select("TD.t4t1 a")[::2]]
    target_broker_data = [int(amount.get_text().replace(',', '')) for amount in soup.select("TD.t3n1")[2::8][:15]]
    flag= False
    day5_trade_amount=0
    for idx in range(len(target_broker_list)):
        if target_broker_list[idx]==data_1[1] and target_broker_data[idx]/data_1[2]['amount']>1.1:
            flag =True
            day5_trade_amount=target_broker_data[idx]
    print(flag,data_1)
    return [flag,day5_trade_amount]
    
df=pd.read_csv('stock_list_test.csv',header=None)
stock_list=[stock.replace(u'\xa0', u'') for stock in df[0].tolist()]
data_dic=[]
error_li=[]

for stock_number in stock_list:
    #stock_number="2453"
    try:     
        data=process_1(stock_number)
        print(stock_number,data[0])
        if data[0]:
            data_handle=process_2(data)
            if data_handle[0]: 
                data_dic.append({"name":stock_number,"broker":data[1],"amount":data[2]['amount'],"percentage":data[2]['percentage'],"5_day_amount":data_handle[1]})
    except:
        print("Error:",stock_number)
        error_li.append(stock_number)
    time.sleep(delay_time)
pd.DataFrame(data_dic).to_csv("result.csv",encoding='utf_8_sig')