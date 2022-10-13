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
    
##target_broker取第三名的券商    
    roll_broker_data = soup.select("TD.t3n1")[:20]
    target_broker_data={"amount":int(roll_broker_data[18].get_text().replace(',', '')),
                        "percentage":float(roll_broker_data[19].get_text().strip('%')),}
    print(target_broker_data)
    today_broker_list = [today_broker.get_text() for today_broker in soup.select("TD.t4t1 a")[::2][:3]]
    today_amount_list = [int(today_amount.get_text().replace(',', '')) for today_amount in soup.select("TD.t3n1")[2::8][:3]]
    today_broker_dataframe=pd.DataFrame({'Broker':today_broker_list, 'Amount':today_amount_list}, columns=['Broker','Amount'])                  
    if target_broker_data['percentage']>threshold and (target_broker_data['amount']*100/target_broker_data['percentage'])>1000:
        ##取前3名來比較5天的量
        return [True,today_broker_dataframe.values.tolist()]
    else:
        return [False]

def process_2(data_1):
    target_url="https://fubon-ebrokerdj.fbs.com.tw/z/zc/zco/zco_"+stock_number+"_2.djhtm"
    r = pip._vendor.requests.get(target_url) 
    soup = BeautifulSoup(r.text,"html.parser")
    target_broker_list = [stock_tag.get_text() for stock_tag in soup.select("TD.t4t1 a")[::2]]
    target_broker_data = [int(amount.get_text().replace(',', '')) for amount in soup.select("TD.t3n1")[2::8][:15]]
    day5_trade_amount_list=[]
    day5_add_percentage_list=[]
    day5_trade_flag_list=[]
    broker_list=[]
    for idx in range(len(target_broker_list)):
        for sdx in range(len(data_1)):
            if target_broker_list[idx]==data_1[sdx][0] and target_broker_data[idx]/data_1[sdx][1]>1.1:
                day5_trade_flag_list.append(True)
                broker_list.append(target_broker_list[idx])
                day5_trade_amount_list.append(target_broker_data[idx])
                day5_add_percentage_list.append(round(target_broker_data[idx]/data_1[sdx][1],2))            
    if len(day5_trade_flag_list)>0:
        print("冠強來了!")
        day5_broker_dataframe=pd.DataFrame({'broker':broker_list ,'Amount':day5_trade_amount_list,'Percentage':day5_add_percentage_list})
        return [True,day5_broker_dataframe.values.tolist()]
    else:
        print("幹!沒有")
        return [False]    

df=pd.read_csv('stock_list_test.csv',header=None)
stock_list=[stock.replace(u'\xa0', u'') for stock in df[0].tolist()]
data_dic=[]
error_li=[]

for stock_number in stock_list:
    #stock_number="2453"
    try:   
        data =process_1(stock_number)
        print(stock_number,data[0])
        if data[0]:
                data_handle=process_2(data[1])
                if data_handle[0]:
                    print(data_handle[1])
                    for dx in range(len(data_handle[1])):
                        data_dic.append({"name":stock_number,"Broker":data_handle[1][dx][0],"5day_Amount":data_handle[1][dx][1],"5day_Add_Percentage":data_handle[1][dx][2]})         
    except:
        print("Error:",stock_number)
        error_li.append(stock_number)
    time.sleep(delay_time)
pd.DataFrame(data_dic).to_csv("result.csv",encoding='utf_8_sig')
print("爽了")