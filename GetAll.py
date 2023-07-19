import requests
import json
import GetToken
import datetime
from datetime import timedelta, datetime
with open("datetime.txt", "r") as file:
    saved_datetime = file.read()
saved_datetime = datetime.strptime(saved_datetime, '%d.%m.%Y')
new_format = "%Y-%m-%dT%H:%M:%SZ"
time1 = saved_datetime.strftime(new_format)
day = timedelta(days= 1)
week = timedelta(weeks= 1)
month = timedelta(weeks= 5)
date_day = saved_datetime - day
time2 = date_day.strftime(new_format)
date_week = saved_datetime - week
time3 = date_week.strftime(new_format)
date_month = saved_datetime - month
time4 = date_month.strftime(new_format)
print(time1)
for a in range (1, 38):
    if a == 1:
        i = 1
    elif a == 2:
        i = 6
    elif a == 3:
        i = 8
    elif a == 4:
        i = 4
    elif a == 5:
        i = 3
    elif a == 6:
        i = 2 #ЗАМЕНИТЬ!!! ГУБЧАТОЕ ЖЕЛЕЗО ОТСУТСТВУЕТ!!!
    elif a == 7:
        i = 5
    elif a == 8:
        i = 67
    if a == 9:
        i = 9
    elif a == 10:
        i = 45
    elif a == 11:
        i = 12
    elif a == 12:
        i = 61 #ЗАМЕНИТЬ РУЛОК Г/К ИНДИЯ FOB ОТСУТСТВУЕТ
    elif a == 13:
        i = 15
    elif a == 14:
        i = 13 
    elif a == 15:
        i = 65 #ЛИБО 66
    elif a == 16:
        i = 16 #ИЛИ НЕТ
    elif a == 17:
        i = 10 #ИЛИ НЕТ
    elif a == 18:
        i = 14
    if a == 19:
        i = 1 #ЗАМЕНИТЬ FeSi 75 Монголия ОТСУТСТВУЕТ
    elif a == 20:
        i = 1 #ЗАМЕНИТЬ FeSi 75 Китай ОТСУТСТВУЕТ
    elif a == 21:
        i = 18
    elif a == 22:
        i = 1 #ЗАМЕНИТЬ FeSi 75 США ОТСУТСТВУЕТ
    elif a == 23:
        i = 1 #ЗАМЕНИТЬ SiMn 65/17 Гуанси ОТСУТСТВУЕТ
    elif a == 24:
        i = 1 #ЗАМЕНИТЬ SiMn 65/16 Индия ОТСУТСТВУЕТ
    elif a == 25:
        i = 1 #ЗАМЕНИТЬ SiMn 65/17 США ОТСУТСТВУЕТ
    elif a == 26:
        i = 19 
    elif a == 27:
        i = 17
    elif a == 28:
        i = 22
    if a == 29:
        i = 1 #ЗАМЕНИТЬ 55%Cr 10%C EXW Монголия ОТСУТСТВУЕТ
    elif a == 30:
        i = 1 #ЗАМЕНИТЬ 70%Cr 0.5%Si Казахстан CIF ОТСУТСТВУЕТ
    elif a == 31:
        i = 1 #ЗАМЕНИТЬ 60%Cr 4%Si Китай CIF ОТСУТСТВУЕТ
    elif a == 32:
        i = 1 #ЗАМЕНИТЬ 48-50%Cr 5%Si Китай CIF ОТСУТСТВУЕТ
    elif a == 33:
        i = 20 
    elif a == 34:
        i = 1 #ЗАМЕНИТЬ 52-60%Cr 0.25%C EXW Монголия ОТСУТСТВУЕТ
    elif a == 35:
        i = 21 
    elif a == 36:
        i = 1 #ЗАМЕНИТЬ 0.1%C EXW США ОТСУТСТВУЕТ
    elif a == 37:
        i = 23 
    payload0 = {"material_source_id": i, "property_id": 1, "start": time1, "finish": time1}
    payload1 = {"material_source_id": i, "property_id": 1, "start": time4, "finish": time1}
    payload2 = {"material_source_id": i, "property_id": 1, "start": time2, "finish": time2}
    payload3 = {"material_source_id": i, "property_id": 1, "start": time3, "finish": time3}
    payload4 = {"material_source_id": i, "property_id": 1, "start": time4, "finish": time4}
    url = 'http://base.metallplace.ru:8080/getValueForPeriod'
    r1 = requests.post(url = url, headers = GetToken.headers, data = json.dumps(payload1))
    data1 = r1.json()
    pretty1 = json.dumps(data1, sort_keys=False, indent=4, ensure_ascii= False, separators=(',', ': '))
    price_feed = data1.get("price_feed", [])
    r2 = requests.post(url = url, headers = GetToken.headers, data = json.dumps(payload2))
    data2 = r2.json()
    prevD = data2.get("prev_price", [])
    print("day")
    print(prevD)
    r3 = requests.post(url = url, headers = GetToken.headers, data = json.dumps(payload3))
    data3 = r3.json()
    prevW = data3.get("prev_price", [])
    print("week")
    print(prevW)
    r4 = requests.post(url = url, headers = GetToken.headers, data = json.dumps(payload4))
    data4 = r4.json()
    prevM = data4.get("prev_price", [])
    r0 = requests.post(url = url, headers = GetToken.headers, data = json.dumps(payload0))
    data0 = r0.json()
    prev0 = data0.get("prev_price", [])
    print("month")
    print(prevM)
    for price in price_feed:
        if price.get("date", "") == time2:
            price_value = price.get("value", None)
            if price_value is not None:
                PR2 = price_value
                print(PR2)
                print("день назад")
                break
        else:
            PR2 = prevD
    for price in price_feed:
        if price.get("date", "") == time3:
            price_value = price.get("value", None)
            if price_value is not None:
                PR3 = price_value
                print(PR3)
                print("неделю назад")
                break
        else:
            PR3 = prevW
    for price in price_feed:
        if price.get("date", "") == time4:
            price_value = price.get("value", None)
            if price_value is not None:
                PR4 = price_value
                print(PR4)
                print("месяц назад")
                break
        else:
            PR4 = prevM
    for price in price_feed:
        if price.get("date", "") == time1:
            price_value = price.get("value", None)
            if price_value is not None:
                PR1 = price_value
                print(PR1)
                print("ща")
                break
        else:
            PR1 = prev0
    if PR1 != 0:
        globals()['PRICE%s' % a] = round(PR1, 2)
    else:
        globals()['PRICE%s' % a] = 0
    if PR1 and PR2 != 0:
        globals()['CH_DAY%s' % a] = round(PR1-PR2, 2)
    else:
        globals()['CH_DAY%s' % a] = 0
    if PR1 and PR2 != 0:
        globals()['CH_DAY_PR%s' % a] = round((((PR1/PR2)*100)-100), 1)
    else:
        globals()['CH_DAY_PR%s' % a] = 0
    if PR1 and PR3 != 0:
        globals()['CH_W%s' % a] = round(PR1 - PR3, 2)
    else:
        globals()['CH_W%s' % a] = 0
    if PR1 and PR3 != 0:
        globals()['CH_W_PR%s' % a] = round((((PR1/PR3)*100)-100), 1)
    else:
        globals()['CH_W_PR%s' % a] = 0
    if PR1 and PR4 != 0:
        globals()['CH_M%s' % a] = round(PR1 - PR4, 2)
    else:
        globals()['CH_M%s' % a] = 0
    if PR1 and PR4 != 0:
        globals()['CH_M_PR%s' % a] = round((((PR1/PR4)*100)-100), 1)
    else:
        globals()['CH_M_PR%s' % a] = 0
    print(globals()['PRICE%s' % a], globals()['CH_DAY%s' % a], globals()['CH_DAY_PR%s' % a], globals()['CH_W%s' % a], globals()['CH_W_PR%s' % a], globals()['CH_M%s' % a], globals()['CH_M_PR%s' % a])
print("Good!")
