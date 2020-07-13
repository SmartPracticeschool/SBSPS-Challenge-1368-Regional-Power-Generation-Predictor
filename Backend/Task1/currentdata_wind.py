import requests
import xlsxwriter
import datetime
import schedule
import time
from pprint import pprint

workbk_out = xlsxwriter.Workbook("Wind.xlsx")
sheet_out = workbk_out.add_worksheet()
sheet_out.write("A1", "Time")
sheet_out.write("B1", "Wind Speed")
sheet_out.write("C1", "Power In Watt")
city = input('City Name :')
n1 = 2
n2 = 2
n3 = 2
now = datetime.datetime.now()
a = now.strftime("%H:%M:%S")
api_address='http://api.openweathermap.org/data/2.5/weather?appid=1222ec2c19edb278b4e39377e4138b42&q='
url = api_address + city
json_data = requests.get(url).json()
wind = json_data['wind']['speed']
print('Wind Speed: {}'.format(wind))
print("Current Time is:", a )
sheet_out.write("B2",wind )
sheet_out.write("A2", a )
sheet_out.write("C2",(0.5*1.23*2826*wind*wind*wind/1000))
#X = input("press 1 to continue")
while 1:
    time.sleep(900)
    n1 += 1
    n2 += 1
    n3 += 1
    now = datetime.datetime.now()
    a = now.strftime("%H:%M:%S")
    api_address='http://api.openweathermap.org/data/2.5/weather?appid=1222ec2c19edb278b4e39377e4138b42&q='
    url = api_address + city
    json_data = requests.get(url).json()
    wind = json_data['wind']['speed']
    print('Wind Speed: {}'.format(wind))
    print("Current Time is:", a )
    sheet_out.write("B"+str(n1),wind )
    sheet_out.write("A"+str(n2), a )
    sheet_out.write("C"+str(n3),(0.5*1.23*2826*wind*wind*wind/1000))
    time.sleep(900)
    n1 += 1
    n2 += 1
    n3 += 1
    now = datetime.datetime.now()
    a = now.strftime("%H:%M:%S")
    api_address='http://api.openweathermap.org/data/2.5/weather?appid=1222ec2c19edb278b4e39377e4138b42&q='
    url = api_address + city
    json_data = requests.get(url).json()
    wind = json_data['wind']['speed']
    print('Wind Speed: {}'.format(wind))
    print("Current Time is:", a )
    sheet_out.write("B"+str(n1),wind )
    sheet_out.write("A"+str(n2), a )
    sheet_out.write("C"+str(n3),(0.5*1.23*2826*wind*wind*wind/1000))
    G = int(input("press 2 to exit"))
    if G != 2:
        print("Thank you")
    else:
        workbk_out.close()
        quit()
