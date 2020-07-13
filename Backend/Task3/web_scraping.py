from bs4 import BeautifulSoup
import pandas as pd
import requests

lis_comment=[]
lis2=[]
lis3=[]
wind_sp=[]
lis_comment=[]
CiTy = input("Enter City:")
Dey = input("Enter Day:")
API = "iaQ8t1MaPqfqkE8VPZR0duvBHMWV3vl1"
countryCode = "IN"
a = "http://dataservice.accuweather.com/locations/v1/cities/"
b = "/search?apikey=iaQ8t1MaPqfqkE8VPZR0duvBHMWV3vl1&q="
c = "&details=true"
search_address = a + countryCode+ b + CiTy + c
json_data1 = requests.get(search_address).json()
location_key = json_data1[0]['Key']
link1 = 'https://www.accuweather.com/en/in/'
symbol = '/'
link2p2 = '/hourly-weather-forecast/'
link2p3 = '?day='
link2 = symbol + location_key + link2p2 + location_key + link2p3
link_final = (link1 + CiTy + link2 + Dey )
print(link_final)
agent = {"User-Agent":"Mozilla/5.0"}
page=requests.get(link_final, headers=agent).text
soup = BeautifulSoup(page,'lxml')
for i in soup.find_all('div', class_='hourly-card-nfl-header'):
	comment = i.find('span', class_='phrase')
	lis_comment.append(comment.text)
wind1=soup.find('div', class_='panel left')
# wind2=wind1.find('span',class_='value')
for i in wind1.find('span', class_='value'):
	lis3.append(i)
for i in soup.find_all('span',class_='value'):
	# lis_comment2=i.find('span',class_='value')
	lis2.append(i.text)
# print(wind2)
index_no=[0,8,16,24,32,40,48,57,66,75,84,93,102,111,120,129,138,147,156,165,173,181,189,197]
for i in index_no:
	wind_sp.append(lis2[i])

hehe=[i.strip('''WSW''') for i in wind_sp]
hehe2=[i.strip("km/h") for i in hehe]
wind_final=[i.strip(" ") for i in hehe2]
#comment final#----------------------------------------------------------------------------------------------------------
comment_final=[i.strip('''\n\t\t\t''') for i in lis_comment]
element=comment_final[0]
del comment_final[0]
comment_final.insert(len(comment_final),element)
print(comment_final)
#wind speed km/h #----------------------------------------------------------------------------------------------------------
element2=wind_final[0]
del wind_final[0]
wind_final.insert(len(wind_final),element2)
df = pd.DataFrame()
df['tes2'] = comment_final
df['test'] = wind_final
df.to_excel('test45.xlsx', index = False)
print(wind_final)
