import xlsxwriter
import datetime
import schedule
import time
import requests
import xlrd
import openpyxl
from xlutils.copy import copy
from bs4 import BeautifulSoup
import pandas as pd
import kivy
from kivy.core.window import Window
from kivy.app import App
from kivy.uix.label import Label
from kivy.uix.gridlayout import GridLayout
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.screenmanager import ScreenManager, Screen, FadeTransition
from kivy.uix.popup import Popup
from kivy.uix.image import AsyncImage
from kivy.graphics import Rectangle
from kivy.graphics import Color
from kivy.uix.widget import Widget

class ScreenManagement(ScreenManager):
    def __init__(self, **kwargs):
        super(ScreenManagement, self).__init__(**kwargs)

class Current_Window(Screen):
    def __init__(self, **kwargs):
        super(Current_Window, self).__init__(**kwargs)
        self.inside1 = Screen()
        self.inside1.cols = 2
        self.add_widget(Label(text='City',color=(0,0,0,1), size_hint=(.45, .1), pos_hint={'x': .05, 'y': .85}))
        self.City = TextInput(multiline=False, size_hint=(.45, .1), pos_hint={'x': .5, 'y': .85})
        self.add_widget(self.City)
        self.add_widget(Label(text='Wind(m/s)',color=(0,0,0,1), size_hint=(.45, .1), pos_hint={'x': .05, 'y': .4}))
        self.wind = TextInput(multiline=False, size_hint=(.45, .1), pos_hint={'x': .5, 'y': .4})
        self.add_widget(self.wind)
        self.add_widget(Label(text='Time',color=(0,0,0,1), size_hint=(.45, .1), pos_hint={'x': .05, 'y': .55}))
        self.time = TextInput(multiline=False, size_hint=(.45, .1), pos_hint={'x': .5, 'y': .55})
        self.add_widget(self.time)
        self.add_widget(Label(text='Power(kw/h)',color=(0,0,0,1), size_hint=(.45, .1), pos_hint={'x': .05, 'y': .25}))
        self.pawar = TextInput(multiline=False, size_hint=(.45, .1), pos_hint={'x': .5, 'y': .25})
        self.add_widget(self.pawar)
        self.btn5 = Button(text='calculate',background_color = (0.9,0.1,0.2,0.8), size_hint=(.9, .1), pos_hint={'center_x': .5, 'y': .7})
        self.add_widget(self.btn5)
        self.btn5.bind(on_press = self.calculate)
        self.inside1.btn8 = Button(text=' Goo back!',background_color = (0.9,0.1,0.2,0.8), size_hint=(.43, .1), pos_hint={'center_x': .75, 'y': .08})
        self.add_widget(self.inside1.btn8)
        self.inside1.btn8.bind(on_press = self.screen_transition)
        self.inside1.btn9 = Button(text='Clear!',background_color = (0.9,0.1,0.2,0.8), size_hint=(.43, .1), pos_hint={'center_x': .3, 'y': .08})
        self.add_widget(self.inside1.btn9)
        self.inside1.btn9.bind(on_press = self.screen_transition2)
    def calculate(self, *args):
        city = self.City.text
        workbk_out = xlsxwriter.Workbook("Wind.xlsx")
        sheet_out = workbk_out.add_worksheet()
        sheet_out.write("A1", "Time")
        sheet_out.write("B1", "Wind Speed")
        sheet_out.write("C1", "Power In Watt")
        n1 = 2
        n2 = 2
        n3 = 2
        now = datetime.datetime.now()
        a = now.strftime("%H:%M:%S")
        api_address='http://api.openweathermap.org/data/2.5/weather?appid=1222ec2c19edb278b4e39377e4138b42&q='
        url = api_address + city
        json_data = requests.get(url).json()
        wind = json_data['wind']['speed']
        # print('Wind Speed: {}'.format(wind))
        # print("Current Time is:", a )
        sheet_out.write("B2",wind )
        sheet_out.write("A2", a )
        sheet_out.write("C2",(0.5*1.23*2826*0.59*wind*wind*wind/1000))
        powar = 0.5*1.23*2826*0.59*wind*wind*wind/1000
        self.time.text = str(a)
        self.pawar.text = str(powar)
        self.wind.text = str(wind)
        workbk_out.close()
    def screen_transition2(self, *args):
        self.time.text = ''
        self.pawar.text = ''
        self.wind.text = ''
        self.City.text = ''
    def screen_transition(self, *args):
        self.time.text = ''
        self.pawar.text = ''
        self.wind.text = ''
        self.City.text = ''
        self.manager.current = 'login'
        Window.clearcolor = (1,1,1,1)

class Prediction_Window(Screen):
    def __init__(self, **kwargs):
        super(Prediction_Window, self).__init__(**kwargs)
        self.inside = Screen()
        self.inside.cols = 2
        self.add_widget(Label(text='City:',color=(0,0,0,1), size_hint=(.45, .08), pos_hint={'x': .05, 'y': 0.85}))
        self.City = TextInput(multiline=False, size_hint=(.45, .08), pos_hint={'x': .5, 'y': 0.85})
        self.add_widget(self.City)
        self.add_widget(Label(text='Which day in Future(eg-2/3):',color=(0,0,0,1), size_hint=(.45, .08), pos_hint={'x': .05, 'y': .75}))
        self.date = TextInput(multiline=False, size_hint=(.45, .08), pos_hint={'x': .5, 'y': .75})
        self.add_widget(self.date)
        self.add_widget(Label(text='Day:',color=(0,0,0,1), size_hint=(.45, .08), pos_hint={'x': .05, 'y': .65}))
        self.day = TextInput(multiline=False, size_hint=(.45, .08), pos_hint={'x': .5, 'y': .65})
        self.add_widget(self.day)
        self.add_widget(Label(text='Wind Mill power Forecast(kw/h):',color=(1,1,1,1), size_hint=(.45, .08), pos_hint={'x': .05, 'y': .27}))
        self.pawar = TextInput(multiline=False, size_hint=(.45, .08), pos_hint={'x': .5, 'y': .27})
        self.add_widget(self.pawar)
        self.add_widget(Label(text='Deficiet:',color=(1,1,1,1), size_hint=(.45, .08), pos_hint={'x': .05, 'y': .14}))
        self.Deficiet = TextInput(multiline=False, size_hint=(.45, .08), pos_hint={'x': .5, 'y': .14})
        self.add_widget(self.Deficiet)
        self.add_widget(Label(text='Alternative Source Available(eg- solar)',color=(0,0,0,1), size_hint=(.45, .08), pos_hint={'x': .05, 'y': .55}))
        self.solarX = TextInput(multiline=False, size_hint=(.45, .08), pos_hint={'x': .5, 'y': .55})
        self.add_widget(self.solarX)
        self.inside.btn7 = Button(text='Predict (Hold to see suggestion)',background_color = (0.9,0.1,0.2,0.9), size_hint=(.43, .08), pos_hint={'center_x': .75, 'y': .40})
        self.add_widget(self.inside.btn7)
        self.inside.btn7.bind(on_press = self.pressed)
        self.inside.btn15 = Button(text='Hold to see Ideal timing to use Wind Mill',background_color = (0.9,0.1,0.2,0.9), size_hint=(.43, .08), pos_hint={'center_x': .3, 'y': .40})
        self.add_widget(self.inside.btn15)
        self.inside.btn15.bind(on_press = self.pressed2)
        self.inside.btn8 = Button(text=' Goo back!',background_color = (0.9,0.1,0.2,0.9), size_hint=(.43, .08), pos_hint={'center_x': .75, 'y': .03})
        self.add_widget(self.inside.btn8)
        self.inside.btn8.bind(on_press = self.screen_tronsition)
        self.inside.btn9 = Button(text='Clear!',background_color = (0.9,0.1,0.2,0.9), size_hint=(.43, .08), pos_hint={'center_x': .3, 'y': .03})
        self.add_widget(self.inside.btn9)
        self.inside.btn9.bind(on_press = self.screen_tronsition2)
    def pressed(self, instance):
        #city = self.city.text
        A = self.City.text
        B = self.date.text
        D = self.day.text
        E = ('Data_'+D)
#---------------------------Scrape--------------------------------------------
        lis_comment=[]
        lis2=[]
        lis3=[]
        wind_sp=[]
        wind_final = []
        lis_comment=[]
        CiTy = self.City.text
        Dey = self.date.text
        API = "pGXdvkVtBFC0SsTgUjBgvHz1ZYPHuO2I"
        countryCode = "IN"
        aaa = "http://dataservice.accuweather.com/locations/v1/cities/"
        bbb = "/search?apikey=pGXdvkVtBFC0SsTgUjBgvHz1ZYPHuO2I&q="
        ccc = "&details=true"
        search_address = aaa + countryCode+ bbb + CiTy + ccc
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
        lis5 = []
        lis7 = []
        for i in lis2:
          if 'km/h' in i:
              lis7.append(i)
          else:
              pass
        index_wind = 0
        while index_wind <= 47:
            lis5.append(lis7[index_wind])
            index_wind += 2

        hehe0 = [i.strip('''WSW''') for i in lis5]
        hehe = [i.strip('''NW''') for i in hehe0]
        hehe2 = [i.strip("km/h") for i in hehe]
        wind_final = [i.strip(" ") for i in hehe2]
        # index_no=[0,8,16,24,32,40,48,57,66,75,84,93,102,111,120,129,138,147,156,165,173,181,189,197]
        # for i in index_no:
        # 	wind_sp.append(lis2[i])
        # hehe=[i.strip('''WSW''') for i in wind_sp]
        # hehe2=[i.strip("km/h") for i in hehe]
        # hehe3=[i.strip("(Low)") for i in hehe2]
        # hehe4=[i.strip("(Moderate)") for i in hehe3]
        # hehe5=[i.strip("(High)") for i in hehe4]
        # hehe6=[i.strip("(99%)") for i in hehe5]
        # hehe7=[i.strip(' (Very') for i in hehe6]
        # hehe8=[i.strip(' (Extrem') for i in hehe7]
        # hehe9=[i.strip('''NW''') for i in hehe8]
        # hehe10=[i.strip('''N''') for i in hehe9]
        # hehe11=[i.strip('''W''') for i in hehe10]
        # hehe12=[i.strip('''S''') for i in hehe11]
        # hehe13=[i.strip('''E''') for i in hehe12]
        # hehe14=[i.strip('''NWN''') for i in hehe13]
        # wind_final=[i.strip("' '") for i in hehe14]
        #comment final#----------------------------------------------------------------------------------------------------------
        comment_final=[i.strip('''\n\t\t\t''') for i in lis_comment]
        element=comment_final[0]
        del comment_final[0]
        comment_final.insert(len(comment_final),element)
        #print(comment_final)
        #wind speed km/h #----------------------------------------------------------------------------------------------------------
        element2=wind_final[0]
        del wind_final[0]
        wind_final.insert(len(wind_final),element2)
        toimee = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24]
        print(wind_final)
        for i in range(0, len(wind_final)):
            wind_final[i] = int(wind_final[i])
        wind_fional = [i * 0.277778 for i in wind_final]
        df = pd.DataFrame()
        df['Time'] = toimee
        df['Conditions'] = comment_final
        df['Wind(m/s)'] = wind_fional
        df['Wind(km/h)'] = wind_final
        df.to_excel(E+'.xlsx', index = False)
        #print(wind_final)
#--------------------------End Scrape-----------------------------------------
        workbook = xlrd.open_workbook(E+'.xlsx')
        sheet = workbook.sheet_by_index(0)
        row_count = sheet.nrows
        lis1=[]
        lis2=[]
        lis3=[]
        lis4=[]
        dict1={}
        dict2={}
        dicttime={}
        time=1
        time2=1
        rw_no=1
        n1=1
        n2=0
        xyz=[]
        abc=[]
        def solar():
            n1 = -1
            n2 = -1
            n = -1 #index of lis3
            for i in range (1,row_count):
                solar5 = sheet.cell_value(i,1)#lis2=theory
                lis2.append(solar5)
            for i in range (1,row_count):
                solar3 = sheet.cell_value(i,0)
                lis3.append(solar3)#lis3=times
        for i in range (1,row_count):
            sv = sheet.cell_value(i,2)
            lis1.append(sv)
        for speed in lis1:
            dict2[time]=speed
            time+=1
        for i in dict2:
            if dict2[i] < 2.5:
                dict2[i] = 0
            elif dict2[i] > 20:
                dict2[i] = 0
            else:
                dict2[i] =  (0.5 * 1.23 * 2826 * 0.59 * dict2[i] * dict2[i] * dict2[i])/1000
        rb = xlrd.open_workbook(E+'.xlsx')
        wb = copy(rb)
        w_sheet = wb.get_sheet(0)
        ml = 1
        for j in range(1,25):
            w_sheet.write(j,4,dict2[ml])
            wb.save(E+'.csv')
            ml += 1
        need = 5000
        sum = 0
        for i in dict2:
            sum += dict2[i]
            self.pawar.text = str(sum)
        if sum <  need:
            self.Deficiet.text = str(need - sum)
            solar()
            for i in lis2:
                if i == "Partly Cloudy":
                    xyz.append(n1)
                n1+=1
        elif sum > need:
            print('Excess Power Generated:', sum - need)
            n1 += 1
        else:
            print('Requirements Stisfied:')
            n1 += 1
        w_sheet.write(27,4,sum)
        wb.save(E+'.csv')
        # print(lis2[n1])
        defoicoit = need - sum
        z = defoicoit/1000
        input_solar = self.solarX.text
        a = ("Use can solar during {} O'clock to compemsate the shortfall".format(xyz))
        b = ('Use generator for',+z,'hours to compensate the shortfall')
        def solar_pop(self):
            layout = GridLayout(cols = 1, padding = 10)
            popupLabel = Label(text = str(a))
            layout.add_widget(popupLabel)
            # Instantiate the modal popup and display
            popup = Popup(title ='Solar Popup',size_hint= (0.9,0.8),size =(200,200), content = layout)
            popup.open()
        def diesel_pop(self):
            layout = GridLayout(cols = 1, padding = 10)
            popupLabel = Label(text = str(b))
            layout.add_widget(popupLabel)
            popup = Popup(title ='Diesel Popup',size =(300,300), content = layout)
            popup.open()
        if input_solar == 'solar':
            solar_pop(self)
        elif input_solar == 'diesel':
            z = defoicoit/1000
            diesel_pop(self)
    def pressed2(self,instance):
        layout = GridLayout(cols = 1, padding = 10)
        popupLabel = Label(text = 'Maximum Output from turbine during: {3,4,5}')
        layout.add_widget(popupLabel)
        # Instantiate the modal popup and display
        popup = Popup(title ='Timer',size_hint= (0.9,0.8),size =(200,200), content = layout)
        popup.open()

    def screen_tronsition2(self, *args):
        self.solarX.text = ''
        self.Deficiet.text = ''
        self.pawar.text = ''
        self.day.text = ''
        self.date.text = ''
        self.City.text = ''
    def screen_tronsition(self, *args):
        self.solarX.text = ''
        self.Deficiet.text = ''
        self.pawar.text = ''
        self.day.text = ''
        self.date.text = ''
        self.City.text = ''
        self.manager.current = 'login'
        Window.clearcolor = (1,1,1,1)

class Setup_Window(Screen):
    def __init__(self, **kwargs):
        super(Setup_Window, self).__init__(**kwargs)
        self.inside = Screen()
        self.inside.cols = 2
        self.add_widget(Label(text='Number of Turbines:',color=(0,0,0,1), size_hint=(.45, .08), pos_hint={'x': .05, 'y': 0.85}))
        self.turbine = TextInput(multiline=False, size_hint=(.45, .08), pos_hint={'x': .5, 'y': 0.85})
        self.add_widget(self.turbine)
        self.add_widget(Label(text='Radius of Turbine:',color=(0,0,0,1), size_hint=(.45, .08), pos_hint={'x': .05, 'y': .75}))
        self.radius = TextInput(multiline=False, size_hint=(.45, .08), pos_hint={'x': .5, 'y': .75})
        self.add_widget(self.radius)
        self.btn7 = Button(text='Submit',background_color = (0.9,0.1,0.2,0.9), size_hint=(.9, .1), pos_hint={'center_x': .5, 'y': .40})
        self.add_widget(self.btn7)
        self.inside.btn8 = Button(text=' Goo back!',background_color = (0.9,0.1,0.2,0.8), size_hint=(.43, .08), pos_hint={'center_x': .75, 'y': .03})
        self.add_widget(self.inside.btn8)
        self.inside.btn8.bind(on_press = self.screen_tronsition)
        self.inside.btn9 = Button(text='Clear!',background_color = (0.9,0.1,0.2,0.8), size_hint=(.43, .08), pos_hint={'center_x': .3, 'y': .03})
        self.add_widget(self.inside.btn9)
        self.inside.btn9.bind(on_press = self.screen_tronsition2)
    def screen_tronsition2(self, *args):
        self.radius.text = ''
        self.turbine.text = ''
    def screen_tronsition(self, *args):
        self.manager.current = 'login'
        Window.clearcolor = (1,1,1,1)

class Guide_Window(Screen):
    def __init__(self, **kwargs):
        super(Guide_Window, self).__init__(**kwargs)
        with open("Guide.txt") as f:
            contents = f.read()
        self.add_widget(Label(text=contents,color=(1,1,1,1),font_size=(20)))
        self.btn6 = Button(text='Goo back!',background_color = (0.9,0.1,0.2,1), size_hint=(.5, .1), pos_hint={'center_x': .5, 'y': .03})
        self.add_widget(self.btn6)
        self.btn6.bind(on_press = self.screen_transition)
    def screen_transition(self, *args):
        self.manager.current = 'login'
        Window.clearcolor = (1,1,1,1)

class LoginWindow(Screen):
    def __init__(self, **kwargs):
        super(LoginWindow, self).__init__(**kwargs)
        self.cols = 2
        self.outside = Screen()
        self.outside.cols = 2
        self.add_widget(AsyncImage(source="chan55.png",size_hint=(.9, .27), pos_hint={'x': .06, 'y': .71}))
        self.btn2 = Button(text='Current Mode',background_color = (0.2,0.5,1,0.8), size_hint=(.9, .2), pos_hint={'center_x': .5, 'y': .5})
        self.add_widget(self.btn2)
        self.btn2.bind(on_press = self.screen_transition)
        self.btn3 = Button(text='Prediction Mode',background_color = (0.2,0.5,1,0.8), size_hint=(.9, .2), pos_hint={'center_x': .5, 'y': .28})
        self.add_widget(self.btn3)
        self.btn3.bind(on_press = self.screen_tronsition)
        self.outside.btn10 = Button(text='Guide',background_color = (0.2,0.5,1,0.8), size_hint=(.45, .2), pos_hint={'center_x': .74, 'y': .06})
        self.add_widget(self.outside.btn10)
        self.outside.btn10.bind(on_press = self.screen_transition1)
        self.outside.btn11 = Button(text='Set Up prerequisites',background_color = (0.2,0.5,1,0.8), size_hint=(.45, .2), pos_hint={'center_x': .27, 'y': .06})
        self.add_widget(self.outside.btn11)
        self.outside.btn11.bind(on_press = self.screen_transition3)
    def screen_tronsition(self, *args):
        self.manager.current = 'Prediction'
    def screen_transition(self, *args):
        #Window.clearcolor = (198/255,74/255,69/255,1)
        self.manager.current = 'Current'
    def screen_transition1(self, *args):
        self.manager.current = 'Guide'
    def screen_transition3(self, *args):
        self.manager.current = 'Setup'

class Application(App):
    def build(self):
        Window.clearcolor = (1,1,1,1)
        Window.add_widget((AsyncImage(source="https://images.unsplash.com/photo-1525723550961-7a8f846d6ba7?ixlib=rb-1.2.1&ixid=eyJhcHBfaWQiOjEyMDd9&auto=format&fit=crop&w=1049&q=80",size_hint=(1.3,1.158))))
        sm = ScreenManagement(transition=FadeTransition())
        sm.add_widget(LoginWindow(name='login'))
        sm.add_widget(Current_Window(name='Current'))
        sm.add_widget(Prediction_Window(name='Prediction'))
        sm.add_widget(Guide_Window(name='Guide'))
        sm.add_widget(Setup_Window(name='Setup'))
        return sm

if __name__ == "__main__":
    Application().run()
