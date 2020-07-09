import kivy
from kivy.app import App
from kivy.uix.label import Label
from kivy.uix.gridlayout import GridLayout
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
import requests
import xlsxwriter
import datetime
import schedule
import time

class MyGrid(GridLayout):
    def __init__(self,**kwargs):
        super(MyGrid, self).__init__(**kwargs)
        self.cols = 2
        self.rows = 4
        self.inside = GridLayout()
        self.inside.cols = 2
        self.inside.add_widget(Label(text="City: "))
        self.name = TextInput(multiline=False)#by default we have multiline
        self.inside.add_widget(self.name)
        self.inside.add_widget(Label(text="State: "))
        self.lastname = TextInput(multiline=False)
        self.inside.add_widget(self.lastname)
        self.inside.add_widget(Label(text="cycles: "))
        self.email = TextInput(multiline=False)
        self.inside.add_widget(self.email)
        self.add_widget(self.inside)
        self.submit = Button(text="Submit", font_size=40)
        self.submit.bind(on_press=self.pressed)
        self.add_widget(self.submit)
        self.add_widget(Label(text="Time: "))
        self.time = TextInput(multiline=False)
        self.add_widget(self.time)
        self.add_widget(Label(text="Wind Speed(m/s): "))
        self.wind = TextInput(multiline=False)
        self.add_widget(self.wind)
        self.add_widget(Label(text="Power Generated(kw/h): "))
        self.power = TextInput(multiline=False)
        self.add_widget(self.power)
    def pressed(self, instance):
        city = self.name.text
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
        sheet_out.write("C2",(0.5*1.23*2826*wind*wind*wind/1000))
        powar = 0.5*1.23*2826*wind*wind*wind/1000
        self.time.text = str(a)
        self.power.text = str(powar)
        self.wind.text = str(wind)
        workbk_out.close()

class MyApp(App):
    def build(self):
        return MyGrid()

if __name__ == "__main__":
    MyApp().run()
