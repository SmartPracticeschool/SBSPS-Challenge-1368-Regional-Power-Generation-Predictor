import kivy
from kivy.app import App
from kivy.uix.label import Label
from kivy.uix.gridlayout import GridLayout
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
import requests
import xlrd
import xlsxwriter
import openpyxl
from xlutils.copy import copy

class MyGrid(GridLayout):
    def __init__(self,**kwargs):
        super(MyGrid, self).__init__(**kwargs)
        self.cols = 2
        self.rows = 4
        self.inside = GridLayout()
        self.inside.cols = 2
        self.inside.add_widget(Label(text="City: "))
        self.city = TextInput(multiline=False)#by default we have multiline
        self.inside.add_widget(self.city)
        self.inside.add_widget(Label(text="Day: "))
        self.day = TextInput(multiline=False)
        self.inside.add_widget(self.day)
        self.inside.add_widget(Label(text="Date: "))
        self.date = TextInput(multiline=False)
        self.inside.add_widget(self.date)
        self.add_widget(self.inside)
        self.submit = Button(text="Submit", font_size=40)
        self.submit.bind(on_press=self.pressed)
        self.add_widget(self.submit)
        self.add_widget(Label(text="Power Generated: "))
        self.gen = TextInput(multiline=False)
        self.add_widget(self.gen)
        self.add_widget(Label(text="Defeciet: "))
        self.defeciet = TextInput(multiline=False)
        self.add_widget(self.defeciet)
        self.add_widget(Label(text="Solar Timing: "))
        self.power = TextInput(multiline=False)
        self.add_widget(self.power)

    def pressed(self, instance):
        #city = self.city.text
        A = self.city.text
        B = self.date.text
        C = 'https://www.wunderground.com/hourly/in/'
        D = self.day.text
        E = ('Data_'+D)
        url = C + A +'/date/' +B
        print(url)
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
            if dict2[i] < 3.3:
                dict2[i] = 0
            elif dict2[i] > 20:
                dict2[i] = 0
            else:
                dict2[i] =  (0.5 * 1.23 *2826 * dict2[i] * dict2[i] * dict2[i])/1000
        rb = xlrd.open_workbook(E+'.xlsx')
        wb = copy(rb)
        w_sheet = wb.get_sheet(0)
        ml = 1
        for j in range(1,25):
            w_sheet.write(j,4,dict2[ml])
            wb.save(E+'.csv')
            ml += 1
        need = 15000
        sum = 0
        for i in dict2:
            sum += dict2[i]
            self.gen.text = str(sum)
        if sum <  need:
            self.defeciet.text = str(need - sum)
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
        print("Use solar at {} O'clock".format(xyz))
        for i in dict2:
            # for j in dict2:
            #     if j != 0:
            #         abc.append(i)
            self.power.text = str(dict2)

class MyApp(App):
    def build(self):
        return MyGrid()

if __name__ == "__main__":
    MyApp().run()
