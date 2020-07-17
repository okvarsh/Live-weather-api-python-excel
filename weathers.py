import xlwings as xw
import requests
import threading
import time

wb1=xw.Book()
wb1.save('weather.xlsx')

ws1=wb1.sheets[0]
ws1.range('A1').value=['City','Humidity', 'Temp','C/F','Update']

cities=["Delhi","London","Mumbai"]
updates=["1","1","1"]
units=["C","C","C"]

ws1.range('A2').options(transpose=True).value= cities
ws1.range('D2').options(transpose=True).value= units
ws1.range('E2').options(transpose=True).value= updates

apikey="b9415f7599fb9b6a40714f9448bc1f24"

def update_temp():
    for i,c in enumerate(cities):
        k=i+2
        cellE="E"+str(k)
        cellD="D"+str(k)
        if ws1.range(cellE).value == 1:
            ws1.cells(k,"B").value=None
            ws1.cells(k,"C").value=None
            if ws1.range(cellD).value == 'C' or ws1.range(cellD).value == 'c':
                unit="metric"
            elif ws1.range(cellD).value == 'F' or ws1.range(cellD).value == 'f':
                unit="imperial"
            print("calling api")
            url='https://api.openweathermap.org/data/2.5/weather?q='+c+'&units='+unit+'&appid='+apikey

            res=requests.get(url).json()
            temp= res['main']['temp']
            humi = res['main']['humidity']
            ws1.cells(k, "B").value = humi
            ws1.cells(k, "C").value = temp
    print("done")
    time.sleep(6)
    print("wake")

while True:
    update_temp()
