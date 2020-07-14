import xlwings as xw
import requests
import threading
import time

wb1=xw.Book()
wb1.save('weather.xlsx')

ws1=wb1.sheets[0]
ws1.range('A1').value=['City','Humidity', '*C', '*F','Update']

cities=["Delhi","London","Karnataka","Mumbai"]
updates=["1","1","1","1"]

ws1.range('A2').options(transpose=True).value= cities
ws1.range('E2').options(transpose=True).value= updates

apikey="b9415f7599fb9b6a40714f9448bc1f24"
#this api is temporary, valid till 14th july 2020

def update_temp():
    for i,c in enumerate(cities):
        k=i+2
        cell="E"+str(k)
        if ws1.range(cell).value == 1:
            ws1.cells(k,"B").value=None
            ws1.cells(k,"C").value=None
            ws1.cells(k,"D").value=None

            unit="metric"
            urlc='https://api.openweathermap.org/data/2.5/weather?q='+c+'&units='+unit+'&appid='+apikey
            unit="imperial"
            urlf='https://api.openweathermap.org/data/2.5/weather?q='+c+'&units='+unit+'&appid='+apikey

            res=requests.get(urlc).json()
            temp1= res['main']['temp']
            humi = res['main']['humidity']
            ws1.cells(k, "B").value = humi
            ws1.cells(k, "C").value = temp1

            res=requests.get(urlf).json()
            temp2= res['main']['temp']
            ws1.cells(k, "C").value = temp1
            ws1.cells(k, "D").value = temp2
    updates=["1","1","0","1"]
    ws1.range('E2').options(transpose=True).value= updates
    print("done")
    time.sleep(3)
    print("wake")

while True:
    update_temp()
