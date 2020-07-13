import xlwings as xw
import requests
import threading
import time

wb1=xw.Book()
wb1.save('weather.xlsx')

ws1=wb1.sheets[0]
ws1.range('A1').value=['City', '*C', '*F']

cities=["Delhi","London","Karnataka","Mumbai"]

ws1.range('A2').options(transpose=True).value= cities

apikey="b9415f7599fb9b6a40714f9448bc1f24"
#this api is temporary, valid till 14th july 2020

def update_temp():
    for i,c in enumerate(cities):
        k=i+2
        
        ws1.cells(k,"B").value=None
        ws1.cells(k,"C").value=None

        unit="metric"
        urlc='https://api.openweathermap.org/data/2.5/weather?q='+c+'&units='+unit+'&appid='+apikey
        unit="imperial"
        urlf='https://api.openweathermap.org/data/2.5/weather?q='+c+'&units='+unit+'&appid='+apikey

        res=requests.get(urlc).json()
        temp= res['main']['temp']
        
        ws1.cells(k, "B").value = temp
        
        res=requests.get(urlf).json()
        temp= res['main']['temp']
        ws1.cells(k, "C").value = temp
        
    print("done")
    time.sleep(30)
    print("wake")
while True:
    update_temp()
