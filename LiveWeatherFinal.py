import requests, json
import xlwings as xw
import time
import sys


def get_temperature(city_name):
    url = "http://api.openweathermap.org/data/2.5/weather?appid=d53280578bb32d47d66ab966197e5574&q="
    complete_url = url+city_name
    req1 = requests.get(complete_url)
    data = req1.json()
    tempr = data['main']['temp']
    c_id = data['id']
    c_name = data['name']
    c_desc = data['weather'][0]['description']
    return tempr,c_id,c_name,c_desc

def kelvin_to_celsius(tmp):
    c =  tmp - 273.15
    return round(c,2)

def kelvin_to_farenheit(tmp):
    f = ((tmp - 273.15) * 1.8) + 32
    return round(f,2)

def weather_xlsx():
    workbook = xw.Book('/Users/ASUS/Desktop/AyushWeatherAssignment/weather.xlsx')
    live = workbook.sheets['Sheet1']

    cities = workbook.sheets['Sheet2']

    print("Presss 'ctrl+c' to end the script! ")
    while True:
        
        total_cities = workbook.sheets[0].range('A' + str(workbook.sheets[0].cells.last_cell.row)).end('up').row
        for i in range(2, total_cities + 1):
            try:
                city = live.range('A' + str(i)).value
                data = get_temperature(city)
                if live.range('D' + str(i)).value == 1:
                    if live.range('C' + str(i)).value == 'C' or live.range('C' + str(i)).value == 'c':
                        live.range('B' + str(i)).value = str(kelvin_to_celsius(data[0])) + 'C'
                    elif live.range('C' + str(i)).value == 'F' or live.range('C' + str(i)).value == 'f':
                        live.range('B' + str(i)).value = str(kelvin_to_farenheit(data[0])) + 'F'
                        
                
                cities.range('A'+str(i)).value = data[1]
                cities.range('B'+str(i)).value = data[2]
                cities.range('D'+str(i)).value = data[3]

                cities.range('C'+str(i)).value = 'Auto Refresh'
                time.sleep(0.5)
                cities.range('C'+str(i)).value = str(kelvin_to_farenheit(data[0])) + 'F' +'/'+ str(kelvin_to_celsius(data[0])) + 'C'
            
            except KeyError:
                live.range('B'+str(i)).value = 'City Not found'
            
            except TypeError:
                live.range('B'+str(i)).value = 'City Not found'
            
            except KeyboardInterrupt:
                print("Thank you for using this Weather Automation Script")
                sys.exit()
            


        time.sleep(0.5) #updates value every 1 seconds


weather_xlsx()
