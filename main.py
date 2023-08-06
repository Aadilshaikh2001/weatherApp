import requests
import json
import win32com.client as wincli

speak = wincli.Dispatch('SAPI.SpVoice')
try:

    city = input("Enter the name of city: ")
    url = f'https://api.weatherapi.com/v1/current.json?key=79087fa565e141539bd100940232806&q={city}'

    r = requests.get(url)
    # print(r.text)
    wdic = json.loads(r.text)
    r = wdic["location"]["region"]
    c = wdic["location"]["country"]
    w = wdic["current"]["temp_c"]
    h = wdic["current"]["humidity"]
    cl = wdic["current"]["cloud"]
    ws = wdic["current"]["wind_kph"]
    wd = wdic["current"]["wind_dir"]
    lu = wdic["current"]["last_updated"]
    print(f" region:- {r},\n country:- {c},\n Temperature:- {w},\n Humidity_level:- {h},\n Cloud_condition:- {cl},\n "
          f"wind_kph:- {ws},\n wind_dir:- {wd},\n Last_updated:- {lu}")
    speak.Speak(
          f"The present atmospheric conditions prevailing in {city}, {r}, {c} indicate a temperature of {w} degrees "
          f"Fahrenheit, accompanied by a humidity level of {h} percent and a cloud coverage ratio of {cl}. At this "
          f"moment in {city}, the wind speed registers {ws} kilometer per hour, blowing from the direction of {wd}. "
          f"This data was recently updated on {lu}."
    )
except Exception as e:
    print('Please enter city name')

speak.Speak('Thank you for using our weather information service. Have a great day!')
