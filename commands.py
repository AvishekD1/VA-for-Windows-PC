from time import ctime # get time details
import webbrowser # open browser
import yfinance as yf # to fetch financial data
import subprocess as sp # to open windows applications
import pyttsx3 as tts

from functions import internet_on, respond, wake_word


def get_commands(voice_data):
    engine = tts.init()
    name=""
    #name
    if "what is your name" in voice_data or "what's your name" in voice_data or "tell me your name" in voice_data:
        if (name!=""):
            engine.say("my name is Avid")
            engine.runAndWait()
        else:
            engine.say("my name is Avid. what's your name?")
            engine.runAndWait()

    if "my name is" in voice_data:
        person_name = voice_data.split("is")[-1].strip()
        engine.say(f"okay, i will remember that {person_name}")
        engine.runAndWait()
        name = person_name # remember name in person object

    #greeting
    if "how are you" in voice_data or "how are you doing" in voice_data:
        engine.say(f"I'm very well, thanks for asking {name}")
        engine.runAndWait()

    #time
    if "what's the time" in voice_data or "tell me the time" in voice_data or "what time is it" in voice_data:
        time = ctime().split(" ")[3].split(":")[0:2]
        if time[0] == "00":
            hours = '12'
        else:
            hours = time[0]
        minutes = time[1]
        time = f'{hours} {minutes}'
        print (time)
        engine.say(time)
        engine.runAndWait()

    #search google
    if "search for" in voice_data and 'youtube' not in voice_data:
        if (internet_on()==True):
            search_term = voice_data.split("for")[-1]
            url = f"https://google.com/search?q={search_term}"
            webbrowser.get().open(url)
            engine.say(f'Here is what I found for {search_term} on google')
        else:
            engine.say(f'Please check your internet connection')
        engine.runAndWait()

    #search youtube
    if "youtube" in voice_data:
        if (internet_on()==True):
            search_term = voice_data.split("for")[-1]
            url = f"https://www.youtube.com/results?search_query={search_term}"
            webbrowser.get().open(url)
            engine.say(f'Here is what I found for {search_term} on youtube')
        else:
            engine.say(f'Please check your internet connection')
        engine.runAndWait()

    #get stock price
    if "price of" in voice_data:
        if (internet_on()==True):
            search_term = voice_data.lower().split(" of ")[-1].strip()
            stocks = {
                "apple":"AAPL",
                "microsoft":"MSFT",
                "facebook":"FB",
                "tesla":"TSLA",
                "bitcoin":"BTC-USD"
            }
            try:
                stock = stocks[search_term]
                stock = yf.Ticker(stock)
                price = stock.info["regularMarketPrice"]

                engine.say(f'price of {search_term} is {price} {stock.info["currency"]} {name}')
                engine.runAndWait()
            except:
                engine.say('oops, something went wrong')
                engine.runAndWait()
        else:
            engine.say(f'Please check your internet connection')
            engine.runAndWait()
    
    #open any website
    if "open" in voice_data and "application" not in voice_data and "ms" not in voice_data and "microsoft" not in voice_data:
        if (internet_on()==True):
            website = (voice_data.split("open")[-1].strip()).split()
            site = "".join(website)
            url = f"www.{site}.com/"
            webbrowser.get().open(url)
            engine.say(f'Opening {site} website')
        else:
            engine.say(f'Please check your internet connection')
        engine.runAndWait()

    #open calculator
    if "calculator" in voice_data:
        sp.Popen('C:\\Windows\\System32\\calc.exe') 
        engine.say(f'Opening calculator')
        engine.runAndWait()

    #open MS Application
    if "ms" in voice_data or "microsoft" in voice_data:
        if "word" in voice_data:
            sp.Popen('C:\Program Files (x86)\Microsoft Office\Office12\WINWORD.exe')
            engine.say("opening Microsoft word")
            engine.runAndWait()

        if "excel" in voice_data:
            sp.Popen('C:\Program Files (x86)\Microsoft Office\Office12\EXCEL.exe')
            engine.say("opening Microsoft excel")
            engine.runAndWait()

        if "power point" in voice_data or "powerpoint" in voice_data:
            sp.Popen('C:\Program Files (x86)\Microsoft Office\Office12\POWERPNT.exe')
            engine.say("opening Microsoft powerpoint")
            engine.runAndWait()

    #standby mode
    if "goodbye" in voice_data or "bye" in voice_data:
        engine.say("Going offline")
        engine.runAndWait()
        ini_comm = wake_word()
        respond(ini_comm)

    #exits the program
    if "exit" in voice_data or "quit" in voice_data:
        engine.say("Closing application")
        engine.runAndWait()
        exit(0)