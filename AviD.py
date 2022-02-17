import speech_recognition as sr # recognise speech
import pyttsx3 as tts
import random
from time import ctime # get time details
import webbrowser # open browser
import yfinance as yf # to fetch financial data
import subprocess as sp # to open windows applications
import time
import socket 
from win10toast import ToastNotifier
from pymongo import MongoClient

import ssl
import certifi
import os 

#connecting database
client = MongoClient('127.0.0.1',27017) 
dbName = "AviD"
db = client[dbName]

#accessing collections
collection1 = db["user"] # user details
tweets = collection1.find()

collection2 = db["preferences"] # user preferences
tweets = collection2.find()

# collection3 = db["action"] # commands
# tweets = collection3.find()

collection4 = db["history"] # commands history
tweets = collection4.find()

class person:
    name = ''
    def setName(self, name):
        self.name = name

r = sr.Recognizer() #initialise a recogniser 
voice = 0 #voice of assistant, default male

toaster = ToastNotifier()
toaster.show_toast("AviD", "Say 'Hello' to continue!",
                icon_path="icons/wave.ico", 
                duration=5)

#listen for audio commands and convert it to text:
def record_audio(ask=False): 
    with sr.Microphone() as source: #microphone as source
        if ask:
            engine.say(ask)
            engine.runAndWait()
        audio = r.listen(source)  #listen for the audio via source
        voice_data = ''
        
        try:
            voice_data = r.recognize_google(audio)  #convert audio to text
        
        except sr.UnknownValueError: #error: recognizer does not understand
            engine.say('I did not get that')
            engine.runAndWait()
        
        except sr.RequestError:
            engine.say('Sorry, the service is down') #error: recognizer is not connected
            engine.runAndWait()
        
        print(f">> {voice_data.lower()}") #print what user said
        return voice_data.lower()

#Recognising the wake word to respond
def wake_word(ask=False):
    with sr.Microphone() as source: # microphone as source
        if ask:
            engine.say(ask) 
            engine.runAndWait()
        audio = r.listen(source)  # listen for the audio via source
        ini_comm = ''
        try:
            ini_comm = r.recognize_google(audio)  # convert audio to text
        except sr.UnknownValueError: # error: recognizer does not understand
            ini_comm = wake_word()
            respond(ini_comm)
        print(f">> {ini_comm.lower()}") # print what user said
    
        return ini_comm.lower()

#Check the availability of internet connection
def internet_on():
    IPaddress=socket.gethostbyname(socket.gethostname())
    if IPaddress=="127.0.0.1":
        return False
    else:
        return True


def get_commands(voice_data):

    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[voice].id)

    #name
    if "what is your name" in voice_data or "what's your name" in voice_data or "tell me your name" in voice_data:
        if person_obj.name:
            engine.say("my name is Avid")
            engine.runAndWait()
        else:
            engine.say("my name is Avid. what's your name?")
            engine.runAndWait()

    if "my name is" in voice_data:
        person_name = voice_data.split("is")[-1].strip()
        engine.say(f"okay, i will remember that {person_name}")
        engine.runAndWait()
        person_obj.setName(person_name) # remember name in person object

    #greeting
    if "how are you" in voice_data or "how are you doing" in voice_data:
        engine.say(f"I'm very well, thanks for asking {person_obj.name}")
        engine.runAndWait()

    # #set personal details
    # if "set personal details" in voice_data:


    #changing voice
    if "to female" in voice_data:
        voice = 1

    #time
    if "what's the time" in voice_data or "tell me the time" in voice_data or "what time is it" in voice_data or "what is the time" in voice_data:
        time = ctime().split(" ")[3].split(":")[0:2]
        if time[0] == "00":
            hours = '12'
        else:
            hours = time[0]
        minutes = time[1]
        time = f'{hours} {minutes}'
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

                engine.say(f'price of {search_term} is {price} {stock.info["currency"]} {person_obj.name}')
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
        toaster.show_toast("AviD", "Standby Mode. \nSay 'Hello' to continue!",
                        icon_path="icons/wave.ico",
                        duration=5)
        engine.runAndWait()
        ini_comm = wake_word()
        respond(ini_comm)

    #exits the program
    if "exit" in voice_data or "quit" in voice_data:
        engine.say("Closing application")
        toaster.show_toast("AviD", "Application Closed",
                        icon_path="icons/wave.ico",
                        duration=5)
        engine.runAndWait()
        exit(0)

def respond(ini_word):
    if ini_word=="hi":
        greetings = [f"hey, how can I help you?", f"hey, what's up?", f"hello, I'm listening"]
        greet = greetings[random.randint(0,len(greetings)-1)]
        engine.say(greet)
        engine.runAndWait()
        while (True):
            voice_data = record_audio()
            get_commands(voice_data)

    else:
        voice_data = wake_word()
        respond(ini_comm)

engine = tts.init()
# voices = engine.getProperty('voices')
# engine.setProperty('voice', voices[voice].id)
engine.setProperty('rate', 150)

time.sleep(1)
voice_data = ''

person_obj = person()

ini_comm = wake_word()
respond(ini_comm)