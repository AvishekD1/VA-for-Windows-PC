import speech_recognition as sr # recognise speech
import pyttsx3 as tts
import random
import socket
import ssl
import certifi
import os # to remove created audio files

from commands import get_commands

ini_comm=""
r = sr.Recognizer() #initialise a recogniser
def internet_on():
    IPaddress=socket.gethostbyname(socket.gethostname())
    if IPaddress=="127.0.0.1":
        return False
    else:
        return True

def respond(ini_word):
    engine = tts.init()
    if ini_word=="hello":
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

def record_audio(ask=False):
    
    engine = tts.init()
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
    engine = tts.init()
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