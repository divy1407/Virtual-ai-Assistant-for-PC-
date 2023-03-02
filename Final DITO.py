from twilio.rest import Client
import win32com.client
import pyttsx3 
import speech_recognition as sr
import wikipedia
import webbrowser
import os
import random
import playsound
import subprocess
import requests
import wolframalpha
from selenium import webdriver
import json
from urllib.request import urlopen
import newsapi
from newsapi import newsapi_client
import pyjokes
import requests
from bs4 import BeautifulSoup
import pygame
import os
import datetime
import psutil
import socket
import threading
from threading import Thread
import pyttsx3
from tkinter import *
from tkinter import simpledialog, Tk
from PIL import ImageTk, Image



destroy= False
state = 0
user_panel_flag = False
REMOTE_SERVER = "1.1.1.1"
app_music = ['open-ended', 'unsure', 'when', 'sorted']
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)


def resource_path(relative_path):
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def speak(audio):
    print('DITO Said:', audio)
    engine.say(audio)
    engine.runAndWait()


def myCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.energy_threshold = 1200
        r.pause_threshold = 0.7
        r.dynamic_energy_adjustment_damping = 0.4
        audio = r.listen(source)
        
        try:
            print("recongnising....")
            query = r.recognize_google(audio, language='en-UK')
            print("User Said:", format (query))
            return query

        except Exception as e:
            print("Exception: Sorry...I couldn't  recognize what u said " + str(e))
            (print('Say that again please ....'))
            speak('Could u please say that again ...')
            said = myCommand()
    return said


def takecommands():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Waiting to be called \ Listening...")
        r.energy_threshold = 1200
        r.pause_threshold = 0.7
        r.dynamic_energy_adjustment_damping = 0.4
        audio = r.listen(source)
        try:
            print("recongnising....")
            query = r.recognize_google(audio, language='en-UK')
            print("User Said:", format (query))
            return query
        except Exception as e:
            print("say that again")
            return "None"
    return "None"


def wakeWord(text):
    WAKE_WORDS = ['hey', 'hi', 'hello','hola']
    text = text.lower()
    for phrase in WAKE_WORDS:
        if phrase in text:
            return True
    return False


def wishMe():
    hour = int(datetime.datetime.now().hour) 
    if hour>=3 and hour <=12:
        speak("Good Morning")
    elif hour>=12 and hour<=16:
        speak("Good Afternoon")
    else:
        speak("Good Evening")
    speak("welcome to the virtual assistant, I'm dito. how may i help you.")


def get_typed_query():
    query = simpledialog.askstring("", "Please type your command")
    #print(query)
    return str(query)



notification_folder = resource_path('Notification_sounds')
IMAGE_PATH = resource_path('D:/Programming Codes/VS Codes/Jarvis/back.png')


class Widget:
  
    def destroy_root(self):
        destroy = True
        sys.exit()
        self.root.destroy()
    

    def clicked(self):      
        print("in clicked")
        speak("what do you want me to search or open")

        self.t_btn.config(state='disabled')
        self.type_query_button.config(state='disabled')
        self.root.update()

        self.userText.set('Listening...')
        query = myCommand()
        print(query)
        self.userText.set(query)
        query = query.lower()

        self.execute(query)

        self.t_btn.config(state='normal')
        self.type_query_button.config(state='normal')
        self.root.update()
        return


    def typed(self):
        print("in typed")
        speak("type here. whatever you want to open or ask")
 
        self.t_btn.config(state='disabled')
        self.mic_button.config(state='disabled')
        self.root.update()

        query = get_typed_query()
        if query != 'None':
            self.execute(query)

        self.t_btn.config(state='normal')
        self.mic_button.config(state='normal')
        self.root.update()
        return


    def voice_command_activation_switch(self):
        
                self.t_btn.config(image=self.voice_command_on)
                self.t_btn.config(state='disabled')
                self.mic_button.config(state='disabled')
                self.enable_keep_listening.config(text="Listening..", font=('Black Ops One', 10, 'bold'))
                self.t_btn.config(state='disabled')
                self.root.update()
                state = 1

                # self.listen_continuously()
                threading.Thread(target=self.listen_continuously()).start()
                self.t_btn.config(state='normal')
                self.mic_button.config(state='normal')
                self.root.update()
                return
       

    def listen_continuously(self):
        global state ,destroy
        #print("state in listen continuouslly fn is :" + str(state))
        while True and not destroy :
            query = takecommands()
            #print("in listen continuously query recieved is:" + query)
            if state == 1 and 'stop listening' not in query:
                #print("state in listen continuouslly fn if loop is :" + str(state))

                if wakeWord(query):
                    print("waked up successfully")
                    speak("How may I help you sir")
                    command = myCommand()
                    stop_flag = self.execute(command)
                    if stop_flag:
                        break
                    else:
                        continue
                else:
                    continue
            else:
                speak('Listening has been stoped now')
                self.t_btn.config(image=self.voice_command_off)
                self.t_btn.config(state='normal')
                self.enable_keep_listening.config(text=" Enable Listening", font=('Black Ops One', 10, 'bold'))
                state = 0
                break
        self.t_btn.config(image=self.voice_command_off)
        self.t_btn.config(state='normal')
        self.enable_keep_listening.config(text=" Enable Listening", font=('Black Ops One', 10, 'bold'))
        self.root.update()
        state = 0

    
    def __init__(self):
      
        self.root = Tk()
        self.root.geometry('800x400+-5+0')
        self.root.iconbitmap(resource_path("D:/Programming Codes/VS Codes/Jarvis/icon.ico"))
        
        img = ImageTk.PhotoImage(Image.open(IMAGE_PATH))

        panel_main = Label(self.root, image=img)
        panel_main.pack(expand="yes", side='top')
        self.root.title('DITO Virtual Assistant')

        mic_img = ImageTk.PhotoImage(Image.open(resource_path("D:/Programming Codes/VS Codes/Jarvis/mic.png")))
        close_button_img = PhotoImage(file=resource_path("D:/Programming Codes/VS Codes/Jarvis/forcestp.png"))

        self.voice_command_on = PhotoImage(file=resource_path("D:/Programming Codes/VS Codes/Jarvis/onn.png"))
        self.voice_command_off = PhotoImage(file=resource_path("D:/Programming Codes/VS Codes/Jarvis/off.png"))

        self.enable_keep_listening = Label(self.root, text="Enable Listening ", bg='gray26',font=('Black Ops One', 10, 'bold'),fg='white')
        self.enable_keep_listening.place(x=5, y=30)
        
        self.t_btn = Button(self.root, image=self.voice_command_off, border=0, bg='gray26',activebackground='gray26', command=self.voice_command_activation_switch)
        self.t_btn.place(x=20, y=75)
        
        self.mic_button = Button(self.root, image=mic_img, font=('Black ops one', 10, 'bold'), bg='#3FE0D0',height='45',width='45', border=3,command= self.clicked)
        self.mic_button.place(x=700, y=230)
       
        self.type_query_button = Button(self.root, text="Type Your Command", font=('Black ops one', 10, 'bold'), bg='#00B0FF',border=3,command=self.typed)
        self.type_query_button.place(x=650, y=190)

        self.close_button = Button(self.root,image=close_button_img, border= 1,font=('Black Ops One', 10, 'bold'), bg='white', fg='white',command=self.destroy_root).place(x=690,y=10)

        self.userText = StringVar()
        self.userText.set('Click on Mic Button')

        userFrame = LabelFrame(self.root, text="User Said:", font=('Black ops one', 10, 'bold'))
        userFrame.pack(fill="both", expand="yes", side='bottom')

        User_Message = Message(userFrame, textvariable=self.userText, bg='#73C2FB', fg='white')
        User_Message.config(font=("Comic Sans MS", 10, 'bold'))
        User_Message.pack(fill='both', expand='yes', )

        wishMe()
        self.root.mainloop()


    def execute(self, query):
        if not destroy:
            query = query.lower()

        if 'exit' in query or 'shutdown' in query or 'bye' in query or 'sleep' in query:
            speak('thank you so much for using me, goodbye')
            self.destroy_root()

        elif "open r studio" in query:
            code_path = "C://Program Files//RStudio//bin//rstudio.exe"
            speak("opening r studio")
            os.startfile(code_path)
        
        elif 'open word' in query:
            code_path = "C://Program Files//Microsoft Office//root//Office16//WINWORD.EXE"
            speak("opening MS Word")
            os.startfile(code_path)
        
        elif 'open powerpoint' in query:
            code_path = "C://Program Files//Microsoft Office//root//Office16//POWERPNT.EXE"
            speak("opening powerpoint")
            os.startfile(code_path)

        elif 'open excel' in query:
            code_path = "C://Program Files//Microsoft Office//root//Office16//EXCEL.EXE"
            speak("opening excel")
            os.startfile(code_path)
        
        elif 'open video editor' in query:
            code_path = "C://Program Files//OpenShot Video Editor//openshot-qt.exe"
            speak("opening video editor")
            os.startfile(code_path)
        
        elif 'open youtube' in query:
            speak("opening youtube")
            webbrowser.open("https://youtube.com/")
        
        elif 'open github' in query:
            speak('opening geet-hub')
            webbrowser.open("https://github.com/")

        elif 'open google' in query:
            speak("opening google")
            webbrowser.open("https://google.com/")

        elif 'play music' in query:
            music_dir = 'D://Songs'
            songs = os.listdir(music_dir)
            Value = random.randint(0,715)
            print(songs[Value])
            speak("playing song")
            os.startfile(os.path.join(music_dir, songs[Value]))

        elif 'the time' in query:
            string_time1 = datetime.datetime.now().strftime("%H")
            string_time2 = datetime.datetime.now().strftime("%M")
            speak(f"the time is {string_time1} hours and {string_time2} minutes")   

        elif 'wikipedia' in query or "who is" in query:
            query = query.replace("wikipedia","")
            results = wikipedia.summary(query, sentences=2)
            speak('according to wikipedia')
            print(results)
            speak(results)

        elif 'news' in query:    
            try: 
                i = 1 
                news_api = os.environ['news_api']
                jsonObj = urlopen("https://newsapi.org/v2/top-headlines?country=in&apiKey="+news_api) 
                data = json.load(jsonObj) 
                speak('here are the top 20 news of india') 
                for item in data['articles']:    
                    print(str(i) + '. ' + item['title'] + '\n') 
                    print(item['description'] + '\n') 
                    speak(str(i) + '. ' + item['title'] + '\n')
                    i += 1
            except Exception as e:      
                print(str(e))

        elif 'joke' in query: 
            speak(pyjokes.get_joke()) 

        elif "weather" in query: 
            weatherapi = os.environ['weather_api']
            speak("which citys weather you want") 
            cityname = myCommand() 
            print(f"City name : {cityname}") 
            completeurl = "http://api.openweathermap.org/data/2.5/weather?q="+cityname+"&appid="+weatherapi
            api_link = requests.get(completeurl)  
            api_data = api_link.json()  
              
            if api_data["cod"] == "404":  
                speak(" City Not Found ")    
            else:  
                xy = api_data['main'] 
                temprature = ((xy["temp"]-273))  
                pressure = xy["pressure"]  
                humidity = xy["humidity"]  
                wind_speed = api_data["wind"]["speed"]
                description = api_data["weather"][0]["description"]
                speak("current temprature is {:.2f} degree celsius".format(temprature))
                print("Temprature is: {:.2f} deg C".format(temprature))
                speak(f"it feels like {description}")
                print(f"Feels Like {description}")
                speak(f"wind speed is {wind_speed} kilometer per hour")
                print(f"Wind-Speed is {wind_speed}kmpl")
                speak(f"current humidity is {humidity} percentage")
                print(f"Humidity {humidity}%")
                speak(f"current atmospheric pressure is {pressure} pascal")
                print(f"{cityname}'s Atmospheric Pressure is {pressure} Pascals")
        
        elif "what is" in query: 
            wolfid = os.environ["wolf_api"]
            client = wolframalpha.Client(wolfid) 
            ressult = client.query(query) 
              
            try: 
                print (next(ressult.results).text)
                speak ("the required answer to your question is ") 
                speak (next(ressult.results).text) 
            except:
                speak("no result found")
                print ("No results")

        elif "send message" in query:
            acc = os.environ["twilio_api"]
            ath = os.environ['twilio_token']
            client = Client(acc, ath)
            speak("what do u want to type")
            message = client.messages \
                .create(
                        body=myCommand(),
                        from_=os.environ['twilio_no'],
                        to=os.environ['my_no']  
                    )
            print(message.sid)
            print("\nSuccessfully Sent")
            speak("message successfully sent")

        elif "how are you" in query:
            speak("i'm fine. how are you")
            webbrowser.open("https://github.com/")

        elif "maps" in query: 
            speak("which city or location you want to search")
            place = myCommand()
            speak("User asked to find") 
            speak(place) 
            webbrowser.open("https://www.google.nl/maps/place/"+place+"/")


if __name__ == '__main__':
    widget = Widget()