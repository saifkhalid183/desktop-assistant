import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import os
import smtplib
import requests
import json
import random
from pygame import mixer               
mixer.init()

engine=pyttsx3.init('sapi5')
voices=engine.getProperty('voices')
engine.setProperty('voice',voices[1].id)

def speaknews(str):
    from win32com.client import Dispatch
    speak= Dispatch("SAPI.SpVoice")
    speak.Speak(str)

    
def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def greetMe():
    hour=int(datetime.datetime.now().hour)
    if(hour>=0 and hour<12):
        speak("Good Morning!")
    elif(hour>=12 and hour<18):
        speak("Good Afternoon!")
    else:
        speak("Good Evening!")
    speak("Hi Saif! I am JUNIOR. How may I help you?")
    
def Command():
    r=sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening....")
        r.pause_threshold=0.5
        audio=r.listen(source)

    try:
        print("Recognizing....")
        query=r.recognize_google(audio,language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        print(e)
        print("Try Again!!!!!!!!!!!!!!!!!!")
        return "None"
    return query
def sendEmail(to,content):
    server=smtplib.SMTP('smtp.gmail.com',587)
    server.ehlo()
    server.starttls()
    server.login('sender email address','sender email password')
    server.sendmail('sender email address',to,content)
    server.close()
if(__name__=="__main__"):
    greetMe()
    while True:
        query=Command().lower()
        if 'hi' in query or 'hello' in query:
            speak("hello and welcome!")

        elif 'wikipedia' in query:
            speak('Searching the WIKIPEDIA....')
            query=query.replace("wikipedia","")
            results=wikipedia.summary(query,sentences=3)
            speak("According to WIKIPEDIA")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            webbrowser.open("google.com")

        elif 'open stack overflow' in query:
            webbrowser.open("stackoverflow.com")

        elif 'play music' in query or 'play song' in query:
            music='G:\\songs'
            songs=os.listdir(music)
            for song in songs:
                print(song)
            nn=len(songs)
            i=random.randint(0,nn-1)
            print(f"Playing : {songs[i]}")
            gaana=os.path.join(music,songs[i])
            mixer.music.load(gaana)
            mixer.music.set_volume(0.7)
            mixer.music.play()
            while(True):
                print("To pause type p, to resume again type r, n for next random song, to exit type e and enter:\n")
                h=input()
                if(h=='p'):
                    mixer.music.pause()
                elif(h=='r'):
                    mixer.music.unpause()
                elif(h=='e'):
                    mixer.music.stop()
                    break
                elif(h=='n'):
                    i=random.randint(0,nn-1)
                    print(f"Playing : {songs[i]}")
                    gaana=os.path.join(music,songs[i])
                    mixer.music.load(gaana)
                    mixer.music.set_volume(0.7)
                    mixer.music.play()    
        
        elif 'time' in query:
            strTime=datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Hey! the time is {strTime}")

        elif 'date' in query:
            tdate=datetime.datetime.now().strftime("%d:%B:%Y:%A")
            speak(f"today's date is : {tdate}")

        elif 'open code' in query:
            path1="C:\\Users\\welcome\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(path1)

        elif 'open chrome' in query:
            path2="C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
            os.startfile(path2)

        elif 'what is your name' in query:
            print("Hi! I am JUNIOR.")
            speak("Hi! I am JUNIOR.")

        elif 'who am i' in query:
            print("You are Saif Khalid.")
            speak("You are Saif Khalid.")

        elif 'gmail' in query:
            try:
                speak("Sure, to type message say 1 or to use voice say 0")
                n=Command()
                if '0' in n:
                    speak("What is the message?")
                    content=Command()
                elif '1' in n:
                    speak("please type")
                    content=input()
                print("Type the receivers mail address")
                to=input()
                sendEmail(to,content)
                print("Done.")
                speak("Done.")
            except Exception as e:
                #print(e)
                print("Sorry some error occurred.")
            
        elif 'sports news' in query :
            url=('http://newsapi.org/v2/top-headlines?'
                'country=in&'
                'category=sports&'
                'apiKey=a14fc394cf6f46f3a452842b4288d4f0')
            r=requests.get(url)
            news=r.text
            parsednews=json.loads(news)
            i=0
            listt=["first","second","third","fourth","fifth","sixth","seventh","eighth","ninth","tenth"]
            v=len(parsednews)
            if v!=0:
                for i in range(v):
                    speaknews(f"{listt[i]} news is:")
                    print(parsednews['articles'][i]['title'])
                    speaknews(parsednews['articles'][i]['title'])
            else:
                speaknews("no new update for news")
                print("no new update for news") 

        elif 'india news' in query or 'news from india' in query:
            url=('http://newsapi.org/v2/top-headlines?'
                'country=in&'
                'apiKey=a14fc394cf6f46f3a452842b4288d4f0')
            r=requests.get(url)
            news=r.text
            parsednews=json.loads(news)
            i=0
            listt=["first","second","third","fourth","fifth","sixth","seventh","eighth","ninth","tenth"]
            v=len(parsednews)
            if v!=0:
                for i in range(v):
                    speaknews(f"{listt[i]} news is:")
                    print(parsednews['articles'][i]['title'])
                    speaknews(parsednews['articles'][i]['title'])
            else:
                speaknews("no new update for news")
                print("no new update for news") 
    
        elif 'news' in query:
            url=('http://newsapi.org/v2/top-headlines?'
                'sources=bbc-news&'
                'apiKey=a14fc394cf6f46f3a452842b4288d4f0')
            r=requests.get(url)
            news=r.text
            parsednews=json.loads(news)
            i=0
            listt=["first","second","third","fourth","fifth","sixth","seventh","eighth","ninth","tenth"]
            newss=parsednews['articles']
            v=len(newss)
            if v!=0:
                for i in range(v):
                    speaknews(f"{listt[i]} news is:")
                    print(parsednews['articles'][i]['title'])
                    speaknews(parsednews['articles'][i]['title'])
            else:
                speaknews("no new update for news")
                print("no new update for news")

        elif 'exit' in query:
            speak("Thank you and gooooooood to work with you.") 
            print("Thank you and gooooooood to work with you.") 
            exit()