import pyttsx3  # pip install pyttsx3
import speech_recognition as sr  # pip install speechRecognition
import datetime
import wikipedia  # pip install wikipedia
import webbrowser
import os
import smtplib
import random
import json
import requests
# import subprocess
# import wolframalpha
# import tkinter
# import random
# import operator
# import datetime
# import wikipedia
# import webbrowser
# import os
# import winshell
# import pyjokes
# import feedparser
# import ctypes
# import time
# import requests
# import shutil
# from twilio.rest import Client
# from clint.textui import progress
# from ecapture import ecapture as ec
# from bs4 import BeautifulSoup
# import win32com.client as wincl
# from urllib.request import urlopen


engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices[1].id)
engine.setProperty('voice', voices[0].id)
newVoiceRate = 180
engine.setProperty('rate', newVoiceRate)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning")

    elif hour >= 12 and hour < 18:
        speak("Good Afternoon")

    else:
        speak("Good Evening")

    print("Hello sir!")
    speak("Hello sir!")


def takeCommand():
    # It takes microphone input from the user and returns string output

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening Sir...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-US')
        print(f"User said: {query}\n")

    except Exception as e:
        # print(e)
        print("Say that again please...")
        return "None"
    return query


def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('youremail@gmail.com', 'your-password')
    server.sendmail('youremail@gmail.com', to, content)
    server.close()


if __name__ == "__main__":
    wishMe()
    while True:
        # if 1:
        query = takeCommand().lower()

        # Logic for executing tasks based on query
        if 'wikipedia' in query:
            speak('Searching Wikipedia... Please Wait Sir!')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=20)
            print("According to Wikipedia\n")
            speak("According to Wikipedia")
            print(results)
            speak(results)
        elif 'open youtube' in query:
            speak("opening youtube Please Wait...")
            webbrowser.open("https://www.youtube.com/")
        elif "how are you" in query:
            speak("I'm fine, glad you me that")
            speak("How are you, Sir")
            print("I'm fine, glad you me that")
            print("How are you, Sir")
        elif 'fine' in query or "good" in query:
            print("It's good to know that your fine")
            speak("It's good to know that your fine")

        elif "what's your name" in query or "What is your name" in query:
            print("I am jarvis")
            speak("I am Jarvis")
        elif "where are you from" in query or "Where do you live" in query:
            print("I am from  internet")
            speak("I am from internet")
        elif 'is love' in query:
            speak("It is 7th sense that destroy all other senses")

        elif "who are you" in query:
            print("I am your virtual assistant created by Rohit Rayaan")
            speak("I am your virtual assistant created by Rohit Rayaan")

        elif "thank you" in query:
            print("you welcome sir!")
            speak("you welcome  sir!")

        elif 'reason for you' in query:
            speak("I was created as a Minor project by Mister Rohit Rayaan ")

        elif "weather" in query:

            # Google Open weather website
            # to get API of Open weather
            api_key = "Api key"
            base_url = "http://api.openweathermap.org / data / 2.5 / weather?"
            speak(" City name ")
            print("City name : ")
            city_name = takeCommand()
            complete_url = base_url + "appid =" + api_key + "&q =" + city_name
            response = requests.get(complete_url)
            x = response.json()

        elif 'open stack' in query:
            webbrowser.open("https://stackoverflow.com/")
        elif 'open google' in query:
            speak("opening google Please Wait...")
            webbrowser.open("https://www.google.com/")
        elif 'open facebook' in query:
            speak("opening facebook Please Wait...")
            webbrowser.open("https://www.facebook.com/")
        elif 'open instagram' in query:
            speak("opening instagram Please Wait...")
            webbrowser.open("https://www.instagram.com/")
        elif 'open twitter' in query:
            speak("opening twitter Please Wait...")
            webbrowser.open("https://www.twitter.com")
        elif 'open amazon' in query:
            speak("opening amazon Please Wait...")
            webbrowser.open("https://www.amazon.com")
        elif 'open wikipedia' in query:
            webbrowser.open("https://www.wikipedia.com")
        elif 'open geeks' in query:
            webbrowser.open("https://www.geeksforgeeks.org")
        elif 'open problem' in query:
            webbrowser.open("https://www.javatpoint.com")
        elif 'khairiyat' in query:
            webbrowser.open("https://www.youtube.com/watch?v=hoNb6HuNmU0")
        elif "map" in query:
            query = query.replace("map", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.open(
                "https://www.google.ca/maps/place/" + location + "")

        elif 'play music' in query:
            music_dir = "D:\\Song"
            songs = os.listdir(music_dir)
            print("Music Starting")
            os.startfile(os.path.join(music_dir, songs[1]))

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            print(f"Sir, the time is {strTime}")
            speak(f"Sir, the time is {strTime}")

        elif 'open code' in query:
            codePath = "C:\\Users\\Artificial\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)
        elif 'open word' in query:
            codePath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word"
            os.startfile(codePath)
        elif 'open excel' in query:
            codePath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel"
            os.startfile(codePath)
        elif 'open firefox' in query:
            codePath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Firefox"
            os.startfile(codePath)
        elif 'open setting' in query:
            codePath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Settings"
            os.startfile(codePath)
        elif 'open outlook' in query:
            codePath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Outlook"
            os.startfile(codePath)
        elif 'android studio' in query:
            codePath = "C:\\Program Files\\Android\\Android Studio\\bin\\studio64.exe"
            os.startfile(codePath)

        elif 'email to rohit' in query:
            try:
                speak("What should I say?")
                # speak(input("Enter your email\n"))
                content = takeCommand()
                to = "youremail@gmail.com"

                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("Sorry sir . I am not able to send this email")
        elif 'exit' in query:
            print("Exit the Program")
            speak("Exit the Program")
            quit()
