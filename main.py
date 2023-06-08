import speech_recognition as sr
import os
import win32com.client
import webbrowser
import openai
import datetime

# for speech
speaker = win32com.client.Dispatch("SAPI.Spvoice")


def say(text):
    speaker.Speak(text)


def takeCommand():
    r = sr.Recognizer()
    
    with sr.Microphone() as source:
        # r.pause_threshold  = 1
        r.adjust_for_ambient_noise(source)
        audio = r.listen(source)

        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Some error Occurred. Sorry from Jarvis"


if __name__ == '__main__' :
    print("Hello ...")
    say("Hello I am Jarvish AI")
    
    while True:
        print("Listening....")
        query = takeCommand()

        # for open site
        sites = [ ["youtube", "https://www.youtube.com"], ["google", "https://www.google.com"], ["wikipedia", "https://www.wikipedia.com"] ]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                say(f"Opening {site[0]} sir...")
                webbrowser.open(site[1])

        # for open music
        if "open music" in query:
            musicPath = "C:/Users/Desai/Downloads/Jai_Shri_Ram.mp3"
            say("Ok I play music")
            os.startfile(musicPath)

        # for show time
        if "the time" in query:
            strfTime = datetime.datetime.now().strftime("%H:%M:%S")
            say(f"Sir time is {strfTime}")

        if "open TeamViewer".lower() in query.lower():
            say("Ok I can open")
            os.system("C:/Users/Desai/Downloads/TeamViewer.lnk")

        say(query)

