import speech_recognition as sr
import os
import win32com.client
import webbrowser
import openai
from config import apikey
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
        

def ai(prompt):
    openai.api_key = apikey
    text = f"OpenAI response for Prompt: {prompt} \n *************************\n\n"

    response = openai.Completion.create(
            model="text-davinci-003",
            prompt=prompt,
            temperature=1,
            max_tokens=256,
            top_p=1,
            frequency_penalty=0,
            presence_penalty=0
        )
    
    # print(response["choise"][0]["text"])
    text += response["choices"][0]["text"]
    if not os.path.exists("D:/Python_for_beg/AI_Desktop_Assistant/Openai"):
        os.mkdir("D:/Python_for_beg/AI_Desktop_Assistant/Openai")
    
    with open(f"Openai/{ ''.join(prompt.split('artificial')[1:]).strip() }.txt", "w") as f:
        f.write(text)


chatStr = " "

def chat(query):
    global chatStr
    openai.api_key = apikey

    chatStr += f"User: {query}\n Jarvis: "
    print(chatStr)

    response = openai.Completion.create(
            model="text-davinci-003",
            prompt=chatStr,
            temperature=1,
            max_tokens=256,
            top_p=1,
            frequency_penalty=0,
            presence_penalty=0
        )
    say(response["choices"][0]["text"])
    chatStr += f"{response['choices'][0]['text']} \n"
    return response["choices"][0]['text']



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
        elif "the time" in query:
            strfTime = datetime.datetime.now().strftime("%H:%M:%S")
            say(f"Sir time is {strfTime}")

        elif "open TeamViewer".lower() in query.lower():
            say("Ok I can open")
            os.system("C:/Users/Desai/Downloads/TeamViewer.lnk")

        elif "Using artificial".lower() in query.lower():
            ai(prompt=query)
        
        elif "Jarvis Quit".lower() in query.lower():
            exit()

        elif "Jarvis Reset".lower() in query.lower():
            chatStr = " "

        else:
            print("Chatting ......")
            chat(query)

        # say(query)

