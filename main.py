import speech_recognition as sr
import os
import webbrowser
import openai
import datetime
import random
import pyttsx3
import win32com.client as wincl
from config import apikey
import wikipedia


engine = wincl.Dispatch("SAPI.SpVoice")

chatStr = ""

def wishMe():
    engine.Speak("Hey There!!!")
    hour=int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        engine.Speak("Good Morning!!!")
    elif hour>=12 and hour<=16:
        engine.Speak("Good Afternoon!!!")
    else:
        engine.Speak("Good Evening!!!") 

    engine.Speak("I am an Intelligent Voice, always ready for your Service!!!. What can I do for you?")

def chat(query):
    global chatStr
    print(chatStr)
    openai.api_key = apikey
    chatStr += f"Prakhar: {query}\n Pro: "
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=chatStr,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    engine.Speak(response["choices"][0]["text"])
    chatStr += f"{response['choices'][0]['text']}\n"
    return response["choices"][0]["text"]


def ai(prompt):
    openai.api_key = apikey
    text = f"OpenAI response for Prompt: {prompt} \n *************************\n\n"

    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=prompt,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    text += response["choices"][0]["text"]
    if not os.path.exists("Openai"):
        os.mkdir("Openai")

    with open(f"Openai/{''.join(prompt.split('intelligence')[1:]).strip()}.txt", "w") as f:
        f.write(text)


def say(text):
    engine.Speak(text)


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Some Error Occurred. Sorry from Pro"


if __name__ == '__main__':
    wishMe()
    
    while True:
        print("Listening...")
        query = takeCommand()
        sites = [["youtube", "https://www.youtube.com"], ["wikipedia", "https://www.wikipedia.com"], ["google", "https://www.google.com"],
                 ["github",""],["creater","https://prakhargupta-portfolio.netlify.app/"]]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                engine.Speak(f"Opening {site[0]} sir...")
                webbrowser.open(site[1])

        if "wiki" in query :
            query=query.replace("wiki","")
            engine.Speak("searching for {} in wikipedia...".format(query))
            results=wikipedia.summary(query,sentences=5)
            engine.Speak("According to results from wikipedia")
            engine.Speak(results)
            engine.Speak("This is about {}".format(query))

        elif "wikipedia" in query:
            query=query.replace("wikipedia","")
            engine.Speak("searching for {} in wikipedia...".format(query))
            results=wikipedia.summary(query,sentences=5)
            engine.Speak("According to results from wikipedia")
            print(results)
            engine.Speak(results)
            engine.Speak("This is about {}".format(query))         
        elif 'tell'in query and 'me'in query and 'about'in query:      
            query=query.replace("tell","")
            query=query.replace("me","")
            query=query.replace("about","")
            engine.Speak("searching for {} in wikipedia...".format(query))
            results=wikipedia.summary(query,sentences=5)
            engine.Speak("According to results from wikipedia")
            print(results)
            engine.Speak(results)
            engine.Speak("This is about {}".format(query))
        
        elif "the time" in query:
            hour = datetime.datetime.now().strftime("%H")
            min = datetime.datetime.now().strftime("%M")
            engine.Speak(f"Sir, the time is {hour} o'clock and {min} minutes")

        elif "open project files".lower() in query.lower():
            os.startfile(r"D:\projects\speech")

        

        elif "Using artificial intelligence".lower() in query.lower():
            ai(prompt=query)

        elif "Quit".lower() in query.lower():
            exit()

        elif "reset chat".lower() in query.lower():
            chatStr = ""

        else:
            print("Chatting...")
            chat(query)
