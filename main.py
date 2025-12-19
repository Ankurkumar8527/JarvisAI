import datetime
import os
import webbrowser
import speech_recognition as sr
import win32com.client
import subprocess
import time
import pyautogui
from openai import OpenAI
from config import api_key

speaker = win32com.client.Dispatch("SAPI.SpVoice")
client = OpenAI(api_key=api_key)
chat_history = []

def say(text):
    print(text)
    speaker.Speak(text)

def chat(query):
    global chat_history
    chat_history.append({"role": "user", "content": query})

    response = client.responses.create(
        model="gpt-4.1-mini",
        input=chat_history
    )

    text = response.output_text
    chat_history.append({"role": "assistant", "content": text})
    say(text)

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 0.8
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query.lower()
        except:
            return ""

def close_app(process):
    subprocess.run(
        ["taskkill", "/f", "/im", process],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )

def close_current_tab():
    time.sleep(0.4)
    pyautogui.hotkey("ctrl", "w")

def google_search(query):
    search_query = (
        query.replace("google search", "")
             .replace("search", "")
             .replace("on google", "")
             .strip()
    )
    if search_query:
        say(f"Searching {search_query} on Google")
        url = f"https://www.google.com/search?q={search_query.replace(' ', '+')}"
        webbrowser.open(url)

if __name__ == "__main__":
    say("Jarvis Activated")

    sites = [
        ["youtube", "https://www.youtube.com"],
        ["google", "https://www.google.com"],
        ["wikipedia", "https://www.wikipedia.org"],
        ["facebook", "https://www.facebook.com"],
        ["instagram", "https://www.instagram.com"],
        ["spotify", "https://open.spotify.com"],
        ["github", "https://github.com"],
        ["linkedin", "https://www.linkedin.com"],
        ["gmail", "https://mail.google.com"]
    ]

    while True:
        handled = False
        query = takeCommand()

        if not query:
            continue

        for site in sites:
            if f"open {site[0]}" in query:
                say(f"Opening {site[0]} sir")
                webbrowser.open(site[1])
                handled = True
                break

        if not handled and ("search" in query or "google search" in query):
            google_search(query)
            handled = True

        elif "close youtube" in query or "close google" in query or "close website" in query or "close tab" in query:
            say("Closing current tab")
            close_current_tab()
            handled = True

        elif "open chrome" in query:
            os.startfile(r"C:\Program Files\Google\Chrome\Application\chrome.exe")
            say("Opening Chrome")
            handled = True

        elif "close chrome" in query:
            close_app("chrome.exe")
            say("Closing Chrome")
            handled = True

        elif "open brave" in query:
            os.startfile(r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe")
            say("Opening Brave")
            handled = True

        elif "close brave" in query:
            close_app("brave.exe")
            say("Closing Brave")
            handled = True

        elif "open word" in query:
            os.startfile(r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE")
            say("Opening Microsoft Word")
            handled = True

        elif "close word" in query:
            close_app("WINWORD.EXE")
            say("Closing Microsoft Word")
            handled = True

        elif "open excel" in query:
            os.startfile(r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE")
            say("Opening Microsoft Excel")
            handled = True

        elif "close excel" in query:
            close_app("EXCEL.EXE")
            say("Closing Microsoft Excel")
            handled = True

        elif "open powerpoint" in query or "open ppt" in query:
            os.startfile(r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE")
            say("Opening PowerPoint")
            handled = True

        elif "close powerpoint" in query or "close ppt" in query:
            close_app("POWERPNT.EXE")
            say("Closing PowerPoint")
            handled = True

        elif "open project ppt" in query or "open third year project" in query:
            os.startfile(r"D:\3rd year Project.pptx")
            say("Opening third year project")
            handled = True

        elif "close project ppt" in query or "close third year project" in query:
            close_app("POWERPNT.EXE")
            say("Closing project presentation")
            handled = True

        elif "open vs code" in query:
            os.startfile(r"C:\Users\Ankur soni\Downloads\VSCodeUserSetup-x64-1.105.1.exe")
            say("Opening Visual Studio Code")
            handled = True

        elif "close vs code" in query:
            close_app("Code.exe")
            say("Closing Visual Studio Code")
            handled = True

        elif "the time" in query:
            now = datetime.datetime.now()
            say(f"Sir time is {now.hour} bajke {now.minute} minutes")
            handled = True

        elif "reset chat" in query:
            chat_history = []
            say("Chat memory reset")
            handled = True

        elif "jarvis quit" in query or "exit jarvis" in query:
            say("Goodbye sir")
            break

        if not handled:
            chat(query)
