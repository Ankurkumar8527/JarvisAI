import os
import datetime
import webbrowser
import subprocess
import time
import urllib.parse

import speech_recognition as sr
import pyautogui
import win32com.client
import requests

from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()

API_KEY = os.getenv("OPENAI_API_KEY")

client = OpenAI(api_key=API_KEY)

speaker = win32com.client.Dispatch("SAPI.SpVoice")

chat_history = []

def say(text):
    print(f"Jarvis: {text}")
    speaker.Speak(text)

def internet_available():
    try:
        requests.get("https://www.google.com", timeout=3)
        return True
    except:
        return False

def chat(query):
    global chat_history

    if not internet_available():
        say("Internet connection is not available.")
        return

    try:
        chat_history.append({
            "role": "user",
            "content": query
        })

        response = client.chat.completions.create(
            model="gpt-4.1-mini",
            messages=[
                {
                    "role": "system",
                    "content": "You are Jarvis, a helpful AI assistant."
                },
                *chat_history
            ]
        )

        reply = response.choices[0].message.content

        chat_history.append({
            "role": "assistant",
            "content": reply
        })

        say(reply)

    except Exception as e:
        print(e)
        say("Sorry sir, I am unable to contact OpenAI.")

def takeCommand():
    r = sr.Recognizer()

    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 0.8

        try:
            r.adjust_for_ambient_noise(source, duration=1)

            audio = r.listen(
                source,
                timeout=5,
                phrase_time_limit=10
            )

            print("Recognizing...")

            query = r.recognize_google(
                audio,
                language="en-in"
            )

            print(f"User: {query}")

            return query.lower()

        except sr.WaitTimeoutError:
            return ""

        except sr.UnknownValueError:
            return ""

        except Exception as e:
            print(e)
            return ""

def close_app(process):
    try:
        subprocess.run(
            ["taskkill", "/f", "/im", process],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
    except:
        pass

def close_current_tab():
    time.sleep(0.5)
    pyautogui.hotkey("ctrl", "w")

def google_search(query):
    search_query = (
        query.replace("google search", "")
        .replace("search", "")
        .replace("on google", "")
        .strip()
    )

    if search_query:
        say(f"Searching {search_query}")

        encoded = urllib.parse.quote(search_query)

        webbrowser.open(
            f"https://www.google.com/search?q={encoded}"
        )

def open_application(path, app_name):
    try:
        os.startfile(path)
        say(f"Opening {app_name}")
    except:
        say(f"{app_name} not found")

if __name__ == "__main__":

    say("Jarvis Activated")

    websites = {
        "youtube": "https://youtube.com",
        "google": "https://google.com",
        "github": "https://github.com",
        "linkedin": "https://linkedin.com",
        "facebook": "https://facebook.com",
        "instagram": "https://instagram.com",
        "spotify": "https://spotify.com",
        "gmail": "https://mail.google.com",
        "wikipedia": "https://wikipedia.org"
    }

    while True:

        try:

            query = takeCommand()

            if not query:
                continue

            handled = False

            for name, url in websites.items():

                if f"open {name}" in query:
                    say(f"Opening {name}")
                    webbrowser.open(url)
                    handled = True
                    break

            if not handled and (
                "google search" in query or
                "search" in query
            ):
                google_search(query)
                handled = True

            elif "close tab" in query:
                close_current_tab()
                say("Closing tab")
                handled = True

            elif "open chrome" in query:
                open_application(
                    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                    "Chrome"
                )
                handled = True

            elif "close chrome" in query:
                close_app("chrome.exe")
                say("Closing Chrome")
                handled = True

            elif "open brave" in query:
                open_application(
                    r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe",
                    "Brave"
                )
                handled = True

            elif "close brave" in query:
                close_app("brave.exe")
                say("Closing Brave")
                handled = True

            elif "open word" in query:
                open_application(
                    r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
                    "Microsoft Word"
                )
                handled = True

            elif "close word" in query:
                close_app("WINWORD.EXE")
                say("Closing Word")
                handled = True

            elif "open excel" in query:
                open_application(
                    r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
                    "Excel"
                )
                handled = True

            elif "close excel" in query:
                close_app("EXCEL.EXE")
                say("Closing Excel")
                handled = True

            elif "open powerpoint" in query or "open ppt" in query:
                open_application(
                    r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
                    "PowerPoint"
                )
                handled = True

            elif "close powerpoint" in query or "close ppt" in query:
                close_app("POWERPNT.EXE")
                say("Closing PowerPoint")
                handled = True

            elif "open vs code" in query:
                open_application(
                    r"C:\Users\Ankur soni\AppData\Local\Programs\Microsoft VS Code\Code.exe",
                    "Visual Studio Code"
                )
                handled = True

            elif "close vs code" in query:
                close_app("Code.exe")
                say("Closing VS Code")
                handled = True

            elif "time" in query:
                now = datetime.datetime.now()
                say(
                    f"It is {now.hour} hours and {now.minute} minutes"
                )
                handled = True

            elif "reset chat" in query:
                chat_history.clear()
                say("Chat memory reset")
                handled = True

            elif (
                "jarvis quit" in query or
                "exit jarvis" in query or
                "goodbye" in query
            ):
                say("Goodbye sir")
                break

            if not handled:
                chat(query)

        except Exception as e:
            print("ERROR:", e)
            say("An unexpected error occurred.")