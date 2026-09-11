import os
import datetime
import webbrowser
import subprocess
import time
import urllib.parse

import speech_recognition as sr
import pyautogui
import win32com.client
import win32gui
import win32process
import requests

from google import genai
from dotenv import load_dotenv


# ============================================================
# ENVIRONMENT & GEMINI
# ============================================================

load_dotenv()

API_KEY = os.getenv("OPENAI_API_KEY")

if not API_KEY:
    print("ERROR: OPENAI_API_KEY is missing from .env")
    print("Please add your API key to the .env file.")
    exit()

client = genai.Client(api_key=API_KEY)


# ============================================================
# TEXT TO SPEECH
# ============================================================

speaker = win32com.client.Dispatch("SAPI.SpVoice")

chat_history = []
previous_interaction_id = None


def say(text):
    print(f"Jarvis: {text}")
    speaker.Speak(text)


# ============================================================
# INTERNET CHECK
# ============================================================

def internet_available():

    try:

        requests.get(
            "https://www.google.com",
            timeout=3
        )

        return True

    except requests.RequestException:

        return False


# ============================================================
# GEMINI CHAT
# ============================================================

def chat(query):

    global chat_history, previous_interaction_id

    if not internet_available():
        say("Internet connection is not available.")
        return

    try:
        # Gemini Interactions API keeps the conversation state using
        # previous_interaction_id, so we do not need to resend the
        # entire chat history on every request.
        interaction_args = {
            "model": "gemini-3.6-flash",
            "input": query,
            "system_instruction": (
                "You are Jarvis, a helpful AI assistant. "
                "Give concise, useful answers. "
                "Address the user politely as sir when appropriate."
            )
        }

        if previous_interaction_id:
            interaction_args["previous_interaction_id"] = previous_interaction_id

        response = client.interactions.create(**interaction_args)

        reply = response.output_text

        if not reply:
            say("Sorry sir, Gemini returned an empty response.")
            return

        previous_interaction_id = response.id

        chat_history.append({
            "role": "user",
            "content": query
        })
        chat_history.append({
            "role": "assistant",
            "content": reply
        })

        say(reply)

    except Exception as e:
        print("Gemini ERROR:", e)

        error_text = str(e)

        if "NOT_FOUND" in error_text or "no longer available" in error_text:
            say("The Gemini model is unavailable. Please check the configured model.")
        elif "401" in error_text or "API key" in error_text or "authentication" in error_text.lower():
            say("Sorry sir, the Gemini API key is invalid or not authorized.")
        elif "429" in error_text or "quota" in error_text.lower():
            say("Sorry sir, the Gemini API quota has been exceeded.")
        else:
            say("Sorry sir, I am unable to contact Gemini.")


# ============================================================
# VOICE COMMAND
# ============================================================

def takeCommand():

    r = sr.Recognizer()

    with sr.Microphone() as source:

        print("Listening...")

        r.pause_threshold = 0.8

        try:

            r.adjust_for_ambient_noise(
                source,
                duration=1
            )

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

        except sr.RequestError as e:

            print("Speech Recognition Error:", e)

            say(
                "Speech recognition service is unavailable."
            )

            return ""

        except Exception as e:

            print("Microphone ERROR:", e)

            return ""


# ============================================================
# CLOSE APPLICATION
# ============================================================

def close_app(process):

    try:

        result = subprocess.run(

            ["taskkill", "/f", "/im", process],

            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )

        return result.returncode == 0

    except Exception as e:

        print("Close application error:", e)

        return False


# ============================================================
# GET BROWSER WINDOWS
# ============================================================

def get_browser_windows(browser):

    """
    Find visible windows belonging to a specific browser.

    Supported:
        chrome
        brave
        edge
        firefox
    """

    browser_processes = {

        "chrome": "chrome.exe",
        "brave": "brave.exe",
        "edge": "msedge.exe",
        "firefox": "firefox.exe"
    }

    process_name = browser_processes.get(
        browser.lower()
    )

    if not process_name:
        return []

    windows = []

    def enum_windows_callback(hwnd, extra):

        if not win32gui.IsWindowVisible(hwnd):
            return

        title = win32gui.GetWindowText(hwnd)

        if not title:
            return

        try:

            _, pid = win32process.GetWindowThreadProcessId(hwnd)

            result = subprocess.check_output(
                [
                    "tasklist",
                    "/FI",
                    f"PID eq {pid}",
                    "/FO",
                    "CSV",
                    "/NH"
                ],
                text=True,
                stderr=subprocess.DEVNULL
            )

            if process_name.lower() in result.lower():

                windows.append(
                    (hwnd, title)
                )

        except Exception:
            pass

    win32gui.EnumWindows(
        enum_windows_callback,
        None
    )

    return windows


# ============================================================
# ACTIVATE BROWSER WINDOW
# ============================================================

def activate_window(hwnd):

    try:

        # Restore if minimized
        win32gui.ShowWindow(
            hwnd,
            9
        )

        time.sleep(0.2)

        win32gui.SetForegroundWindow(
            hwnd
        )

        time.sleep(0.4)

        return True

    except Exception as e:

        print("Window activation error:", e)

        return False


# ============================================================
# FIND WINDOW BY PAGE TITLE
# ============================================================

def find_browser_window_by_title(
    page_name,
    browser=None
):

    """
    Find a browser window whose title contains page_name.

    Example:
        YouTube - Google Chrome

    page_name:
        youtube
        github
        gmail
        google
    """

    browsers = (
        [browser]
        if browser
        else ["chrome", "brave", "edge", "firefox"]
    )

    page_name = page_name.lower()

    for current_browser in browsers:

        windows = get_browser_windows(
            current_browser
        )

        for hwnd, title in windows:

            if page_name in title.lower():

                return hwnd, title, current_browser

    return None, None, None


# ============================================================
# CLOSE SPECIFIC WEBSITE TAB
# ============================================================

def close_website_tab(website_name):

    hwnd, title, browser = find_browser_window_by_title(
        website_name
    )

    if hwnd:

        print(
            f"Found {website_name}: "
            f"{title} ({browser})"
        )

        if activate_window(hwnd):

            pyautogui.hotkey(
                "ctrl",
                "w"
            )

            time.sleep(0.3)

            say(
                f"Closing {website_name}"
            )

            return True

    say(
        f"I could not find an open {website_name} window."
    )

    return False


# ============================================================
# CLOSE TAB IN SPECIFIC BROWSER
# ============================================================

def close_browser_tab(browser_name):

    windows = get_browser_windows(
        browser_name
    )

    if not windows:

        say(
            f"{browser_name} is not currently open."
        )

        return False

    # Select the first visible browser window
    hwnd, title = windows[0]

    print(
        f"Target browser window: {title}"
    )

    if activate_window(hwnd):

        pyautogui.hotkey(
            "ctrl",
            "w"
        )

        time.sleep(0.3)

        say(
            f"Closing {browser_name} tab"
        )

        return True

    return False


# ============================================================
# CLOSE CURRENT TAB
# ============================================================

def close_current_tab():

    time.sleep(0.3)

    pyautogui.hotkey(
        "ctrl",
        "w"
    )


# ============================================================
# GOOGLE SEARCH
# ============================================================

def google_search(query):

    search_query = (
        query
        .replace("google search", "")
        .replace("search", "")
        .replace("on google", "")
        .strip()
    )

    if search_query:

        say(
            f"Searching {search_query}"
        )

        encoded = urllib.parse.quote(
            search_query
        )

        webbrowser.open(
            f"https://www.google.com/search?q={encoded}"
        )


# ============================================================
# OPEN APPLICATION
# ============================================================

def open_application(
    path,
    app_name
):

    try:

        if not os.path.exists(path):

            say(
                f"{app_name} is not installed at the expected location."
            )

            return

        os.startfile(path)

        say(
            f"Opening {app_name}"
        )

    except Exception as e:

        print(
            f"{app_name} ERROR:",
            e
        )

        say(
            f"Unable to open {app_name}"
        )


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":

    say(
        "Jarvis Activated"
    )


    # ========================================================
    # WEBSITES
    # ========================================================

    websites = {

        "youtube":
            "https://youtube.com",

        "google":
            "https://google.com",

        "github":
            "https://github.com",

        "linkedin":
            "https://linkedin.com",

        "facebook":
            "https://facebook.com",

        "instagram":
            "https://instagram.com",

        "spotify":
            "https://spotify.com",

        "gmail":
            "https://mail.google.com",

        "wikipedia":
            "https://wikipedia.org"
    }


    # ========================================================
    # APPLICATION PATHS
    # ========================================================

    applications = {

        "chrome": (
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            "Chrome"
        ),

        "brave": (
            r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe",
            "Brave"
        ),

        "word": (
            r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
            "Microsoft Word"
        ),

        "excel": (
            r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
            "Excel"
        ),

        "powerpoint": (
            r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
            "PowerPoint"
        ),

        "vs code": (
            r"C:\Users\Ankur soni\AppData\Local\Programs\Microsoft VS Code\Code.exe",
            "Visual Studio Code"
        )
    }


    # ========================================================
    # CLOSE APPLICATION COMMANDS
    # ========================================================

    close_commands = {

        "chrome":
            "chrome.exe",

        "brave":
            "brave.exe",

        "edge":
            "msedge.exe",

        "firefox":
            "firefox.exe",

        "word":
            "WINWORD.EXE",

        "excel":
            "EXCEL.EXE",

        "powerpoint":
            "POWERPNT.EXE",

        "ppt":
            "POWERPNT.EXE",

        "vs code":
            "Code.exe",

        "visual studio code":
            "Code.exe"
    }


    # ========================================================
    # MAIN LOOP
    # ========================================================

    while True:

        try:

            query = takeCommand()

            if not query:
                continue

            handled = False


            # =================================================
            # EXIT JARVIS
            # =================================================

            if (
                "jarvis quit" in query
                or "exit jarvis" in query
                or "goodbye" in query
                or "shutdown jarvis" in query
            ):

                say(
                    "Goodbye sir"
                )

                break

            handled = True


            # =================================================
            # OPEN WEBSITE
            # =================================================

            if handled:

                matched_open_website = False

                for name, url in websites.items():

                    if (
                        f"open {name}" in query
                        or f"launch {name}" in query
                        or f"start {name}" in query
                    ):

                        say(
                            f"Opening {name}"
                        )

                        webbrowser.open(url)

                        matched_open_website = True

                        break

                if matched_open_website:
                    continue


            # =================================================
            # CLOSE SPECIFIC WEBSITE
            # =================================================

            matched_close_website = False

            for name in websites:

                if (
                    f"close {name}" in query
                    or f"exit {name}" in query
                ):

                    close_website_tab(
                        name
                    )

                    matched_close_website = True

                    break

            if matched_close_website:
                continue


            # =================================================
            # GOOGLE SEARCH
            # =================================================

            if (
                "google search" in query
                or query.startswith("search ")
            ):

                google_search(query)

                continue


            # =================================================
            # CLOSE SPECIFIC BROWSER TAB
            # =================================================

            if "close chrome tab" in query:

                close_browser_tab(
                    "chrome"
                )

                continue


            if "close brave tab" in query:

                close_browser_tab(
                    "brave"
                )

                continue


            if "close edge tab" in query:

                close_browser_tab(
                    "edge"
                )

                continue


            if "close firefox tab" in query:

                close_browser_tab(
                    "firefox"
                )

                continue


            # =================================================
            # CLOSE CURRENT TAB
            # =================================================

            if (
                "close tab" in query
                or "close current tab" in query
            ):

                close_current_tab()

                say(
                    "Closing tab"
                )

                continue


            # =================================================
            # OPEN APPLICATION
            # =================================================

            opened_application = False

            for name, (path, app_name) in applications.items():

                if (
                    f"open {name}" in query
                    or f"launch {name}" in query
                    or f"start {name}" in query
                ):

                    open_application(
                        path,
                        app_name
                    )

                    opened_application = True

                    break

            if opened_application:
                continue


            # =================================================
            # CLOSE APPLICATION
            # =================================================

            closed_application = False

            for name, process in close_commands.items():

                if (
                    f"close {name}" in query
                    or f"exit {name}" in query
                ):

                    if close_app(process):

                        say(
                            f"Closing {name}"
                        )

                    else:

                        say(
                            f"{name} is not currently running."
                        )

                    closed_application = True

                    break

            if closed_application:
                continue


            # =================================================
            # TIME
            # =================================================

            if (
                "what time" in query
                or "tell me the time" in query
                or query == "time"
                or "current time" in query
            ):

                now = datetime.datetime.now()

                say(
                    f"It is {now.strftime('%I:%M %p')}"
                )

                continue


            # =================================================
            # RESET CHAT
            # =================================================

            if (
                "reset chat" in query
                or "clear chat" in query
                or "forget conversation" in query
            ):

                chat_history.clear()
                previous_interaction_id = None

                say(
                    "Chat memory reset"
                )

                continue


            # =================================================
            # GENERAL AI QUESTION
            # =================================================

            # Only here do we call Gemini.

            chat(query)


        except KeyboardInterrupt:

            print(
                "\nJarvis stopped."
            )

            break

        except Exception as e:

            print(
                "ERROR:",
                e
            )

            say(
                "An unexpected error occurred."
            )
