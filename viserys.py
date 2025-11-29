try:
    from win32com.client import Dispatch
except Exception:
    Dispatch = None
from datetime import datetime
import speech_recognition as sr
import wikipedia
import webbrowser as wb
import os
import random
from PIL import ImageGrab
import psutil
import platform
import numpy as np
from typing import Optional

JOKES = [
    "I told my computer I needed a break, and it said 'No problem — I'll go to sleep.'",
    "Why do programmers prefer dark mode? Because light attracts bugs.",
    "How many programmers does it take to change a light bulb? None — it's a hardware problem.",
    "Why did the programmer quit his job? Because he didn't get arrays.",
    "A SQL query walks into a bar, walks up to two tables and asks, 'Can I join you?'",
    "There are 10 types of people in the world: those who understand binary and those who don't.",
    "I've got a really good UDP joke to tell you, but I don't know if you'll get it.",
    "Debugging: Being the detective in a crime movie where you are also the murderer.",
]

tts = None
pyttsx3_engine = None
if Dispatch:
    try:
        tts = Dispatch("SAPI.SpVoice")
    except Exception:
        tts = None

if not tts:
    try:
        import pyttsx3
        pyttsx3_engine = pyttsx3.init()
        voices = pyttsx3_engine.getProperty('voices')
        if len(voices) > 1:
            pyttsx3_engine.setProperty('voice', voices[1].id)
        pyttsx3_engine.setProperty('rate', 150)
        pyttsx3_engine.setProperty('volume', 1)
    except Exception:
        pyttsx3_engine = None

VOICE_INPUT_AVAILABLE = True
try:
    import pyaudio
except Exception:
    VOICE_INPUT_AVAILABLE = False

SD_AVAILABLE = False
try:
    import sounddevice as sd
    SD_AVAILABLE = True
except Exception:
    SD_AVAILABLE = False

INPUT_GAIN = 1.8
DEFAULT_RECORD_SECONDS = 4
SELECTED_DEVICE_INDEX = None


def choose_input_device() -> Optional[int]:
    """Interactively choose an input device. Returns device index or None."""
    devices = list_input_devices()
    if not devices:
        speak("No audio input devices found.")
        return None

    speak("I found the following input devices:")
    for idx, name in devices:
        speak(f"Device {idx}: {name}")
        print(f"Device {idx}: {name}")

    choice = takecommand("Please say or type the device number to use, or say default to use the first one.")
    if not choice:
        return devices[0][0]

    digits = ''.join(filter(str.isdigit, choice))
    if digits:
        try:
            idx = int(digits)
            for d, _ in devices:
                if d == idx:
                    return idx
        except Exception:
            pass

    return devices[0][0]


def list_input_devices() -> list:
    """Return a list of available input devices (name and index)."""
    devices = []
    if not SD_AVAILABLE:
        return devices
    try:
        all_dev = sd.query_devices()
        for i, dev in enumerate(all_dev):
            if dev.get('max_input_channels', 0) > 0:
                devices.append((i, dev.get('name', 'Unknown')))
    except Exception:
        pass
    return devices


def speak(audio) -> None:
    """Speak the given text using SAPI, pyttsx3, or fallback to printing.

    Guarantees an audio attempt at each step; falls back to printing when TTS is unavailable.
    """
    try:
        if tts:
            tts.Speak(str(audio))
        elif pyttsx3_engine:
            pyttsx3_engine.say(str(audio))
            pyttsx3_engine.runAndWait()
        else:
            if platform.system() == "Windows":
                try:
                    escaped = str(audio).replace('"', '\\"')
                    cmd = f"Add-Type –AssemblyName System.Speech; (New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak(\"{escaped}\")"
                    os.system(f"powershell -Command \"{cmd}\"")
                except Exception:
                    print(str(audio))
            else:
                print(str(audio))
    except Exception:
        try:
            print(str(audio))
        except Exception:
            pass


def time() -> None:
    """Tells the current time."""
    current_time = datetime.now().strftime("%I:%M:%S %p")
    speak("The current time is")
    speak(current_time)
    print("The current time is", current_time)


def date() -> None:
    """Tells the current date."""
    now = datetime.now()
    speak("The current date is")
    speak(f"{now.day} {now.strftime('%B')} {now.year}")
    print(f"The current date is {now.day}/{now.month}/{now.year}")


def wishme() -> None:
    """Greets the user based on the time of day."""
    speak("Welcome back")
    print("Welcome back")

    hour = datetime.now().hour
    if 4 <= hour < 12:
        speak("Good morning!")
        print("Good morning!")
    elif 12 <= hour < 16:
        speak("Good afternoon!")
        print("Good afternoon!")
    elif 16 <= hour < 24:
        speak("Good evening!")
        print("Good evening!")
    else:
        speak("Good night, see you tomorrow.")

    assistant_name = load_name()
    speak(f"{assistant_name} at your service. Please tell me how may I assist you.")
    print(f"{assistant_name} at your service. Please tell me how may I assist you.")


def screenshot() -> None:
    """Takes a screenshot and saves it."""
    try:
        img = ImageGrab.grab()
        script_dir = os.path.dirname(os.path.abspath(__file__))
        screenshots_dir = os.path.join(script_dir, "screenshots")
        os.makedirs(screenshots_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        img_path = os.path.join(screenshots_dir, f"screenshot_{timestamp}.png")
        img.save(img_path)
        speak(f"Screenshot saved as {img_path}.")
        print(f"Screenshot saved as {img_path}.")
    except Exception as e:
        speak("Failed to take screenshot.")
        print(f"Screenshot error: {e}")

def takecommand(prompt: Optional[str] = None, timeout: int = 6) -> Optional[str]:
    """Takes microphone input from the user and returns it as lowercase text.

    If `prompt` is provided, it will be spoken before listening.
    """
    if prompt:
        speak(prompt)
    if VOICE_INPUT_AVAILABLE:
        r = sr.Recognizer()
        try:
            r.dynamic_energy_threshold = True
            r.energy_threshold = 300
        except Exception:
            pass
        try:
            with sr.Microphone() as source:
                try:
                    r.adjust_for_ambient_noise(source, duration=0.5)
                except Exception:
                    pass
                speak("Listening now.")
                r.pause_threshold = 1

                try:
                    audio = r.listen(source, timeout=timeout)
                except sr.WaitTimeoutError:
                    speak("Timeout occurred. Please try again.")
                    return None

        except Exception as e:
            print(f"Microphone error: {e}")
            audio = None

        if audio:
            try:
                speak("Recognizing.")
                query = r.recognize_google(audio, language="en-in")
                print(f"Heard: {query}")
                return query.lower()
            except sr.UnknownValueError:
                speak("Sorry, I did not understand that.")
                return None
            except sr.RequestError:
                speak("Speech recognition service is unavailable.")
                return None
            except Exception as e:
                speak(f"An error occurred while recognizing speech: {e}")
                print(f"Error: {e}")

    if SD_AVAILABLE:
        try:
            devices = list_input_devices()
            if not devices:
                speak("No input audio devices detected. Please connect a microphone and ensure it's enabled in Windows settings.")
            else:
                dev_index, dev_name = devices[0]
                try:
                    dev_info = sd.query_devices(dev_index)
                    samplerate = int(dev_info.get('default_samplerate', 16000))
                except Exception:
                    samplerate = 16000

                duration = max(DEFAULT_RECORD_SECONDS, timeout)
                speak(f"Recording from {dev_name} for {duration} seconds.")
                try:
                    recording = sd.rec(int(duration * samplerate), samplerate=samplerate, channels=1, dtype='int16', device=dev_index)
                    sd.wait()
                    try:
                        arr = recording.astype('float32')
                        arr *= INPUT_GAIN
                        arr = np.clip(arr, -32768, 32767)
                        int16_arr = arr.astype('int16')
                        data = int16_arr.tobytes()
                    except Exception:
                        data = recording.tobytes()

                    audio_data = sr.AudioData(data, samplerate, 2)
                    speak("Recognizing.")
                    query = sr.Recognizer().recognize_google(audio_data, language="en-in")
                    print(f"Heard: {query}")
                    return query.lower()
                except Exception as e:
                    print(f"sounddevice capture failed: {e}")
                    speak("I couldn't capture audio with sounddevice. Please check microphone permissions, select the correct input device, or try restarting the program.")
        except Exception as e:
            print(f"sounddevice recording failed: {e}")

    speak("Voice input is not available. Please type your response.")
    try:
        typed = input((prompt + "\n> ") if prompt else "> ")
        return typed.lower().strip() if typed else None
    except Exception:
        return None

    r = sr.Recognizer()
    with sr.Microphone() as source:
        try:
            r.adjust_for_ambient_noise(source, duration=0.5)
        except Exception:
            pass
        speak("Listening now.")
        r.pause_threshold = 1

        try:
            audio = r.listen(source, timeout=timeout)
        except sr.WaitTimeoutError:
            speak("Timeout occurred. Please try again.")
            return None

    try:
        speak("Recognizing.")
        query = r.recognize_google(audio, language="en-in")
        print(f"Heard: {query}")
        return query.lower()
    except sr.UnknownValueError:
        speak("Sorry, I did not understand that.")
        return None
    except sr.RequestError:
        speak("Speech recognition service is unavailable.")
        return None
    except Exception as e:
        speak(f"An error occurred while recognizing speech: {e}")
        print(f"Error: {e}")
        return None

def open_notepad() -> None:
    """Opens Windows Notepad."""
    try:
        os.system("notepad")
        speak("Opening Notepad")
    except Exception as e:
        speak("Sorry, I couldn't open Notepad")
        print(f"Error: {e}")

def scan_installed_apps() -> dict:
    """Scans common directories for installed applications."""
    installed_apps = {}
    search_locations = [
        os.path.expandvars(r"%ProgramFiles%"),
        os.path.expandvars(r"%ProgramFiles(x86)%"),
        os.path.expandvars(r"%LocalAppData%"),
        os.path.expandvars(r"%AppData%"),
        r"C:\Windows\System32",
        os.path.expandvars(r"%LocalAppData%\WhatsApp"),
        os.path.expandvars(r"%LocalAppData%\Microsoft\WindowsApps"),
        os.path.expandvars(r"%LocalAppData%\Programs")
    ]

    extensions = ['.exe']

    for location in search_locations:
        if os.path.exists(location):
            for root, _, files in os.walk(location):
                for file in files:
                    if any(file.lower().endswith(ext) for ext in extensions):
                        app_name = os.path.splitext(file)[0].lower()
                        installed_apps[app_name] = os.path.join(root, file)

    return installed_apps

INSTALLED_APPS = scan_installed_apps()

def open_app(app_name: str) -> None:
    """Opens a Windows application by name."""
    common_apps = {
        "chrome": "start chrome",
        "edge": "start msedge",
        "word": "start winword",
        "excel": "start excel",
        "powerpoint": "start powerpnt",
        "calculator": "calc",
        "paint": "mspaint",
        "cmd": "start cmd",
        "control panel": "control",
        "task manager": "taskmgr",
        "explorer": "explorer",
        "notepad": "notepad",
        "whatsapp": os.path.expandvars(r"%LocalAppData%\WhatsApp\WhatsApp.exe"),
    }

    try:
        app_name_lower = app_name.lower()

        if app_name_lower in ("whatsapp", "whatsapp desktop", "whatsapp messenger"):
            matches = [(n, p) for n, p in INSTALLED_APPS.items() if 'whatsapp' in n]
            if matches:
                try:
                    os.startfile(matches[0][1])
                    speak(f"Opening {matches[0][0]}")
                    print(f"Opening application: {matches[0][1]}")
                    return
                except Exception:
                    pass

            try:
                try:
                    os.startfile("whatsapp://")
                    speak("Opening WhatsApp")
                    print("Opening WhatsApp via URI protocol")
                    return
                except Exception:
                    os.system('start "" "whatsapp://"')
                    speak("Opening WhatsApp")
                    print("Opening WhatsApp via start command")
                    return
            except Exception as e:
                speak("Couldn't open WhatsApp automatically. Please open it from the Start menu or provide the path.")
                print(f"WhatsApp open error: {e}")
                return

        if app_name_lower in common_apps:
            command = common_apps[app_name_lower]
            if command.endswith('.exe'):
                os.startfile(command)
            else:
                os.system(command)
            speak(f"Opening {app_name}")
            print(f"Opening {app_name}")
            return

        possible_matches = []
        search_term = app_name_lower.replace(" ", "").replace(".", "")

        for installed_app, path in INSTALLED_APPS.items():
            if search_term in installed_app.replace(" ", "").replace(".", ""):
                possible_matches.append((installed_app, path))

        if len(possible_matches) == 1:
            app_path = f'"{possible_matches[0][1]}"'
            os.startfile(possible_matches[0][1])
            speak(f"Opening {possible_matches[0][0]}")
            print(f"Opening application: {app_path}")

        elif len(possible_matches) > 1:
            speak("I found multiple possible matches. Please choose one of the following options.")
            print("Multiple matches found:")
            for i, (app, _) in enumerate(possible_matches[:8], 1):
                print(f"{i}. {app}")
                speak(f"Option {i}: {app}")

            choice = takecommand("Please say the number or name of the application you want to open.")
            if choice:
                digits = ''.join(filter(str.isdigit, choice))
                if digits:
                    idx = int(digits) - 1
                    if 0 <= idx < len(possible_matches):
                        os.startfile(possible_matches[idx][1])
                        speak(f"Opening {possible_matches[idx][0]}")
                        print(f"Opening application: {possible_matches[idx][1]}")
                        return

                choice_name = choice.replace(' ', '').lower()
                for app, path in possible_matches:
                    if choice_name in app.replace(' ', '').lower():
                        os.startfile(path)
                        speak(f"Opening {app}")
                        print(f"Opening application: {path}")
                        return

            speak("Couldn't understand your choice. Aborting open request.")

        else:
            command = f"start {app_name_lower}"
            os.system(command)
            speak(f"Attempting to open {app_name}")
            print(f"Executing command: {command}")

    except Exception as e:
        speak(f"Sorry, I couldn't open {app_name}")
        print(f"Error: {e}")
        print("Try using the exact application name or provide more specific keywords.")

def get_system_info() -> None:
    """Gets and speaks system information."""
    system = platform.system()
    processor = platform.processor()
    memory = psutil.virtual_memory()
    disk_path = os.path.abspath(os.sep)
    try:
        disk = psutil.disk_usage(disk_path)
    except Exception:
        try:
            disk = psutil.disk_usage('C:\\')
        except Exception:
            disk = None

    info = f"Operating System: {system}\n"
    info += f"Processor: {processor}\n"
    info += f"RAM Usage: {memory.percent}%\n"
    info += f"Disk Usage: {disk.percent}%" if disk else "Disk Usage: N/A"

    speak("Here's your system information:")
    print(info)
    speak(info)

def get_battery_status() -> None:
    """Gets and speaks battery information."""
    battery = psutil.sensors_battery()
    if battery:
        status = "plugged in" if battery.power_plugged else "not plugged in"
        info = f"Battery percentage is {battery.percent}% and power is {status}"
    else:
        info = "Battery information not available"

    speak(info)
    print(info)

def get_running_processes() -> None:
    """Gets and speaks information about running processes."""
    processes = []
    for proc in psutil.process_iter(['name', 'memory_percent']):
        try:
            mem = proc.info.get('memory_percent') or 0.0
            name = proc.info.get('name') or 'Unknown'
            processes.append((name, mem))
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass

    processes.sort(key=lambda x: x[1], reverse=True)
    top_processes = processes[:5]

    speak("Here are the top 5 processes by memory usage:")
    for proc_name, memory in top_processes:
        info = f"{proc_name}: {memory:.1f}% memory"
        print(info)
        speak(info)

def play_music(song_name=None) -> None:
    """Plays music from the user's Music directory."""
    song_dir = os.path.expanduser("~\\Music")
    if not os.path.exists(song_dir):
        speak("Music folder not found on this system.")
        print(f"Music folder not found: {song_dir}")
        return

    songs = [f for f in os.listdir(song_dir) if os.path.isfile(os.path.join(song_dir, f))]

    if song_name:
        songs = [song for song in songs if song_name.lower() in song.lower()]

    if songs:
        song = random.choice(songs)
        try:
            os.startfile(os.path.join(song_dir, song))
            speak(f"Playing {song}.")
            print(f"Playing {song}.")
        except Exception as e:
            speak("Couldn't play the song.")
            print(f"Error playing song: {e}")
    else:
        speak("No song found.")
        print("No song found.")

def set_name() -> None:
    """Sets a new name for the assistant."""
    speak("What would you like to name me?")
    name = takecommand()
    if name:
        with open("assistant_name.txt", "w") as file:
            file.write(name)
        speak(f"Alright, I will be called {name} from now on.")
    else:
        speak("Sorry, I couldn't catch that.")

def load_name() -> str:
    """Loads the assistant's name from a file, or uses a default name."""
    try:
        with open("assistant_name.txt", "r") as file:
            return file.read().strip()
    except FileNotFoundError:
        return "viserys"


def search_wikipedia(query):
    """Searches Wikipedia and returns a summary."""
    try:
        speak("Searching Wikipedia...")
        result = wikipedia.summary(query, sentences=2)
        speak(result)
        print(result)
    except wikipedia.exceptions.DisambiguationError:
        speak("Multiple results found. Please be more specific.")
    except Exception:
        speak("I couldn't find anything on Wikipedia.")


if __name__ == "__main__":
    wishme()

    while True:
        query = takecommand("Listening for your command.")
        if not query:
            continue

        if "time" in query:
            time()

        elif "date" in query:
            date()

        elif "wikipedia" in query:
            wiki_query = query.replace("wikipedia", "").strip()
            if not wiki_query:
                wiki_query = takecommand("What would you like to search on Wikipedia?")
            if wiki_query:
                search_wikipedia(wiki_query)
            else:
                speak("No search query provided for Wikipedia.")

        elif "play music" in query:
            song_name = query.replace("play music", "").strip()
            if not song_name:
                song_name = takecommand("Which song would you like to play? Say part of the name.")
            play_music(song_name)

        elif "open youtube" in query:
            wb.open("youtube.com")

        elif "open google" in query:
            wb.open("google.com")

        elif "change your name" in query:
            set_name()

        elif "screenshot" in query:
            screenshot()
            speak("I've taken screenshot, please check it")

        elif "tell me a joke" in query:
            joke = random.choice(JOKES)
            speak(joke)
            print(joke)

        elif "open notepad" in query:
            open_notepad()

        elif "open" in query:
            app_name = query.replace("open", "").strip()
            if not app_name:
                app_name = takecommand("Which application would you like me to open?")
            if app_name:
                open_app(app_name)
            else:
                speak("Please specify which application to open.")

        elif "system info" in query or "system information" in query:
            get_system_info()

        elif "battery" in query or "battery status" in query:
            get_battery_status()

        elif "running processes" in query or "top processes" in query:
            get_running_processes()

        elif "shutdown" in query:
            speak("Shutting down the system, goodbye!")
            os.system("shutdown /s /f /t 1")
            break

        elif "restart" in query:
            speak("Restarting the system, please wait!")
            os.system("shutdown /r /f /t 1")
            break

        elif "offline" in query or "exit" in query:
            speak("Going offline. Have a good day!")
            break
