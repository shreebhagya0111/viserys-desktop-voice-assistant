# Python Voice Assistant

A Windows-based voice assistant built using **Python 3.14** that can recognize speech, speak responses, open apps, fetch information, and perform system tasks.

---

## 🚀 Features
- Voice input (mic) with typing fallback  
- Text-to-Speech using Windows SAPI / pyttsx3  
- Time & date announcements  
- Wikipedia search & jokes  
- Open apps (Notepad, Chrome, WhatsApp, etc.)  
- Play music & take screenshots  
- System info, battery status, running processes  
- Shutdown, restart & exit commands  

---

## 📦 Requirements

Install dependencies:

```bash
pip install speechrecognition wikipedia pyttsx3 pyaudio sounddevice pillow psutil numpy pywin32
✅ Works on Windows with Python 3.14

Save your code as:
assistant.py

Run:
python assistant.py

Example Commands:
what is the time
what is the date
search wikipedia python
tell me a joke
open notepad
play music
take screenshot
system info
battery status
shutdown
exit

Notes:
Microphone must be enabled in Windows privacy settings.
Install pyaudio properly for voice recognition support.
App launcher auto-detects installed programs.
