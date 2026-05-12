# Pink Assistant — AI Voice Controlled Desktop Assistant

Pink Assistant is a Python-based AI desktop assistant for Windows that combines:

- Voice Recognition
- Local LLM Intelligence
- Desktop Automation
- Spotify Controls
- System Controls
- Offline + Online Speech Recognition

The assistant listens for the wake word **"pink"** and executes commands like opening applications, controlling Spotify, adjusting system settings, searching YouTube, setting alarms, and more.

# Features

## Voice Assistant
- Wake word detection ("pink")
- Online speech recognition using Google Speech API
- Offline speech recognition using Vosk
- Text-to-speech responses


## Spotify Controls
Supports voice-controlled Spotify automation:

- Search songs
- Play tracks
- Pause music
- Next/Previous track
- Volume controls
- Shuffle
- Like songs

### Example Commands
pink search spotify faded
pink play
pink select 2
pink next
pink spotify volume up
pink shuffle

## Browser & YouTube Controls
Example Commands
pink open youtube
pink search youtube python tutorial
pink play despacito on youtube
pink open gmail

## System Controls
Supported Operations
Brightness control
Volume control
Open settings
Battery status
Lock laptop
Take screenshots
Open camera
Task manager
Open downloads
Create folders
Delete files
CPU/RAM monitoring
Example Commands
pink increase brightness
pink decrease volume
pink battery status
pink lock my laptop
pink take screenshot

## AI Integration

Pink Assistant uses:

Ollama
Llama 3

for natural language understanding.

The LLM converts user commands into structured actions.

Example:
{
  "action": "open_web",
  "url": "https://mail.google.com"
}

## Architecture
Microphone
   ↓
Speech Recognition
(Google / Vosk)
   ↓
Wake Word Detection
   ↓
Command Parser
   ↓
LLM Intent Processing
   ↓
Execution Engine
   ↓
Desktop Automation / System Controls

## Technologies Used
| Technology        | Purpose                        |
| ----------------- | ------------------------------ |
| Python            | Main programming language      |
| SpeechRecognition | Online speech recognition      |
| Vosk              | Offline speech recognition     |
| pyttsx3           | Text-to-speech                 |
| pyautogui         | Desktop automation             |
| pygetwindow       | Window handling                |
| psutil            | System monitoring              |
| Ollama            | Local LLM runtime              |
| Llama 3           | Natural language understanding |
| win32api          | Media/system key controls      |
| threading         | Background alarm monitoring    |

## Project Structure
pink-assistant/
│
├── main.py
├── run_pink.bat
├── requirements.txt
├── pink_app_index.json
│
├── models/
│   └── vosk-model/
│
├── sounds/
│   └── mustang.wav
│
└── assets/

## Installation
Clone Repository
git clone <your-repo-url>
cd pink-assistant

Create Virtual Environment
python -m venv venv
Activate:
Windows
venv\Scripts\activate

Install Dependencies
pip install -r requirements.txt
📦 Required Packages
SpeechRecognition
pyttsx3
psutil
pyautogui
screen-brightness-control
vosk
sounddevice
pywin32
ollama
pygetwindow

Setup Offline Speech Recognition
Download Vosk Model
Download model from:
https://alphacephei.com/vosk/models
Extract into:
models/vosk-model

Setup Ollama + Llama 3
Install Ollama
https://ollama.com
Pull Llama 3 Model
ollama run llama3:8b

Running the Assistant
Option 1
Double click:
run_pink.bat

Option 2
Run manually:
python main.py

Example Commands
Applications
pink open chrome
pink open vscode
pink close spotify
pink switch to chrome
Spotify
pink search spotify believer
pink play
pink pause
pink next
pink previous
pink spotify volume up
YouTube
pink open youtube
pink play coding music on youtube
System
pink battery status
pink open wifi settings
pink increase brightness
pink take screenshot

## AI Command Routing
Pink Assistant uses hybrid command processing:
Rule-Based Commands
Fast execution for:
Spotify
System controls
Browser automation
LLM-Based Commands
Natural language understanding using:
Ollama
Llama 3




