"""
Pink Assistant — full Spotify controls added (list-select + top-play update)

Usage:
- Double-click run_pink.bat (installs dependencies then runs this file)
- Say "pink" as the wake word before commands:
    - "pink search spotify faded"
    - "pink select 2"
    - "pink play"
    - "pink play 3"`
    - "pink pause"
    - "pink next"
    - "pink previous"
    - "pink spotify volume up"
    - "pink spotify volume down"
    - "pink shuffle"
    - "pink like"
- Works on Windows. Relies on pyautogui and (optionally) pygetwindow for window positioning.
"""
import threading
import pyautogui
import json
from vosk import Model, KaldiRecognizer
import sounddevice as sd
import json
import socket
import win32com.client
import ollama
import os
import keyboard
import sys
import time
import re
import subprocess
from datetime import datetime

# External libs (may be installed by run_pink.bat)
try:
    import speech_recognition as sr
    import pyttsx3
    import psutil
    import pyautogui
    from screen_brightness_control import get_brightness, set_brightness
    import win32api
    import winsound
except Exception as e:
    print("Missing libraries or partial install. Run run_pink.bat to auto-install requirements.")
    print("Error:", e)
    raise

# optional for reliable window coordinates
try:
    import pygetwindow as gw
except Exception:
    gw = None

# ========== CONFIG ==========
CONFIG = {
    "wake_word": "pink",        # assistant will ONLY respond when this word appears
    "user_name": "sir",

    # Startup sound
    "mustang_sound": r"C:\Users\alles\Downloads\mustang-7-87449 (1).wav",

    # App aliases (used for "switch to", "open")
    "app_aliases": {
        "chrome": ["chrome", "google"],
        "edge": ["edge", "microsoft edge"],
        "vscode": ["vscode", "code", "visual studio", "visual studio code"],
        "cmd": ["cmd", "command prompt", "terminal"],
        "notepad": ["notepad"],
        "spotify": ["spotify"],
        "whatsapp": ["whatsapp"]
    },
    
    # Actual executables
    "app_paths": {
        "chrome": "chrome.exe",
        "edge": "msedge.exe",
        "vscode": "code.exe",
        "cmd": "cmd.exe",
        "notepad": "notepad.exe",
        "spotify": "spotify.exe",
        "whatsapp": 'explorer.exe "shell:AppsFolder\\5319275A.WhatsAppDesktop_cv1g1gvanyjgm!App"'
    }
}
APP_INDEX_FILE = os.path.join(os.path.dirname(__file__), "pink_app_index.json")

# ========== Number parsing helpers ==========
_number_words = {
    "zero":0,"one":1,"two":2,"three":3,"four":4,"five":5,"six":6,"seven":7,"eight":8,"nine":9,
    "ten":10,"eleven":11,"twelve":12,"thirteen":13,"fourteen":14,"fifteen":15,"sixteen":16,
    "seventeen":17,"eighteen":18,"nineteen":19,"twenty":20,"thirty":30,"forty":40,"fifty":50,
    "sixty":60,"seventy":70,"eighty":80,"ninety":90,"hundred":100
}

def _extract_number_from_text(text):
    if not text:
        return None
    m = re.search(r'(\d{1,3})\s*%?', text)
    if m:
        try:
            val = int(m.group(1))
            return max(0, min(100, val))
        except:
            pass
    words = re.findall(r"[a-z]+", text.lower())
    if not words:
        return None
    total = 0
    current = 0
    found_any = False
    for w in words:
        if w in _number_words:
            found_any = True
            scale = _number_words[w]
            if scale == 100:
                if current == 0:
                    current = 100
                else:
                    current *= 100
            else:
                current += scale
        else:
            if current:
                total += current
                current = 0
    total += current
    if found_any:
        return max(0, min(100, int(total)))
    return None

def internet_available():
    try:
        socket.create_connection(("8.8.8.8", 53), timeout=2)
        return True
    except:
        return False

# ========== Voice Engine ==========
class VoiceEngine:
    def __init__(self):
        import speech_recognition as sr
        import pyttsx3

        self.recognizer = sr.Recognizer()   # 🔥 REQUIRED
        self.engine = pyttsx3.init()

        self.engine.setProperty("rate", 170)
        self.engine.setProperty("volume", 1.0)

        # Offline Vosk
        model_path = os.path.join(
            os.path.dirname(__file__),
            "models",
            "vosk-model"
            "vosk-model-small-en-us-0.15"
        )

        if not os.path.exists(model_path):
            raise RuntimeError(f"Vosk model folder not found at: {model_path}")

        self.vosk_model = Model(model_path)

        self.vosk_rec = KaldiRecognizer(self.vosk_model, 16000)

    def speak(self, text):
        if not text:
            return

        print(f"PINK: {text}")

        try:
            self.engine.stop()
            self.engine.say(text)
            self.engine.runAndWait()
        except Exception as e:
            print("TTS ERROR:", e)
    def stop(self):
        try:
            self.engine.stop()
        except:
            pass

    def listen_google(self, timeout=6, phrase_time_limit=6):
        with self.mic as source:
            self.recognizer.adjust_for_ambient_noise(source, duration=0.5)
            audio = self.recognizer.listen(source, timeout=timeout, phrase_time_limit=phrase_time_limit)
        return self.recognizer.recognize_google(audio).lower()

    def listen_vosk(self, duration=5):
        print("Listening (offline)...")
        audio = sd.rec(int(duration * 16000), samplerate=16000, channels=1, dtype="int16")
        sd.wait()

        if self.vosk_rec.AcceptWaveform(audio.tobytes()):

            result = json.loads(self.vosk_rec.Result())
            return result.get("text", "").lower()
        return ""

    def listen(self):
        import speech_recognition as sr

        try:
            with sr.Microphone() as source:
                print("Using Google Speech API")

                # 🔥 FIX HERE (correct place)
                self.recognizer.adjust_for_ambient_noise(source, duration=1)

                audio = self.recognizer.listen(source, phrase_time_limit=5)

            text = self.recognizer.recognize_google(audio)
            return text.lower()

        except Exception as e:
            print("Speech error:", e)
            return ""


# ========== App Controller ==========
class AppController:
    def __init__(self):
        self.search_results = []     # list of strings (query or placeholder titles)
        self.result_positions = []   # list of (x,y) coords for each result row
        self.spotify_window = None
        self.top_play_pos = None     # (x,y) for green play in Top result
        self.last_selected = None    # index of selected list item (1-based)
        self.start_menu = StartMenuIndexer()
        self.last_search_app = None      # "spotify", "youtube", etc.
        self.last_search_query = None 
        
        
    def get_active_window_title(self):
        if not gw:
            return ""
        try:
            win = gw.getActiveWindow()
            return win.title.lower() if win else ""
        except:
            return ""

    def load_app_index(self):
        if os.path.exists(APP_INDEX_FILE):
            with open(APP_INDEX_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        return {}

    
    
    
    def scan_apps(self):
        """
        Scans common Windows locations for .exe files
        and builds an app index.
        """
        search_dirs = [
            os.environ.get("ProgramFiles", ""),
            os.environ.get("ProgramFiles(x86)", ""),
            os.environ.get("LOCALAPPDATA", ""),
            os.environ.get("APPDATA", ""),
            os.path.expanduser("~/Desktop"),
            os.path.expanduser("~/Downloads")
        ]

        app_index = {}

        for base in search_dirs:
            if not base or not os.path.exists(base):
                continue

            for root, dirs, files in os.walk(base):
                for file in files:
                    if file.lower().endswith(".exe"):
                        name = file.lower().replace(".exe", "")
                        full_path = os.path.join(root, file)

                        # Store first occurrence only
                        if name not in app_index:
                            app_index[name] = full_path

        # Save to file
        with open(APP_INDEX_FILE, "w", encoding="utf-8") as f:
            json.dump(app_index, f, indent=2)

        return len(app_index)

    
    def open_as_website(self, name):
        import webbrowser

        name = name.lower().strip().replace(" ", "")

        # special cases
        SPECIAL_SITES = {
            "chatgpt": "https://chat.openai.com",
            "whatsapp": "https://web.whatsapp.com"
        }

        if name in SPECIAL_SITES:
            webbrowser.open(SPECIAL_SITES[name])
            return True

        # default
        if "." in name:
            url = f"https://{name}"
        else:
            url = f"https://www.{name}.com"

        try:
            webbrowser.open(url)
            return True
        except:
            return False


    def open_and_click_first_result(self, query):
        import webbrowser
        import time
        import pyautogui

        # open google search
        url = f"https://www.google.com/search?q={query.replace(' ', '+')}"
        webbrowser.open(url)

        time.sleep(3)

        # focus browser (reuse your existing function if possible)
        try:
            self.switch_to_app("chrome")
        except:
            pass

        time.sleep(1)

        # navigate to first result
        for _ in range(6):   # safer than fixed 5
            pyautogui.press("tab")
            time.sleep(0.1)

        pyautogui.press("enter")
        return True
    
    
    
    
    
    def switch_to_app(self, app_name):
        app_name = app_name.lower().strip()

        # Try to activate existing window
        if gw:
            for w in gw.getAllWindows():
                if app_name in (w.title or "").lower():
                    try:
                        if w.isMinimized:
                            w.restore()
                        w.activate()
                        return True
                    except:
                        pass

        # If not found → open it
        return self.open_app(app_name)







    
    
    
    def _find_spotify_window(self):
        if gw:
            try:
                wins = gw.getWindowsWithTitle("Spotify")
                if wins:
                    w = wins[0]
                    self.spotify_window = (w.left, w.top, w.width, w.height)
                    return self.spotify_window
            except Exception:
                pass
        screen_w, screen_h = pyautogui.size()
        self.spotify_window = (0, 0, screen_w, screen_h)
        return self.spotify_window

    def open_app(self, name):
        name = name.lower().strip()

        # 🌐 Known Web Apps
        web_apps = {
            "gmail": "https://mail.google.com",
            "youtube": "https://www.youtube.com",
            "google": "https://www.google.com",
            "chatgpt": "https://chat.openai.com",
            "instagram": "https://www.instagram.com",
            "facebook": "https://www.facebook.com"
        }

        if name in web_apps:
            os.system(f"start {web_apps[name]}")
            return True

        # 1️⃣ Exact match in CONFIG
        if name in CONFIG.get("app_paths", {}):
            try:
                subprocess.Popen(CONFIG["app_paths"][name], shell=True)
                return True
            except:
                pass

        # 2️⃣ Start Menu search
        path = self.start_menu.find(name)
        if path:
            try:
                subprocess.Popen(path)
                return True
            except:
                pass

        # 3️⃣ Scanned app index
        app_index = self.load_app_index()
        for app_name, path in app_index.items():
            if name in app_name:
                try:
                    subprocess.Popen(path)
                    return True
                except:
                    pass

       # 4️⃣ Try opening as direct website
        if self.open_as_website(name):
            return True

        # 5️⃣ FINAL fallback → Google search + open first result
        return self.open_and_click_first_result(name)
        
        
        

    def close_app(self, name):
        name = name.lower().strip()
        if name in CONFIG['app_paths']:
            proc = CONFIG['app_paths'][name]
        else:
            proc = name
        try:
            os.system(f"taskkill /f /im {proc} >nul 2>&1")
            return True
        except Exception as e:
            print("Close app error:", e)
            return False

    # ----- UPDATED: search_spotify stores top-play coordinates and list column coords -----
    def search_spotify(self, text):
        """
        Opens Spotify, searches for text, and stores approximate coordinates for result rows
        and the top-result play button (so play will click that green button directly).
        Resets previous selection.
        """
        if not self.open_app("spotify"):
            return False
        time.sleep(2.5)

        try:
            # focus search bar and type
            pyautogui.hotkey('ctrl', 'l')
            time.sleep(0.2)
            pyautogui.write(text, interval=0.05)
            pyautogui.press('enter')
            time.sleep(2.5)  # wait for results to load

            # locate Spotify window to compute coordinates
            left, top, width, height = self._find_spotify_window()

            # --- compute top-result play button position (heuristic)
            # These multipliers work well for normal/maximized Spotify windows; adjust if needed.
            play_x = left + int(width * 0.60)   # x inside the top-result card (toward right)
            play_y = top  + int(height * 0.18)  # y near top of window where the top-result sits
            self.top_play_pos = (play_x, play_y)

            # Determine start point for song rows for select_result (list on the right column)
            # SHIFTED to right column compared to earlier naive values.
            start_x = left + int(width * 0.47)  # approx column where song titles appear
            start_y = top + int(height * 0.24)  # first song row a little below the top result
            gap = max(40, int(height * 0.06))                # row height   # vertical gap between rows
            count = 8
            self.result_positions = []
            for i in range(count):
                y = start_y + i * gap
                self.result_positions.append((start_x, y))

            # store placeholder titles (can't OCR offline)
            self.search_results = [f"{text} (result {i+1})" for i in range(count)]

            # reset last selected index (user must select again)
            self.last_selected = None

            print(f"Stored {len(self.result_positions)} result positions and top-play {self.top_play_pos} for Spotify.")
            return True
        except Exception as e:
            print("Spotify search error:", e)
            return False
        self.last_search_app = "spotify"
        self.last_search_query = text

    def select_result(self, n):
        """
        Move the mouse to the nth search-result row (1-based) and single-click to focus it.
        Records the selection so 'pink play' will play the selected row.
        """
        try:
            if not self.result_positions:
                print("No stored results — do a search first.")
                return False
            if n <= 0 or n > len(self.result_positions):
                print("Requested index out of range.")
                return False
            x, y = self.result_positions[n-1]
            pyautogui.moveTo(x, y, duration=0.3)
            pyautogui.click()
            # store selection for play() to act on
            self.last_selected = n
            print(f"Moved to and selected result {n} at ({x},{y})")
            return True
        except Exception as e:
            print("Spotify select error:", e)
            return False

    # ----- UPDATED: play_first_result plays selected list item when present, else top-play -----
    def play_first_result(self):
        """
        Play either the currently selected list item (if any) OR click the Top-result play button.
        """
        try:
            if not self.result_positions and not self.top_play_pos:
                print("No stored results — do a search first.")
                return False

            # Try to activate Spotify window
            if gw:
                try:
                    wins = gw.getWindowsWithTitle("Spotify")
                    if wins:
                        wins[0].activate()
                        time.sleep(0.35)
                except Exception:
                    pass

            # If user selected a list item previously, play that one
            if self.last_selected:
                idx = self.last_selected
                if 1 <= idx <= len(self.result_positions):
                    x, y = self.result_positions[idx-1]
                    pyautogui.moveTo(x, y, duration=0.2)
                    pyautogui.click()
                    time.sleep(0.12)
                    pyautogui.press('enter')
                    print(f"Played selected result {idx} at ({x},{y}).")
                    return True
                else:
                    # invalid index stored, clear it
                    self.last_selected = None





            # No selection — click top-play if we have coordinates
                        # No selection — click top-play if we have coordinates
            if self.top_play_pos:
                px, py = self.top_play_pos
                # Try to ensure Spotify window is active first
                try:
                    if gw:
                        wins = gw.getWindowsWithTitle("Spotify")
                        if wins:
                            wins[0].activate()
                            time.sleep(0.2)
                except:
                    pass
                pyautogui.moveTo(px, py, duration=0.25)
                pyautogui.click()
                time.sleep(0.25)
                return True

            
            
            
            
            

            # final fallback: select first row and press Enter
            x, y = self.result_positions[0]
            pyautogui.moveTo(x, y, duration=0.25)
            pyautogui.click()
            time.sleep(0.15)
            pyautogui.press('enter')
            print("Fallback: pressed Enter on first result.")
            return True
        except Exception as e:
            print("Spotify play error:", e)
            return False

    def play_nth_result(self, n):
        try:
            ok = self.select_result(n)
            if not ok:
                return False
            time.sleep(0.15)
            pyautogui.press('enter')
            return True
        except Exception as e:
            print("play_nth_result error:", e)
            return False

    # Media key controls using win32api (virtual-key codes)
    def spotify_play_pause(self):
        try:
            win32api.keybd_event(0xB3, 0)  # VK_MEDIA_PLAY_PAUSE
            return True
        except Exception as e:
            print("play_pause error:", e)
            return False

    def spotify_next(self):
        try:
            win32api.keybd_event(0xB0, 0)  # VK_MEDIA_NEXT_TRACK
            return True
        except Exception as e:
            print("next error:", e)
            return False

    def spotify_previous(self):
        try:
            win32api.keybd_event(0xB1, 0)  # VK_MEDIA_PREV_TRACK
            return True
        except Exception as e:
            print("previous error:", e)
            return False

    def spotify_volume_up(self, steps=3):
        try:
            for _ in range(max(1, int(steps))):
                win32api.keybd_event(0xAF, 0)  # Volume Up
                time.sleep(0.06)
            return True
        except Exception as e:
            print("volume_up error:", e)
            return False

    def spotify_volume_down(self, steps=3):
        try:
            for _ in range(max(1, int(steps))):
                win32api.keybd_event(0xAE, 0)  # Volume Down
                time.sleep(0.06)
            return True
        except Exception as e:
            print("volume_down error:", e)
            return False

    def spotify_mute_toggle(self):
        try:
            win32api.keybd_event(0xAD, 0)  # Mute
            return True
        except Exception as e:
            print("mute error:", e)
            return False

    # Heuristic toggle for shuffle — clicks approximate shuffle button in Spotify bottom-left
    def spotify_toggle_shuffle(self):
        try:
            left, top, width, height = self._find_spotify_window()
            shuffle_x = left + int(width * 0.10)
            shuffle_y = top + int(height * 0.90)
            pyautogui.moveTo(shuffle_x, shuffle_y, duration=0.2)
            pyautogui.click()
            print("Clicked shuffle (heuristic).")
            return True
        except Exception as e:
            print("shuffle error:", e)
            return False

    # Heuristic like/unlike current track: click near bottom-left area of track details
    def spotify_like_unlike(self):
        try:
            left, top, width, height = self._find_spotify_window()
            candidates = [
                (left + int(width * 0.35), top + int(height * 0.86)),
                (left + int(width * 0.75), top + int(height * 0.20)),
                (left + int(width * 0.88), top + int(height * 0.14)),
            ]
            for (x, y) in candidates:
                pyautogui.moveTo(x, y, duration=0.15)
                pyautogui.click()
                time.sleep(0.12)
            print("Attempted like/unlike clicks (heuristic).")
            return True
        except Exception as e:
            print("like/unlike error:", e)
            return False

# ========== System Controller (brightness/volume etc.) ==========
class SystemController:
    
    
    def __init__(self, voice_engine):
        self.voice = voice_engine
        self.apps = AppController()
        self.alarm = AlarmManager(self.voice)
    
    def contextual_search(self, query):
        title = self.apps.get_active_window_title()

        def clear_and_search():
            pyautogui.hotkey("ctrl", "a")
            time.sleep(0.05)
            pyautogui.press("backspace")
            time.sleep(0.05)
            pyautogui.write(query)
            pyautogui.press("enter")

        # YOUTUBE
        if "youtube" in title:
            pyautogui.press("/")
            time.sleep(0.2)
            clear_and_search()
            return "Searching on YouTube."

        # SPOTIFY
        if "spotify" in title:
            pyautogui.hotkey("ctrl", "l")
            time.sleep(0.2)
            clear_and_search()
            return "Searching on Spotify."

        # BROWSER (Google / Edge)
        if "chrome" in title or "edge" in title:
            pyautogui.hotkey("ctrl", "l")
            time.sleep(0.2)
            clear_and_search()
            return "Searching on the web."

        # FALLBACK
        os.system(f"start https://www.google.com/search?q={query.replace(' ', '+')}")
        return "Searching on Google."


    def contextual_play(self, query):
        title = self.apps.get_active_window_title()

        if "spotify" in title:
            self.apps.search_spotify(query)
            time.sleep(1.2)
            self.apps.play_first_result()
            return "Playing on Spotify."

        if "youtube" in title or "chrome" in title or "edge" in title:
            self.open_youtube(query)
            return "Playing on YouTube."

        os.system(f"start https://www.google.com/search?q={query.replace(' ', '+')}")
        return "Searching and playing on Google."

    
        
        
    def open_youtube(self, query=None):
        url = "https://www.youtube.com"
        if query:
            q = query.replace(" ", "+")
            url = f"https://www.youtube.com/results?search_query={q}"
        # open URL in default browser
        os.system(f"start {url}")
        time.sleep(1.7)

        # Try to focus common browsers so keyboard presses do something
        self.try_focus_browser()

        # If query, attempt to open top result by pressing Tab/Enter (best-effort)
        if query:
            time.sleep(1.2)
            # reduce number of tabs — some browsers focus the address bar first
            pyautogui.press("tab", presses=4, interval=0.12)
            pyautogui.press("enter")
        return "YouTube opened."

    def try_focus_browser(self):
        # helper to attempt to focus an opened browser window
        browsers = ["chrome", "edge", "firefox"]
        for b in browsers:
            try:
                if self.apps.switch_to_app(b):
                    time.sleep(0.2)
                    return True
            except:
                pass
        return False


    def youtube_control(self, action):
        if action == "play":
            pyautogui.press("space")
        elif action == "next":
            pyautogui.hotkey("shift", "n")
        elif action == "forward":
            pyautogui.press("l")
        elif action == "rewind":
            pyautogui.press("j")

    def play_sound(self, path):
        try:
            if path and os.path.exists(path):
                winsound.PlaySound(path, winsound.SND_FILENAME | winsound.SND_ASYNC)
        except Exception:
            pass

    def check_battery(self):
        try:
            b = psutil.sensors_battery()
            if not b:
                return "Battery information not available."
            pct = int(b.percent)
            plugged = b.power_plugged
            state = "plugged in" if plugged else "not plugged in"
            note = ""
            if pct < 20 and not plugged:
                note = " You should plug in the charger immediately."
            return f"Battery is at {pct}% and {state}.{note}"
        except Exception as e:
            return "Unable to read battery status."

    def adjust_zoom(self, command):
        cmd = command.lower()

        # extract percentage (10, 20, etc.)
        num = _extract_number_from_text(cmd)
        if num is None:
            num = 10  # default step

        # each Ctrl + press ≈ 10%
        steps = max(1, num // 10)
        if "reset" in cmd:
            pyautogui.hotkey("ctrl", "0")
            return "Zoom reset to default."

        if any(w in cmd for w in ["out", "decrease", "reduce"]):
            for _ in range(steps):
                pyautogui.hotkey("ctrl", "-")
                time.sleep(0.05)
            return f"Zoomed out by {num}%."

        if any(w in cmd for w in ["in", "increase"]):
            for _ in range(steps):
                pyautogui.hotkey("ctrl", "+")
                time.sleep(0.05)
            return f"Zoomed in by {num}%."

        return "I couldn't determine zoom direction."

    
    
    
    def adjust_brightness(self, command):
        try:
            try:
                cur = get_brightness()
                cur_val = cur[0] if isinstance(cur, (list, tuple)) else cur
                cur_val = int(cur_val)
            except Exception:
                cur_val = None

            cmd = (command or "").lower()

            if " to " in cmd:
                num = _extract_number_from_text(cmd.split(" to ")[-1])
                if num is None:
                    num = _extract_number_from_text(cmd)
                if num is None:
                    return "Couldn't parse target brightness amount."
                try:
                    set_brightness(num)
                    return f"Brightness set to {num}%."
                except Exception as e:
                    return f"Couldn't set brightness: {e}"

            if " by " in cmd:
                num = _extract_number_from_text(cmd.split(" by ")[-1])
                if num is None:
                    num = _extract_number_from_text(cmd)
                if num is None:
                    num = 20
                if any(w in cmd for w in ["increase", "up", "raise", "brighten"]):
                    if cur_val is None:
                        return "Couldn't read current brightness."
                    new = min(100, cur_val + num)
                else:
                    if cur_val is None:
                        return "Couldn't read current brightness."
                    new = max(0, cur_val - num)
                try:
                    set_brightness(new)
                    return f"Brightness set to {new}%."
                except Exception as e:
                    return f"Couldn't set brightness: {e}"

            if any(w in cmd for w in ["increase", "up", "raise", "brighten"]):
                step = 20
                if cur_val is None:
                    return "Couldn't read current brightness."
                new = min(100, cur_val + step)
                set_brightness(new)
                return f"Brightness set to {new}%."
            if any(w in cmd for w in ["decrease", "down", "lower", "dim"]):
                step = 20
                if cur_val is None:
                    return "Couldn't read current brightness."
                new = max(0, cur_val - step)
                set_brightness(new)
                return f"Brightness set to {new}%."

            num = _extract_number_from_text(cmd)
            if num is not None:
                try:
                    set_brightness(num)
                    return f"Brightness set to {num}%."
                except Exception as e:
                    return f"Couldn't set brightness: {e}"

            return "I couldn't determine how much to change the brightness by."
        except Exception as e:
            return f"Couldn't adjust brightness: {e}"

    def adjust_volume(self, command):
        try:
            cmd = (command or "").lower()
            num = None
            mode = "delta"
            if " to " in cmd:
                num = _extract_number_from_text(cmd.split(" to ")[-1])
                mode = "set"
            elif " by " in cmd:
                num = _extract_number_from_text(cmd.split(" by ")[-1])
                mode = "delta"
            else:
                num = _extract_number_from_text(cmd)
                mode = "delta"

            if num is None:
                num = 10

            def press_volume_up(times):
                for _ in range(max(1, int(times))):
                    win32api.keybd_event(0xAF, 0)
                    time.sleep(0.06)

            def press_volume_down(times):
                for _ in range(max(1, int(times))):
                    win32api.keybd_event(0xAE, 0)
                    time.sleep(0.06)

            if mode == "set":
                assumed_current = 50
                delta = num - assumed_current
                presses = max(1, round(abs(delta) / 2))
                if delta > 0:
                    press_volume_up(presses)
                elif delta < 0:
                    press_volume_down(presses)
                return f"Attempted to set volume to {num}% (approximation)."

            presses = max(1, round(abs(num) / 2))
            if any(w in cmd for w in ["increase", "up", "louder", "raise"]):
                press_volume_up(presses)
                return f"Increased volume by ~{num}%."
            elif any(w in cmd for w in ["decrease", "down", "lower", "quieter"]):
                press_volume_down(presses)
                return f"Decreased volume by ~{num}%."
            elif "mute" in cmd:
                win32api.keybd_event(0xAD, 0)
                return "Toggled mute."
            else:
                press_volume_up(presses)
                return f"Adjusted volume by ~{num}%."
        except Exception as e:
            return f"Couldn't adjust volume: {e}"

    def open_settings(self, page=""):
        try:
            pages = {
                "wifi": "ms-settings:network-wifi",
                "bluetooth": "ms-settings:bluetooth",
                "display": "ms-settings:display",
                "sound": "ms-settings:sound",
                "battery": "ms-settings:batterysaver"
            }
            if page in pages:
                os.system(f"start {pages[page]}")
                return f"Opened {page} settings."
            os.system("start ms-settings:")
            return "Opened Windows settings."
        except Exception as e:
            return f"Couldn't open settings: {e}"

    def get_time(self):
        now = datetime.now()
        return now.strftime("%I:%M %p")
class AlarmManager:
    def __init__(self, voice):
        self.voice = voice
        self.alarms = []
        self.running = True
        threading.Thread(target=self._monitor, daemon=True).start()

    def add_alarm(self, alarm_time):
        self.alarms.append(alarm_time)

    def clear_alarms(self):
        self.alarms.clear()

    def list_alarms(self):
        return self.alarms

    def _monitor(self):
        while self.running:
            now = datetime.now().strftime("%H:%M")
            for alarm in self.alarms[:]:
                if alarm == now:
                    self.voice.speak("Wake up sir. Alarm ringing.")
                    winsound.Beep(1500, 4000)
                    self.alarms.remove(alarm)
            time.sleep(20)
class StartMenuIndexer:
    def __init__(self):
        self.apps = {}
        self.build_index()

    def build_index(self):
        paths = [
            os.path.join(os.environ["APPDATA"], "Microsoft", "Windows", "Start Menu", "Programs"),
            os.path.join(os.environ["PROGRAMDATA"], "Microsoft", "Windows", "Start Menu", "Programs")
        ]

        shell = win32com.client.Dispatch("WScript.Shell")

        for base in paths:
            for root, _, files in os.walk(base):
                for f in files:
                    if f.endswith(".lnk"):
                        name = os.path.splitext(f)[0].lower()
                        full = os.path.join(root, f)
                        try:
                            shortcut = shell.CreateShortcut(full)
                            target = shortcut.TargetPath
                            if target:
                                self.apps[name] = target
                        except:
                            pass

    def find(self, spoken_name):
        spoken_name = spoken_name.lower()
        for name, path in self.apps.items():
            if spoken_name in name:
                return path
        return None
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

class SimpleRAG:
    def __init__(self, filepath):
        self.filepath = filepath
        self.docs = []
        self.vectorizer = TfidfVectorizer()
        self.vectors = None
        self.load_data()

    def load_data(self):
        if not os.path.exists(self.filepath):
            return
        
        with open(self.filepath, "r", encoding="utf-8") as f:
            self.docs = [line.lower().strip() for line in f.readlines() if line.strip()]
            print("RAG loaded docs:", self.docs)

        if self.docs:
            self.vectors = self.vectorizer.fit_transform(self.docs)

    def retrieve(self, query, top_k=3):
        if not self.docs:
            return []

        query_vec = self.vectorizer.transform([query.lower()])
        sims = cosine_similarity(query_vec, self.vectors).flatten()

        top_indices = sims.argsort()[-top_k:][::-1]
        return [self.docs[i] for i in top_indices]
    
    
        
        
        
class ActivityTracker:
    def __init__(self, app_controller):
        self.apps = app_controller
        self.file = os.path.join(os.path.dirname(__file__), "memory", "activity.json")

    def log_activity(self):
        title = self.apps.get_active_window_title()
        now = datetime.now().strftime("%H:%M")

        entry = {
            "time": now,
            "title": title
        }

        data = []
        if os.path.exists(self.file):
            with open(self.file, "r") as f:
                data = json.load(f)

        data.append(entry)

        # keep last 50 entries only
        data = data[-50:]

        with open(self.file, "w") as f:
            json.dump(data, f, indent=2)



import tkinter as tk
import threading


class PopupUI:
    def __init__(self):
        self.root = None

    def show(self, message):
        def run():
            if self.root:
                try:
                    self.root.destroy()
                except:
                    pass

            self.root = tk.Tk()
            self.root.overrideredirect(True)
            self.root.configure(bg="black")
            self.root.attributes("-topmost", True)

            # 🔥 SIZE + POSITION
            w, h = 900, 300
            screen_w = self.root.winfo_screenwidth()
            screen_h = self.root.winfo_screenheight()

            x = (screen_w // 2) - (w // 2)
            y = screen_h - h - 50

            self.root.geometry(f"{w}x{h}+{x}+{y}")

            # 🔥 🔥 🔥 PASTE YOUR HUD CODE HERE 🔥 🔥 🔥

            container = tk.Frame(self.root, bg="black")
            container.pack(fill="both", expand=True)

            # LEFT
            left = tk.Frame(container, bg="black")
            left.pack(side="left", fill="y", padx=10)

            tk.Label(left, text="SYSTEM", fg="#00ffff", bg="black",
                    font=("Consolas", 12, "bold")).pack(anchor="w")

            tk.Label(left, text="STATUS: ACTIVE", fg="#00ffff", bg="black").pack(anchor="w")
            tk.Label(left, text="AI: ONLINE", fg="#00ffff", bg="black").pack(anchor="w")

            # CENTER
            center = tk.Frame(container, bg="black")
            center.pack(side="left", expand=True, fill="both")

            tk.Label(center,
                    text="PINK ASSISTANT",
                    fg="#00ffff",
                    bg="black",
                    font=("Consolas", 16, "bold")).pack(pady=5)

            tk.Label(center,
                    text=message,
                    fg="#00ffff",
                    bg="black",
                    wraplength=600,
                    font=("Consolas", 14)).pack(pady=20)

            # RIGHT
            right = tk.Frame(container, bg="black")
            right.pack(side="right", fill="y", padx=10)

            tk.Label(right, text="WELLBEING", fg="#00ffff", bg="black",
                    font=("Consolas", 12, "bold")).pack(anchor="e")

            stats = get_wellbeing_stats()
            score = calculate_focus_score(stats)

            for k, v in stats.items():
                tk.Label(
                    right,
                    text=f"{k}: {v}m",
                    fg="#00ffff",
                    bg="black"
                ).pack(anchor="e")

            tk.Label(
                right,
                text=f"Focus: {score}%",
                fg="#00ffff",
                bg="black",
                font=("Consolas", 12, "bold")
            ).pack(anchor="e", pady=5)
            
            
            

            self.root.mainloop()

        threading.Thread(target=run, daemon=True).start()

    def close(self):
        if self.root:
            try:
                self.root.quit()
                self.root.destroy()
            except:
                pass
            self.root = None

def clean_title(title):
    title = title.lower()

    # remove browser junk
    for junk in ["- microsoft edge", "- google chrome", "- personal"]:
        title = title.replace(junk, "")

    return title.strip()


import re

def clean_title(title):
    title = title.lower()

    for junk in ["- microsoft edge", "- google chrome", "- personal"]:
        title = title.replace(junk, "")

    title = re.sub(r"and \d+ more pages", "", title)
    title = re.sub(r"[|]", " ", title)

    return " ".join(title.split()).strip()

def classify_activity(title):
    title = title.lower()

    if any(k in title for k in ["code", "vscode", "github", "leetcode", "python"]):
        return "coding"

    elif any(k in title for k in ["tutorial", "course", "learn", "lecture", "ai"]):
        return "learning"

    elif any(k in title for k in ["youtube", "song", "movie", "ipl", "comedy"]):
        return "entertainment"

    return "other"

def get_wellbeing_stats():
    path = os.path.join(os.path.dirname(__file__), "memory", "activity.json")

    if not os.path.exists(path):
        return {}

    with open(path, "r") as f:
        data = json.load(f)

    stats = {
        "coding": 0,
        "learning": 0,
        "entertainment": 0,
        "other": 0
    }

    for d in data:
        title = clean_title(d["title"])
        category = classify_activity(title)

        stats[category] += 10   # each log = 10 sec

    # convert to minutes
    for k in stats:
        stats[k] = round(stats[k] / 60, 1)

    return stats

def calculate_focus_score(stats):
    productive = stats["coding"] + stats["learning"]
    total = sum(stats.values())

    if total == 0:
        return 0

    return int((productive / total) * 100)


def activity_to_text():
    path = os.path.join(os.path.dirname(__file__), "memory", "activity.json")

    if not os.path.exists(path):
        return []

    with open(path, "r") as f:
        data = json.load(f)

    if not data:
        return []

    # latest activity
    latest = clean_title(data[-1]["title"])

    # approx 10 min ago (60 logs × 10 sec)
    past_index = max(0, len(data) - 60)
    past = clean_title(data[past_index]["title"])

    return [latest, past]


def get_yesterday_activity():
    path = os.path.join(os.path.dirname(__file__), "memory", "activity.json")

    if not os.path.exists(path):
        return []

    with open(path, "r") as f:
        data = json.load(f)

    if not data:
        return []

    # get current time
    now = datetime.now()

    results = []

    for d in data:
        try:
            t = datetime.strptime(d["time"], "%H:%M")

            # assume same day → simulate "yesterday" using older logs
            # (since we don't store date yet)
            results.append(clean_title(d["title"]))
        except:
            continue

    # return last meaningful 3 activities
    return list(dict.fromkeys(results))[-5:]  
    
        
    
# ========== Main Assistant =====
class PinkAssistant:
    def __init__(self):
        self.voice = VoiceEngine()
        self.system = SystemController(self.voice)
        self.rag = SimpleRAG(
            os.path.join(os.path.dirname(__file__), "memory", "knowledge.txt")
        )
        self.tracker = ActivityTracker(self.system.apps)
        self.ui = PopupUI()
        self.ui_active = False
        threading.Thread(target=self.track_loop, daemon=True).start()
        self.boot()

    def boot(self):
        if CONFIG.get("mustang_sound"):
            self.system.play_sound(CONFIG["mustang_sound"])
        self.voice.speak(f"All systems operational. Good {self.get_time_of_day()}, {CONFIG.get('user_name','sir')}.")
        try:
            batt = self.system.check_battery()
            self.voice.speak(batt)
        except Exception:
            pass

    def get_time_of_day(self):
        h = datetime.now().hour
        if h < 12:
            return "morning"
        if h < 18:
            return "afternoon"
        return "evening"
    
    
    def track_loop(self):
        while True:
            self.tracker.log_activity()
            time.sleep(10)

    def parse_and_execute(self, command):
        c = (command or "").lower().strip()

        show_ui = False

        if "show me" in c or "display" in c:
            show_ui = True
            c = c.replace("show me", "").replace("display", "").strip()


        # 🔥 NOW THIS WILL WORK
        if "digital wellbeing" in c or "focus report" in c:
            stats = get_wellbeing_stats()
            score = calculate_focus_score(stats)

            msg = f"You spent {stats['coding']} minutes coding, {stats['learning']} minutes learning, and your focus score is {score} percent."

            if show_ui:
                self.ui.show(msg)
                self.ui_active = True

            self.voice.speak(msg)
            return
        
        # ===== STRIP WAKE WORD (CRITICAL FIX) =====
        if c.startswith(CONFIG["wake_word"] + " "):
            c = c[len(CONFIG["wake_word"]) + 1:]
        
        
        show_ui = False

        if "show me" in c or "display" in c:
            show_ui = True
            c = c.replace("show me", "").replace("display", "").strip()
            
        if c.startswith("search "):
            query = c.replace("search", "").strip()
            res = self.system.contextual_search(query)
            self.voice.speak(res)
            return

          
        if c in ("scan apps", "rescan apps", "update apps"):
            count = self.system.apps.scan_apps()
            self.voice.speak(f"I indexed {count} applications.")
            return

        if c.startswith("open "):
            app = c.replace("open", "").strip()

            # 🚫 youtube is NOT an app
            if app == "youtube":
                self.voice.speak(self.system.open_youtube())
                return

            ok = self.system.apps.open_app(app)
            self.voice.speak(
                f"Opened {app}." if ok else f"I could not find {app} on this system."
            )
            return


          
        # ----- EXACT SETTINGS COMMANDS (MUST COME FIRST) -----
        if c in ("battery settings", "open battery settings", "open battery"):
            res = self.system.open_settings("battery")
            self.voice.speak(res)
            return

        if c in ("display settings", "open display settings", "open display"):
            res = self.system.open_settings("display")
            self.voice.speak(res)
            return

        if c in ("wifi settings", "open wifi settings", "open wifi"):
            res = self.system.open_settings("wifi")
            self.voice.speak(res)
            return
  
          
          
          
          
            
        

        if not c or c == "unrecognized":
            return

        if "battery" in c or "charge" in c:
            self.voice.speak(self.system.check_battery()); return
        if "time" in c:
            t = self.system.get_time()
            self.voice.speak(f"The time is {t}"); return
            
        # ----- Alarm commands -----
        if "alarm" in c or "wake me up" in c:
            if "cancel" in c or "stop" in c or "remove" in c:
                self.system.alarm.clear_alarms()
                self.voice.speak("All alarms cancelled.")
                return

            if "list" in c:
                alarms = self.system.alarm.list_alarms()
                if not alarms:
                    self.voice.speak("No alarms are set.")
                else:
                    self.voice.speak("Alarms set at " + ", ".join(alarms))
                return

            # set alarm
            match = re.search(r'(\d{1,2})\s*(\d{2})?\s*(am|pm)?', c)
            if match:
                hour = int(match.group(1))
                minute = int(match.group(2) or 0)
                mer = match.group(3)

                if mer == "pm" and hour != 12:
                    hour += 12
                if mer == "am" and hour == 12:
                    hour = 0

                alarm_time = f"{hour:02d}:{minute:02d}"
                self.system.alarm.add_alarm(alarm_time)
                self.voice.speak(f"Alarm set for {alarm_time}.")
                return

    
        if any(w in c for w in ["brightness", "display", "bright"]):
            res = self.system.adjust_brightness(c)
            self.voice.speak(res); return
        if any(w in c for w in ["volume", "louder", "quieter", "mute"]) and ("spotify" not in c):
            res = self.system.adjust_volume(c)
            self.voice.speak(res); return
        if "settings" in c:
            page = ""
            for p in ["wifi", "bluetooth", "display", "sound", "battery"]:
                if p in c:
                    page = p; break
            res = self.system.open_settings(page)
            self.voice.speak(res); return

        # ----- Switch to app (generic) -----
        if c.startswith("switch to "):
            app = c.replace("switch to", "").strip()
            ok = self.system.apps.switch_to_app(app)
            self.voice.speak(f"Switched to {app}." if ok else f"Couldn't switch to {app}.")
            return

        # ----- YouTube commands -----
        if "youtube" in c:
            if "open" in c:
                self.voice.speak(self.system.open_youtube())
                return

            if "search" in c or "play" in c:
                query = c.replace("youtube", "").replace("search", "").replace("play", "").strip()
                self.voice.speak(self.system.open_youtube(query))
                return

            if "next" in c:
                self.system.youtube_control("next")
                self.voice.speak("Next video.")
                return

            if "forward" in c:
                self.system.youtube_control("forward")
                self.voice.speak("Forwarding video.")
                return

            if "rewind" in c or "back" in c:
                self.system.youtube_control("rewind")
                self.voice.speak("Rewinding video.")
                return

            if "close" in c:
                pyautogui.hotkey("ctrl", "w")
                self.voice.speak("Closed YouTube tab.")
                return
        
        # generic app open / close
        if c.startswith("open "):
            app = c.split("open",1)[1].strip()
            ok = self.system.apps.open_app(app)
            self.voice.speak(f"Opened {app}." if ok else f"Couldn't open {app}.")
            return
        # ----- Zoom commands -----
        if "zoom" in c:
            res = self.system.adjust_zoom(c)
            self.voice.speak(res)
            return

        if "minimise window" in c:
            pyautogui.hotkey("win", "down")
            self.voice.speak("Window minimized.")
            return
        if "maximize window" in c:
            pyautogui.hotkey("win", "up")
            self.voice.speak("Window maximized.")
            return
        if "close current window" in c:
            pyautogui.hotkey("alt", "f4")
            self.voice.speak("Current window closed.")
            return
        if "show desktop" in c:
            pyautogui.hotkey("win", "d")
            self.voice.speak("Desktop shown.")
            return
        if "task manager" in c:
            pyautogui.hotkey("ctrl", "shift", "esc")
            self.voice.speak("Task Manager opened.")
            return
        if "open download" in c or "open downloads" in c:
            os.system("explorer %USERPROFILE%\\Downloads")
            self.voice.speak("Opened Downloads folder.")
            return

        if "create folder" in c:
            match = re.search(r"create folder named (.+)", c)

            if match:
                folder = match.group(1)
                path = os.path.join(os.path.expanduser("~/Desktop"), folder)
                os.makedirs(path, exist_ok=True)
                self.voice.speak(f"Folder {folder} created.")
            else:
                self.voice.speak("Please say a folder name.")
            return
        # trigger deletion flow
        if "delete file" in c:
            # store the target to confirm; extract filename if possible
            filename = c.replace("delete file", "").replace(" dot ", ".").strip()
            if filename:
                self.pending_delete = filename
                self.voice.speak(f"Do you want to delete {filename}? Say yes or no.")
            else:
                self.voice.speak("Please say the file name to delete.")
            return

        # confirm deletion
        if hasattr(self, "pending_delete"):
            if c.strip() == "yes":
                desktop = os.path.expanduser("~/Desktop")
                filepath = os.path.join(desktop, self.pending_delete)
                try:
                    if os.path.exists(filepath):
                        os.remove(filepath)
                        self.voice.speak("File deleted.")
                    else:
                        self.voice.speak("File not found on Desktop.")
                except Exception as e:
                    self.voice.speak("Couldn't delete the file.")
                del self.pending_delete
                return
            if c.strip() == "no":
                self.voice.speak("Deletion cancelled.")
                del self.pending_delete
                return

        if "search file" in c:
            query = c.replace("search file", "").strip()
            os.system(f'explorer /search,"{query}"')
            self.voice.speak(f"Searching for files named {query}.")
            return
        
        
        if "scroll down" in c:
            pyautogui.scroll(-500)
            self.voice.speak("Scrolling down.")
            return

        if "scroll up" in c:
            pyautogui.scroll(500)
            self.voice.speak("Scrolling up.")
            return

        if "refresh page" in c:
            pyautogui.press("f5")
            self.voice.speak("Page refreshed.")
            return
        if "open new tab" in c:
            # try to focus browser; prefer chrome then edge
            if not self.system.apps.switch_to_app("chrome"):
                self.system.apps.switch_to_app("edge")
            time.sleep(0.25)
            pyautogui.hotkey("ctrl", "t")
            self.voice.speak("New tab opened.")
            return

        if "close" in c and "tab" in c:
            pyautogui.hotkey("ctrl", "w")
            self.voice.speak("Tab closed.")
            return
        if "change output" in c or "headphones" in c:
            os.system("start ms-settings:sound")
            self.voice.speak("Please select headphones from sound settings.")
            return
        if "mute notifications" in c:
            os.system("start ms-settings:notifications")
            self.voice.speak("Notification settings opened.")
            return
        if "mute microphone" in c:
            pyautogui.hotkey("win", "alt", "k")
            self.voice.speak("Microphone toggled.")
            return
        if "do not disturb" in c:
            os.system("start ms-settings:quietmoments")
            self.voice.speak("Do Not Disturb enabled.")
            return

        if "lock my laptop" in c:
            pyautogui.hotkey("win", "l")
            return
        
        if "log out" in c:
            os.system("shutdown -l")
            return
        if "take screenshot" in c:
            pyautogui.hotkey("win", "prtsc")
            self.voice.speak("Screenshot taken.")
            return

        if "open camera" in c:
            os.system("explorer.exe shell:AppsFolder\\Microsoft.WindowsCamera_8wekyb3d8bbwe!App")
            self.voice.speak("Camera opened.")
            return

        if "cpu usage" in c:
            cpu = psutil.cpu_percent()
            ram = psutil.virtual_memory().percent
            self.voice.speak(f"CPU usage is {cpu} percent. RAM usage is {ram} percent.")
            return
        if "check storage" in c:
            disk = psutil.disk_usage('/')
            self.voice.speak(f"Storage used {disk.percent} percent.")
            return
        if "open terminal" in c:
            os.system("start cmd")
            self.voice.speak("Terminal opened.")
            return

        if "run my code" in c:
            title = self.system.apps.get_active_window_title()

            if "visual studio code" in title or "vs code" in title:
                pyautogui.hotkey("ctrl", "shift", "b")
                self.voice.speak("Triggered build or run in VS Code.")
                return

            if "command prompt" in title or "powershell" in title:
                self.voice.speak("You are already in terminal. Please run your command.")
                return

            if "chrome" in title or "edge" in title or "google" in title:
                self.voice.speak("Running code in browser is not supported.")
                return

            self.voice.speak("I don't know how to run code in this application.")
            return

        
        


        
        # Spotify flow
        if "search" in c and "spotify" in c:
            try:
                q = c.split("search")[-1].replace("spotify","").strip()
                if not q:
                    self.voice.speak("Please tell me what to search for in Spotify.")
                    return
                ok = self.system.apps.search_spotify(q)
                self.voice.speak("Search done. Say 'pink select number N' to move to result N, 'pink play N' to play N, or 'pink play' to play the top result." if ok else "Couldn't perform Spotify search.")
                return
            except Exception:
                self.voice.speak("Couldn't parse your Spotify search command.")
                return
        if any(w in c for w in ["stop", "hold"]):
            self.system.apps.spotify_play_pause()
            self.voice.speak("Music paused.")
            return

        # "select N" or "select number N"
        if "select" in c and any(ch.isdigit() for ch in c) :
            try:
                # allow "select 2" or "select number 2"
                ntext = re.search(r'(\d+)', c)
                if not ntext:
                    n = _extract_number_from_text(c.split("select")[-1])
                else:
                    n = int(ntext.group(1))
                if not n:
                    self.voice.speak("Please say a valid number after select.")
                    return
                ok = self.system.apps.select_result(n)
                self.voice.speak(f"Selected result {n}." if ok else "Couldn't select that result.")
            except Exception:
                self.voice.speak("Please say a valid number after select.")
            return

        # play nth or plain play
        if re.search(r'\bplay\b', c):
            # 1️⃣ If Spotify search just happened → play Spotify
            if self.system.apps.last_search_app == "spotify":
                ok = self.system.apps.play_first_result()
                self.voice.speak(
                    "Playing the top Spotify result."
                    if ok else
                    "I couldn't play the song. Try selecting a number."
                )
                return

            # 2️⃣ Otherwise fallback
            self.voice.speak("Nothing to play right now.")
            return


        # pause / resume / playpause
        if any(w in c for w in ["pause", "resume", "playpause", "play/pause", "pause song"]):
            ok = self.system.apps.spotify_play_pause()
            self.voice.speak("Toggled play/pause." if ok else "Couldn't toggle play/pause.")
            return

        # next / previous
        if "next" in c:
            ok = self.system.apps.spotify_next()
            self.voice.speak("Skipped to next track." if ok else "Couldn't skip to next track.")
            return
        if "previous" in c or "back" in c:
            ok = self.system.apps.spotify_previous()
            self.voice.speak("Went to previous track." if ok else "Couldn't go to previous track.")
            return

        # spotify volume controls
        if "spotify" in c and any(w in c for w in ["volume", "louder", "quieter", "mute", "up", "down"]):
            if "up" in c or "louder" in c or "increase" in c:
                # try to extract amount
                num = _extract_number_from_text(c)
                steps = max(1, (num // 2) if num else 3)
                ok = self.system.apps.spotify_volume_up(steps)
                self.voice.speak("Increased Spotify volume." if ok else "Couldn't change Spotify volume.")
                return
            if "down" in c or "lower" in c or "decrease" in c or "quieter" in c:
                num = _extract_number_from_text(c)
                steps = max(1, (num // 2) if num else 3)
                ok = self.system.apps.spotify_volume_down(steps)
                self.voice.speak("Decreased Spotify volume." if ok else "Couldn't change Spotify volume.")
                return
            if "mute" in c:
                ok = self.system.apps.spotify_mute_toggle()
                self.voice.speak("Toggled mute." if ok else "Couldn't toggle mute.")
                return

        # shuffle toggle
        if "shuffle" in c:
            ok = self.system.apps.spotify_toggle_shuffle()
            self.voice.speak("Toggled shuffle." if ok else "Couldn't toggle shuffle.")
            return

        # like/unlike
        if "like" in c or "save" in c or "heart" in c:
            ok = self.system.apps.spotify_like_unlike()
            self.voice.speak("Toggled like on current track." if ok else "Couldn't like the track.")
            return

        # open spotify
        if "open spotify" in c:
            ok = self.system.apps.open_app("spotify")
            self.voice.speak("Spotify opened." if ok else "Couldn't open Spotify.")
            return

        
        if c.startswith("close "):
            app = c.split("close",1)[1].strip()
            ok = self.system.apps.close_app(app)
            self.voice.speak(f"Closed {app}." if ok else f"Couldn't close {app}.")
            return

        if "shutdown" in c or "sleep" in c:
            self.voice.speak("Shutting down. Goodbye.")
            if sys.platform == "win32":
                os.system("shutdown /s /t 5")
            return

        
        if "what was i doing" in c:
            activity = activity_to_text()

            if activity:
                latest = activity[0]
                past = activity[1]

                if latest == past:
                    msg = f"You were working on {latest} recently."
                else:
                    msg = f"You were recently working on {latest}, and about 10 minutes ago you were working on {past}"

                # 🔥 FORCE UI ALWAYS
                self.ui.show(msg)
                self.ui_active = True

                self.voice.speak(msg)
            else:
                self.voice.speak("No activity found.")

            return
        
        
        n
        
        # ---------- RAG BLOCK ----------
        context = self.rag.retrieve(c)

        # add activity memory
        context = []

        if "what was i doing" in c:
            context = activity_to_text()
        else:
            context = self.rag.retrieve(c)

        print("RAG context:", context)

        if context:
            rag_prompt = f"""
        You MUST answer ONLY using the given context.

        Context:
        {chr(10).join(context)}

        Question:
        {c}

        Answer:
        """

            response = ollama.chat(
                model="llama3:8b",
                messages=[
                    {"role": "system", "content": "You are a precise assistant."},
                    {"role": "user", "content": rag_prompt}
                ]
            )

            answer = response["message"]["content"]
            self.voice.speak(answer)
            return
        # ---------- END RAG ----------
        
        
        
        generated = self.ask_llm(c)
        print("LLM Generated:", generated)

        if not generated:
            self.voice.speak("I couldn't understand that.")
            return

        # Prevent infinite loop if LLM returns same text
        if generated.strip() == c.strip():
            self.voice.speak("I couldn't understand that.")
            return

        # Add wake word back and execute translated command
        translated_command = CONFIG["wake_word"] + " " + generated
        self.parse_and_execute(translated_command)
        return

        action = data.get("action")

        # ===== ACTION HANDLING =====
        if action == "open_web":
            url = data.get("url")
            if url:
                os.system(f"start {url}")
                self.voice.speak("Opening website.")
            return
        
        if action == "open_app":
            target = data.get("target")
            ok = self.system.apps.open_app(target)
            self.voice.speak(f"Opened {target}." if ok else f"Couldn't open {target}.")
            return

        if action == "close_app":
            target = data.get("target")
            ok = self.system.apps.close_app(target)
            self.voice.speak(f"Closed {target}." if ok else f"Couldn't close {target}.")
            return

        if action == "check_battery":
            self.voice.speak(self.system.check_battery())
            return

        if action == "get_time":
            self.voice.speak(f"The time is {self.system.get_time()}")
            return

        if action == "search_youtube":
            query = data.get("query")
            self.voice.speak(self.system.open_youtube(query))
            return

        if action == "search_spotify":
            query = data.get("query")
            self.system.apps.search_spotify(query)
            self.system.apps.play_first_result()
            self.voice.speak("Playing on Spotify.")
            return

        if action == "shutdown":
            self.voice.speak("Shutting down.")
            os.system("shutdown /s /t 5")
            return

        if action == "chat":
            self.voice.speak(data.get("response"))
            return

        # fallback
        self.voice.speak("I couldn't understand the command.")
        
        
        
        
        

        print("LLM Generated:", generated)

        if generated.startswith("CHAT:"):
            self.voice.speak(generated.replace("CHAT:", "").strip())
        else:
            # Recursively execute generated command
            self.parse_and_execute(generated)

    def ask_llm(self, prompt):
        response = ollama.chat(
            model="llama3:8b",
            messages=[
                {
                    "role": "system",
                    "content": """
    You are a command translator.

    Your job:
    Convert the user's sentence into a valid Pink Assistant command.

    Rules:
    - Return ONLY a command.
    - No JSON.
    - No explanation.
    - No notes.
    - No extra text.

    Examples:

    User: check my emails
    Output: open gmail

    User: any new mails
    Output: open gmail

    User: play faded song
    Output: search spotify faded

    User: open telegram
    Output: open telegram

    User: what time is it
    Output: time
    """
                },
                {
                    "role": "user",
                    "content": prompt
                }
            ]
        )

        return response["message"]["content"].strip().lower()
    
    def run(self):
        print("🎤 Listening...")

        while True:
            try:
                text = self.voice.listen()

                if not text:
                    continue

                text = (text or "").lower().strip()
                print("Heard:", text)

                # 🔥 IGNORE NOISE
                if len(text.split()) <= 1:
                    print("Ignored noise:", text)
                    continue

                # 🔥 INTERRUPT (SPACE)
                if keyboard.is_pressed("space"):
                    print("INTERRUPT TRIGGERED")
                    self.voice.stop()
                    time.sleep(0.3)
                    continue

                # 🔥 GLOBAL CLOSE UI
                if any(word in text for word in ["ok", "okay", "close", "fine", "good", "nice"]):
                    if self.ui_active:
                        print("FORCED UI CLOSE")
                        if hasattr(self, "ui") and self.ui:
                            self.ui.close()
                        self.ui_active = False
                        self.voice.speak("Alright.")
                        continue

                # 🔥 UI MODE
                if self.ui_active:
                    self.parse_and_execute(text)
                    continue

                # 🔥 NORMAL MODE
                if CONFIG["wake_word"] in text:
                    self.parse_and_execute(text)
                else:
                    print("No wake word detected")

            except Exception as e:
                print("Loop error:", e)

# ========== Start ==========
if __name__ == "__main__":
    missing = []
    try:
        import speech_recognition
    except Exception:
        missing.append("SpeechRecognition")
    try:
        import pyttsx3
    except Exception:
        missing.append("pyttsx3")
    try:
        import psutil
    except Exception:
        missing.append("psutil")
    try:
        import screen_brightness_control
    except Exception:
        missing.append("screen-brightness-control")
    try:
        import pyautogui
    except Exception:
        missing.append("pyautogui")

    if missing:
        print("Missing packages detected:", missing)
        print("Run run_pink.bat to auto-install required packages, or install them manually via pip.")
    print("Starting Pink Assistant...")
    assistant = PinkAssistant()
    assistant.run()