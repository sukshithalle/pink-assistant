"""
Pink Assistant — full Spotify controls added (list-select + top-play update) + YouTube commands

Usage:
- Double-click run_pink.bat (installs dependencies then runs this file)
- Say "pink" as the wake word before commands:
    - "pink search spotify faded"
    - "pink select 2"
    - "pink play"
    - "pink play 3"
    - "pink hold"            # <-- changed from "pause"
    - "pink stop"            # <-- changed from "pause"
    - "pink next"
    - "pink previous"
    - "pink spotify volume up"
    - "pink spotify volume down"
    - "pink shuffle"
    - "pink like"
    - "pink play faded on youtube"
    - "pink search faded on youtube"
    - "pink open youtube"
- Works on Windows. Relies on pyautogui and (optionally) pygetwindow for window positioning.
"""

import os
import sys
import time
import re
import subprocess
import webbrowser
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
    "mustang_sound": None,     # optional: put a .wav/.mp3 file path to play at boot
    "app_paths": {
        "chrome": "chrome.exe",
        "whatsapp": "whatsapp.exe",
        "edge": "msedge.exe",
        "spotify": "spotify.exe",
        "vscode": "code.exe",
        "cmd": "cmd.exe",
        "notepad": "notepad.exe"
    }
}

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

# ========== Voice Engine ==========
class VoiceEngine:
    def __init__(self):
        self.recognizer = sr.Recognizer()
        self.mic = sr.Microphone()
        self.engine = None
        self.last_command_time = 0
        self._init_tts()

    def _init_tts(self):
        try:
            self.engine = pyttsx3.init(driverName='sapi5')
        except Exception:
            try:
                self.engine = pyttsx3.init()
            except Exception:
                self.engine = None
        if self.engine:
            try:
                self.engine.setProperty('rate', 170)
                self.engine.setProperty('volume', 1.0)
            except Exception:
                pass

    def speak(self, text):
        if not text:
            return
        print(f"PINK: {text}")
        try:
            if self.engine:
                self.engine.say(text)
                self.engine.runAndWait()
                return
        except Exception:
            pass
        try:
            if sys.platform == "win32":
                os.system(f'PowerShell -Command "Add-Type -AssemblyName System.Speech; (New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak(\'{text}\');"')
            else:
                os.system(f'echo "{text}"')
        except Exception:
            pass

    def listen(self, timeout=6, phrase_time_limit=6):
        with self.mic as source:
            try:
                self.recognizer.adjust_for_ambient_noise(source, duration=0.7)
            except Exception:
                pass
            print("Listening...")
            try:
                audio = self.recognizer.listen(source, timeout=timeout, phrase_time_limit=phrase_time_limit)
            except sr.WaitTimeoutError:
                return ""
            try:
                text = self.recognizer.recognize_google(audio)
                text = text.lower()
                print("User said:", text)
                self.last_command_time = time.time()
                return text
            except sr.UnknownValueError:
                return "unrecognized"
            except sr.RequestError as e:
                print("Speech API request error:", e)
                return "unrecognized"
            except Exception as e:
                print("Listen error:", e)
                return "unrecognized"

# ========== App Controller ==========
class AppController:
    def __init__(self):
        self.search_results = []     # list of strings (query or placeholder titles)
        self.result_positions = []   # list of (x,y) coords for each result row
        self.spotify_window = None
        self.top_play_pos = None     # (x,y) for green play in Top result
        self.last_selected = None    # index of selected list item (1-based)

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
        if name in CONFIG['app_paths']:
            try:
                os.system(f"start {CONFIG['app_paths'][name]}")
                time.sleep(1.5)
                return True
            except Exception as e:
                print("Open app error:", e)
                return False
        else:
            try:
                os.system(f"start {name}")
                time.sleep(1.5)
                return True
            except Exception as e:
                print("Open app fallback error:", e)
                return False

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
            start_y = top + int(height * 0.22)
            start_x = left + int(width * 0.47)  # X moved right to match the song list column
            gap = max(40, int(height * 0.06))   # vertical gap between rows
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
            if self.top_play_pos:
                px, py = self.top_play_pos
                pyautogui.moveTo(px, py, duration=0.2)
                pyautogui.click()
                print(f"Clicked top-play button at ({px},{py}).")
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

# ========== Main Assistant ==========
class PinkAssistant:
    def __init__(self):
        self.voice = VoiceEngine()
        self.system = SystemController(self.voice)
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

    def parse_and_execute(self, command):
        c = (command or "").lower()
        if not c or c == "unrecognized":
            return
        if CONFIG["wake_word"] in c:
            c = c.replace(CONFIG["wake_word"], "").strip()

        # ----------------- YouTube commands (added) -----------------
        # play <query> on youtube  OR play <query> youtube
        m = re.search(r'play\s+(.+?)\s+(?:on\s+)?youtube\b', c)
        if m:
            query = m.group(1).strip()
            query_url = query.replace(" ", "+")
            self.voice.speak(f"Searching YouTube for {query}")
            webbrowser.open(f"https://www.youtube.com/results?search_query={query_url}")
            return

        # search <query> on youtube
        m = re.search(r'search\s+(.+?)\s+(?:on\s+)?youtube\b', c)
        if m:
            query = m.group(1).strip()
            query_url = query.replace(" ", "+")
            self.voice.speak(f"Searching YouTube for {query}")
            webbrowser.open(f"https://www.youtube.com/results?search_query={query_url}")
            return

        # open youtube
        if "open youtube" in c or c.strip() == "youtube":
            self.voice.speak("Opening YouTube")
            webbrowser.open("https://www.youtube.com")
            return
        # ------------------------------------------------------------

        if "battery" in c or "charge" in c:
            self.voice.speak(self.system.check_battery()); return
        if "time" in c:
            t = self.system.get_time()
            self.voice.speak(f"The time is {t}"); return
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
            # check for "play N"
            m = re.search(r'play\s+(\d+)', c)
            if m:
                n = int(m.group(1))
                ok = self.system.apps.play_nth_result(n)
                self.voice.speak(f"Playing result {n}." if ok else f"Couldn't play result {n}.")
                return
            # otherwise play first / selected
            ok = self.system.apps.play_first_result()
            self.voice.speak("Playing result." if ok else "Couldn't play the song. Try 'pink search spotify <song>' first.")
            return

        # hold / stop / resume / playpause
        if any(w in c for w in ["hold", "stop", "resume", "playpause", "play/pause"]):
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

        # generic app open / close
        if c.startswith("open "):
            app = c.split("open",1)[1].strip()
            ok = self.system.apps.open_app(app)
            self.voice.speak(f"Opened {app}." if ok else f"Couldn't open {app}.")
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

        self.voice.speak("I didn't understand that command.")

    def run(self):
        print(f"Pink Assistant running. Say the wake word exactly: '{CONFIG['wake_word']}' before your command.")
        while True:
            text = self.voice.listen()
            if not text:
                time.sleep(0.4)
                continue
            if CONFIG["wake_word"] in text:
                self.parse_and_execute(text)
            else:
                print("No wake word detected; ignoring.")
            time.sleep(0.2)

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
    #print hello