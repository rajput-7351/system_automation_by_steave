#function written by me search_target_element and press_keyboard write_script.


# Configuration
SPOTIFY_LAUNCH_COMMAND = "spotify"
DELAY_AFTER_LAUNCH = 8
SEARCH_DELAY = 3
CONFIRMATION_TIMEOUT = 10  # Seconds to wait for confirmation


import speech_recognition as sr
import pyttsx3
import os
import time
import platform
import pyautogui
import subprocess
import psutil
import shutil
import winshell
from pathlib import Path
import json
import win32gui
import win32con
import win32api
import win32process
import ctypes
from ctypes import wintypes
import re
from difflib import SequenceMatcher


import sys
import pyautogui
import time
import threading
import speech_recognition as sr
import pyttsx3
from PyQt5 import QtWidgets, QtGui, QtCore
import time


# Initialize text-to-speech engine
engine = pyttsx3.init()

# Configure pyautogui for safety
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5

# Common Windows paths
DESKTOP_PATH = Path.home() / "Desktop"
DOCUMENTS_PATH = Path.home() / "Documents"
DOWNLOADS_PATH = Path.home() / "Downloads"
MUSIC_PATH = Path.home() / "Music"
PICTURES_PATH = Path.home() / "Pictures"
VIDEOS_PATH = Path.home() / "Videos"

# Current directory tracking
current_directory = Path.home()

def speak(text):
    """Make the assistant speak"""
    print(f"Assistant: {text}")
    engine.say(text)
    engine.runAndWait()

def listen():
    """Listen for user input through microphone"""
    recognizer = sr.Recognizer()
    
    with sr.Microphone() as source:
        # start=time.time()
        print("Listening_regular...")
        recognizer.adjust_for_ambient_noise(source, duration=1)
        audio = recognizer.listen(source, timeout=5, phrase_time_limit=10)
        # end=time.time()-start
        # print(f"Listening time: {end} seconds")

    try:
        # startr=time.time()
        print("Recognizing...")
        # print("listening next...")
        query = recognizer.recognize_google(audio).lower()
        print(f"User said: {query}")
        # endr=time.time()-startr
        # print(f"Recognition time: {endr} seconds")
        return query
    except sr.UnknownValueError:
        print("Sorry, I didn't catch that. Could you repeat?")
        return ""
    except sr.RequestError as e:
        print(f"Could not request results; {e}")
        return ""
    except Exception as e:
        print(f"Error: {e}")
        return ""
    



# -------------------------- PyQt5 Transparent Grid --------------------------- #
class TransparentGridOverlay(QtWidgets.QWidget):
    def __init__(self, grid_size=30):
        super().__init__()
        self.grid_size = grid_size
        self.screen_width, self.screen_height = pyautogui.size()
        self.cell_width = self.screen_width // grid_size
        self.cell_height = self.screen_height // grid_size

        # Transparent fullscreen window
        self.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint | QtCore.Qt.FramelessWindowHint)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.setGeometry(0, 0, self.screen_width, self.screen_height)
        self.is_visible = False

    def paintEvent(self, event):
        qp = QtGui.QPainter(self)
        pen = QtGui.QPen(QtCore.Qt.white, 1, QtCore.Qt.SolidLine)
        qp.setPen(pen)

        for row in range(self.grid_size):
            for col in range(self.grid_size):
                x1 = col * self.cell_width
                y1 = row * self.cell_height
                rect = QtCore.QRect(x1, y1, self.cell_width, self.cell_height)
                qp.drawRect(rect)

                # Draw RGB sections
                section_width = self.cell_width // 3
                qp.fillRect(x1, y1, section_width, self.cell_height, QtGui.QColor(255, 0, 0, 25))
                qp.fillRect(x1 + section_width, y1, section_width, self.cell_height, QtGui.QColor(0, 255, 0, 25))
                qp.fillRect(x1 + 2 * section_width, y1, section_width, self.cell_height, QtGui.QColor(0, 0, 255, 25))

                # Add cell coordinates
                if self.cell_width > 20:
                    qp.setPen(QtGui.QColor(255, 255, 255))
                    if row ==0 or col == 0:  # Only show text for first row and column
                        if row ==0:
                            qp.drawText(x1 + 5, y1 + 15, f"{col}")
                        elif col == 0:
                            qp.drawText(x1 + 5, y1 + 15, f"{row}")
                        # qp.drawText(x1 + 5, y1 + 15, f"{row},{col}")

    def show_grid(self):
        self.showFullScreen()
        self.is_visible = True

    def hide_grid(self):
        self.hide()
        self.is_visible = False


# ----------------------- Voice Assistant and Automation ---------------------- #
class ScreenAutomation:
    def __init__(self):
        self.app = QtWidgets.QApplication(sys.argv)
        self.grid_overlay = TransparentGridOverlay()
        self.engine = pyttsx3.init()
        self.saved_coordinates = {}

    def speak(self, text):
        print(f"Assistant: {text}")
        self.engine.say(text)
        self.engine.runAndWait()

    def listen(self):
        recognizer = sr.Recognizer()
        with sr.Microphone() as source:
            print("Listening...")
            recognizer.adjust_for_ambient_noise(source, duration=1)
            try:
                audio = recognizer.listen(source, timeout=5, phrase_time_limit=8)
                command = recognizer.recognize_google(audio).lower()
                print(f"User said: {command}")
                return command
            except sr.WaitTimeoutError:
                print("Listening timed out.")
            except sr.UnknownValueError:
                print("Sorry, I didn’t catch that.")
            except Exception as e:
                print(f"Error: {e}")
            return ""

    def click_grid_block(self, row, col, color):
        """Click on grid block's color section"""
        if row >= self.grid_overlay.grid_size or col >= self.grid_overlay.grid_size:
            self.speak("Invalid grid coordinates")
            return

        x = col * self.grid_overlay.cell_width
        y = row * self.grid_overlay.cell_height
        section_width = self.grid_overlay.cell_width // 3

        if color == "red":
            x += section_width // 2
        elif color == "green":
            x += section_width + section_width // 2
        elif color == "blue":
            x += 2 * section_width + section_width // 2
        else:
            self.speak("Invalid color. Use red, green, or blue.")
            return

        y += self.grid_overlay.cell_height // 2
        pyautogui.click(x, y)
        print(f"Clicked on {row},{col} {color} section at ({x},{y})")
        self.speak(f"Clicked on block {row}, {col}, {color}")

    def process_voice_command(self, command):
        # command=command.instring()

        words = command.split()
        if "open grid" in command:
            self.grid_overlay.show_grid()
            self.speak("Grid overlay opened")

        elif "close grid" in command:
            self.grid_overlay.hide_grid()
            self.speak("Grid overlay closed")
    #   example like click block 29 slash 15 green
        elif "click block" in command:
            try:
                parts = command.split()
                row, col, color = int(parts[2]), int(parts[4]), parts[5]
                self.click_grid_block(row, col, color)
            except Exception:
                self.speak("Invalid format. Say click block row col color")

        elif "exit" in command or "quit" in command:
            self.speak("Goodbye!")
            sys.exit()

    
    

automation = ScreenAutomation()
# search functionaliy to find file folder any thing in this pc or my pc
# we call this fun. when in my voice have search target element X.....

def search_target_element(target_element):
    """Search for target element using Ctrl+F"""
    try:
        pyautogui.hotkey('ctrl', 'f')
        time.sleep(0.5)
        pyautogui.hotkey('ctrl', 'a')   
        pyautogui.press('backspace')  # Clear any previous search
        pyautogui.write(target_element)
        time.sleep(3)
        pyautogui.press('enter')
        time.sleep(0.5)
        pyautogui.press('right') 
        time.sleep(0.5)
        pyautogui.press('left')  
        time.sleep(0.5)
        pyautogui.press('enter')  
        # pyautogui.press('esc')  # Close search dialog
        pyautogui.press('escape')  # Close search dialog
        speak(f"Searched for {target_element}")
        return True
    except Exception as e:
        print(f"Error in search_target_element: {e}")
        return False

def press_keyboard(key_combination):
    """Press keyboard keys or key combinations"""
    try:
        key_combination = key_combination.strip()
        
        if not key_combination:
            print("No key combination provided")
            return False
            
        # Split the key combination
        keys = key_combination.split()
        
        if len(keys) == 1:
            # Single key press
            pyautogui.press(keys[0])
            speak(f'Pressed {keys[0]}')
        elif len(keys) == 2:
            # Two key combination (e.g., ctrl a)
            pyautogui.hotkey(keys[0], keys[1])
            speak(f'Pressed {keys[0]} {keys[1]}')
        elif len(keys) == 3:
            # Three key combination (e.g., ctrl shift a)
            pyautogui.hotkey(keys[0], keys[1], keys[2])
            speak(f'Pressed {keys[0]} {keys[1]} {keys[2]}')
        else:
            # More than 3 keys - handle as hotkey sequence
            pyautogui.hotkey(*keys)
            speak(f'Pressed {" ".join(keys)}')
            
        return True
    except Exception as e:
        print(f"Error in press_keyboard: {e}")
        return False

def write_script(script):
    """Write script text"""
    try:
        print('Writing...')
        pyautogui.write(script)
        print('Done')
        return True
    except Exception as e:
        print(f"Error writing script: {e}")
        return False

def enter_write_script_mode():
    """Enter script writing mode with voice commands"""
    print('Entered script mode')
    speak('Entered script mode. Say "exit script mode" to exit.')
    
    while True:
        try:
            print('Listening_script....')
            task = listen()

            if not task:  # Handle empty input
                continue
                
            task = task.lower().strip()
            
            # Exit condition
            if 'exit' in task and 'script' in task and 'mode' in task:
                break
            
            # Handle specific press commands BEFORE general keyboard command
            elif 'press' in task and 'enter' in task:
                pyautogui.press('enter')
                speak('Pressed enter')
            elif 'press' in task and 'tab' in task:
                pyautogui.press('tab')
                speak('Pressed tab')
            elif 'new' in task and 'line' in task:
                pyautogui.press('enter')
                speak('New line')
            elif 'press' in task and 'space' in task:
                pyautogui.press('space')
                speak('Pressed space')
            elif 'press' in task and 'backspace' in task:
                pyautogui.press('backspace')
                speak('Pressed backspace')
            elif 'press' in task and 'delete' in task:
                pyautogui.press('delete')
                speak('Pressed delete')
            
            # General keyboard command (for combinations like ctrl+a, ctrl+c, etc.)
            elif 'press' in task and 'keyboard' in task:
                try:
                    key = task.split("keyboard")[-1].strip()
                    if key:
                        press_keyboard(key)
                    else:
                        print("No key specified after 'keyboard'")
                except Exception as e:
                    print(f"Error pressing key in script mode: {e}")
            
            # Write the text
            else:
                write_script(task)
                
        except Exception as e:
            print(f"Error in script mode: {e}")
            
    speak('Exited script mode')
    return True



# Additional utility function for confidence-based confirmation
def confirm_action_with_speech(self,action_description, heard_command):
    """Confirm action when speech recognition confidence is low"""
    speak(f"I heard '{heard_command}'. Do you want to {action_description}?")
    
    confirmation_responses = ['yes', 'yeah', 'yep', 'confirm', 'correct', 'right', 'true']
    denial_responses = ['no', 'nope', 'cancel', 'wrong', 'false', 'stop']
    
    response = self.listen()
    normalized_response = normalize_speech_input(response)
    
    for confirm_word in confirmation_responses:
        if confirm_word in normalized_response:
            return True
    
    for deny_word in denial_responses:
        if deny_word in normalized_response:
            return False
    
    # If unclear, ask again
    speak("I didn't understand. Please say yes or no.")
    return confirm_action_with_speech(action_description, heard_command)



def normalize_speech_input(text):
    """Normalize speech input to handle common misrecognitions"""
    # Convert to lowercase
    text = text.lower()
    
    # Common drive letter corrections
    drive_corrections = {
        'sea': 'c', 'see': 'c', 'tea': 'c', 'se': 'c', 'si': 'c',
        'bee': 'b', 'be': 'b', 'beach': 'b',
        'dee': 'd', 'de': 'd', 'the': 'd',
        'eat': 'e', 'he': 'e', 'easy': 'e',
        'if': 'f', 'half': 'f', 'ef': 'f',
        'g': 'g', 'gee': 'g', 'gene': 'g',
        'h': 'h', 'age': 'h', 'each': 'h',
        'i': 'i', 'eye': 'i', 'ay': 'i',
        'j': 'j', 'jay': 'j', 'jane': 'j',
        'k': 'k', 'kay': 'k', 'okay': 'k',
        'l': 'l', 'el': 'l', 'hell': 'l',
        'm': 'm', 'em': 'm', 'him': 'm',
        'n': 'n', 'en': 'n', 'and': 'n',
        'o': 'o', 'oh': 'o', 'owe': 'o',
        'p': 'p', 'pee': 'p', 'peace': 'p',
        'q': 'q', 'queue': 'q', 'cue': 'q',
        'r': 'r', 'are': 'r', 'our': 'r',
        's': 's', 'es': 's', 'yes': 's',
        't': 't', 'tea': 't', 'tee': 't',
        'u': 'u', 'you': 'u', 'new': 'u',
        'v': 'v', 'vee': 'v', 'we': 'v',
        'w': 'w', 'double': 'w', 'you': 'w',
        'x': 'x', 'ex': 'x', 'eggs': 'x',
        'y': 'y', 'why': 'y', 'wine': 'y',
        'z': 'z', 'zee': 'z', 'easy': 'z'
    }
    
    # Application name corrections
    app_corrections = {
        'not bad': 'notepad', 'note bad': 'notepad', 'no pad': 'notepad',
        'calculator': 'calculator', 'calc': 'calculator', 'calculate': 'calculator',
        'pain': 'paint', 'paint': 'paint',
        'chrome': 'chrome', 'grow': 'chrome', 'room': 'chrome',
        'firefox': 'firefox', 'fire fox': 'firefox', 'fireforks': 'firefox',
        'edge': 'edge', 'age': 'edge', 'hedge': 'edge',
        'word': 'word', 'world': 'word', 'ward': 'word',
        'excel': 'excel', 'ex cell': 'excel', 'axel': 'excel',
        'powerpoint': 'powerpoint', 'power point': 'powerpoint', 'power': 'powerpoint',
        'outlook': 'outlook', 'out look': 'outlook', 'look': 'outlook',
        'spotify': 'spotify', 'spot if i': 'spotify', 'spot': 'spotify',
        'discord': 'discord', 'dis cord': 'discord', 'this cord': 'discord',
        'skype': 'skype', 'sky': 'skype', 'skip': 'skype',
        'teams': 'teams', 'team': 'teams', 'teens': 'teams',
        'cmd': 'cmd', 'command': 'cmd', 'command prompt': 'cmd',
        'powershell': 'powershell', 'power shell': 'powershell', 'shell': 'powershell',
        'task manager': 'task manager', 'task': 'task manager', 'manager': 'task manager',
        'control panel': 'control panel', 'control': 'control panel', 'panel': 'control panel'
    }
    
    # Folder name corrections
    folder_corrections = {
        'desktop': 'desktop', 'desk top': 'desktop', 'desk': 'desktop',
        'documents': 'documents', 'document': 'documents', 'dock': 'documents',
        'downloads': 'downloads', 'download': 'downloads', 'down': 'downloads',
        'music': 'music', 'music': 'music', 'musics': 'music',
        'pictures': 'pictures', 'picture': 'pictures', 'pics': 'pictures',
        'videos': 'videos', 'video': 'videos', 'vids': 'videos'
    }
    
    # Apply corrections
    words = text.split()
    corrected_words = []
    
    for word in words:
        # Check drive corrections
        if word in drive_corrections:
            corrected_words.append(drive_corrections[word])
        # Check app corrections
        elif word in app_corrections:
            corrected_words.append(app_corrections[word])
        # Check folder corrections
        elif word in folder_corrections:
            corrected_words.append(folder_corrections[word])
        else:
            corrected_words.append(word)
    
    return ' '.join(corrected_words)

def fuzzy_match_command(user_input, command_list, threshold=0.6):
    """Find the best matching command using fuzzy matching"""
    best_match = None
    best_ratio = 0
    
    for command in command_list:
        ratio = SequenceMatcher(None, user_input.lower(), command.lower()).ratio()
        if ratio > best_ratio and ratio >= threshold:
            best_ratio = ratio
            best_match = command
    
    return best_match, best_ratio

def extract_drive_letter(text):
    """Extract drive letter from text with better error handling"""
    # Normalize the text first
    text = normalize_speech_input(text)
    
    # Look for drive patterns
    drive_patterns = [
        r'drive\s+([a-z])',  # "drive c"
        r'([a-z])\s+drive',  # "c drive"
        r'([a-z]):\s*$',     # "c:" at end
        r'([a-z])\s*colon',  # "c colon"
    ]
    
    for pattern in drive_patterns:
        match = re.search(pattern, text)
        if match:
            return match.group(1)
    
    # If no pattern found, try to find single letters
    words = text.split()
    for word in words:
        if len(word) == 1 and word.isalpha():
            return word
    
    # Default to C drive if nothing found
    return 'c'


def get_confirmation(self,action_description):
    """Get user confirmation for destructive operations"""
    speak(f"Are you sure you want to {action_description}? Say yes to confirm or no to cancel.")
    
    start_time = time.time()
    while time.time() - start_time < CONFIRMATION_TIMEOUT:
        response = self.listen()
        if response:
            if 'yes' in response or 'confirm' in response or 'proceed' in response:
                return True
            elif 'no' in response or 'cancel' in response or 'stop' in response:
                return False
    
    speak("No confirmation received. Operation cancelled.")
    return False

# =========================== SYSTEM NAVIGATION ===========================

def go_to_desktop():
    """Navigate to desktop"""
    try:
        # Windows key + D to show desktop
        pyautogui.hotkey('win', 'd')
        time.sleep(1)
        speak("Now on desktop")
        return True
    except Exception as e:
        speak(f"Could not go to desktop: {e}")
        return False

def open_this_pc():
    """Open This PC / File Explorer"""
    try:
        # Windows key + E to open File Explorer
        pyautogui.hotkey('win', 'e')
        time.sleep(2)
        speak("Opened This PC")
        return True
    except Exception as e:
        speak(f"Could not open This PC: {e}")
        return False

def open_recycle_bin():
    """Open Recycle Bin"""
    try:
        # Use Windows Run command
        pyautogui.hotkey('win', 'r')
        time.sleep(0.5)
        pyautogui.write('shell:RecycleBinFolder')
        pyautogui.press('enter')
        time.sleep(2)
        speak("Opened Recycle Bin")
        return True
    except Exception as e:
        speak(f"Could not open Recycle Bin: {e}")
        return False

def enhanced_navigate_to_drive(drive_input):
    """Enhanced drive navigation with better speech recognition"""
    try:
        # Extract drive letter with improved recognition
        drive_letter = extract_drive_letter(drive_input)
        
        # Validate drive letter
        if not drive_letter.isalpha() or len(drive_letter) != 1:
            speak("Invalid drive letter. Using C drive as default.")
            drive_letter = 'c'
        
        # Navigate to drive
        pyautogui.hotkey('win', 'r')
        time.sleep(0.5)
        pyautogui.write(f'{drive_letter}:')
        pyautogui.press('enter')
        time.sleep(2)
        speak(f"Opened {drive_letter.upper()} drive")
        return True
        
    except Exception as e:
        speak(f"Could not open drive: {e}")
        return False

def navigate_to_folder(folder_path):
    """Navigate to specific folder"""
    try:
        if folder_path.lower() == 'desktop':
            path = str(DESKTOP_PATH)
        elif folder_path.lower() == 'documents':
            path = str(DOCUMENTS_PATH)
        elif folder_path.lower() == 'downloads':
            path = str(DOWNLOADS_PATH)
        elif folder_path.lower() == 'music':
            path = str(MUSIC_PATH)
        elif folder_path.lower() == 'pictures':
            path = str(PICTURES_PATH)
        elif folder_path.lower() == 'videos':
            path = str(VIDEOS_PATH)
        else:
            path = folder_path
        
        pyautogui.hotkey('win', 'r')
        time.sleep(0.5)
        pyautogui.write(path)
        pyautogui.press('enter')
        time.sleep(2)
        speak(f"Opened {folder_path}")
        return True
    except Exception as e:
        speak(f"Could not open {folder_path}: {e}")
        return False

# =========================== APPLICATION MANAGEMENT ===========================

def enhanced_launch_application(app_input):
    """Enhanced application launching with better speech recognition"""
    try:
        # Normalize the app name
        normalized_input = normalize_speech_input(app_input)
        
        # Common applications with multiple possible names
        app_commands = {
            'notepad': ['notepad', 'not bad', 'note bad', 'no pad'],
            'calculator': ['calc', 'calculator', 'calculate'],
            'paint': ['mspaint', 'paint', 'pain'],
            'chrome': ['chrome', 'grow', 'room'],
            'firefox': ['firefox', 'fire fox', 'fireforks'],
            'edge': ['msedge', 'edge', 'age', 'hedge'],
            'word': ['winword', 'word', 'world', 'ward'],
            'excel': ['excel', 'ex cell', 'axel'],
            'powerpoint': ['powerpnt', 'powerpoint', 'power point', 'power'],
            'outlook': ['outlook', 'out look', 'look'],
            'spotify': ['spotify', 'spot if i', 'spot'],
            'discord': ['discord', 'dis cord', 'this cord'],
            'skype': ['skype', 'sky', 'skip'],
            'teams': ['teams', 'team', 'teens'],
            'cmd': ['cmd', 'command', 'command prompt'],
            'powershell': ['powershell', 'power shell', 'shell'],
            'task manager': ['taskmgr', 'task manager', 'task', 'manager'],
            'control panel': ['control', 'control panel', 'panel']
        }
        
        # Find matching application
        command_to_use = None
        app_name = None
        
        for cmd, variations in app_commands.items():
            for variation in variations:
                if variation in normalized_input:
                    command_to_use = cmd
                    app_name = variation
                    break
            if command_to_use:
                break
        
        # If no exact match, try fuzzy matching
        if not command_to_use:
            all_apps = list(app_commands.keys())
            best_match, confidence = fuzzy_match_command(normalized_input, all_apps)
            if best_match and confidence > 0.6:
                command_to_use = best_match
                app_name = best_match
                speak(f"Did you mean {best_match}? Launching it now.")
        
        # Use original input if no match found
        if not command_to_use:
            command_to_use = normalized_input
            app_name = normalized_input
        
        # Try to launch using subprocess
        try:
            subprocess.Popen([command_to_use])
            time.sleep(2)
            speak(f"Launched {app_name}")
            return True
        except FileNotFoundError:
            # Try using Windows Run command
            pyautogui.hotkey('win', 'r')
            time.sleep(0.5)
            pyautogui.write(command_to_use)
            pyautogui.press('enter')
            time.sleep(2)
            speak(f"Launched {app_name}")
            return True
            
    except Exception as e:
        speak(f"Could not launch application: {e}")
        return False

def close_application(app_name):
    """Close application by name"""
    try:
        # Find and close the application
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if app_name.lower() in proc.info['name'].lower():
                    if get_confirmation(f"close {app_name}"):
                        proc.terminate()
                        speak(f"Closed {app_name}")
                        return True
                    else:
                        speak("Operation cancelled")
                        return False
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        
        speak(f"Could not find {app_name} running")
        return False
        
    except Exception as e:
        speak(f"Could not close {app_name}: {e}")
        return False

# =========================== FILE OPERATIONS ===========================

def create_file(filename, file_type="txt"):
    """Create a new file"""
    try:
        if get_confirmation(f"create file {filename}.{file_type}"):
            file_path = DESKTOP_PATH / f"{filename}.{file_type}"
            
            if file_path.exists():
                speak(f"File {filename}.{file_type} already exists")
                return False
            
            file_path.touch()
            speak(f"Created file {filename}.{file_type} on desktop")
            return True
        else:
            speak("File creation cancelled")
            return False
            
    except Exception as e:
        speak(f"Could not create file: {e}")
        return False

def create_folder(folder_name):
    """Create a new folder"""
    try:
        if get_confirmation(f"create folder {folder_name}"):
            folder_path = DESKTOP_PATH / folder_name
            
            if folder_path.exists():
                speak(f"Folder {folder_name} already exists")
                return False
            
            folder_path.mkdir()
            speak(f"Created folder {folder_name} on desktop")
            return True
        else:
            speak("Folder creation cancelled")
            return False
            
    except Exception as e:
        speak(f"Could not create folder: {e}")
        return False

def delete_file(filename):
    """Delete a file"""
    try:
        # Search for file in common locations
        search_paths = [DESKTOP_PATH, DOCUMENTS_PATH, DOWNLOADS_PATH]
        
        for path in search_paths:
            for file_path in path.glob(f"*{filename}*"):
                if file_path.is_file():
                    if get_confirmation(f"delete file {file_path.name}"):
                        winshell.delete_file(str(file_path))
                        speak(f"Deleted file {file_path.name}")
                        return True
                    else:
                        speak("Delete operation cancelled")
                        return False
        
        speak(f"Could not find file {filename}")
        return False
        
    except Exception as e:
        speak(f"Could not delete file: {e}")
        return False

def delete_folder(folder_name):
    """Delete a folder"""
    try:
        # Search for folder in common locations
        search_paths = [DESKTOP_PATH, DOCUMENTS_PATH, DOWNLOADS_PATH]
        
        for path in search_paths:
            for folder_path in path.glob(f"*{folder_name}*"):
                if folder_path.is_dir():
                    if get_confirmation(f"delete folder {folder_path.name} and all its contents"):
                        shutil.rmtree(str(folder_path))
                        speak(f"Deleted folder {folder_path.name}")
                        return True
                    else:
                        speak("Delete operation cancelled")
                        return False
        
        speak(f"Could not find folder {folder_name}")
        return False
        
    except Exception as e:
        speak(f"Could not delete folder: {e}")
        return False

def open_file(filename):
    """Open a file"""
    try:
        # Search for file in common locations
        search_paths = [DESKTOP_PATH, DOCUMENTS_PATH, DOWNLOADS_PATH]
        
        for path in search_paths:
            for file_path in path.glob(f"*{filename}*"):
                if file_path.is_file():
                    os.startfile(str(file_path))
                    speak(f"Opened file {file_path.name}")
                    return True
        
        speak(f"Could not find file {filename}")
        return False
        
    except Exception as e:
        speak(f"Could not open file: {e}")
        return False

# =========================== SYSTEM CONTROLS ===========================

def adjust_volume(action, level=None):
    """Adjust system volume"""
    try:
        if action == "increase":
            pyautogui.press('volumeup', presses=8)
            speak("Volume increased")
        elif action == "decrease":
            print("reach")
            pyautogui.press('volumedown', presses=5)
            speak("Volume decreased")
        elif action == "mute":
            pyautogui.press('volumemute')
            speak("Volume muted")
        elif action == "unmute":
            pyautogui.press('volumemute')
            speak("Volume unmuted")
        return True
    except Exception as e:
        speak(f"Could not adjust volume: {e}")
        return False

def system_control(action):
    """System control operations"""
    try:
        if action == "shutdown":
            if get_confirmation("shutdown the computer"):
                os.system("shutdown /s /t 1")
                speak("Shutting down computer")
                return True
        elif action == "restart":
            if get_confirmation("restart the computer"):
                os.system("shutdown /r /t 1")
                speak("Restarting computer")
                return True
        elif action == "sleep":
            if get_confirmation("put computer to sleep"):
                os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
                speak("Putting computer to sleep")
                return True
        elif action == "lock":
            pyautogui.hotkey('win', 'l')
            speak("Computer locked")
            return True
        
        speak("Operation cancelled")
        return False
        
    except Exception as e:
        speak(f"Could not perform system control: {e}")
        return False

def window_management(action):
    """Window management operations"""
    try:
        if action == "minimize":
            pyautogui.hotkey('win', 'down')
            speak("Window minimized")
        elif action == "maximize":
            pyautogui.hotkey('win', 'up')
            speak("Window maximized")
        elif action == "switch":
            pyautogui.hotkey('alt', 'tab')
            speak("Switched window")
        elif action == "close":
            pyautogui.hotkey('alt', 'f4')
            speak("Closed window")
        return True
    except Exception as e:
        speak(f"Could not manage window: {e}")
        return False

# =========================== SPOTIFY FUNCTIONS (From Original Code) ===========================

def is_spotify_running():
    """Check if Spotify is already running"""
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if 'spotify' in proc.info['name'].lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return False

def launch_spotify():
    """Launch Spotify application"""
    try:
        system = platform.system()
        
        if system == "Windows":
            try:
                subprocess.Popen([SPOTIFY_LAUNCH_COMMAND])
            except FileNotFoundError:
                os.system("start spotify:")
        elif system == "Darwin":
            os.system("open -a Spotify")
        elif system == "Linux":
            os.system("spotify &")
        
        time.sleep(DELAY_AFTER_LAUNCH)
        return True
    except Exception as e:
        print(f"Error launching Spotify: {e}")
        return False

def focus_spotify_window():
    """Focus on Spotify window"""
    try:
        windows = pyautogui.getWindowsWithTitle("Spotify")
        
        if not windows:
            possible_titles = ["Spotify Premium", "Spotify Free", "Spotify"]
            for title in possible_titles:
                windows = pyautogui.getWindowsWithTitle(title)
                if windows:
                    break
        
        if windows:
            spotify_window = windows[0]
            
            if hasattr(spotify_window, 'restore'):
                spotify_window.restore()
            
            spotify_window.activate()
            time.sleep(1)
            return True
        else:
            print("No Spotify window found")
            return False
            
    except Exception as e:
        print(f"Error focusing Spotify window: {e}")
        return False

def search_and_play_song(song_name):
    """Search for a song and play it using GUI automation"""
    try:
        speak(f"Searching for {song_name}")
        
        pyautogui.hotkey('ctrl', 'l')
        time.sleep(1)
        
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.2)
        pyautogui.write(song_name, interval=0.05)
        time.sleep(0.5)
        
        pyautogui.press('enter')
        time.sleep(SEARCH_DELAY)
        
        success = False
        
        try:
            pyautogui.press('tab', presses=4, interval=0.3)
            time.sleep(0.5)
            
            pyautogui.press('enter')
            time.sleep(2)
            
            # Try to click the first song
            screen_width, screen_height = pyautogui.size()
            song_x = int(screen_width * 0.6)
            song_y = int(screen_height * 0.565)
            
            pyautogui.doubleClick(song_x, song_y)
            time.sleep(1)
            
            success = True
            
        except Exception as e1:
            print(f"Search method failed: {e1}")
            success = False
        
        if success:
            speak(f"Now playing {song_name}")
            return True
        else:
            speak("I found the song but couldn't start playing it")
            return False
        
    except Exception as e:
        print(f"Error in search and play: {e}")
        speak("Sorry, I couldn't play that song")
        return False

def play_song(song_name):
    """Main function to play a song on Spotify"""
    try:
        if not is_spotify_running():
            speak("Spotify is not running. Let me open it for you.")
            if not launch_spotify():
                speak("Failed to launch Spotify")
                return
        
        if not focus_spotify_window():
            speak("Could not focus on Spotify window")
            return
        
        search_and_play_song(song_name)
        
    except Exception as e:
        print(f"Error in play_song: {e}")
        speak("Sorry, I encountered an error while trying to play the song")

# =========================== COMMAND PROCESSING ===========================

def enhanced_process_command(query):
    """Enhanced command processing with better speech recognition"""
    if not query:
        return True
    
    # Normalize the entire query
    normalized_query = normalize_speech_input(query)
    words = normalized_query.split()
    
    # Navigation commands
    if 'go' in words and 'desktop' in words:
        go_to_desktop()
    elif 'open' in words and 'this' in words and 'pc' in words:
        open_this_pc()
    elif 'open' in words and 'recycle' in words and 'bin' in words:
        open_recycle_bin()
    elif 'navigate' in words and 'drive' in words:
        enhanced_navigate_to_drive(normalized_query)
    elif 'navigate' in words and 'folder' in words:
        if len(words) > words.index('folder') + 1:
            folder_name = ' '.join(words[words.index('folder') + 1:])
            navigate_to_folder(folder_name)
        else:
            speak("Please specify folder name")
    
    # Application management with enhanced recognition
    elif 'launch' in words or ('open' in words and 'application' in words):
        if 'application' in words:
            app_name = ' '.join(words[words.index('application') + 1:])
        else:
            app_name = ' '.join(words[words.index('launch') + 1:])
        if app_name:
            enhanced_launch_application(app_name)
        else:
            speak("Please specify application name")
    elif 'close' in words and 'application' in words:
        app_name = ' '.join(words[words.index('application') + 1:])
        if app_name:
            close_application(app_name)
        else:
            speak("Please specify application name")
    
    # File operations
    elif 'create' in words and 'file' in words:
        if len(words) > words.index('file') + 1:
            filename = words[words.index('file') + 1]
            file_type = words[words.index('file') + 2] if len(words) > words.index('file') + 2 else "txt"
            create_file(filename, file_type)
        else:
            speak("Please specify filename")
    elif 'create' in words and 'folder' in words:
        if len(words) > words.index('folder') + 1:
            folder_name = ' '.join(words[words.index('folder') + 1:])
            create_folder(folder_name)
        else:
            speak("Please specify folder name")
    elif 'delete' in words and 'file' in words:
        if len(words) > words.index('file') + 1:
            filename = ' '.join(words[words.index('file') + 1:])
            delete_file(filename)
        else:
            speak("Please specify filename")
    elif 'delete' in words and 'folder' in words:
        if len(words) > words.index('folder') + 1:
            folder_name = ' '.join(words[words.index('folder') + 1:])
            delete_folder(folder_name)
        else:
            speak("Please specify folder name")
    elif 'open' in words and 'file' in words:
        if len(words) > words.index('file') + 1:
            filename = ' '.join(words[words.index('file') + 1:])
            open_file(filename)
        else:
            speak("Please specify filename")
    
    # System controls
    elif 'volume' in words:
        if 'up' in words or 'increase' in words:
            adjust_volume('increase')
        elif 'down' in words or 'decrease' in words:
            adjust_volume('decrease')
        elif 'mute' in words:
            adjust_volume('mute')
        elif 'unmute' in words:
            adjust_volume('unmute')
    elif 'system' in words:
        if 'shutdown' in words:
            system_control('shutdown')
        elif 'restart' in words:
            system_control('restart')
        elif 'sleep' in words:
            system_control('sleep')
        elif 'lock' in words:
            system_control('lock')
    
    # Window management
    elif 'window' in words:
        if 'minimize' in words or 'minimise' in words:
            window_management('minimize')
        elif 'maximize' in words:
            window_management('maximize')
        elif 'switch' in words:
            window_management('switch')
        elif 'close' in words:
            window_management('close')
    
    # Spotify commands
    elif 'play' in words and 'song' in words:
        song_name = ' '.join(words[words.index('song') + 1:])
        if song_name:
            play_song(song_name)
        else:
            speak("Please specify song name")
    elif 'open' in words and 'spotify' in words:
        speak("Opening Spotify")
        if launch_spotify():
            speak("Spotify opened successfully")
        else:
            speak("Failed to open Spotify")
    elif 'pause' in words or 'stop' in words:
        speak("Pausing music")
        pyautogui.press('space')
    elif 'next' in words or 'skip' in words:
        speak("Playing next song")
        pyautogui.hotkey('ctrl', 'right')
    elif 'previous' in words or 'back' in words:
        speak("Playing previous song")
        pyautogui.hotkey('ctrl', 'left')
    
    # Updated parsing logic for the three function calls
    elif 'search' in words and 'target' in words and 'element' in words:
        try:
            # Convert words to string if it's a list
            if isinstance(words, list):
                words_str = ' '.join(words)
            else:
                words_str = words
                
            # Better parsing for target element
            if 'element' in words_str:
                target = words_str.split('element')[-1].strip()
                if target:  # Check if target is not empty
                    search_target_element(target)
                else:
                    print("No target element specified")
        except Exception as e:
            print(f"Error searching target element: {e}")

    # press key
    elif 'press' in words and 'keyboard' in words:
        try:
            # Convert words to string if it's a list
            if isinstance(words, list):
                words_str = ' '.join(words)
            else:
                words_str = words
                
            key_combination = words_str.split('keyboard')[-1].strip()
            if key_combination:  # Check if key combination is not empty
                press_keyboard(key_combination)
            else:
                print("No key combination specified")
        except Exception as e:
            print(f"Error pressing keyboard: {e}")

    # enter write script mode
    elif 'enter' in words and 'script' in words and 'mode' in words:
        try:
            enter_write_script_mode()
        except Exception as e:
            print(f"Error entering script mode: {e}")

    elif "open" in words and "grid" in words:
            if isinstance(words, list):
                words_str = ' '.join(words)
            else:
                words_str = words
            automation.process_voice_command(words_str)
            print("Grid overlay opened")
            # automation.speak("Grid overlay opened")

    elif "close" in words and "grid" in words:
            if isinstance(words, list):
                words_str = ' '.join(words)
            else:
                words_str = words
            automation.process_voice_command(words_str)
            print("Grid overlay closed")
            # automation.speak("Grid overlay closed")
            # example like click block 29 slash 15 green
    elif "click" in words and "block" in words:
            if isinstance(words, list):
                    words_str = ' '.join(words)
            else:
                words_str = words
            automation.process_voice_command(words_str)
         

    # Exit command
    elif 'exit' in words or 'quit' in words or 'goodbye' in words:
        speak("Goodbye!")
        return False
    

    else:
        print("Sorry, I didn't understand that command")
    
    return True


def main():
    """Main function to run the voice assistant"""
    # speak("Hello! papa I'm your son. works as Windows voice assistant. Here are some commands you can use:")
    # speak("Navigation: go desktop, open this pc, open recycle bin, navigate drive c, navigate folder documents")
    # speak("Applications: launch notepad, open application chrome, close application spotify")
    # speak("Files: create file test, create folder new folder, delete file test, open file document")
    # speak("System: volume up, volume down, system shutdown, system lock")
    # speak("Window: window minimize, window maximize, window switch, window close")
    # speak("Spotify: play song despacito, open spotify, pause, next, previous")
    # speak("some other: search target element, pess key X..., enter script mode")
    # speak("Say exit to quit")
    
    
    # Run PyQt app in a separate thread
    # qt_thread = threading.Thread(target=automation.app.exec_)
    # qt_thread.start()
    # Run voice automation in main thread
    # automation.run_automation()

    while True:
        try:
            query = listen()
            # print(type(query))
            if not enhanced_process_command(query):
                break
        except KeyboardInterrupt:
            speak("Goodbye!")
            break
        except Exception as e:
            print(f"Unexpected error: {e}")
            print("Sorry, something went wrong. Please try again.")

if __name__ == "__main__":
    main()
    
    # automation = ScreenAutomation()

    # # Run PyQt app in a separate thread
    # qt_thread = threading.Thread(target=automation.app.exec_)
    # qt_thread.start()

    # # Run voice automation in main thread
    # automation.run_automation()

    
