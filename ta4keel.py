import keyboard
import pyautogui
import tkinter as tk
from tkinter import simpledialog
import time
import sys

# Tashkeel dictionary with combined forms
tashkeel_map = {
    'fatha': 'َ',
    'damma': 'ُ',
    'kasra': 'ِ',
    'shadda': 'ّ',
    'sukun': 'ْ',
    'tanween_fath': 'ً',
    'tanween_damm': 'ٌ',
    'tanween_kasr': 'ٍ',
    'shadda_fatha': 'َّ',  # Shadda then Fatha
    'shadda_damma': 'ُّ',
    'shadda_kasra': 'ِّ'
}

# Ensure Word is focused before typing (bring Word to the front)
def focus_on_word():
    # Check if Word is running and bring it to the front
    windows = pyautogui.getWindowsWithTitle("Word")
    if windows:
        word_window = windows[0]
        word_window.activate()
        print("Microsoft Word is now focused.")
    else:
        print("Word not found. Please open Word.")
        sys.exit()

# Function to show input dialog in the main thread
def ask_for_tashkeel():
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Show a prompt with available options
    options = "\n".join([f"{key}: {value}" for key, value in tashkeel_map.items()])
    choice = simpledialog.askstring("Tashkeel Input", f"Choose a tashkeel:\n{options}\n\nType the *name* (e.g. 'shadda_fatha'):")

    if choice and choice.strip().lower() in tashkeel_map:
        tashkeel = tashkeel_map[choice.strip().lower()]
        print(f"Selected Tashkeel: {tashkeel}")  # Debugging output
        focus_on_word()  # Make sure Word is focused
        pyautogui.write(tashkeel, interval=0.1)  # Add a small interval between keystrokes
        time.sleep(0.2)  # Additional delay after typing the Tashkeel
    else:
        pyautogui.write("⛔ Invalid choice")
        time.sleep(0.2)

# Function to trigger the dialog on hotkey press
def on_hotkey():
    ask_for_tashkeel()  # Call the function directly (in the main thread)

# Register the hotkey
keyboard.add_hotkey('ctrl+shift+t', on_hotkey)

print("Script running. Press Ctrl+Shift+T in Word to input Tashkeel...")
keyboard.wait()  # Keeps the script running and listens for hotkeys
