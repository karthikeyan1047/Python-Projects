import webbrowser, os, sys, time, psutil
from pynput import keyboard
from pynput.keyboard import Controller, Key

LOCK_FILE = os.path.join(os.path.dirname(__file__), "locks" , "__cust_shortcuts2.lock")

def kill_previous_instance():
    if not os.path.exists(LOCK_FILE):
        return

    try:
        with open(LOCK_FILE, 'r') as f:
            old_pid = int(f.read())
        if psutil.pid_exists(old_pid):
            p = psutil.Process(old_pid)
            if 'python' in p.name().lower():
                p.terminate()
                p.wait(timeout=3)
    except Exception as e:
        pass
kill_previous_instance()

with open(LOCK_FILE, 'w') as f:
    f.write(str(os.getpid()))

import atexit
atexit.register(lambda: os.remove(LOCK_FILE) if os.path.exists(LOCK_FILE) else None)


# =========================================
# # # ACTIONS FUNCTIONS

typed_chars = ""
kb_controller = Controller()

def open_target(path):
    if path.startswith("http://") or path.startswith("https://"):
        webbrowser.open(path)
    elif os.path.exists(path):
        os.startfile(path)
    else:
        pass

def clear_word(word):
    for _ in range(len(word)+1):
        kb_controller.press(Key.backspace)
        kb_controller.release(Key.backspace)

def on_press(key):
    global typed_chars
    try:
        if key.char.isalnum():
            typed_chars += key.char
    except AttributeError:
        if key == Key.space or key == Key.enter:
            check_trigger()
        elif key == Key.esc:
            return False

def check_trigger():
    global typed_chars

    for trigger_word, target in open_triggers.items():
        if typed_chars.endswith(trigger_word):
            clear_word(trigger_word)
            open_target(target)
            typed_chars = ""
            return
    typed_chars = ""

def hotstring_open():
    with keyboard.Listener(on_press=on_press) as listener:
        listener.join()

# =========================================
# # # ACTIONS

open_triggers = {
    "ggg": "https://www.google.com",
    "doc": r"C:\Users\karthikeyans\Documents\picture",
    "yt": "https://youtube.com"
}
hotstring_open()
