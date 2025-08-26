import keyboard, mouse, time, os, threading, webbrowser, sys, time, psutil
from mouse import ButtonEvent
from keyboard import KeyboardEvent

LOCK_FILE = os.path.join(os.path.dirname(__file__), "locks" ,"__cust_shortcuts1.lock")

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



# ==========================================================
# # SHORTCUTS FUNCTIONS
def mouse_key(mouse_button, key, func):
    def check_combo():
        def key_silent(e):
            if ( 
                mouse.is_pressed(mouse_button) and
                e.name == key 
            ):
                return False
            return True
    
        keyboard.hook_key(key, key_silent, suppress=True)
        while True:
            if mouse.is_pressed(mouse_button):
                if keyboard.is_pressed(key):
                    func()
                    time.sleep(1)
            time.sleep(0.01)
    threading.Thread(target=check_combo, daemon=True).start()   # since multiple mouse_key shortcuts i may clash. That's why threading is required

shortcuts = []
state = {
    "mouse_down": False,
    "pressed_keys": set(),
    "current_button": None
}
def mouse_and_keys(shortcut_keys, callback):
    mouse_button = shortcut_keys[0].lower()
    required_keys = set(k.lower() for k in shortcut_keys[1:])
    shortcuts.append({
        "mouse_button": mouse_button,
        "required_keys": required_keys,
        "callback": callback,
        "key_count": len(required_keys)
    })

def mouse_event_handler(event):
    if not isinstance(event, ButtonEvent):
        return

    if event.event_type == 'down':
        state["mouse_down"] = True
        state["current_button"] = event.button.lower()

    elif event.event_type == 'up':
        state["mouse_down"] = False
        state["pressed_keys"].clear()
        state["current_button"] = None

def key_press_handler(event: KeyboardEvent):
    key = event.name.lower()
    state["pressed_keys"].add(key)

    if state["mouse_down"]:
        matching_shortcuts = [
            s for s in shortcuts
            if s["mouse_button"] == state["current_button"]
            and s["required_keys"].issubset(state["pressed_keys"])
        ]

        if matching_shortcuts:
            matching_shortcuts.sort(key=lambda s: s["key_count"], reverse=True)
            most_specific = matching_shortcuts[0]
            most_specific["callback"]()

            list0 = matching_shortcuts[0]['required_keys']
            if all([k in state["pressed_keys"] for k in list0]):
                return False  # Suppress this key
    return True

def key_release_handler(event: KeyboardEvent):
    key = event.name.lower()
    state["pressed_keys"].discard(key)

def start_listener():
    def run():
        mouse.hook(mouse_event_handler)
        keyboard.on_press(key_press_handler, suppress=True)
        keyboard.on_release(key_release_handler)
        keyboard.wait()

    t = threading.Thread(target=run)
    t.daemon = True
    t.start()

def hotstring_replace(hotstring_map:dict):
    buffer = ""
    while True:
        event = keyboard.read_event()
        if event.event_type != keyboard.KEY_DOWN:
            continue

        key = event.name
        if len(key) == 1:
            buffer += key
        elif key == 'backspace':
            buffer = buffer[:-1]
        elif key in ["space", "enter", "tab"]:
            if buffer in hotstring_map:
                for _ in range(len(buffer)+1):
                    keyboard.press_and_release('backspace')
                    time.sleep(0.02)
                keyboard.write(hotstring_map[buffer])
            buffer = ""
        else:
            continue
buffer = ""
# ==========================================================
# # # FUNCTIONS

def open_path(path):
    os.startfile(path)

def on_combination_detected():
    open_path(os.path.dirname(__file__))

def on_combination_detected1():
    open_path(os.path.dirname(__file__))

# ==========================================================
# # # CODES

fpath = r"C:\Users\karthikeyans\Desktop\Missing Mapping\ActivityCode_Mapping_Missing - RAK.csv"
keyboard.add_hotkey('shift+alt+1', lambda: open_path(path=fpath), suppress=True)

fpath1 = r"C:\Users\karthikeyans\Desktop\Missing Mapping\Clinician_Mapping_Missing - RAK.csv"
keyboard.add_hotkey('shift+alt+7', lambda: open_path(path=fpath1), suppress=True)

mouse_key('left', 'a', lambda : open_path('notepad.exe'))
mouse_key('middle', 'a', lambda : open_path(fpath))

mouse_and_keys(['left', 'alt', 'f'], lambda : open_path(fpath))
mouse_and_keys(['left', 'left windows', 'f'], lambda : open_path('calc.exe'))
mouse_and_keys(['left', 'ctrl', 'shift', 'q'], lambda : open_path('calc.exe'))
mouse_and_keys(['left', 'ctrl', 'q'], lambda : open_path('notepad.exe'))
mouse_and_keys(['left', 'ctrl', 'alt'], lambda : open_path('calc.exe'))                 
mouse_and_keys(['middle', 'ctrl', 'alt', 'shift'] , on_combination_detected1)           

# lambda - function has no function inside
# no lambda - function has another function inside
start_listener()

replace_triggers = {
    'ww': 'Water',
    'ee': 'Energy'
}
hotstring_replace(replace_triggers)
keyboard.wait()