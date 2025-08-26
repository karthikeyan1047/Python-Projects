import tkinter as tk
import keyboard, mouse, time, os, threading, webbrowser, sys, psutil
from mouse import ButtonEvent

LOCK_FILE = os.path.join(os.path.dirname(__file__), "locks" , "__cust_shortcuts3.lock")

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

# ==================================================
# # # SHORTCUT FUNCTIONS

def mouse_key(mouse_button, key, func):
    def check_combo():
        def key_silent(e):
            if e.name == key and mouse.is_pressed(mouse_button):
                return False
            return True
    
        keyboard.hook_key(key, key_silent, suppress=True)
        while True:
            if mouse.is_pressed(mouse_button):
                if keyboard.is_pressed(key):
                    func()
                    time.sleep(1)
            time.sleep(0.01)
    threading.Thread(target=check_combo, daemon=True).start()

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

def key_event_handler(event):
    key = event.name.lower()

    if event.event_type == 'down':
        state["pressed_keys"].add(key)
    elif event.event_type == 'up':
        prev_keys = state["pressed_keys"].copy()
        state["pressed_keys"].discard(key)
    
        if state["mouse_down"]:
            matching_shortcuts = [
                s for s in shortcuts
                if s["mouse_button"] == state["current_button"]
                and s["required_keys"].issubset(prev_keys )
            ]

            if matching_shortcuts:
                matching_shortcuts.sort(key=lambda s: s["key_count"], reverse=True)
                most_specific = matching_shortcuts[0]
                most_specific["callback"]()

def start_listener():
    def run():
        mouse.hook(mouse_event_handler)
        keyboard.hook(key_event_handler)
        keyboard.wait()

    t = threading.Thread(target=run)
    t.daemon = True
    t.start()


def button_shortcuts():
    def main_ui():
        root = tk.Tk()
        root.attributes('-topmost', True)
        root.title("File Operations")
        window_width = 800
        window_height = 100

        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = int((screen_width / 2) - (window_width / 2))
        y = int((screen_height / 2) - (window_height / 2))
        root.geometry(f"{window_width}x{window_height}+{x}+{y}")

        tk.Label(root, text="Choose :", font=('Copperplate Gothic Bold', 12, 'bold'), fg='black').grid(row=0, column=0, padx=5, pady=5)
        tk.Button(root, text="Create Files", width=12, font=('Bahnschrift SemiLight SemiConde', 10, 'bold'), fg='white', bg='black', command=lambda: [create_files(), root.destroy()]).grid(row=0, column=1, padx=5, pady=5)
        tk.Button(root, text="Transfer Files", width=12, font=('Bahnschrift SemiLight SemiConde', 10, 'bold'), fg='white', bg='black', command=lambda: [transfer_files(), root.destroy()]).grid(row=0, column=2, padx=5, pady=5)
        tk.Button(root, text="Delete Files", width=12, font=('Bahnschrift SemiLight SemiConde', 10, 'bold'), fg='white', bg='black', command=lambda: [delete_files(), root.destroy()]).grid(row=0, column=3, padx=5, pady=5)

        tk.Label(root, text="Pick :", font=('Copperplate Gothic Bold', 12, 'bold'), fg='black').grid(row=1, column=0, padx=5, pady=5)
        tk.Button(root, text="Create Files", width=12, font=('Bahnschrift SemiLight SemiConde', 10, 'bold'), fg='white', bg='black', command=lambda: [create_files(), root.destroy()]).grid(row=1, column=1, padx=5, pady=5)
        tk.Button(root, text="Transfer Files", width=12, font=('Bahnschrift SemiLight SemiConde', 10, 'bold'), fg='white', bg='black', command=lambda: [transfer_files(), root.destroy()]).grid(row=1, column=2, padx=5, pady=5)
        tk.Button(root, text="Delete Files", width=12, font=('Bahnschrift SemiLight SemiConde', 10, 'bold'), fg='white', bg='black', command=lambda: [delete_files(), root.destroy()]).grid(row=1, column=3, padx=5, pady=5)

        root.mainloop()
    main_ui()

# ==================================================
# # ACTION FUNCTIONS

def open_path(path):
    os.startfile(path)

def create_files():
    open_path(os.path.dirname(__file__))

def transfer_files():
    open_path(os.path.dirname(__file__))

def delete_files():
    open_path(os.path.dirname(__file__))



# ==================================================
# # # ACTIONS


mouse_and_keys(['middle', 'enter'], button_shortcuts)
start_listener()

keyboard.wait()




