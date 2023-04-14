import threading
import tkinter as tk
from time import sleep
import win32con
import win32gui
import win32api


class Keyboard:
    def __init__(self, hwnd):
        self.hwnd = hwnd

    def press_key(self, key):
        vk_code = win32api.MapVirtualKey(key, 0)
        win32api.PostMessage(self.hwnd, win32con.WM_KEYDOWN, key, vk_code)
        win32api.PostMessage(self.hwnd, win32con.WM_KEYUP, key, vk_code)
        sleep(0.1)


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("GUI with Button")
        self.delay = tk.DoubleVar(value=3)
        self.looping = False
        self.window = tk.StringVar(value="Mir4G[1]")
        self.only_r = tk.BooleanVar(value=False)
        self.delay_label = tk.Label(self.root, text="Delay (seconds):")
        self.delay_entry = tk.Entry(self.root, textvariable=self.delay)
        self.window_label = tk.Label(self.root, text="Window:")
        self.window_option1 = tk.Radiobutton(self.root, text="Mir4G[0]", variable=self.window, value="Mir4G[0]")
        self.window_option2 = tk.Radiobutton(self.root, text="Mir4G[1]", variable=self.window, value="Mir4G[1]")
        self.window_option3 = tk.Radiobutton(self.root, text="Mir4G[2]", variable=self.window, value="Mir4G[2]")
        self.only_r_check = tk.Checkbutton(self.root, text="Press R only", variable=self.only_r)
        self.button_start = tk.Button(self.root, text="Start", command=self.start_loop)
        self.button_stop = tk.Button(self.root, text="Stop", command=self.stop_loop, state=tk.DISABLED)
        self.delay_label.pack(fill=tk.X, expand=True)
        self.delay_entry.pack(fill=tk.X, expand=True)
        self.window_label.pack(fill=tk.X, expand=True)
        self.window_option1.pack(fill=tk.X, expand=True)
        self.window_option2.pack(fill=tk.X, expand=True)
        self.window_option3.pack(fill=tk.X, expand=True)
        self.only_r_check.pack(fill=tk.X, expand=True)
        self.button_start.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.button_stop.pack(side=tk.LEFT, fill=tk.X, expand=True)


    def run(self):
        self.root.mainloop()

    def start_loop(self):
        self.looping = True
        self.button_start.config(state=tk.DISABLED)
        self.button_stop.config(state=tk.NORMAL)
        delay = self.delay.get()
        window_title = self.window.get()
        only_r = self.only_r.get()
        hwnd = win32gui.FindWindow(None, window_title)
        keyboard = Keyboard(hwnd)
        if not hwnd:
            message = f"Could not find window with title: {window_title}"
            win32api.MessageBox(0, message, "Window Handle", 0)
            self.stop_loop()
            self.button_start.config(state=tk.NORMAL)
            self.button_stop.config(state=tk.DISABLED)
            return
        thread = threading.Thread(target=self.loop, args=(keyboard, delay, only_r), daemon=True)
        thread.start()

    def loop(self, keyboard, delay, only_r):
        while self.looping:
            if not only_r:
                keyboard.press_key(win32con.VK_TAB)
                keyboard.press_key(win32con.VK_NEXT)
                keyboard.press_key(win32con.VK_NEXT)
                keyboard.press_key(win32con.VK_NEXT)
                keyboard.press_key(ord('F'))
            keyboard.press_key(ord('R'))
            sleep(delay)

        self.button_start.config(state=tk.NORMAL)
        self.button_stop.config(state=tk.DISABLED)

    def stop_loop(self):
        self.looping = False




app = App()
app.run()
