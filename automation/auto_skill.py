import sys
import threading
from time import sleep

import win32con
import win32api
import win32gui
from PyQt5 import QtWidgets
from qt_material import apply_stylesheet
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication


class Keyboard:
    def __init__(self, hwnd):
        self.hwnd = hwnd

    def press_key(self, key):
        vk_code = win32api.MapVirtualKey(key, 0)
        win32api.PostMessage(self.hwnd, win32con.WM_KEYDOWN, key, vk_code)
        win32api.PostMessage(self.hwnd, win32con.WM_KEYUP, key, vk_code)
        sleep(0.1)


class App(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.key_presses = None
        self.setWindowTitle("鯊塵暴")
        self.setGeometry(100, 100, 100, 100)
        self.setWindowIcon(QIcon('../icon/sand-storm.ico'))
        self.delay = QtWidgets.QDoubleSpinBox(self)
        self.delay.setValue(10)
        self.delay_label = QtWidgets.QLabel("Delay (seconds):", self)
        self.window_label = QtWidgets.QLabel("Window:", self)
        self.window = QtWidgets.QButtonGroup(self)
        self.window_option = QtWidgets.QComboBox(self)
        self.window_option.addItems(["Mir4G[0]", "Mir4G[1]", "Mir4G[2]"])
        self.window_option.setCurrentIndex(1)
        self.hotkey_label = QtWidgets.QLabel("Hotkey:", self)
        self.search_monster = QtWidgets.QCheckBox("Search Monster", self)
        self.press_r_check = QtWidgets.QCheckBox("Press R", self)
        self.start_button = QtWidgets.QPushButton("Start", self)
        self.stop_button = QtWidgets.QPushButton("Stop", self)
        self.stop_button.setDisabled(True)
        self.checkboxes = []

        self.layout1 = QtWidgets.QVBoxLayout(self)
        self.layout2 = QtWidgets.QGridLayout(self)
        self.layout1.addWidget(self.delay_label)
        self.layout1.addWidget(self.delay)
        self.layout1.addWidget(self.window_label)
        self.layout1.addWidget(self.window_option)
        self.layout1.addWidget(self.hotkey_label)
        self.layout1.addLayout(self.layout2)
        self.layout2.addWidget(self.search_monster, 4, 0)
        self.layout2.addWidget(self.press_r_check, 4, 1)

        for i in range(6):
            checkbox = QtWidgets.QCheckBox(f"Press {i + 1}", self)
            checkbox.setChecked(False)
            self.checkboxes.append(checkbox)
            self.layout2.addWidget(checkbox, i % 3, i // 3)
        self.layout1.addWidget(self.start_button)
        self.layout1.addWidget(self.stop_button)

        self.start_button.clicked.connect(self.start_loop)
        self.stop_button.clicked.connect(self.stop_loop)

        self.looping = False

    def start_loop(self):

        delay = self.delay.value()
        window_title = self.window_option.currentText()
        hwnd = win32gui.FindWindow(None, window_title)
        keyboard = Keyboard(hwnd)

        if not hwnd:
            message = f"Could not find window with title: {window_title}"
            win32api.MessageBox(0, message, "Window Handle", 0)
            self.stop_loop()
            return

        self.looping = True
        self.start_button.setDisabled(True)
        self.stop_button.setDisabled(False)
        self.delay.setDisabled(True)
        self.window_option.setDisabled(True)

        self.looping = True
        self.start_button.setDisabled(True)
        self.stop_button.setDisabled(False)
        self.delay.setDisabled(True)
        self.window_option.setDisabled(True)

        self.thread = threading.Thread(target=self.loop, args=(keyboard, delay), daemon=True)
        self.thread.start()

    def stop_loop(self):
        self.looping = False
        self.start_button.setDisabled(False)
        self.stop_button.setDisabled(True)
        self.delay.setDisabled(False)
        self.window_option.setDisabled(False)
        self.stop_button.setDisabled(True)

    def loop(self, keyboard, delay):
        while self.looping:

            search_monster = self.search_monster.isChecked()
            press_r = self.press_r_check.isChecked()
            self.key_presses = []
            for i, checkbox in enumerate(self.checkboxes):
                if checkbox.isChecked():
                    self.key_presses.append(ord(str(i + 1)))

            if search_monster:
                keyboard.press_key(win32con.VK_TAB)
                keyboard.press_key(win32con.VK_NEXT)
                keyboard.press_key(win32con.VK_NEXT)
                keyboard.press_key(win32con.VK_NEXT)
                keyboard.press_key(ord('F'))
            if press_r:
                keyboard.press_key(ord('R'))
            if self.key_presses:
                for key_press in self.key_presses:
                    keyboard.press_key(key_press)
                    sleep(1)
            for i in range(int(delay / 0.1)):
                if not self.looping:
                    break
                sleep(0.1)

        self.start_button.setDisabled(False)
        self.stop_button.setDisabled(True)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    apply_stylesheet(app, theme='dark_cyan.xml')
    window = App()
    window.show()
    sys.exit(app.exec_())
