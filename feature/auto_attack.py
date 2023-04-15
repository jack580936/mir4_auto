import threading
from time import sleep
import win32con
import win32gui
import win32api
import sys

from PyQt5 import QtCore, QtGui, QtWidgets
from qt_material import apply_stylesheet
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QRadioButton, \
    QCheckBox, QPushButton, QLineEdit, QVBoxLayout, QHBoxLayout, QWidget


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

        self.setWindowTitle("鯊塵暴")
        self.setGeometry(100, 100, 300, 200)

        self.delay = QtWidgets.QDoubleSpinBox(self)
        self.delay.setValue(3)
        self.delay.setReadOnly(False)

        self.delay_label = QtWidgets.QLabel("Delay (seconds):", self)
        self.window_label = QtWidgets.QLabel("Window:", self)
        self.window = QtWidgets.QButtonGroup(self)
        self.window_option1 = QtWidgets.QRadioButton("Mir4G[0]", self)
        self.window_option2 = QtWidgets.QRadioButton("Mir4G[1]", self)
        self.window_option2.setChecked(True)
        self.window_option3 = QtWidgets.QRadioButton("Mir4G[2]", self)
        self.search_monster = QtWidgets.QCheckBox("Search Monster", self)
        self.search_monster.setChecked(True)
        self.press_r_check = QtWidgets.QCheckBox("Press R", self)
        self.start_button = QtWidgets.QPushButton("Start", self)
        self.stop_button = QtWidgets.QPushButton("Stop", self)
        self.stop_button.setDisabled(True)
        self.start_button.setFocus()

        # Connect the stateChanged signal to the enable_thread function
        self.press_r_check.stateChanged.connect(self.enable_thread)
        self.search_monster.stateChanged.connect(self.enable_thread)
        self.delay.editingFinished.connect(self.restart_thread)
        self.window.buttonClicked.connect(self.restart_thread)

        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.addWidget(self.delay_label)
        self.layout.addWidget(self.delay)
        self.layout.addWidget(self.window_label)
        self.layout.addWidget(self.window_option1)
        self.layout.addWidget(self.window_option2)
        self.layout.addWidget(self.window_option3)
        self.layout.addWidget(self.search_monster)
        self.layout.addWidget(self.press_r_check)
        self.layout.addWidget(self.start_button)
        self.layout.addWidget(self.stop_button)

        self.window.addButton(self.window_option1)
        self.window.addButton(self.window_option2)
        self.window.addButton(self.window_option3)

        self.start_button.clicked.connect(self.start_loop_threaded)
        self.stop_button.clicked.connect(self.stop_loop_threaded)

        self.looping = None
        self.thread = None

    def enable_thread(self, state):
        if state == QtCore.Qt.Checked:
            self.start_loop_threaded()
        else:
            self.stop_loop_threaded()
            self.start_loop_threaded()

    def restart_thread(self):
        self.stop_loop_threaded()
        self.start_loop_threaded()

    def start_loop_threaded(self):
        self.looping = True
        self.start_button.setDisabled(True)
        self.stop_button.setDisabled(False)

        delay = self.delay.value()
        window_title = self.window.checkedButton().text()
        search_monster = self.search_monster.isChecked()
        press_r = self.press_r_check.isChecked()
        hwnd = win32gui.FindWindow(None, window_title)
        keyboard = Keyboard(hwnd)

        if not hwnd:
            message = f"Could not find window with title: {window_title}"
            win32api.MessageBox(0, message, "Window Handle", 0)
            self.stop_loop_threaded()
            return

        self.thread = threading.Thread(target=self.loop, args=(keyboard, delay, search_monster, press_r), daemon=True)
        self.thread.start()

    def stop_loop_threaded(self):
        self.looping = False
        if self.thread is not None:
            self.thread.join()
            self.thread = None
        self.start_button.setDisabled(False)
        self.stop_button.setDisabled(True)

    def loop(self, keyboard, delay, search_monster, press_r):
        while self.looping:
            if search_monster:
                keyboard.press_key(win32con.VK_TAB)
                keyboard.press_key(win32con.VK_NEXT)
                keyboard.press_key(win32con.VK_NEXT)
                keyboard.press_key(win32con.VK_NEXT)
                keyboard.press_key(ord('F'))
            if press_r:
                keyboard.press_key(ord('R'))
            sleep(delay)

        self.start_button.setDisabled(False)
        self.stop_button.setDisabled(True)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    apply_stylesheet(app, theme='dark_cyan.xml')
    icon = QIcon('../icon/sand-storm.ico')
    app.setWindowIcon(icon)
    window = App()
    window.show()
    sys.exit(app.exec_())
