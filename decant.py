import pyautogui
import time
import win32gui
import re
import PySimpleGUI as sg

#This class defines some things that will be needed to call excel windows into the foreground
class WindowMgr:
    """Encapsulates some calls to the winapi for window management"""

    def __init__ (self):
        """Constructor"""
        self._handle = None

    def find_window(self, class_name, window_name=None):
        """find a window by its class_name"""
        self._handle = win32gui.FindWindow(class_name, window_name)

    def _window_enum_callback(self, hwnd, wildcard):
        """Pass to win32gui.EnumWindows() to check all the opened windows"""
        if re.match(wildcard, str(win32gui.GetWindowText(hwnd))) is not None:
            self._handle = hwnd

    def find_window_wildcard(self, wildcard):
        """find a window whose title matches the wildcard regex"""
        self._handle = None
        win32gui.EnumWindows(self._window_enum_callback, wildcard)

    def set_foreground(self):
        """put the window in the foreground"""
        win32gui.SetForegroundWindow(self._handle)

#This section of the couds pops up a simple GUI were the user is going to input the name of the Excel file they will be working with
layout = [[sg.Text('Name of Excel File you will be using (case sensitive)')],
                 [sg.InputText()],
                 [sg.Submit(), sg.Cancel()]]

window = sg.Window('SunQuest Label Automation', layout)

event, values = window.read()
window.close()

text_input = values[0]
sg.popup('You entered', text_input)

#This calls the window of interest into the foreground once the user intputs what window they will be working with
w = WindowMgr()
w.find_window_wildcard(".*" + text_input + ".*")
w.set_foreground()
#this code only works on screen sizes that have a width=1920, height=1080
print(pyautogui.size())
try:
    while True:
        w = WindowMgr()
        w.find_window_wildcard(".*A10960.*")
        w.set_foreground()

        #the following function clicks in location
        pyautogui.click(1040,246, duration = 2)
        time.sleep(.5)
        pyautogui.hotkey('ctrlleft', 'c')
        time.sleep(.5)
        pyautogui.scroll(-120)
        time.sleep(1)
        pyautogui.moveTo(164,92, duration = 1)
        pyautogui.click(164,92, duration = 1)
        time.sleep(1)
        pyautogui.hotkey('ctrlleft', 'v')
        time.sleep(.5)
        pyautogui.click(795,89, duration = 1)
        pyautogui.click(790,259, duration = 1)
        pyautogui.click(1014,661, duration = 1)
        time.sleep(2)

        w = WindowMgr()
        w.find_window_wildcard(".*A10960.*")
        w.set_foreground()

except KeyboardInterrupt:
    pass
