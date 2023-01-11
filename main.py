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

#This section of the code pops up a simple GUI were the user is going to input the name of the Excel file they will be working with
layout = [[sg.Text('Name of Excel File you will be using (case sensitive)')],
                 [sg.InputText()],
                 [sg.Submit(), sg.Cancel()]]

window = sg.Window('SunQuest Label Automation', layout)

event, values = window.read()
window.close()

text_input = values[0]
sg.popup('You entered', text_input)

#this code only works on screen sizes that have a width=1920, height=1080 and this functions tells you the messurments of your screen
print(pyautogui.size())

try:
    while True:
        #This sets the window of interest into the foreground using regular expression search based on your gui inputs (Must be changed manually to match the excel file you will use)
        w = WindowMgr()
        w.find_window_wildcard(".*A10960.*")
        w.set_foreground()

        #the following copy and pastes the first Excel Cell (excluding headers), and copy pastes it into the search bar on SunQuest
        pyautogui.click(1040,252, duration = 2)
        time.sleep(.5)
        pyautogui.hotkey('ctrlleft', 'c')
        time.sleep(.5)
        pyautogui.scroll(-120)
        time.sleep(1)
        pyautogui.moveTo(904,293, duration = 1)
        pyautogui.click(904,293, duration = 1)
        time.sleep(1)
        pyautogui.hotkey('ctrlleft', 'v')
        time.sleep(.5)
        pyautogui.click(1079,291, duration = 1)

        #This sections hits safe on the pop up screen
        pyautogui.click(962,711, duration = 1)
        pyautogui.click(776,669, duration = 1)
        time.sleep(2)

        #Change this to print a specific order on the list based on it's coordinate position.
        pyautogui.click(689, 380, duration=1)

        #This clicks date and time (which auto fill with each click) and saves it
        pyautogui.click(735,651, duration = 1)
        pyautogui.click(735,667, duration = 1)
        pyautogui.click(735,683, duration = 1)
        pyautogui.click(1047,799, duration = 1)
        time.sleep(1)

        #This Routs the inputs for the order entery
        pyautogui.click(1020,763, duration = 1)
        time.sleep(3)

        #Fills out the source and saves it
        pyautogui.click(1005,747, duration = 1)
        pyautogui.click(1132, 799, duration=1)

        #Brings the Excel sheet to the foreground and repeats the cycle
        pyautogui.click(1040, 246, duration=2)

#This is used to Interrupt the code. The user should aggresively move their mouse for about 2 seconds for the code to break off
except KeyboardInterrupt:
    pass
