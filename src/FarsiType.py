from pynput import keyboard
from plyer import notification
import subprocess
import pyautogui
import time
import sys
import psutil
import ctypes
import win32process
import win32gui

class AdobeConnectFarsiType:
    def __init__(self):

        if(self.ProcessExists('connect.exe') == False):
            print('Your Adobe Connect is not running!')
            notification.notify(
                title='Adobe Connect Farsi Type',
                message='Your Adobe Connect is not running!',
                app_name='Adobe Connect Farsi Type',
            )
            time.sleep(4)
            sys.exit(0)
        
        self.COMBINATIONS = []
        for com in self.Farsi_Combinations():
            self.COMBINATIONS.append({keyboard.KeyCode(char=com)})
        self.current = set()

        if(self.GetActiveKeyboardLanguage() == '0x409'):
        else:

        with keyboard.Listener(on_press=self.OnAnyKeyPressed, on_release=self.OnAnyKeyReleased) as listener:
            listener.join()

    def Farsi_Combinations(self):
        return {
            'd':b'\xd9\x8a',
            'ی':b'\xd9\x8a'
        }

    def ProcessExists(self, Process):
        call = 'TASKLIST', '/FI', 'imagename eq %s' % Process
        output = subprocess.check_output(call).decode()
        last_line = output.strip().split('\r\n')[-1]
        return last_line.lower().startswith(Process.lower())

    def Farsi_Formatter(self, key):
        if(self.GetActiveWindow() == 'connect.exe'):
            time.sleep(0.015)
            pyautogui.press('backspace')
            pyautogui.hotkey('shift', 'x')
    
    def GetActiveWindow(self):
        fgw = win32gui.GetForegroundWindow()
        pid = win32process.GetWindowThreadProcessId(fgw)
        return psutil.Process(pid[-1]).name()
    
    def GetActiveKeyboardLanguage(self):
        user32 = ctypes.WinDLL('user32', use_last_error=True)
        curr_window = user32.GetForegroundWindow()
        thread_id = user32.GetWindowThreadProcessId(curr_window, 0)
        klid = user32.GetKeyboardLayout(thread_id)
        lid = klid & (2**16 - 1)
        return hex(lid)

    def OnAnyKeyPressed(self, key):
        if any([key in COMBO for COMBO in self.COMBINATIONS]):
            self.current.add(key)
            if any(all(k in self.current for k in COMBO) for COMBO in self.COMBINATIONS):
                self.Farsi_Formatter(key)

    def OnAnyKeyReleased(self, key):
        if any([key in COMBO for COMBO in self.COMBINATIONS]):
            self.current.remove(key)

app = AdobeConnectFarsiType()