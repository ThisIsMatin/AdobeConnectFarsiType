from pynput import keyboard
from plyer import notification
import os
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

        self.on_alt_pressed = False
        self.keyboard_language = ''

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
            pyautogui.hotkey('alt', 'shift')
            self.keyboard_language = 'fa'

        elif(self.GetActiveKeyboardLanguage() == '0x429'):
            pyautogui.hotkey('alt', 'shift')
            print('Reloading your keyboard data...')
            self.RestartProgram()

        else:
            error_message = 'Unknown keyboard language! Please change your keyboard language to persian or english'
            print(error_message)
            notification.notify(
                title='Adobe Connect Farsi Type',
                message=error_message,
                app_name='Adobe Connect Farsi Type',
            )
            time.sleep(4)
            sys.exit(0)

        print('Adobe Connect Farsi Type Enabled !')
        
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
        if(self.GetActiveWindow() == 'connect.exe' and self.keyboard_language == 'fa'):
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
    
    def RestartProgram(self):
        try:
            p = psutil.Process(os.getpid())
            for handler in p.get_open_files() + p.connections():
                os.close(handler.fd)
        except:
            pass

        python = sys.executable
        os.execl(python, python, *sys.argv)

    def OnAnyKeyPressed(self, key):
        # Chcek for disabling tool, if alt and shift are held and the keyboard language is changed
        if(key == keyboard.Key.alt_l or key == keyboard.Key.alt_l or key == keyboard.Key.alt or key == keyboard.Key.alt_gr):
            self.on_alt_pressed = True
            
        if(key == keyboard.Key.shift or key == keyboard.Key.shift_l or key == keyboard.Key.shift_r):
            if(self.on_alt_pressed == True):
                self.keyboard_language = 'en' if self.keyboard_language == 'fa' else 'fa'
                self.on_alt_pressed = False

        if any([key in COMBO for COMBO in self.COMBINATIONS]):
            self.current.add(key)
            if any(all(k in self.current for k in COMBO) for COMBO in self.COMBINATIONS):
                self.Farsi_Formatter(key)

    def OnAnyKeyReleased(self, key):
        if(key == keyboard.Key.alt_l or key == keyboard.Key.alt_l):
            self.on_alt_pressed = False
        
        if any([key in COMBO for COMBO in self.COMBINATIONS]):
            self.current.remove(key)

app = AdobeConnectFarsiType()