# Copyright 2025 Adobe. All rights reserved.
# This file is licensed to you under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License. You may obtain a copy
# of the License at http://www.apache.org/licenses/LICENSE-2.0

# Unless required by applicable law or agreed to in writing, software distributed under
# the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR REPRESENTATIONS
# OF ANY KIND, either express or implied. See the License for the specific language
# governing permissions and limitations under the License.
import pyperclip,pyautogui
import time,os


def switchAgain():
    os.system("./Utilities/changeInputSource/changeInputSource select com.apple.keylayout.ABC")
    time.sleep(0.5)
    pyautogui.hotkey('ctrl','space',interval=0.2)
    print("Typed....")
    time.sleep(0.5)

def jpTypingCheck():
    pyperclip.copy('')
    pyautogui.hotkey('command', 'a',interval=0.5)
    time.sleep(0.2)
    pyautogui.write("asahi",interval=0.5)
        
    pyautogui.press("enter")
        
    time.sleep(0.5)
    pyautogui.hotkey('command', 'a')
    pyautogui.hotkey('command', 'x')
    rendered=pyperclip.paste()
    if(rendered=="asahi" or rendered=="" or rendered==None):
        print("ja_JP IME Not Changed")
        print("Trying Again.....")
        switchAgain()
        pyautogui.hotkey('command', 'a',interval=0.5)
        time.sleep(0.2)
        pyautogui.write("asahi",interval=0.5)
            
        pyautogui.press("enter")
            
        time.sleep(0.5)
        pyautogui.hotkey('command', 'a')
        pyautogui.hotkey('command', 'x')
        rendered=pyperclip.paste()
        if(rendered=="asahi" or rendered=="" or rendered==None):
        # keyboard_listener.stop()
        # JpEnglish[0]=True
        # return [111]
            return True
        return False
    else:
        return False

def koTypingCheck():
    pyperclip.copy('')
    pyautogui.hotkey('command', 'a',interval=0.5)
    time.sleep(0.2)
    pyautogui.write("rka.",interval=0.5)
        
    # pyautogui.press("enter")
        
    time.sleep(0.5)
    pyautogui.hotkey('command', 'a')
    pyautogui.hotkey('command', 'x')
    rendered=pyperclip.paste()
    if(rendered=="rka." or rendered=="" or rendered==None):
        print("ko_KR IME Not Changed")
        print("Trying Again....")
        # keyboard_listener.stop()
        # JpEnglish[0]=True
        # return [111]
        switchAgain()
        pyautogui.hotkey('command', 'a',interval=0.5)
        time.sleep(0.2)
        pyautogui.write("rka.",interval=0.5)
            
        # pyautogui.press("enter")
            
        time.sleep(0.5)
        pyautogui.hotkey('command', 'a')
        pyautogui.hotkey('command', 'x')
        rendered=pyperclip.paste()
        if(rendered=="rka." or rendered=="" or rendered==None):
            return True

        return False
    else:
        return False

def cnTypingCheck():
    pyperclip.copy('')
    pyautogui.hotkey('command', 'a',interval=0.5)
    time.sleep(0.2)
    pyautogui.write("xian",interval=0.5)
    pyautogui.press("space")  
    # pyautogui.press("enter")
        
    time.sleep(0.5)
    pyautogui.hotkey('command', 'a')
    pyautogui.hotkey('command', 'x')
    rendered=pyperclip.paste()
    if(rendered=="xian " or rendered=="" or rendered==None or rendered=="xian"):
        print("zh_CN IME Not Changed")
        print("Trying Again....")
        switchAgain()
        pyautogui.hotkey('command', 'a',interval=0.5)
        time.sleep(0.2)
        pyautogui.write("xian",interval=0.5)
        pyautogui.press("space")  
        # pyautogui.press("enter")
            
        time.sleep(0.5)
        pyautogui.hotkey('command', 'a')
        pyautogui.hotkey('command', 'x')
        rendered=pyperclip.paste()
        if(rendered=="xian " or rendered=="" or rendered==None or rendered=="xian"):
            return True
        # keyboard_listener.stop()
        # JpEnglish[0]=True
        # return [111]
        return False
    else:
        return False

def twTypingCheck():
    pyperclip.copy('')
    pyautogui.hotkey('command', 'a',interval=0.5)
    time.sleep(0.2)
    pyautogui.write("t",interval=0.5)
    pyautogui.press("space")  
    # pyautogui.press("enter")
        
    time.sleep(0.5)
    pyautogui.hotkey('command', 'a')
    pyautogui.hotkey('command', 'x')
    rendered=pyperclip.paste()
    if(rendered=="t " or rendered=="" or rendered==None or rendered=="t"):
        print("zh_TW IME Not Changed")
        print("Trying Again....")
        switchAgain()
        pyautogui.hotkey('command', 'a',interval=0.5)
        time.sleep(0.2)
        pyautogui.write("t",interval=0.5)
        pyautogui.press("space")  
        # pyautogui.press("enter")
            
        time.sleep(0.5)
        pyautogui.hotkey('command', 'a')
        pyautogui.hotkey('command', 'x')
        rendered=pyperclip.paste()
        if(rendered=="t " or rendered=="" or rendered==None or rendered=="t"):
            return True
        # keyboard_listener.stop()
        # JpEnglish[0]=True
        # return [111]
        return False
    else:
        return False