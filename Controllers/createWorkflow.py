# Copyright 2025 Adobe. All rights reserved.
# This file is licensed to you under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License. You may obtain a copy
# of the License at http://www.apache.org/licenses/LICENSE-2.0

# Unless required by applicable law or agreed to in writing, software distributed under
# the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR REPRESENTATIONS
# OF ANY KIND, either express or implied. See the License for the specific language
# governing permissions and limitations under the License.

from pynput import mouse
from pynput import keyboard
from pynput.mouse import Button, Controller as MouseController
from pynput.keyboard import Key, Controller as KeyboardController
import time
import json
import sys,os,shutil,subprocess
from flask import *
from Utilities.Utilities import *
from datetime import datetime
import pyautogui
import pyperclip
from Utilities.keyboardLanguage import *
from Utilities.SendMail.newSendMail import *
import webbrowser
import threading,_thread
from threading import Thread
from xlwt import Workbook
import xlwt
import xlsxwriter
import pyautogui
from Utilities.listToExcel import *
import webbrowser
from flask_cors import CORS, cross_origin
import signal,atexit
from Utilities.functionSCPToServer import *
from Utilities.checkDisplayOff import *


def createWorkflow():
    if(request.form['installation']=="2"):
        productName=request.form['product']
        productPath="/Applications/"+productName+"/"+productName+".app"
        print(productName)
        print(productPath)
        applicationLocale=request.form['installationLocale']
        print(applicationLocale)
        # Check ProductName Folder Present Else Create It for Coordinates
        checkandcreatefolder(productName)
        if(os.path.exists(productPath)==True):
            os.system("open '{}'".format(productPath))
            ##Here adding Screen in Full Size
            width,height=pyautogui.size()
            print(width,height,sep="-")
       
            now=datetime.now()
            name_of_recording = productName+"-"+now.strftime("%d-%m-%Y %H-%M-%S")+"_"+str(width)+"_"+str(height)+"_"+applicationLocale
            record_all = True

            userPath=os.path.expanduser('~')
            print(userPath)

            desktopPath=userPath+"/Desktop/"
            downloadPath=userPath+"/Downloads/"
            print(desktopPath)
            workflowPath=downloadPath+'IME Framework/Workflows/Desktop'
            checkandCreateScratch(workflowPath)
            storage = []
            count = 0
            def start_keyboard_listener():
                keyboard_listener = keyboard.Listener(
                            on_press=on_press,
                            on_release=on_release)
                keyboard_listener.start()
            try:
                def on_press(key):
                    try:
                        json_object = {'action':'pressed_key', 'key':key.char, '_time': time.time()}
                    except AttributeError:
                        if key == keyboard.Key.esc:
                            with open('{0}/{1}.txt'.format(workflowPath,name_of_recording), 'w') as outfile:
                                json.dump(storage, outfile)
                            # sys.exit(0)
                            mouse_listener.stop()
                            # keyboard_listener.stop()
                            keyboard_listener.stop()
                            os.system('open "/Applications/Google Chrome.app"')
                            os.system("pkill -f '{}'".format(productPath))
                        json_object = {'action':'pressed_key', 'key':str(key), '_time': time.time()}
                    storage.append(json_object)

                def on_release(key):
                    try:
                        json_object = {'action':'released_key', 'key':key.char, '_time': time.time()}
                    except AttributeError:
                        json_object = {'action':'released_key', 'key':str(key), '_time': time.time()}
                    storage.append(json_object)
                        

                def on_move(x, y):
                    if (record_all) == True:
                        if len(storage) >= 1:
                            if storage[-1]['action'] != "moved":
                                json_object = {'action':'moved', 'x':x, 'y':y, '_time':time.time()}
                                storage.append(json_object)
                            elif storage[-1]['action'] == "moved" and time.time() - storage[-1]['_time'] > 0.02:
                                json_object = {'action':'moved', 'x':x, 'y':y, '_time':time.time()}
                                storage.append(json_object)
                        else:
                            json_object = {'action':'moved', 'x':x, 'y':y, '_time':time.time()}
                            storage.append(json_object)
                    else:
                        if len(storage) >= 1:
                            if (storage[-1]['action'] == "pressed" and storage[-1]['button'] == 'Button.left') or (storage[-1]['action'] == "moved" and time.time() - storage[-1]['_time'] > 0.02):
                                json_object = {'action':'moved', 'x':x, 'y':y, '_time':time.time()}
                                storage.append(json_object)

                def on_click(x, y, button, pressed):
                    json_object = {'action':'pressed' if pressed else 'released', 'button':str(button), 'x':x, 'y':y, '_time':time.time()}
                    storage.append(json_object)
                    if len(storage) > 1:
                        if storage[-1]['action'] == 'released' and storage[-1]['button'] == 'Button.right' and storage[-1]['_time'] - storage[-2]['_time'] > 2:
                            with open('data/{}.txt'.format(name_of_recording), 'w') as outfile:
                                json.dump(storage, outfile)
                            return False


                def on_scroll(x, y, dx, dy):
                    json_object = {'action': 'scroll', 'vertical_direction': int(dy), 'horizontal_direction': int(dx), 'x':x, 'y':y, '_time': time.time()}
                    storage.append(json_object)



                
                keyboard_listener = keyboard.Listener(
                    on_press=on_press,
                    on_release=on_release)

                mouse_listener = mouse.Listener(
                        on_click=on_click,
                        on_scroll=on_scroll,
                        on_move=on_move)

                mouse_listener.start()
                time.sleep(0.6)
                keyboard_listener.start()
                
                
                keyboard_listener.join()
                mouse_listener.join()


    
            except AttributeError:
                print("Sorry Some Issue With the Installer")
                return [100]
                
        else:
            return [101]
           
    elif(request.form['installation']=="3"):
        productName=request.form['product']
        productPath=request.form['product_url']
        print(productName)
        print(productPath)
        applicationLocale=request.form['installationLocale']
        print(applicationLocale)

        checkandcreatefolder(productName)
        if(os.path.exists(productPath)==True and productPath[-4:]==".app"):
            os.system("open '{}'".format(productPath))
            ##Here adding Screen in Full Size
            width,height=pyautogui.size()
            print(width,height,sep="-")
           
            now=datetime.now()
            name_of_recording = productName+"-"+now.strftime("%d-%m-%Y %H-%M-%S")+"_"+str(width)+"_"+str(height)+"_"+applicationLocale
            record_all = True
            userPath=os.path.expanduser('~')
            print(userPath)

            desktopPath=userPath+"/Desktop/"
            downloadPath=userPath+"/Downloads/"
            print(desktopPath)
            workflowPath=downloadPath+'IME Framework/Workflows'
            checkandCreateScratch(workflowPath)

            

            storage = []
            count = 0
            def start_keyboard_listener():
                keyboard_listener = keyboard.Listener(
                            on_press=on_press,
                            on_release=on_release)
                keyboard_listener.start()

            def on_press(key):
                try:
                    json_object = {'action':'pressed_key', 'key':key.char, '_time': time.time()}
                except AttributeError:
                    if key == keyboard.Key.esc:
                        with open('{0}/{1}.txt'.format(workflowPath,name_of_recording), 'w') as outfile:
                            json.dump(storage, outfile)
                        # sys.exit(0)
                        mouse_listener.stop()
                        # keyboard_listener.stop()
                        keyboard_listener.stop()
                        os.system('open "/Applications/Google Chrome.app"')
                        os.system("pkill -f '{}'".format(productPath))
                    json_object = {'action':'pressed_key', 'key':str(key), '_time': time.time()}
                storage.append(json_object)

            def on_release(key):
                try:
                    json_object = {'action':'released_key', 'key':key.char, '_time': time.time()}
                except AttributeError:
                    json_object = {'action':'released_key', 'key':str(key), '_time': time.time()}
                storage.append(json_object)
                    

            def on_move(x, y):
                if (record_all) == True:
                    if len(storage) >= 1:
                        if storage[-1]['action'] != "moved":
                            json_object = {'action':'moved', 'x':x, 'y':y, '_time':time.time()}
                            storage.append(json_object)
                        elif storage[-1]['action'] == "moved" and time.time() - storage[-1]['_time'] > 0.02:
                            json_object = {'action':'moved', 'x':x, 'y':y, '_time':time.time()}
                            storage.append(json_object)
                    else:
                        json_object = {'action':'moved', 'x':x, 'y':y, '_time':time.time()}
                        storage.append(json_object)
                else:
                    if len(storage) >= 1:
                        if (storage[-1]['action'] == "pressed" and storage[-1]['button'] == 'Button.left') or (storage[-1]['action'] == "moved" and time.time() - storage[-1]['_time'] > 0.02):
                            json_object = {'action':'moved', 'x':x, 'y':y, '_time':time.time()}
                            storage.append(json_object)

            def on_click(x, y, button, pressed):
                json_object = {'action':'pressed' if pressed else 'released', 'button':str(button), 'x':x, 'y':y, '_time':time.time()}
                storage.append(json_object)
                if len(storage) > 1:
                    if storage[-1]['action'] == 'released' and storage[-1]['button'] == 'Button.right' and storage[-1]['_time'] - storage[-2]['_time'] > 2:
                        with open('data/{}.txt'.format(name_of_recording), 'w') as outfile:
                            json.dump(storage, outfile)
                        return False


            def on_scroll(x, y, dx, dy):
                json_object = {'action': 'scroll', 'vertical_direction': int(dy), 'horizontal_direction': int(dx), 'x':x, 'y':y, '_time': time.time()}
                storage.append(json_object)

            
            keyboard_listener = keyboard.Listener(
                on_press=on_press,
                on_release=on_release)

            mouse_listener = mouse.Listener(
                    on_click=on_click,
                    on_scroll=on_scroll,
                    on_move=on_move)

            keyboard_listener.start()
            time.sleep(0.4)
            mouse_listener.start()
            
            
            
            
            keyboard_listener.join()
            mouse_listener.join()
        else:
            return [101]
    return [productPath,workflowPath,name_of_recording]
 