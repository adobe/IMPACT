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
from Utilities.enUSHandling.functionsenUSHandling import *
import random
from Utilities.DBIntegration import *
from urllib.parse import urlparse
def webWorkflowIME():
    
    def createZipofFolder(folderName):
        print("FOLDER NAME : "+folderName)
        shutil.make_archive(folderName, 'zip', folderName)
    
    
    commacoordFileName=request.form['formFileSm']
    productPath=request.form['executeWebURL']
    testAccountPassword=request.form['testAccountPassword']
    if ',' in commacoordFileName:
        multiplecoordFileName = commacoordFileName.split(',')
    else:
        multiplecoordFileName = [commacoordFileName]
    localegroup=request.form.getlist('locale-group')
    notificationMail=request.form['notificationMail']
    langCode={
    # "ja_JP":["Kotoeri RomajiTyping"],
    # "ko_KR":["2SetKorean"],
    # "ko_KR231":["2SetKorean"],
    # "zh_CN":["Pinyin Simplified "],
    # "zh_TW1":["Cangjie Traditional"],
    # "zh_TW":["Cangjie Traditional"]
    "ja_JP":["com.apple.inputmethod.Kotoeri.RomajiTyping.Japanese","com.apple.inputmethod.Kotoeri.Japanese"],
    "Arabic":"com.apple.keylayout.Arabic-QWERTY",
    "ko_KR":["com.apple.inputmethod.Korean.2SetKorean","com.apple.inputmethod.Korean.2SetKorean"],
    "ko_KR231":["com.apple.inputmethod.Korean.2SetKorean","com.apple.inputmethod.Korean.2SetKorean"],
    # "zh_CN":["com.apple.inputmethod.TCIM.Pinyin","com.apple.inputmethod.TCIM.Pinyin"],
    "zh_CN":["com.apple.inputmethod.SCIM.ITABC","com.apple.inputmethod.SCIM.ITABC "],
    "zh_TW1":["com.apple.inputmethod.TCIM.Cangjie","com.apple.inputmethod.TCIM.Cangjie"],
    "zh_TW":["com.apple.inputmethod.TCIM.Cangjie","com.apple.inputmethod.TCIM.Cangjie"]
}
    for locale in localegroup:
        print("Insert Time...")
        getLangCode = verifyIMEpresent(locale)
        print(getLangCode)
        if(getLangCode==locale or len(getLangCode)<=1):
            # try:
            #     sendIMEFailedMail(productPath,notificationMail,localegroup)
            # except:
            #     pass
            #     #return "<h4 style='text-align:center'>You are Not in Adobe Network. Kindly Connect with GlobalProtect VPN</h4><div style='display: flex;justify-content: center;align-items: center;'><a href='/' style='background-color: #4CAF50; border: none;color: white;  padding:8px 30px; text-align: center;text-decoration: none;display: inline-block;font-size: 16px;margin: 4px 2px;cursor: pointer'>Home</a></div>"
            # os.system('pkill -f {}'.format(productPath))
            # # keyboard_listener.stop()
            # os.system("open /Applications/Google\ Chrome.app")
            # return [106,locale,langCode[locale]]
            try:
                # sendIMEFailedMail(productPath,notificationMail,localegroup)
                # print("./Utilities/changeInputSource/changeInputSource enable {}".format(langCode[locale][0]))
                # print("./Utilities/changeInputSource/changeInputSource enable {}".format(langCode[locale][1]))
                os.system(f"./Utilities/changeInputSource/changeInputSource enable {langCode[locale][0]}")
                os.system(f"./Utilities/changeInputSource/changeInputSource enable {langCode[locale][1]}")
            except:
                pass
            
            #return "<h4 style='text-align:center'>Input Method Source is Not Added, Kindly Add the Input source of {} Locale from <strong> System Preference > Keyboard </strong> </h4>".format(locale)+"<div style='display: flex;justify-content: center;align-items: center;'><a href='/' style='background-color: #4CAF50; border: none;color: white;  padding:8px 30px; text-align: center;text-decoration: none;display: inline-block;font-size: 16px;margin: 4px 2px;cursor: pointer'>Home</a></div>"
    for coordFileName in multiplecoordFileName:
        print(coordFileName)
        ## Coord File CHeck
        width,height=pyautogui.size()
        
        print(localegroup)
        IMEChange="Pass"
        TypingChars="Pass"
        appCrashed="False"
        productName=productPath.split('/')[2]
        userPath=os.path.expanduser('~')
        print(userPath)

        desktopPath=userPath+"/Desktop/"
        downloadPath=userPath+"/Downloads/"
        print(desktopPath)
        now=datetime.now()
        savePath=downloadPath+'IME Framework/{}'.format(now.strftime("%d-%m-%y-%H-%M-%S"))
        checkandCreateScratch(savePath)
        fileName=savePath+'/'+'Report-' + now.strftime("%d-%m-%y-%H-%M-%S") +'.xlsx'
        print(fileName)
        screenshotFolderName=savePath+'/Screenshots/'
        checkandCreateScratch(screenshotFolderName)
        def takeScreenshot(index):
            screen=pyautogui.screenshot()

            image_now=datetime.now()
            index=index.replace("?","")
            filename=screenshotFolderName+'/Screen-'+image_now.strftime("%H-%M-%S")+'-'+str(index).replace(" ",'')+'.png'
            scp_filename="https://imegeneric.corp.adobe.com/"+'IME Framework/{}'.format(now.strftime("%d-%m-%y-%H-%M-%S"))+'/Screenshots/'+'Screen-'+image_now.strftime("%H-%M-%S")+'-'+str(index).replace(" ",'')+'.png'
            
            screen.save(filename)
            # return filename
            return scp_filename

        f = open(coordFileName)
        data = json.load(f)
        # print(data)
        len_of_file=len(data)
        countMoved=0
        countPressed=0
        countReleased=0
        for i in range(len_of_file):
            
            if(data[i]["action"]=="moved"):
                countMoved+=1
            elif(data[i]["action"]=="pressed_key"):
                countPressed+=1
            elif(data[i]["action"]=="released_key"):
                countReleased+=1
        if(len_of_file > 2 and (data[0]["action"]=="moved" or data[len_of_file-1]["action"]=="released_key") and countPressed>=1 and countReleased>=1):
            print("Valid")
        else:
            print("Not Valid Coord File")
            f.close()
            return [102]
            return "<h4 style='text-align:center'> Sorry ! Some Issue With the Workflow File : " + coordFileName + "<br>Kindly Create Workflow Again <br></h4>" + "<div style='display: flex;justify-content: center;align-items: center;'><a href='/' style='  background-color: #4CAF50; border: none;color: white;  padding: 15px 32px; text-align: center;text-decoration: none;display: inline-block;font-size: 16px;margin: 4px 2px;cursor: pointer;'>Home</a></div>"
        f.close()
        cmd = """osascript -e 'tell application "QuickTime Player" to start (new screen recording)
        delay 1
        tell application "System Events" to key code 36
        '"""
        startSleepCount()
        os.system(cmd)


        getRandomProfile=random.randint(1,20)
        print("Chrome Profile", getRandomProfile)
        command = 'open -na "/Applications/Google Chrome.app" --args --user-data-dir=/tmp/NewProfile{0} --incognito --no-first-run --disable-prompt-on-repost --new-window {1}'.format(getRandomProfile,productPath)

        os.system(command)
        
        time.sleep(18)
        ##Here adding Screen in Full Size
        width,height=pyautogui.size()
        print(width,height,sep="-")
        time.sleep(0.5)
        JpEnglish=[False]
        KoEnglish=[False]
        CnEnglish=[False]
        TwEnglish=[False]

        testStringList=["asahi","gakkou"]
        name_of_recording = coordFileName
        number_of_plays = 1
        insert=[]
        output={}
        manualList=[]

        JPintlExcelList=['INTL 164 : Typing in different modes of JP IME','INTL 163 : Switch between different modes of JP IME','INTL 162 : Enable IME, switch between languages, type using IME','INTL 169 : Commit text when focus is moved','INTL 224 : Dakuon Characters (Accented)','INTL 167 : ESC/Delete Key']
        # JPswitchingimeresult=[None]
        # JPfocusmoving=[None]
        JPresults=[None]*len(JPintlExcelList)
        ## KOREAN Test Cases
        JP224TestStringlist=["ji","pa","ga"]
        JP167TestStringlist=["konotekisutowosakujosuru"]
        KO176TestStringlist=["rkatkgkqslek.","wkf wlsoskdy?","akssktj qksrkdnjdy!"]
        KO174TestStringlist=["Shift Q","Shift W","Shift E","Shift R","Shift T","Shift O"]
        KO231TestStringlist=["eogksalsrnrrhk qmfkwlfdml 16rkd wjsdl duffuTek."]
        KOintlExcelList=['Typing of characters in KO IME','Enable IME, switch between languages, type using IME',"INTL 176 : Korean Punctuations","INTL 174 : Korean Complex Consonant & Vowels"]

        KOresults=[None]*len(KOintlExcelList)
        row_index=0
        zhCNintlExcelList=['INTL 158 : Apostrophe - The Syllable Dividing Mark','INTL 157 : Special Characters','INTL 153 : Select Option from Candidate','INTL 156 : Enter Chinese Punctutaions and Symbols']
        zhCNresults=[None]*len(zhCNintlExcelList)
        zhCN158Stringlist=["xi'an","huan"]
        zhCN157Stringlist=["nv","nu","diannao"]

        zhTWintlExcelList=['INTL 184 : Input Chinese using Cangjie','INTL 180 : Enter Chinese Punctuations and Special Characters']
        zhTWresults=[None]*len(zhTWintlExcelList)
        zhTW184Stringlist=["tyhc","ci","smr","tzhc","tzhz"]
        primary_dict=json.load(open('testData/testData.json'))
        passwordUsed=[False]
        row_index=0

        for locale in localegroup:
            output[locale]=dict()
        with open(name_of_recording) as json_file:
            data = json.load(json_file)

        special_keys = {"Key.shift": Key.shift, "Key.tab": Key.tab, "Key.caps_lock": Key.caps_lock, "Key.ctrl": Key.ctrl, "Key.alt": Key.alt, "Key.cmd": Key.cmd, "Key.cmd_r": Key.cmd_r, "Key.alt_r": Key.alt_r, "Key.ctrl_r": Key.ctrl_r, "Key.shift_r": Key.shift_r, "Key.enter": Key.enter, "Key.backspace": Key.backspace, "Key.f19": Key.f19, "Key.f18": Key.f18, "Key.f17": Key.f17, "Key.f16": Key.f16, "Key.f15": Key.f15, "Key.f14": Key.f14, "Key.f13": Key.f13, "Key.media_volume_up": Key.media_volume_up, "Key.media_volume_down": Key.media_volume_down, "Key.media_volume_mute": Key.media_volume_mute, "Key.media_play_pause": Key.media_play_pause, "Key.f6": Key.f6, "Key.f5": Key.f5, "Key.right": Key.right, "Key.down": Key.down, "Key.left": Key.left, "Key.up": Key.up, "Key.page_up": Key.page_up, "Key.page_down": Key.page_down, "Key.home": Key.home, "Key.end": Key.end, "Key.delete": Key.delete, "Key.space": Key.space}

        mouse = MouseController()
        keyboard = KeyboardController()

        for loop in range(number_of_plays):
            for index, obj in enumerate(data):
                action, _time= obj['action'], obj['_time']
                try:
                    next_movement = data[index+1]['_time']
                    pause_time = next_movement - _time
                except IndexError as e:
                    pause_time = 1
            
                    
                if action == "pressed_key" or action == "released_key":
                    key = obj['key'] if 'Key.' not in obj['key'] else special_keys[obj['key']]
                    print("action: {0}, time: {1}, key: {2}".format(action, _time, str(key)))
                    if action == "released_key" and key == Key.shift:
                        print("Here inside Shift....")
                # Check if the next action is mouse movement or Enter key press
                        if index < len(data) - 1:
                            next_action = data[index + 1]['action']
                     
                            print("Before.... Here enter Password", next_action)
                            if next_action == "moved":
                                print("Here enter Password")
                                pyautogui.typewrite(testAccountPassword,interval=0.4)
                                passwordUsed=[True]

                    if(action=="pressed_key" and key=="r" and passwordUsed[0]==True):
                        manualList.append(["","","","",""])
                        for locale in localegroup:
                            print("Insert Time...")
                            getLangCode = verifyIMEpresent(locale)
                            print(getLangCode)
                            if(getLangCode=="" or len(getLangCode)<=1):
                                IMEChange="Fail"
                          
                                return [103]
                                return "<h1>Sorry, Currently IME Is Not Changing</h1>"
                    

                            os.system(f"./Utilities/changeInputSource/changeInputSource select {getLangCode}")
                            time.sleep(3)
                            if(locale=="ja_JP"):
                                checkOutput=jpTypingCheck()
                                if(checkOutput==True):
                                    # keyboard_listener.stop()
                                    JpEnglish[0]=True
                                    stop=""" osascript -e 'tell application "System Events" to key code 53 using {control down, command down}'"""
                                    os.system(stop)

                                    os.system("pkill -f 'QuickTime Player'")
                                    pyautogui.hotkey("command","w",interval=0.2)
                                    os.system("open /Applications/Google\ Chrome.app")
                                    os.system(f"./Utilities/changeInputSource/changeInputSource select com.apple.keylayout.ABC-India")
                                    os.system(f"./Utilities/changeInputSource/changeInputSource select com.apple.keylayout.ABC")
                                    return [111]
                              
                                for i in testStringList:
                                    temp=[]
                                    row_index+=1
                                    output[locale][i]=[]
                                    time.sleep(0.5)
                                    pyautogui.hotkey('command', 'a',interval=0.5)
                                    time.sleep(0.2)
                                    pyautogui.write(i,interval=0.5)
                                    
                                    pyautogui.press("enter")
                                    screenname=takeScreenshot(i)
                                    time.sleep(0.5)
                                    pyautogui.hotkey('command', 'a')
                                    pyautogui.hotkey('command', 'x')
                                            # insert.append(pyperclip.paste())
                                    rendered=pyperclip.paste()
                                    output[locale][i].append(pyperclip.paste())
                                    print("Current ",locale,output,sep="->")
                                
                                    try:
                                        expected=primary_dict[locale][i][0]
                                    except KeyError:
                                        expected=i
                                        
                                    temp=[locale+".Hiragana",i,rendered,expected,"Pass" if rendered in expected else "Fail",screenname]
                                    JPintlresult="Pass" if rendered in expected else "Fail"
                                    JPresults[0]=JPresults[1]=JPresults[2]=JPintlresult
                                    manualList.append(temp)
                                
                                time.sleep(0.5)
                                if(is_display_off()):
                                    print("System is in Sleep Mode... !!")
                                    stopFlask(notificationMail,productPath,localegroup)
                                    return [105]
                                for jp224string in JP224TestStringlist:
                                    temp=[]
                                    row_index+=1
                                    output[locale][jp224string]=[]
                                    time.sleep(0.5)
                                    pyautogui.hotkey('command', 'a',interval=0.5)
                                    time.sleep(0.2)
                                    pyautogui.write(jp224string,interval=0.5)
                                    pyautogui.press('enter')
                                    screenname=takeScreenshot(jp224string)
                                    pyautogui.hotkey('command', 'a')
                                    pyautogui.hotkey('command', 'x')
                                    rendered=pyperclip.paste()
                                    output[locale][jp224string].append(pyperclip.paste())
                                    print("Current ",locale,output,sep="->")
                                            
                                    try:
                                        expected=str(primary_dict[locale][jp224string])
                                    except KeyError:
                                        expected=jp224string
                                            
                                    temp=[locale,jp224string,rendered,expected,"Pass" if rendered == expected else "Fail",screenname]
                                    JPresults[4]="Pass" if rendered == expected else "Fail"
                                    manualList.append(temp)
                                    time.sleep(0.5)
                                # time.sleep(0.5)
                                if(is_display_off()):
                                    print("System is in Sleep Mode... !!")
                                    stopFlask(notificationMail,productPath,localegroup)
                                    return [105]

                                time.sleep(0.5)
                                pyperclip.copy('')
                                for jp167string in JP167TestStringlist:
                                    temp=[]
                                    row_index+=1
                                    output[locale][jp167string]=[]
                                    time.sleep(0.5)
                                    pyautogui.hotkey('command', 'a',interval=0.5)
                                    time.sleep(0.2)
                                    pyautogui.write(jp167string,interval=0.5)
                                        
                                    time.sleep(0.7)
                                    pyautogui.press("space")
                                    time.sleep(0.7)
                                    pyautogui.press("esc")
                                    time.sleep(0.5)
                                    # pyautogui.click()
                                    # time.sleep(0.3)
                                    pyautogui.press("enter")
                                    time.sleep(0.2)
                                    screenname=takeScreenshot(jp167string)
                                    pyautogui.hotkey('command', 'a')
                                    pyautogui.hotkey('command', 'x')
                                    rendered=pyperclip.paste()
                                    output[locale][jp167string].append(pyperclip.paste())
                                    print("Current ",locale,output,sep="->")
                                        
                                    try:
                                        expected=str(primary_dict[locale][jp167string])
                                    except KeyError:
                                        expected=jp167string
                                        
                                    temp=[locale,jp167string,rendered,expected,"Pass" if rendered == expected else "Fail",screenname]
                                    JPresults[5]="Pass" if rendered == expected else "Fail"
                                    manualList.append(temp)
                                    time.sleep(0.5)
                                if(is_display_off()):
                                    print("System is in Sleep Mode... !!")
                                    stopFlask(notificationMail,productPath,localegroup)
                                    return [105]
                                pyperclip.copy('')
                                for i in testStringList:
                                    pyperclip.copy('')
                                    temp=[]
                                    row_index+=1
                                    output[locale][i]=[]
                                    pyautogui.hotkey('command', 'a',interval=0.5)
                                    time.sleep(0.2)
                                    pyautogui.write(i,interval=0.5)
                                    
                                        # pyautogui.press("enter")
                                    os.system("open '{}'".format("Downloads")) # Switch to Product
                                    # os.system('open /Applications/Google\ Chrome.app') # Switch to Chrome
                                    
                                    time.sleep(1)
                                    screenname=takeScreenshot(i)
                                    time.sleep(0.5)
                                    pyautogui.click()
                                    time.sleep(0.5)
                                    pyautogui.hotkey('command', 'a')
                                    time.sleep(0.5)
                                    pyautogui.hotkey('command', 'x')
                                            # insert.append(pyperclip.paste())
                                    rendered=pyperclip.paste()
                                    output[locale][i].append(pyperclip.paste())
                                    print("Current ",locale,output,sep="->")
                                        # sheet1.write(row_index,0,locale)
                                        # sheet1.write(row_index,1,i)
                                        # sheet1.write(row_index,2,rendered)
                                    try:
                                        expected=primary_dict[locale][i][0]
                                    except KeyError:
                                        expected=i
                                        # sheet1.write(row_index,3,expected[0])
                                        # sheet1.write(row_index,4,"Pass" if rendered in expected else "Fail")
                                    temp=[locale + ".Hiragana",i,rendered,expected,"Pass" if rendered in expected else "Fail",screenname]
                                    JPintlresult="Pass" if rendered in expected else "Fail"
                                    JPresults[3]=JPintlresult
                                    manualList.append(temp)
                                    time.sleep(0.5)

    ## Korean Test Cases        
                            
                            elif(locale=="ko_KR"):
                                checkOutput=koTypingCheck()
                                if(checkOutput==True):
                                    # keyboard_listener.stop()
                                    KoEnglish[0]=True
                                    stop=""" osascript -e 'tell application "System Events" to key code 53 using {control down, command down}'"""
                                    os.system(stop)

                                    os.system("pkill -f 'QuickTime Player'")
                                    pyautogui.hotkey("command","w",interval=0.2)
                                    os.system("open /Applications/Google\ Chrome.app")
                                    os.system(f"./Utilities/changeInputSource/changeInputSource select com.apple.keylayout.ABC-India")
                                    os.system(f"./Utilities/changeInputSource/changeInputSource select com.apple.keylayout.ABC")
                                    return [111]
                                pyautogui.click()
                                for koString in KO176TestStringlist:
                                    temp=[]
                                    row_index+=1
                                    output[locale][koString]=[]
                                    time.sleep(0.5)
                                    pyautogui.hotkey('command', 'a',interval=0.5)
                                    time.sleep(0.2)
                                    pyautogui.write(koString,interval=0.5)
                                    screenname=takeScreenshot(koString)
                                    pyautogui.hotkey('command', 'a')
                                    pyautogui.hotkey('command', 'x')
                                                # insert.append(pyperclip.paste())
                                    rendered=pyperclip.paste()
                                    output[locale][koString].append(pyperclip.paste())
                                    print("Current ",locale,output,sep="->")
                                    
                                    try:
                                        expected=str(primary_dict[locale][koString])
                                    except KeyError:
                                        expected=koString
                                
                                    temp=[locale,koString,rendered,expected,"Pass" if rendered == expected else "Fail",screenname]
                                    KOresults[0]=KOresults[1]=KOresults[2]="Pass" if rendered == expected else "Fail"
                                    manualList.append(temp)
                                    time.sleep(0.5)
                                # time.sleep(0.5)
                                if(is_display_off()):
                                    print("System is in Sleep Mode... !!")
                                    stopFlask(notificationMail,productPath,localegroup)
                                    return [105]
                                for koString2 in KO174TestStringlist:
                                    temp=[]
                                    row_index+=1
                                    output[locale][koString2]=[]
                                    time.sleep(0.5)
                                    pyautogui.hotkey('command', 'a',interval=0.5)
                                    time.sleep(0.2)
                                    hotkeys=koString2.split()
                                    pyautogui.hotkey(hotkeys[0],hotkeys[1],interval=0.5)
                                    screenname=takeScreenshot(koString2)
                                    pyautogui.press('space')
                                    pyautogui.hotkey('command', 'a')
                                    pyautogui.hotkey('command', 'x')
                                                # insert.append(pyperclip.paste())
                                    rendered=pyperclip.paste()
                                    rendered=rendered.strip()
                                    output[locale][koString2].append(pyperclip.paste())
                                    print("Current ",locale,output,sep="->")
                                    
                                    try:
                                        expected=str(primary_dict[locale][koString2])
                                    except KeyError:
                                        expected=koString2
                                        
                                    temp=[locale,koString2,rendered,expected,"Pass" if rendered == expected else "Fail",screenname]
                                    KOresults[3]="Pass" if rendered == expected else "Fail"
                                    manualList.append(temp)
                                    time.sleep(0.5)
                                
                                if(is_display_off()):
                                    print("System is in Sleep Mode... !!")
                                    stopFlask(notificationMail,productPath,localegroup)
                                    return [105]
                            
                            ## Korean 231 Only
                            elif(locale=="ko_KR231"):
                                KOintlExcelList.append("INTL 231 : Enter key for Line Break")
                                for ko231String in KO231TestStringlist:
                                    temp=[]
                                    row_index+=1
                                    output[locale][ko231String]=[]
                                    time.sleep(0.5)
                                    pyautogui.hotkey('command', 'a',interval=0.5)
                                    time.sleep(0.2)
                                    pyautogui.write(ko231String,interval=0.5)
                                    time.sleep(0.3)
                                    pyautogui.press("enter")
                                    time.sleep(0.3)
                                    pyautogui.write(ko231String,interval=0.5)
                                    screenname=takeScreenshot(ko231String)
                                    pyautogui.hotkey('command', 'a')
                                    pyautogui.hotkey('command', 'x')
                                                # insert.append(pyperclip.paste())
                                    rendered=pyperclip.paste()
                                    print("KO 231 : ", rendered)
                                    output[locale][ko231String].append(pyperclip.paste())
                                    print("Current ",locale,output,sep="->")
                                        
                                    try:
                                        expected=str(primary_dict[locale][ko231String])
                                    except KeyError:
                                        expected=ko231String
                                        
                                    temp=[locale,ko231String,rendered,expected,"Pass" if rendered == expected else "Fail",screenname]
                                    KOresults.append("Pass" if rendered == expected else "Fail")
                                    manualList.append(temp)
                                    time.sleep(0.5)
                                time.sleep(0.5)

                                

                            ## Chinese Simplified
                            elif(locale=="zh_CN"):
                                checkOutput=cnTypingCheck()
                                if(checkOutput==True):
                                    # keyboard_listener.stop()
                                    CnEnglish[0]=True
                                    stop=""" osascript -e 'tell application "System Events" to key code 53 using {control down, command down}'"""
                                    os.system(stop)

                                    os.system("pkill -f 'QuickTime Player'")
                                    pyautogui.hotkey("command","w",interval=0.2)
                                    os.system("open /Applications/Google\ Chrome.app")
                                    os.system(f"./Utilities/changeInputSource/changeInputSource select com.apple.keylayout.ABC-India")
                                    os.system(f"./Utilities/changeInputSource/changeInputSource select com.apple.keylayout.ABC")
                                    return [111]
                                pyautogui.click()
                                pyautogui.press('space')

                                #pass ## 158,157,153
                                tmp158pass=[]
                                for zhCN158String in zhCN158Stringlist:
                                    temp=[]
                                    row_index+=1
                                    output[locale][zhCN158String]=[]
                                    time.sleep(0.5)
                                    pyautogui.hotkey('command', 'a',interval=0.5)
                                    time.sleep(0.2)
                                    pyautogui.write(zhCN158String,interval=0.5)
                                    time.sleep(0.5)
                                    pyautogui.press("space")
                                    time.sleep(0.5)
                                    screenname=takeScreenshot(zhCN158String)
                                    pyautogui.hotkey('command', 'a')
                                    pyautogui.hotkey('command', 'x')
                                                
                                    rendered=pyperclip.paste()
                                    output[locale][zhCN158String].append(pyperclip.paste())
                                    print("Current ",locale,output,sep="->")
                                        
                                    try:
                                        expected=str(primary_dict[locale][zhCN158String])
                                    except KeyError:
                                        expected=zhCN158String
                                        
                                    temp=[locale,zhCN158String,rendered,expected,"Pass" if rendered == expected else "Fail",screenname]
                                    zhCNresults[0]="Pass" if rendered == expected else "Fail"
                                    tmp158pass.append(zhCNresults[0])
                                    if("Pass" in tmp158pass):
                                        zhCNresults[2]="Pass"
                                    else:
                                        zhCNresults[2]="Fail"
                                    manualList.append(temp)
                                    time.sleep(0.5)
                                # time.sleep(0.5)
                                if(is_display_off()):
                                    print("System is in Sleep Mode... !!")
                                    stopFlask(notificationMail,productPath,localegroup)
                                    return [105]
                                for zhCN157String in zhCN157Stringlist:
                                    temp=[]
                                    row_index+=1
                                    output[locale][zhCN157String]=[]
                                    time.sleep(0.5)
                                    pyautogui.hotkey('command', 'a',interval=0.5)
                                    time.sleep(0.2)
                                    pyautogui.write(zhCN157String,interval=0.5)
                                    time.sleep(0.5)
                                    pyautogui.press("space")
                                    time.sleep(0.5)
                                    screenname=takeScreenshot(zhCN157String)
                                    pyautogui.hotkey('command', 'a')
                                    pyautogui.hotkey('command', 'x')
                                                # insert.append(pyperclip.paste())
                                    rendered=pyperclip.paste()
                                    output[locale][zhCN157String].append(pyperclip.paste())
                                    print("Current ",locale,output,sep="->")
                                
                                    try:
                                        expected=str(primary_dict[locale][zhCN157String])
                                    except KeyError:
                                        expected=zhCN157String
                                        
                                    temp=[locale,zhCN157String,rendered,expected,"Pass" if rendered == expected else "Fail",screenname]
                                    spacePass="Pass" if rendered == expected else "Fail"
                                    
                                    manualList.append(temp)
                                    
                                    time.sleep(0.5)
                                    pyautogui.hotkey('command', 'a',interval=0.5)
                                    time.sleep(0.2)
                                    pyautogui.write(zhCN157String,interval=0.5)
                                    time.sleep(0.5)
                                    pyautogui.press("enter")
                                    time.sleep(0.5)
                                    screenname=takeScreenshot(zhCN157String)
                                    pyautogui.hotkey('command', 'a')
                                    pyautogui.hotkey('command', 'x')
                                                
                                    rendered=pyperclip.paste()
                                    enterPass="Pass" if rendered == zhCN157String else "Fail"
                                    if(spacePass=="Pass" and enterPass=="Pass"):
                                        zhCNresults[1]="Pass"
                                    else:
                                        zhCNresults[1]="Fail"
                                    time.sleep(0.5)
                                # time.sleep(0.5)
                                if(is_display_off()):
                                    print("System is in Sleep Mode... !!")
                                    stopFlask(notificationMail,productPath,localegroup)
                                    return [105]
                                ## 153 INTL Chinese Punctuations
                                pyautogui.hotkey('option','shift','b',interval=0.5)
                                time.sleep(0.2)
                                pyautogui.press("down")
                                time.sleep(0.2)
                                screenname=takeScreenshot("zhCN153-Symbols")
                                pyautogui.press('enter')
                                time.sleep(0.2)
                                pyautogui.hotkey('command', 'a')
                                pyautogui.hotkey('command', 'x')
                                rendered=pyperclip.paste()
                                if(rendered!=""):
                                    zhCNresults[3]="Pass"
                                else:
                                    zhCNresults[3]="Fail"
                                time.sleep(0.5)
                            

                            
                            ## Chinese Tradiotion zh_TW
                            elif(locale=="zh_TW"):
                                checkOutput=twTypingCheck()
                                if(checkOutput==True):
                                    # keyboard_listener.stop()
                                    TwEnglish[0]=True
                                    stop=""" osascript -e 'tell application "System Events" to key code 53 using {control down, command down}'"""
                                    os.system(stop)

                                    os.system("pkill -f 'QuickTime Player'")
                                    pyautogui.hotkey("command","w",interval=0.2)
                                    os.system("open /Applications/Google\ Chrome.app")
                                    os.system(f"./Utilities/changeInputSource/changeInputSource select com.apple.keylayout.ABC-India")
                                    os.system(f"./Utilities/changeInputSource/changeInputSource select com.apple.keylayout.ABC")
                                    return [111]
                                pyautogui.click()
                                pyautogui.press('space')
                                for zhTW184String in zhTW184Stringlist:
                                    temp=[]
                                    row_index+=1
                                    output[locale][zhTW184String]=[]
                                    time.sleep(0.5)
                                    pyautogui.hotkey('command', 'a',interval=0.5)
                                    time.sleep(0.2)
                                    pyautogui.write(zhTW184String,interval=0.5)
                                    time.sleep(0.5)
                                    
                                    screenname=takeScreenshot(zhTW184String)
                                    pyautogui.hotkey('command', 'a')
                                    pyautogui.hotkey('command', 'x')
                                                # insert.append(pyperclip.paste())
                                    rendered=pyperclip.paste()
                                    output[locale][zhTW184String].append(pyperclip.paste())
                                    print("Current ",locale,output,sep="->")
                                    
                                    try:
                                        expected=str(primary_dict[locale][zhTW184String])
                                    except KeyError:
                                        expected=zhTW184String
                                        
                                    temp=[locale,zhTW184String,rendered,expected,"Pass" if rendered == expected else "Fail",screenname]
                                    zhTWresults[0]="Pass" if rendered == expected else "Fail"
                                    manualList.append(temp)
                                    time.sleep(0.5)
                                # time.sleep(0.5)
                                if(is_display_off()):
                                    print("System is in Sleep Mode... !!")
                                    stopFlask(notificationMail,productPath,localegroup)
                                    return [105]
                                    ## 180 INTL Chinese Punctuations
                                pyautogui.hotkey('command', 'a')
                                time.sleep(0.2)
                                pyautogui.hotkey('option','shift','b')
                                time.sleep(0.2)
                                pyautogui.press("down")
                                time.sleep(0.2)
                                screenname=takeScreenshot("zhTW180-Symbols")
                                pyautogui.press('enter')
                                time.sleep(0.2)
                                
                                pyautogui.hotkey('option','shift','b',interval=0.5)
                                time.sleep(0.2)
                                pyautogui.press("down")
                                time.sleep(0.2)
                                screenname=takeScreenshot("zhTW180-Symbols")
                                pyautogui.press('enter')
                                time.sleep(0.2)
                                pyautogui.hotkey('command', 'a')
                                pyautogui.hotkey('command', 'x')
                                rendered=pyperclip.paste()
                                if(rendered!=""):
                                    zhTWresults[1]="Pass"
                                else:
                                    zhTWresults[1]="Fail"
                                time.sleep(0.5)
                            time.sleep(1)
                        os.system(f"./Utilities/changeInputSource/changeInputSource select com.apple.keylayout.ABC-India")
                        os.system(f"./Utilities/changeInputSource/changeInputSource select com.apple.keylayout.ABC")
                        time.sleep(0.6)


                    elif(action=="released_key" and key=="r"):
                        print("Fetch Time...")
                        
                    elif action == "pressed_key":
                        keyboard.press(key)
                    else:
                        keyboard.release(key)
                    time.sleep(pause_time)


                else:
                    move_for_scroll = True
                    x, y = obj['x'], obj['y']
                    if action == "scroll" and index > 0 and (data[index - 1]['action'] == "pressed" or data[index - 1]['action'] == "released"):
                        if x == data[index - 1]['x'] and y == data[index - 1]['y']:
                            move_for_scroll = False
                    print("x: {0}, y: {1}, action: {2}, time: {3}".format(x, y, action, _time))
                    mouse.position = (x, y)
                    if action == "pressed" or action == "released" or action == "scroll" and move_for_scroll == True:
                        time.sleep(0.1)
                    if action == "pressed":
                        mouse.press(Button.left if obj['button'] == "Button.left" else Button.right)
                    elif action == "released":
                        mouse.release(Button.left if obj['button'] == "Button.left" else Button.right)
                    elif action == "scroll":
                        horizontal_direction, vertical_direction = obj['horizontal_direction'], obj['vertical_direction']
                        mouse.scroll(horizontal_direction, vertical_direction)
                    time.sleep(pause_time)
        os.system(f"./Utilities/changeInputSource/changeInputSource select com.apple.keylayout.ABC-India")
        os.system(f"./Utilities/changeInputSource/changeInputSource select com.apple.keylayout.ABC")
            ## END QT
        stop=""" osascript -e 'tell application "System Events" to key code 53 using {control down, command down}'"""
        os.system(stop)

        os.system("pkill -f 'QuickTime Player'")

        generateExcelFromListJP_KO_CN_TW(manualList,fileName,JPintlExcelList,JPresults,KOintlExcelList,KOresults,zhCNintlExcelList,zhCNresults,zhTWintlExcelList,zhTWresults)

        print(output)
        print("FOLDER NAME : "+screenshotFolderName)
        # createZipofFolder(screenshotFolderName)
        pyautogui.hotkey("command","w",interval=0.2)
        os.system("open /Applications/Google\ Chrome.app")
        # os.system("pkill -f '{}'".format(productPath))
        screenshot=screenshotFolderName+'.zip'
        try:
                source=screenshotFolderName
                parentSource=screenshotFolderName[:-1]
                print(os.path.dirname(parentSource))
                newsource=os.path.dirname(parentSource).strip()
                print("Source : ",source)
                print("New Source : ",newsource)
                destination='C:\inetpub\wwwroot\IME Framework'
                print("Destination : ", destination)
                try :
                    new_copy_folder_via_scp(newsource, destination)
                    
                except Exception as e:
                    print(e.__doc__)
                    print (e)
                newWebsendonereportMail(productPath,notificationMail,localegroup,fileName,IMEChange,TypingChars,appCrashed,name_of_recording)
                productSubject=urlparse(productPath).netloc
                dbEntryList=[productSubject,notificationMail,"Workflow",str(localegroup),"Utility","Desktop"]
                insertLogDatainDB(dbEntryList)
                # return "<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css' integrity='sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm' crossorigin='anonymous'> \
                #     <body style='background-color:17202A'>\
                #     <div class='d-flex align-items-center justify-content-center' style='margin:41px;'>\
                #     <br>"+"<div class='alert alert-success' role='alert'><h4 class='alert-heading'>WoW !!</h4><hr>\
                #         <h6 style='color:black'> Complete Processing of Build : "+ productPath  + "<br><br>"+"Send Detailed Report to : " + notificationMail +"</h4>" "</h6><div style='display: flex;justify-content: center;align-items: center;'><a href='/download?filename={0}'>Download Report</a></div><div style='display: flex;justify-content: center;align-items: center;'><br><a href='/openfolder?filename={1}' target='_blank'>Open Screen Folder</a></div><div style='display: flex;justify-content: center;align-items: center;'><a href='/' style='background-color: #4CAF50; border: none;color: white;  padding:8px 30px; text-align: center;text-decoration: none;display: inline-block;font-size: 16px;margin: 4px 2px;cursor: pointer'>Home</a></div>".format(fileName,screenshotFolderName)+"</div></div></body>"
            
            
        except:
                source=screenshotFolderName
                parentSource=screenshotFolderName[:-1]
                print(os.path.dirname(parentSource))
                newsource=os.path.dirname(parentSource).strip()
                print("Source : ",source)
                print("New Source : ",newsource)
                destination='C:\inetpub\wwwroot\IME Framework'
                print("Destination : ", destination)
                try :
                    new_copy_folder_via_scp(newsource, destination)
                    
                except Exception as e:
                    print(e.__doc__)
                    print (e)
                print(fileName)
            
    return [107,productPath,notificationMail,localegroup,multiplecoordFileName]


