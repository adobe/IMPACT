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
def desktopliveime1():

    productPath=request.form['productName1']
    langCode={

    "ja_JP":["com.apple.inputmethod.Kotoeri.RomajiTyping.Japanese","com.apple.inputmethod.Kotoeri.Japanese"],
    "Arabic":"com.apple.keylayout.Arabic-QWERTY",
    "ko_KR":["com.apple.inputmethod.Korean.2SetKorean","com.apple.inputmethod.Korean.2SetKorean"],
    "ko_KR231":["com.apple.inputmethod.Korean.2SetKorean","com.apple.inputmethod.Korean.2SetKorean"],
    # "zh_CN":["com.apple.inputmethod.TCIM.Pinyin","com.apple.inputmethod.TCIM.Pinyin"],
    "zh_CN":["com.apple.inputmethod.SCIM.ITABC","com.apple.inputmethod.SCIM.ITABC "],
    "zh_TW1":["com.apple.inputmethod.TCIM.Cangjie","com.apple.inputmethod.TCIM.Cangjie"],
    "zh_TW":["com.apple.inputmethod.TCIM.Cangjie","com.apple.inputmethod.TCIM.Cangjie"]
}
    # coordFileName=request.form['formFileSm']
    localegroup=request.form.getlist('locale-group')
    notificationMail=request.form['notificationMail']
    for locale in localegroup:
        print("Insert Time...")
        getLangCode = verifyIMEpresent(locale)
        print(getLangCode)
        if(getLangCode==locale or len(getLangCode)<=1):
            try:
   
                os.system(f"./Utilities/changeInputSource/changeInputSource enable {langCode[locale][0]}")
                os.system(f"./Utilities/changeInputSource/changeInputSource enable {langCode[locale][1]}")
            except:
                pass
                # return [100]
                #return "<h4 style='text-align:center'>You are Not in Adobe Network. Kindly Connect with GlobalProtect VPN or Join Adobe Network</h4><div style='display: flex;justify-content: center;align-items: center;'><a href='/' style='background-color: #4CAF50; border: none;color: white;  padding:8px 30px; text-align: center;text-decoration: none;display: inline-block;font-size: 16px;margin: 4px 2px;cursor: pointer'>Home</a></div>"
        #     os.system('pkill -f {}'.format(productPath))
        #     # keyboard_listener.stop()
        #     os.system("open /Applications/Google\ Chrome.app")
        #     return [101,locale,langCode[locale]]
        #     return "<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css' integrity='sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm' crossorigin='anonymous'> \
        #     <body style='background-color:17202A'>\
        #     <div class='d-flex align-items-center justify-content-center' style='margin:41px'><br>\
        #     <br>"+"<div class='alert alert-warning' role='alert'><h4 class='alert-heading'>Holy guacamole!!</h4>\
        # <p>Input Method Source is Not Added. Kindly Add the Input source of <strong>{}</strong> Locale from <strong> System Preference > Keyboard > {}</strong> </h4>".format(locale,langCode[locale])+"<div style='display: flex;justify-content: center;align-items: center;'><a href='/' style='background-color: #4CAF50; border: none;color: white;  padding:8px 30px; text-align: center;text-decoration: none;display: inline-block;font-size: 16px;margin: 4px 2px;cursor: pointer'>Home</a></div>'\
        #     </p>\
        # </div></div></body>"
    print(localegroup)
    IMEChange="Pass"
    TypingChars="Pass"
    appCrashed="False"
    storage = []
    count = 0
    record_all=True
    if(os.path.exists(productPath)==False):
        return [102]
        return "<h3 style='text-align:center;'> Sorry Please Check Product Path <br> Product Path Not Found </h3>"
    
        ## ADDING QT INTEGRATION
    cmd = """osascript -e 'tell application "QuickTime Player" to start (new screen recording)
    delay 1
    tell application "System Events" to key code 36
    '"""

    os.system(cmd)
    os.system("open '{}'".format(productPath))
    time.sleep(2)
    testStringList=["asahi","gakkou"]
    # name_of_recording = coordFileName
    number_of_plays = 1
    output={}
    for locale in localegroup:
        output[locale]=dict()
    now = datetime.now()
    primary_dict=json.load(open('testData/testData.json'))

    userPath=os.path.expanduser('~')
    print(userPath)

    desktopPath=userPath+"/Desktop/"
    downloadPath=userPath+"/Downloads/"
    print(desktopPath)
    savePath=downloadPath+'IME Framework/{}'.format(now.strftime("%d-%m-%y-%H-%M-%S"))
    checkandCreateScratch(savePath)
    fileName=savePath+'/'+'Report-' + now.strftime("%d-%m-%y-%H-%M-%S") +'.xlsx'
    print(fileName)
    # wb = xlsxwriter.Workbook(fileName)
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
    # KOswitchingimeresult=[None]
    # KOfocusmoving=[None]
    KOresults=[None]*len(KOintlExcelList)
    row_index=0
    zhCNintlExcelList=['INTL 158 : Apostrophe - The Syllable Dividing Mark','INTL 157 : Special Characters','INTL 153 : Select Option from Candidate','INTL 156 : Enter Chinese Punctutaions and Symbols']
    zhCNresults=[None]*len(zhCNintlExcelList)
    zhCN158Stringlist=["xi'an","huan"]
    zhCN157Stringlist=["nv","nu","diannao"]

    zhTWintlExcelList=['INTL 184 : Input Chinese using Cangjie','INTL 180 : Enter Chinese Punctuations and Special Characters']
    zhTWresults=[None]*len(zhTWintlExcelList)
    zhTW184Stringlist=["tyhc","ci","smr","tzhc","tzhz"]

    productName=productPath.split('/')[2]
    productSubject=productPath.split('/')[-1]
    screenshotFolderName=savePath+'/Screenshots/'
    JpEnglish=[False]
    KoEnglish=[False]
    CnEnglish=[False]
    TwEnglish=[False]
    checkandCreateScratch(screenshotFolderName)
    def takeScreenshot(index):
        image_now=datetime.now()
        screen=pyautogui.screenshot()
    
        index=index.replace("?","")
        filename=screenshotFolderName+'/Screen-'+image_now.strftime("%H-%M-%S")+'-'+str(index).replace(" ",'')+'.png'
        scp_filename="https://imegeneric.corp.adobe.com/"+'IME Framework/{}'.format(now.strftime("%d-%m-%y-%H-%M-%S"))+'/Screenshots/'+'Screen-'+image_now.strftime("%H-%M-%S")+'-'+str(index).replace(" ",'')+'.png'
        
        screen.save(filename)
        # return filename
        return scp_filename
    
    def on_press(key):
        global row_index
        row_index=0
        
        try:
            json_object = {'action':'pressed_key', 'key':key.char, '_time': time.time()}
            if key.char == 'm':
                manualList.append(["","","","",""])
                for locale in localegroup:
                        print("Insert Time...")
                        getLangCode = verifyIMEpresent(locale)
                        print(getLangCode)
                        if(getLangCode=="" or len(getLangCode)<=1):
                            IMEChange="Fail"
                            sendIMEFailedMail(productPath,notificationMail,localegroup)
                            os.system('pkill -f {}').format(productPath)
                            return "<h1>Sorry, Currently IME Is Not Changing</h1>"
                    

                        os.system(f"./Utilities/changeInputSource/changeInputSource select {getLangCode}")
                        time.sleep(3)
                        if(locale=="ja_JP"):
                            checkOutput=jpTypingCheck()
                            if(checkOutput==True):
                                keyboard_listener.stop()
                                JpEnglish[0]=True
                                return [111]
                    
                            for i in testStringList:
                                temp=[]
                                row_index+=1
                                output[locale][i]=[]
                                time.sleep(1)
                                pyautogui.hotkey('command', 'a',interval=0.5)
                                time.sleep(0.5)
                                pyautogui.write(i,interval=0.5)
                                
                                pyautogui.press("enter")
                                screenname=takeScreenshot(i)
                                time.sleep(1)
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
                            
                            time.sleep(1)
                            for jp224string in JP224TestStringlist:
                                temp=[]
                                row_index+=1
                                output[locale][jp224string]=[]
                                time.sleep(1)
                                pyautogui.hotkey('command', 'a',interval=0.5)
                                time.sleep(0.5)
                                pyautogui.write(jp224string,interval=0.5)
                                pyautogui.press("enter")
                                screenname=takeScreenshot(jp224string)
                                pyautogui.hotkey('command', 'a')
                                pyautogui.hotkey('command', 'x')
                                rendered=pyperclip.paste()
                                output[locale][jp224string].append(pyperclip.paste())
                                print("Current ",locale,output,sep="->")
                                        # sheet1.write(row_index,0,locale)
                                        # sheet1.write(row_index,1,i)
                                        # sheet1.write(row_index,2,rendered)
                                try:
                                    expected=str(primary_dict[locale][jp224string])
                                except KeyError:
                                    expected=jp224string
                                        # sheet1.write(row_index,3,expected)
                                        # sheet1.write(row_index,4,"Pass" if rendered==expected else "Fail")
                                temp=[locale,jp224string,rendered,expected,"Pass" if rendered == expected else "Fail",screenname]
                                JPresults[4]="Pass" if rendered == expected else "Fail"
                                manualList.append(temp)
                                time.sleep(1)
                            time.sleep(1)
                            pyperclip.copy('')
                            for jp167string in JP167TestStringlist:
                                temp=[]
                                row_index+=1
                                output[locale][jp167string]=[]
                                time.sleep(1)
                                pyautogui.hotkey('command', 'a',interval=0.5)
                                time.sleep(0.5)
                                pyautogui.write(jp167string,interval=0.5)
                                    
                                time.sleep(0.7)
                                pyautogui.press("space")
                                time.sleep(0.7)
                                pyautogui.press("esc")
                                time.sleep(1)
                                pyautogui.click()
                                time.sleep(0.3)
                                pyautogui.press("enter")
                                time.sleep(0.5)
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
                                        # sheet1.write(row_index,3,expected)
                                        # sheet1.write(row_index,4,"Pass" if rendered==expected else "Fail")
                                temp=[locale,jp167string,rendered,expected,"Pass" if rendered == expected else "Fail",screenname]
                                JPresults[5]="Pass" if rendered == expected else "Fail"
                                manualList.append(temp)
                                time.sleep(1)
                            
                            time.sleep(1)
                            pyperclip.copy('')
                            for i in testStringList:
                                pyperclip.copy('')
                                temp=[]
                                row_index+=1
                                output[locale][i]=[]
                                pyautogui.hotkey('command', 'a',interval=0.5)
                                time.sleep(0.5)
                                pyautogui.write(i,interval=0.5)
                                
                                    # pyautogui.press("enter")
                                os.system('open /Applications/Google\ Chrome.app') # Switch to Chrome
                                os.system("open '{}'".format(productPath)) # Switch to Product
                                time.sleep(1)
                                screenname=takeScreenshot(i)
                                time.sleep(1)
                                pyautogui.click()
                                time.sleep(1)
                                pyautogui.hotkey('command', 'a')
                                time.sleep(1)
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
                                time.sleep(1)
                        
## Korean Test Cases        
                        elif(locale=="ko_KR"):
                            checkOutput=koTypingCheck()
                            if(checkOutput==True):
                                keyboard_listener.stop()
                                KoEnglish[0]=True
                                return [111]
                            for koString in KO176TestStringlist:
                                temp=[]
                                row_index+=1
                                output[locale][koString]=[]
                                time.sleep(1)
                                pyautogui.hotkey('command', 'a',interval=0.5)
                                time.sleep(0.5)
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
                                time.sleep(1)
                            time.sleep(1)
                            for koString2 in KO174TestStringlist:
                                temp=[]
                                row_index+=1
                                output[locale][koString2]=[]
                                time.sleep(1)
                                pyautogui.hotkey('command', 'a',interval=0.5)
                                time.sleep(0.5)
                                hotkeys=koString2.split()
                                pyautogui.hotkey(hotkeys[0],hotkeys[1],interval=0.5)
                                screenname=takeScreenshot(koString2)
                                pyautogui.hotkey('command', 'a')
                                pyautogui.hotkey('command', 'x')
                                            # insert.append(pyperclip.paste())
                                rendered=pyperclip.paste()
                                output[locale][koString2].append(pyperclip.paste())
                                print("Current ",locale,output,sep="->")
                                
                                try:
                                    expected=str(primary_dict[locale][koString2])
                                except KeyError:
                                    expected=koString2
                                    
                                temp=[locale,koString2,rendered,expected,"Pass" if rendered == expected else "Fail",screenname]
                                KOresults[3]="Pass" if rendered == expected else "Fail"
                                manualList.append(temp)
                                time.sleep(1)
                            time.sleep(1)
                        ## Korean 231 Only
                        elif(locale=="ko_KR231"):
                            KOintlExcelList.append("INTL 231 : Enter key for Line Break")
                            for ko231String in KO231TestStringlist:
                                temp=[]
                                row_index+=1
                                output[locale][ko231String]=[]
                                time.sleep(1)
                                pyautogui.hotkey('command', 'a',interval=0.5)
                                time.sleep(0.5)
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
                                time.sleep(1)
                            time.sleep(1)

                            # print("Will Exceute Later...")

                        ## Chinese Simplified
                        elif(locale=="zh_CN"):
                            checkOutput=cnTypingCheck()
                            if(checkOutput==True):
                                keyboard_listener.stop()
                                CnEnglish[0]=True
                                return [111]
                            tmp158pass=[]
                            for zhCN158String in zhCN158Stringlist:
                                temp=[]
                                row_index+=1
                                output[locale][zhCN158String]=[]
                                time.sleep(1)
                                pyautogui.hotkey('command', 'a',interval=0.5)
                                time.sleep(0.5)
                                pyautogui.write(zhCN158String,interval=0.5)
                                time.sleep(1)
                                pyautogui.press("space")
                                time.sleep(1)
                                screenname=takeScreenshot(zhCN158String)
                                pyautogui.hotkey('command', 'a')
                                pyautogui.hotkey('command', 'x')
                                            # insert.append(pyperclip.paste())
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
                                time.sleep(1)
                            time.sleep(1)
                            for zhCN157String in zhCN157Stringlist:
                                temp=[]
                                row_index+=1
                                output[locale][zhCN157String]=[]
                                time.sleep(1)
                                pyautogui.hotkey('command', 'a',interval=0.5)
                                time.sleep(0.5)
                                pyautogui.write(zhCN157String,interval=0.5)
                                time.sleep(1)
                                pyautogui.press("space")
                                time.sleep(1)
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
                                # zhCNresults[1]="Pass" if rendered == expected else "Fail"
                                manualList.append(temp)
                                ## Check ENter also as its in 157
                                time.sleep(1)
                                pyautogui.hotkey('command', 'a',interval=0.5)
                                time.sleep(0.5)
                                pyautogui.write(zhCN157String,interval=0.5)
                                time.sleep(1)
                                pyautogui.press("enter")
                                time.sleep(1)
                                screenname=takeScreenshot(zhCN157String)
                                pyautogui.hotkey('command', 'a')
                                pyautogui.hotkey('command', 'x')
                                            # insert.append(pyperclip.paste())
                                rendered=pyperclip.paste()
                                enterPass="Pass" if rendered == zhCN157String else "Fail"
                                if(spacePass=="Pass" and enterPass=="Pass"):
                                    zhCNresults[1]="Pass"
                                else:
                                    zhCNresults[1]="Fail"
                                time.sleep(1)
                            time.sleep(1)
                            ## 153 INTL Chinese Punctuations
                            pyautogui.hotkey('option','shift','b',interval=0.5)
                            time.sleep(0.5)
                            pyautogui.press("down")
                            time.sleep(0.5)
                            screenname=takeScreenshot("zhCN153-Symbols")
                            pyautogui.press('enter')
                            time.sleep(0.5)
                            pyautogui.hotkey('command', 'a')
                            pyautogui.hotkey('command', 'x')
                            rendered=pyperclip.paste()
                            if(rendered!=""):
                                zhCNresults[3]="Pass"
                            else:
                                zhCNresults[3]="Fail"
                            time.sleep(1)

                        
                        ## Chinese Tradiotion zh_TW
                        elif(locale=="zh_TW"):
                            checkOutput=twTypingCheck()
                            if(checkOutput==True):
                                keyboard_listener.stop()
                                TwEnglish[0]=True
                                return [111]
                            for zhTW184String in zhTW184Stringlist:
                                temp=[]
                                row_index+=1
                                output[locale][zhTW184String]=[]
                                time.sleep(1)
                                pyautogui.hotkey('command', 'a',interval=0.5)
                                time.sleep(0.5)
                                pyautogui.write(zhTW184String,interval=0.5)
                                time.sleep(1)
                                
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
                                time.sleep(1)
                            time.sleep(1)
                                ## 180 INTL Chinese Punctuations
                            pyautogui.hotkey('command', 'a')
                            time.sleep(0.5)
                            pyautogui.hotkey('option','shift','b')
                            time.sleep(0.5)
                            pyautogui.press("down")
                            time.sleep(0.5)
                            screenname=takeScreenshot("zhTW180-Symbols")
                            pyautogui.press('enter')
                            time.sleep(0.5)
                            
                            pyautogui.hotkey('option','shift','b',interval=0.5)
                            time.sleep(0.5)
                            pyautogui.press("down")
                            time.sleep(0.5)
                            screenname=takeScreenshot("zhTW180-Symbols")
                            pyautogui.press('enter')
                            time.sleep(0.5)
                            pyautogui.hotkey('command', 'a')
                            pyautogui.hotkey('command', 'x')
                            rendered=pyperclip.paste()
                            if(rendered!=""):
                                zhTWresults[1]="Pass"
                            else:
                                zhTWresults[1]="Fail"
                            time.sleep(1)
                        time.sleep(1)
                os.system(f"./Utilities/changeInputSource/changeInputSource select com.apple.keylayout.ABC-India")
                os.system(f"./Utilities/changeInputSource/changeInputSource select com.apple.keylayout.ABC")
                time.sleep(0.6)
                pyautogui.write("Press esc to generate report..",interval=0.1)
        except AttributeError:
            if key == keyboard.Key.esc:
                
                keyboard_listener.stop()
                
                print(output)
                print("Manual List : ",manualList)
                generateExcelFromListJP_KO_CN_TW(manualList,fileName,JPintlExcelList,JPresults,KOintlExcelList,KOresults,zhCNintlExcelList,zhCNresults,zhTWintlExcelList,zhTWresults)
                os.system("open /Applications/Google\ Chrome.app")
                os.system(f"./Utilities/changeInputSource/changeInputSource select com.apple.keylayout.ABC-India")
                os.system(f"./Utilities/changeInputSource/changeInputSource select com.apple.keylayout.ABC")
                # wb.close()
                
                return output
            

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
          
                return False


    def on_scroll(x, y, dx, dy):
        json_object = {'action': 'scroll', 'vertical_direction': int(dy), 'horizontal_direction': int(dx), 'x':x, 'y':y, '_time': time.time()}
        storage.append(json_object)


    keyboard_listener = keyboard.Listener(
        on_press=on_press,
        on_release=on_release)

    

    keyboard_listener.start()

    keyboard_listener.join()
    if(JpEnglish[0]==True or KoEnglish[0]==True or CnEnglish[0]==True or TwEnglish[0]==True):

        stop=""" osascript -e 'tell application "System Events" to key code 53 using {control down, command down}'"""
        os.system(stop)

        os.system("pkill -f 'QuickTime Player'")
        os.system("open /Applications/Google\ Chrome.app")
        os.system("pkill -f '{}'".format(productPath))
        os.system(f"./Utilities/changeInputSource/changeInputSource select com.apple.keylayout.ABC-India")
        os.system(f"./Utilities/changeInputSource/changeInputSource select com.apple.keylayout.ABC")
        return [111]
    else:
        ## END QT
        stop=""" osascript -e 'tell application "System Events" to key code 53 using {control down, command down}'"""
        os.system(stop)

        os.system("pkill -f 'QuickTime Player'")
        return [screenshotFolderName,productPath,notificationMail,localegroup,fileName]
