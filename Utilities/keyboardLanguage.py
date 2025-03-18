# Copyright 2025 Adobe. All rights reserved.
# This file is licensed to you under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License. You may obtain a copy
# of the License at http://www.apache.org/licenses/LICENSE-2.0

# Unless required by applicable law or agreed to in writing, software distributed under
# the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR REPRESENTATIONS
# OF ANY KIND, either express or implied. See the License for the specific language
# governing permissions and limitations under the License.
import os


langCode={
    "ja_JP":["com.apple.inputmethod.Kotoeri.RomajiTyping.Japanese","com.apple.inputmethod.Kotoeri.Japanese"],
    "Arabic":"com.apple.keylayout.Arabic-QWERTY",
    "ko_KR":["com.apple.inputmethod.Korean.2SetKorean","com.apple.inputmethod.Korean.2SetKorean"],
    "ko_KR231":["com.apple.inputmethod.Korean.2SetKorean","com.apple.inputmethod.Korean.2SetKorean"],
    "zh_CN":["com.apple.inputmethod.SCIM.ITABC","com.apple.inputmethod.SCIM.ITABC "],
    "zh_TW1":["com.apple.inputmethod.TCIM.Cangjie","com.apple.inputmethod.TCIM.Cangjie"],
    "zh_TW":["com.apple.inputmethod.TCIM.Cangjie","com.apple.inputmethod.TCIM.Cangjie"]
}



def changeIME(language):
    try:
        langInputCode=langCode[language]
        return langInputCode
    except KeyError:
        return ""

def verifyIMEpresent(language):
    import subprocess
    try:
        langInputCode=langCode[language]
        print(langInputCode)
        command1='./Utilities/changeInputSource/changeInputSource list-enabled | grep {} | wc -l '.format(langInputCode[0])
        imePresent1=subprocess.check_output(command1, shell=True)
        command2='./Utilities/changeInputSource/changeInputSource list-enabled | grep {} | wc -l '.format(langInputCode[1])
        imePresent2=subprocess.check_output(command2, shell=True)
        print(imePresent1,imePresent2)
        imePresent1=imePresent1.decode().replace('\n','').replace(' ','')
        imePresent2=imePresent2.decode().replace('\n','').replace(' ','')
        print(imePresent1,imePresent2)
        if(int(imePresent1)==0 and int(imePresent2)==0):
            return language
        elif(int(imePresent1)!=0):
            return langInputCode[0]
        elif(int(imePresent2!=0)):
            return langInputCode[1]
    except KeyError:
        return language



