# Copyright 2025 Adobe. All rights reserved.
# This file is licensed to you under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License. You may obtain a copy
# of the License at http://www.apache.org/licenses/LICENSE-2.0

# Unless required by applicable law or agreed to in writing, software distributed under
# the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR REPRESENTATIONS
# OF ANY KIND, either express or implied. See the License for the specific language
# governing permissions and limitations under the License.
import subprocess,time,os
from Utilities.SendMail.sendDisplayOffMail import *
import pyautogui

count=0
def startSleepCount():
    global count
    command = 'pmset -g log | grep -c "Display is turned off"'
    output = subprocess.check_output(command, shell=True, stderr=subprocess.DEVNULL)
    output = output.decode().strip()
    count = int(output)
    print("At Start : ", count)
def is_display_off():
    global count
    try:
        command = 'pmset -g log | grep -c "Display is turned off"'
        output = subprocess.check_output(command, shell=True, stderr=subprocess.DEVNULL)
        output = output.decode().strip()
        newcount=int(output)
        if(newcount>count):
            return True
    except (subprocess.CalledProcessError, OSError):
        return False


def stopFlask(notificationMail,productPath,locale):
    stop=""" osascript -e 'tell application "System Events" to key code 53 using {control down, command down}'"""
    os.system(stop)
    time.sleep(0.5)
    os.system("pkill -f 'QuickTime Player'")
    os.system(f"./Utilities/changeInputSource/changeInputSource select com.apple.keylayout.ABC-India")
    os.system(f"./Utilities/changeInputSource/changeInputSource select com.apple.keylayout.ABC")
    sendDisplayOFFMail(productPath,notificationMail,locale)
