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
import json,logging
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
import logging
from Utilities.functionSCPToServer import *
from Utilities.checkDisplayOff import *
from Controllers.webliveIME import *
from Controllers.desktopliveIME import *
from Controllers.desktopMultipleWorkflow import *
from Controllers.createWorkflow import *
from Controllers.webcreateworkflow import *
from Controllers.webworkflow import *
from Utilities.DBIntegration import *
from urllib.parse import urlparse
app = Flask(__name__)

        #####
        # Main Route
        #####

@app.route('/')

def index_page():
    return render_template("homeRoute.html")
  
        #####
        # Desktop IME Route
        #####
@app.route('/desktopime')
def desktopIME():
    return render_template("desktopRoute.html")

@app.route('/webime')
def webIME():
    return render_template("webRoute.html")

@app.route('/check_file', methods=['GET'])
def check_file():
    file_path = request.args.get('filepath')
    if os.path.exists(file_path):
        return jsonify({'exists': True})
    else:
        return jsonify({'exists': False})

        #####
        # Download File
        #####
@app.route('/download', methods=['GET'])
def download_file():
    filename= request.args.get('filename')
    print(filename)
    return send_file(filename,as_attachment=True)
@app.route('/openfolder')
def openfolder():
    foldername=request.args.get('filename')
    os.system('open "{}"'.format(foldername))
    return 'Checked Finder'

        #####
        # Web Live IME 
        #####

@app.route('/webliveime', methods=['GET', 'POST'])
def webliveIME():
    if (request.method == 'POST'):
        print("Current.....")
        valueReturns=webliveIME1()
        if(valueReturns[0]==100):
            return "<h4 style='text-align:center'>You are Not in Adobe Network. Kindly Connect with GlobalProtect VPN or Join Adobe Network</h4><div style='display: flex;justify-content: center;align-items: center;'><a href='/' style='background-color: #4CAF50; border: none;color: white;  padding:8px 30px; text-align: center;text-decoration: none;display: inline-block;font-size: 16px;margin: 4px 2px;cursor: pointer'>Home</a></div>"
        
        elif(valueReturns[0]==101):
            return "<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css' integrity='sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm' crossorigin='anonymous'> \
                    <body style='background-color:17202A'>\
                    <div class='d-flex align-items-center justify-content-center' style='margin:41px'><br>\
                    <br>"+"<div class='alert alert-warning' role='alert'><h4 class='alert-heading'>Holy guacamole!!</h4>\
                <p>Input Method Source is Not Added. Kindly Add the Input source of <strong>{}</strong> Locale from <strong> System Preference > Keyboard > {}</strong> </h4>".format(valueReturns[1],valueReturns[2])+"<div style='display: flex;justify-content: center;align-items: center;'><a href='/' style='background-color: #4CAF50; border: none;color: white;  padding:8px 30px; text-align: center;text-decoration: none;display: inline-block;font-size: 16px;margin: 4px 2px;cursor: pointer'>Home</a></div>'\
                    </p>\
                </div></div></body>"
        elif (valueReturns[0]==105):
            return "<h1> System is is Sleep Mode"
        elif (valueReturns[0]==111):
                print("Got Here- IME Not Changed")
                return "<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css' integrity='sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm' crossorigin='anonymous'> \
                        <body style='background-color:17202A'>\
                        <div class='d-flex align-items-center justify-content-center' style='margin:41px'><br>\
                        <br>"+"<div class='alert alert-warning' role='alert'><h4 class='alert-heading'>Holy guacamole!!</h4>\
                    <p>Sorry ! Some Problem Seems in IME of System. Please Check the Application or Try After Some Time ! </h4>"+"<div style='display: flex;justify-content: center;align-items: center;'><a href='/' style='background-color: #4CAF50; border: none;color: white;  padding:8px 30px; text-align: center;text-decoration: none;display: inline-block;font-size: 16px;margin: 4px 2px;cursor: pointer'>Home</a></div>'\
                        </p>\
                    </div></div></body>"
        else:
            screenshotFolderName,productPath,notificationMail,localegroup,fileName=valueReturns
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
                productPath=urlparse(productPath).netloc
                newsendwebliveMail(productPath,notificationMail,localegroup,fileName)
                dbEntryList=[productPath,notificationMail,"Live",str(localegroup),"Utility","Web"]
                insertLogDatainDB(dbEntryList)
                return "<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css' integrity='sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm' crossorigin='anonymous'> \
                    <body style='background-color:17202A'>\
                    <div class='d-flex align-items-center justify-content-center' style='margin:41px;'>\
                    <br>"+"<div class='alert alert-success' role='alert'><h4 class='alert-heading'>WoW !!</h4><hr>\
                        <h6 style='color:black'> Complete Processing of : "+ productPath  + "<br><br>"+"Send Detailed Report to : " + notificationMail +"</h4>" "</h6><div style='display: flex;justify-content: center;align-items: center;'><a href='/download?filename={0}'>Download Report</a></div><div style='display: flex;justify-content: center;align-items: center;'><br><a href='/openfolder?filename={1}' target='_blank'>Open Screen Folder</a></div><div style='display: flex;justify-content: center;align-items: center;'><a href='/' style='background-color: #4CAF50; border: none;color: white;  padding:8px 30px; text-align: center;text-decoration: none;display: inline-block;font-size: 16px;margin: 4px 2px;cursor: pointer'>Home</a></div>".format(fileName,screenshotFolderName)+"</div></div></body>"

            
            except:
                print(fileName)
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
                return "<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css' integrity='sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm' crossorigin='anonymous'> \
                    <body style='background-color:17202A'>\
                    <div class='d-flex align-items-center justify-content-center' style='margin:41px;'>\
                    <br>"+"<div class='alert alert-success' role='alert'><h4 class='alert-heading'>WoW !!</h4><hr>\
                        <h6 style='color:black'> Complete Processing of : "+ productPath  +"<br><br>"+ "Some Network Problem with Mail to : " + notificationMail +"</h4>" "</h6><div style='display: flex;justify-content: center;align-items: center;'><a href='/download?filename={0}'>Download Report</a></div><div style='display: flex;justify-content: center;align-items: center;'><br><a href='/openfolder?filename={1}' target='_blank'>Open Screen Folder</a></div><div style='display: flex;justify-content: center;align-items: center;'><a href='/' style='background-color: #4CAF50; border: none;color: white;  padding:8px 30px; text-align: center;text-decoration: none;display: inline-block;font-size: 16px;margin: 4px 2px;cursor: pointer'>Home</a></div>".format(fileName,screenshotFolderName)+"</div></div></body>"

            

@app.route('/liveime', methods=['GET', 'POST'])
def desktoplive():
    if (request.method == 'POST'):
        valueReturns=desktopliveime1()
        if(valueReturns[0]==100):
            return "<h4 style='text-align:center'>You are Not in Adobe Network. Kindly Connect with GlobalProtect VPN or Join Adobe Network</h4><div style='display: flex;justify-content: center;align-items: center;'><a href='/' style='background-color: #4CAF50; border: none;color: white;  padding:8px 30px; text-align: center;text-decoration: none;display: inline-block;font-size: 16px;margin: 4px 2px;cursor: pointer'>Home</a></div>"
        
        elif(valueReturns[0]==101):
            return "<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css' integrity='sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm' crossorigin='anonymous'> \
                    <body style='background-color:17202A'>\
                    <div class='d-flex align-items-center justify-content-center' style='margin:41px'><br>\
                    <br>"+"<div class='alert alert-warning' role='alert'><h4 class='alert-heading'>Holy guacamole!!</h4>\
                <p>Input Method Source is Not Added. Kindly Add the Input source of <strong>{}</strong> Locale from <strong> System Preference > Keyboard > {}</strong> </h4>".format(valueReturns[1],valueReturns[2])+"<div style='display: flex;justify-content: center;align-items: center;'><a href='/' style='background-color: #4CAF50; border: none;color: white;  padding:8px 30px; text-align: center;text-decoration: none;display: inline-block;font-size: 16px;margin: 4px 2px;cursor: pointer'>Home</a></div>'\
                    </p>\
                </div></div></body>"
        elif(valueReturns[0]==102):
            return "<h3 style='text-align:center;'> Sorry Please Check Product Path <br> Product Path Not Found </h3>"
        elif (valueReturns[0]==105):
            return "<h1> System is is Sleep Mode"
        elif (valueReturns[0]==111):
            print("Got Here- IME Not Changed")
            return "<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css' integrity='sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm' crossorigin='anonymous'> \
                    <body style='background-color:17202A'>\
                    <div class='d-flex align-items-center justify-content-center' style='margin:41px'><br>\
                    <br>"+"<div class='alert alert-warning' role='alert'><h4 class='alert-heading'>Holy guacamole!!</h4>\
                <p>Sorry ! Some Problem Seems in IME of System. Please Check the Application or Try After Some Time ! </h4>"+"<div style='display: flex;justify-content: center;align-items: center;'><a href='/' style='background-color: #4CAF50; border: none;color: white;  padding:8px 30px; text-align: center;text-decoration: none;display: inline-block;font-size: 16px;margin: 4px 2px;cursor: pointer'>Home</a></div>'\
                    </p>\
                </div></div></body>"
        else:
            screenshotFolderName,productPath,notificationMail,localegroup,fileName=valueReturns
            try:
                print("In Block...")
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
                newsendlivereportMail(productPath,notificationMail,localegroup,fileName)
                productSubject=productPath.split('/')[-1].replace(".app","")
                dbEntryList=[productSubject,notificationMail,"Live",str(localegroup),"Utility","Desktop"]
                insertLogDatainDB(dbEntryList)
                return "<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css' integrity='sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm' crossorigin='anonymous'> \
                    <body style='background-color:17202A'>\
                    <div class='d-flex align-items-center justify-content-center' style='margin:41px;'>\
                    <br>"+"<div class='alert alert-success' role='alert'><h4 class='alert-heading'>WoW !!</h4><hr>\
                        <h6 style='color:black'> Complete Processing of Build : "+ productPath  + "<br><br>"+"Send Detailed Report to : " + notificationMail +"</h4>" "</h6><div style='display: flex;justify-content: center;align-items: center;'><a href='/download?filename={0}'>Download Report</a></div><div style='display: flex;justify-content: center;align-items: center;'><br><a href='/openfolder?filename={1}' target='_blank'>Open Screen Folder</a></div><div style='display: flex;justify-content: center;align-items: center;'><a href='/' style='background-color: #4CAF50; border: none;color: white;  padding:8px 30px; text-align: center;text-decoration: none;display: inline-block;font-size: 16px;margin: 4px 2px;cursor: pointer'>Home</a></div>".format(fileName,screenshotFolderName)+"</div></div></body>"
            
            
            except:
                print("In Exception..",fileName)
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
                return "<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css' integrity='sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm' crossorigin='anonymous'> \
                    <body style='background-color:17202A'>\
                    <div class='d-flex align-items-center justify-content-center' style='margin:41px;'>\
                    <br>"+"<div class='alert alert-success' role='alert'><h4 class='alert-heading'>WoW !!</h4><hr>\
                        <h6 style='color:black'> Complete Processing of Build : "+ productPath  +"<br><br>"+"Please download below report </h4>" "</h6><div style='display: flex;justify-content: center;align-items: center;'><a href='/download?filename={0}'>Download Report</a></div><div style='display: flex;justify-content: center;align-items: center;'><br><a href='/openfolder?filename={1}' target='_blank'>Open Screen Folder</a></div><div style='display: flex;justify-content: center;align-items: center;'><a href='/' style='background-color: #4CAF50; border: none;color: white;  padding:8px 30px; text-align: center;text-decoration: none;display: inline-block;font-size: 16px;margin: 4px 2px;cursor: pointer'>Home</a></div>".format(fileName,screenshotFolderName)+"</div></div></body>"
    else:
        return redirect(url_for('index_page'))

@app.route('/check', methods=['GET', 'POST'])
def desktopWorkflow():
    if (request.method == 'POST'):
        ## Check in Adobe Network Or Not
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(1)   
            result = sock.connect_ex(('authrelay.corp.adobe.com',587))
            if result != 0:
                raise TimeoutError
            
            sock.close()
        except:
            return "<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css' integrity='sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm' crossorigin='anonymous'> \
            <body style='background-color:17202A'>\
                <div class='d-flex align-items-center justify-content-center' style='margin:41px'><br>\
            <br>"+"<div class='alert alert-warning' role='alert'><h4 class='alert-heading' >Holy guacamole!!</h4>\
                Sorry ! You are not in Adobe Network, Please connect to VPN First \
                    <div style='display: flex;justify-content: center;align-items: center;'><a href='/' style='background-color: #4CAF50; border: none;color: white;  padding:8px 30px; text-align: center;text-decoration: none;display: inline-block;font-size: 16px;margin: 4px 2px;cursor: pointer'>Home</a></div></body>"
        valueReturns=desktopWorkflowIME()
        if(valueReturns[0]==100):
            return "<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css' integrity='sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm' crossorigin='anonymous'> \
            <body style='background-color:17202A'>\
                <div class='d-flex align-items-center justify-content-center' style='margin:41px'><br>\
            <br>"+"<div class='alert alert-warning' role='alert'><h4 class='alert-heading' >Holy guacamole!!</h4>\
                Sorry This Workflow File is not Correct ! Kindly Create New workflow\
        </div></div></body>"

        elif(valueReturns[0]==102):
            return "<h4 style='text-align:center'> Sorry ! Some Issue With the Workflow File : "+ "<br>Kindly Create Workflow Again <br></h4>" + "<div style='display: flex;justify-content: center;align-items: center;'><a href='/' style='  background-color: #4CAF50; border: none;color: white;  padding: 15px 32px; text-align: center;text-decoration: none;display: inline-block;font-size: 16px;margin: 4px 2px;cursor: pointer;'>Home</a></div>"
        elif(valueReturns[0]==103):
            return "<h1>Sorry, Currently IME Is Not Changing</h1>"
        elif(valueReturns[0]==105):
            return "<h1> System is is Sleep Mode"
        elif (valueReturns[0]==106):
            return "<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css' integrity='sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm' crossorigin='anonymous'> \
            <body style='background-color:17202A'>\
            <div class='d-flex align-items-center justify-content-center' style='margin:41px'><br>\
            <br>"+"<div class='alert alert-warning' role='alert'><h4 class='alert-heading' >Holy guacamole!!</h4>\
            <p>Input Method Source is Not Added. Kindly Add the Input source of <strong>{}</strong> Locale from <strong> System Preference > Keyboard > {}</strong> </h4>".format(valueReturns[1],valueReturns[2])+"<div style='display: flex;justify-content: center;align-items: center;'><a href='/' style='background-color: #4CAF50; border: none;color: white;  padding:8px 30px; text-align: center;text-decoration: none;display: inline-block;font-size: 16px;margin: 4px 2px;cursor: pointer'>Home</a></div>'\
                </p>\
            </div></div></body>"
        elif (valueReturns[0]==107):
            ide,productPath,notificationMail,localegroup,multiplecoordFileName=valueReturns
            list_text = render_template_string('{% for element in my_list %}{{ element }}<br>{% endfor %}', my_list=multiplecoordFileName)

            return "<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css' integrity='sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm' crossorigin='anonymous'> \
                   <body style='background-color:17202A'>\
                   <div class='d-flex align-items-center justify-content-center' style='margin:41px;'>\
                    <br>"+"<div class='alert alert-success' role='alert'><h4 class='alert-heading'>WoW !!</h4><hr>\
                      <h6 style='color:black'> Complete Processing of Build : "+ productPath  + "<br><br>"+"Executed Workflow : "+"<br><br>"+list_text+ "<br><br>"+"Send Detailed Report to : " + notificationMail +"</h4>"+"<a href='/' style='background-color: #4CAF50; border: none;color: white;  padding:8px 30px; text-align: center;text-decoration: none;display: inline-block;font-size: 16px;margin: 4px 2px;cursor: pointer'>Home</a></div></div></body>"
        elif (valueReturns[0]==111):
            print("Got Here- IME Not Changed")
            return "<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css' integrity='sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm' crossorigin='anonymous'> \
                    <body style='background-color:17202A'>\
                    <div class='d-flex align-items-center justify-content-center' style='margin:41px'><br>\
                    <br>"+"<div class='alert alert-warning' role='alert'><h4 class='alert-heading'>Holy guacamole!!</h4>\
                <p>Sorry ! Some Problem Seems in IME of System. Please Check the Application or Try After Some Time ! </h4>"+"<div style='display: flex;justify-content: center;align-items: center;'><a href='/' style='background-color: #4CAF50; border: none;color: white;  padding:8px 30px; text-align: center;text-decoration: none;display: inline-block;font-size: 16px;margin: 4px 2px;cursor: pointer'>Home</a></div>'\
                    </p>\
                </div></div></body>"
   
        return redirect(url_for('index_page'))
@app.route('/webexecuteworkflow', methods=['GET', 'POST'])
def webWorkflow():
    if (request.method == 'POST'):

        print("Before....")
        valueReturns=webWorkflowIME()
        print(valueReturns)
        if(valueReturns[0]==100):
            return "<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css' integrity='sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm' crossorigin='anonymous'> \
            <body style='background-color:17202A'>\
                <div class='d-flex align-items-center justify-content-center' style='margin:41px'><br>\
            <br>"+"<div class='alert alert-warning' role='alert'><h4 class='alert-heading' >Holy guacamole!!</h4>\
                Sorry This Workflow File is not Correct ! Kindly Create New workflow\
        </div></div></body>"
       
        elif(valueReturns[0]==102):
            return "<h4 style='text-align:center'> Sorry ! Some Issue With the Workflow File : "+ "<br>Kindly Create Workflow Again <br></h4>" + "<div style='display: flex;justify-content: center;align-items: center;'><a href='/' style='  background-color: #4CAF50; border: none;color: white;  padding: 15px 32px; text-align: center;text-decoration: none;display: inline-block;font-size: 16px;margin: 4px 2px;cursor: pointer;'>Home</a></div>"
        elif(valueReturns[0]==103):
            return "<h1>Sorry, Currently IME Is Not Changing</h1>"
        elif(valueReturns[0]==105):
            return "<h1> System is is Sleep Mode"
        elif (valueReturns[0]==106):
            return "<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css' integrity='sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm' crossorigin='anonymous'> \
            <body style='background-color:17202A'>\
            <div class='d-flex align-items-center justify-content-center' style='margin:41px'><br>\
            <br>"+"<div class='alert alert-warning' role='alert'><h4 class='alert-heading' >Holy guacamole!!</h4>\
            <p>Input Method Source is Not Added. Kindly Add the Input source of <strong>{}</strong> Locale from <strong> System Preference > Keyboard > {}</strong> </h4>".format(valueReturns[1],valueReturns[2])+"<div style='display: flex;justify-content: center;align-items: center;'><a href='/' style='background-color: #4CAF50; border: none;color: white;  padding:8px 30px; text-align: center;text-decoration: none;display: inline-block;font-size: 16px;margin: 4px 2px;cursor: pointer'>Home</a></div>'\
                </p>\
            </div></div></body>"
        elif (valueReturns[0]==107):
            ide,productPath,notificationMail,localegroup,multiplecoordFileName=valueReturns
            list_text = render_template_string('{% for element in my_list %}{{ element }}<br>{% endfor %}', my_list=multiplecoordFileName)

            return "<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css' integrity='sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm' crossorigin='anonymous'> \
                   <body style='background-color:17202A'>\
                   <div class='d-flex align-items-center justify-content-center' style='margin:41px;'>\
                    <br>"+"<div class='alert alert-success' role='alert'><h4 class='alert-heading'>WoW !!</h4><hr>\
                      <h6 style='color:black'> Complete Processing of Build : "+ productPath  + "<br><br>"+"Executed Workflow : "+"<br><br>"+list_text+ "<br><br>"+"Send Detailed Report to : " + notificationMail +"</h4>"+"<a href='/' style='background-color: #4CAF50; border: none;color: white;  padding:8px 30px; text-align: center;text-decoration: none;display: inline-block;font-size: 16px;margin: 4px 2px;cursor: pointer'>Home</a></div></div></body>"
        elif (valueReturns[0]==111):
            print("Got Here- IME Not Changed")
            return "<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css' integrity='sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm' crossorigin='anonymous'> \
                    <body style='background-color:17202A'>\
                    <div class='d-flex align-items-center justify-content-center' style='margin:41px'><br>\
                    <br>"+"<div class='alert alert-warning' role='alert'><h4 class='alert-heading'>Holy guacamole!!</h4>\
                <p>Sorry ! Some Problem Seems in IME of System. Please Check the Application or Try After Some Time ! </h4>"+"<div style='display: flex;justify-content: center;align-items: center;'><a href='/' style='background-color: #4CAF50; border: none;color: white;  padding:8px 30px; text-align: center;text-decoration: none;display: inline-block;font-size: 16px;margin: 4px 2px;cursor: pointer'>Home</a></div>'\
                    </p>\
                </div></div></body>"

    else:
        return redirect(url_for('index_page'))
@app.route('/processing', methods=['GET', 'POST'])
def createWorkflow1():
    if (request.method == 'POST'):
        valueReturns=createWorkflow()
        if(valueReturns[0]==100):
            return "<h4 style='style='text-align:center;'>Sorry, Some Issue with the Installer <br> Please try After Some time !!</h4>"
        elif(valueReturns[0]==101):
            return "<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css' integrity='sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm' crossorigin='anonymous'> \
            <body style='background-color:17202A'>\
            <div class='d-flex align-items-center justify-content-center' style='margin:41px'><br>\
            <br>"+"<div class='alert alert-warning' role='alert'><h4 class='alert-heading' >Holy guacamole!!</h4>\
            <p>Sorry Please Verify Application Path, Application Path is Not Correct</p>\
            </div></div></body>"
        else:
            productPath,workflowPath,name_of_recording=valueReturns
            return render_template('desktopRoute.html',appPath=productPath,fileName='{0}/{1}.txt'.format(workflowPath,name_of_recording))
    else:
        return redirect(url_for('index_page'))

@app.route('/webprocessing', methods=['GET', 'POST'])
def webcreateWorkflow1():
    if (request.method == 'POST'):
        valueReturns=webcreateWorkflow()
        if(valueReturns[0]==100):
            return "<h4 style='style='text-align:center;'>Sorry, Some Issue with the Installer <br> Please try After Some time !!</h4>"
        elif(valueReturns[0]==101):
            return "<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css' integrity='sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm' crossorigin='anonymous'> \
            <body style='background-color:17202A'>\
            <div class='d-flex align-items-center justify-content-center' style='margin:41px'><br>\
            <br>"+"<div class='alert alert-warning' role='alert'><h4 class='alert-heading' >Holy guacamole!!</h4>\
            <p>Sorry Please Verify Application Path, Application Path is Not Correct</p>\
            </div></div></body>"
        else:
            productPath,workflowPath,name_of_recording=valueReturns
            return render_template('webRoute.html',appPath=productPath,fileName='{0}/{1}.txt'.format(workflowPath,name_of_recording))
    else:
        return redirect(url_for('index_page'))
@app.errorhandler(404)
def not_found(e):
  return render_template("404.html")
## App Closing Handlings
def goodbye():
    stop=""" osascript -e 'tell application "System Events" to key code 53 using {control down, command down}'"""
    os.system(stop)

    os.system("pkill -f 'QuickTime Player'")
   
def handle_sigterm(signum, frame):
    goodbye()
    exit(0)

atexit.register(goodbye)

signal.signal(signal.SIGHUP, handle_sigterm)
atexit.register(goodbye)
if __name__ == '__main__':

    logging.getLogger('werkzeug').disabled = True
    chrome_path = 'open -a /Applications/Google\ Chrome.app %s'
    webbrowser.get(chrome_path).open_new('http://localhost:3000')
    app.run(debug=False,port=3000)
   
