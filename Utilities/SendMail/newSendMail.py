# Copyright 2025 Adobe. All rights reserved.
# This file is licensed to you under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License. You may obtain a copy
# of the License at http://www.apache.org/licenses/LICENSE-2.0

# Unless required by applicable law or agreed to in writing, software distributed under
# the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR REPRESENTATIONS
# OF ANY KIND, either express or implied. See the License for the specific language
# governing permissions and limitations under the License.

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from string import Template
from datetime import datetime
import socket
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from string import Template
from datetime import datetime
import socket
from xlsx2html import xlsx2html
def read_template(filename):
 with open(filename, 'r', encoding='utf-8') as template_file:
  template_file_content = template_file.read()
 return Template(template_file_content)

def getMailPassword():
    import requests
    json_url = 'Password URL'

    try:
        response = requests.get(json_url)
        if response.status_code == 200:
            config = response.json()

            # Access the password from the JSON data
            password = config.get('password')

            if password is not None:
                # Use the 'password' variable in your code
                return password
    except Exception as e:
        print(f"An error occurred: {e}")

def newWebsendonereportMail(productPath,notificationMail,locale,filename,IMEChange,TypingChars,appCrashed,name_of_recording):
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.settimeout(1)   
    result = sock.connect_ex(('authrelay Server',587))
    if result != 0:
        raise TimeoutError
    
    sock.close()
    convertedHTML=filename+"-Convert.html"
    productSubject=extract_text_after_https(productPath)
    xlsx2html(filename, convertedHTML,sheet=1)
    productName=productPath
    password=getMailPassword()
    fromaddr=""
    toaddrlist=notificationMail
    namelist="All"
    cc = ""
    mix=toaddrlist+","+cc

    msg = MIMEMultipart()

    msg['From'] = fromaddr

    msg['To'] = toaddrlist
    msg['Cc'] = cc
    now=datetime.now()

    msg['Subject'] = productSubject+" -Web Workflow IME Automation Report - " + "Locale - " + str(locale) + "-"+ now.strftime("%d-%m-%Y %H-%M-%S")

    message_template = read_template('Utilities/SendMail/template.txt')
    message_template2=read_template(convertedHTML)
    message = message_template.substitute(PERSON_NAME=namelist,PRODUCT_NAME=productName,LOCALE=locale,APP_CRASHED=appCrashed,IME_STATUS=IMEChange,CHAR_TYPING=TypingChars,WORKFLOW_NAME=name_of_recording)
    message2=message_template2.substitute(PERSON_NAME=namelist,PRODUCT_NAME=productName,LOCALE=locale)

    combined=message+message2

    msg.attach(MIMEText(combined, 'html'))
    filelist=[filename]
    for filename in filelist:
        print(filename)
        attachment = open(filename, "rb")


        p = MIMEBase('application', 'octet-stream')

            # To change the payload into encoded form
        p.set_payload((attachment).read())

            # encode into base64
        encoders.encode_base64(p)

        p.add_header('Content-Disposition', "attachment; filename= %s" % "IME Framework Report.xlsx")
            # attach the instance 'p' to instance 'msg'
        msg.attach(p)


    s = smtplib.SMTP('authrelay server', 587)

    s.starttls()

    s.login(fromaddr, password)

    text = msg.as_string()

    s.sendmail(fromaddr, mix.split(','), text)
    s.quit()

    print("--Completed--")
def newsendonereportMail(productPath,notificationMail,locale,filename,IMEChange,TypingChars,appCrashed,name_of_recording):
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.settimeout(1)   
    result = sock.connect_ex(('Relay Server',587))
    if result != 0:
        raise TimeoutError
    
    sock.close()
    convertedHTML=filename+"-Convert.html"
    xlsx2html(filename, convertedHTML,sheet=1)
    productName=productPath
    password=getMailPassword()
    # stringRendering=output
    fromaddr=""
    toaddrlist=notificationMail
    namelist="All"
    # cc = ",,
    cc = ""
    mix=toaddrlist+","+cc

    msg = MIMEMultipart()


    msg['From'] = fromaddr

        # storing the receivers email address
    msg['To'] = toaddrlist
    msg['Cc'] = cc
    now=datetime.now()
    productSubject=productPath.split('/')[-1].replace(".app","")
        # storing the subject
    msg['Subject'] = productSubject+" -Desktop IME Automation Report - " + "Locale - " + str(locale) + "-"+ now.strftime("%d-%m-%Y %H-%M-%S")

        # string to store the body of the mail
    message_template = read_template('Utilities/SendMail/template.txt')
    message_template2=read_template(convertedHTML)
    message = message_template.substitute(PERSON_NAME=namelist,PRODUCT_NAME=productName,LOCALE=locale,APP_CRASHED=appCrashed,IME_STATUS=IMEChange,CHAR_TYPING=TypingChars,WORKFLOW_NAME=name_of_recording)
    message2=message_template2.substitute(PERSON_NAME=namelist,PRODUCT_NAME=productName,LOCALE=locale)
        # attach the body with the msg instance
    combined=message+message2

    msg.attach(MIMEText(combined, 'html'))
  
    filelist=[filename]
    for filename in filelist:
        print(filename)
        attachment = open(filename, "rb")

            # instance of MIMEBase and named as p
        p = MIMEBase('application', 'octet-stream')

            # To change the payload into encoded form
        p.set_payload((attachment).read())

            # encode into base64
        encoders.encode_base64(p)

        p.add_header('Content-Disposition', "attachment; filename= %s" % "IME Framework Report.xlsx")
            # attach the instance 'p' to instance 'msg'
        msg.attach(p)


    s = smtplib.SMTP('Relay Server', 587)

        # start TLS for security
    s.starttls()

        # Authentication
    s.login(fromaddr, password)

        # Converts the Multipart msg into a string
    text = msg.as_string()

        # sending the mail
    s.sendmail(fromaddr, mix.split(','), text)

        # terminating the session
    s.quit()

    print("--Completed--")



def sendReportMail(productPath,output,notificationMail,locale,filename,IMEChange,TypingChars,appCrashed,screenshotzip):

    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.settimeout(1)   
    result = sock.connect_ex(('Relay Server',587))
    if result != 0:
        raise TimeoutError
    
    sock.close()
    productName=productPath
    stringRendering=output
    fromaddr=""
    toaddrlist=notificationMail
    cc=""
    password=getMailPassword()
    mix=toaddrlist+","+cc
    namelist=notificationMail.split('@')[0]

    # instance of MIMEMultipart
    msg = MIMEMultipart()

        # storing the senders email address
    msg['From'] = fromaddr

        # storing the receivers email address
    msg['To'] = toaddrlist
    msg['Cc'] = cc
    now=datetime.now()
        # storing the subject
    msg['Subject'] = " IME Automation Report - " + "Locale - " + str(locale) + "-"+ now.strftime("%d-%m-%Y %H-%M-%S")

        # string to store the body of the mail
    message_template = read_template('Utilities/SendMail/template.txt')
    message = message_template.substitute(PERSON_NAME=namelist,STRING_RENDERING=stringRendering,PRODUCT_NAME=productName,LOCALE=locale,APP_CRASHED=appCrashed,IME_STATUS=IMEChange,CHAR_TYPING=TypingChars)
    message2 = """\
    <html>
      <head></head>
      <body>
        <strong>{}</strong>
        </p>
      </body>
    </html>
    """.format(stringRendering)
        # attach the body with the msg instance
    msg.attach(MIMEText(message, 'html'))

    filelist=[filename,screenshotzip]
    for filename in filelist:
        print(filename)
        attachment = open(filename, "rb")

            # instance of MIMEBase and named as p
        p = MIMEBase('application', 'octet-stream')

            # To change the payload into encoded form
        p.set_payload((attachment).read())

            # encode into base64
        encoders.encode_base64(p)

        p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
            # attach the instance 'p' to instance 'msg'
        msg.attach(p)


    s = smtplib.SMTP('Relay Server', 587)

        # start TLS for security
    s.starttls()

        # Authentication
    s.login(fromaddr, password)

        # Converts the Multipart msg into a string
    text = msg.as_string()

        # sending the mail
    s.sendmail(fromaddr, mix.split(','), text)

        # terminating the session
    s.quit()

    print("--Completed--")

def sendIMEFailedMail(productPath,notificationMail,locale):
    productName=productPath
    # stringRendering=output
    fromaddr=""
    toaddrlist=notificationMail
    namelist=notificationMail.split('@')[0]
    password=getMailPassword()
    cc = ""
    mix=toaddrlist+","+cc
    # instance of MIMEMultipart
    msg = MIMEMultipart()

        # storing the senders email address
    msg['From'] = fromaddr

        # storing the receivers email address
    msg['To'] = toaddrlist
    msg['Cc'] = cc
    now=datetime.now()
        # storing the subject
    msg['Subject'] = " IME Automation Failed - " + "Locale - " + str(locale) + "-"+ now.strftime("%d-%m-%Y %H-%M-%S")

        # string to store the body of the mail
    message_template = read_template('Utilities/SendMail/IMEFailedTemplate.txt')
    message = message_template.substitute(PERSON_NAME=namelist,PRODUCT_NAME=productName,LOCALE=locale)
        # attach the body with the msg instance
    msg.attach(MIMEText(message, 'html'))
   
    s = smtplib.SMTP('Relay Server', 587)

        # start TLS for security
    s.starttls()

        # Authentication
    s.login(fromaddr, password)

        # Converts the Multipart msg into a string
    text = msg.as_string()

        # sending the mail
    s.sendmail(fromaddr, mix.split(','), text)

        # terminating the session
    s.quit()

    print("--Completed--")


def newsendlivereportMail(productPath,notificationMail,locale,filename):
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.settimeout(1)   
    result = sock.connect_ex(('Relay Server',587))
    if result != 0:
        raise TimeoutError
    
    sock.close()
    ### Create HTML
    convertedHTML=filename+"-Convert.html"
    xlsx2html(filename, convertedHTML,sheet=1)
    productName=productPath
    password=getMailPassword()
    fromaddr=""
    toaddrlist=notificationMail
    namelist="All"

    cc = ""
    mix=toaddrlist+","+cc
    # instance of MIMEMultipart
    msg = MIMEMultipart()

        # storing the senders email address
    msg['From'] = fromaddr

        # storing the receivers email address
    msg['To'] = toaddrlist
    msg['Cc'] = cc
    now=datetime.now()
    productSubject=productPath.split('/')[-1].replace(".app","")
        # storing the subject
    msg['Subject'] = productSubject +" -Desktop IME Framework Report - " + "Locale - " + str(locale) + "-"+ now.strftime("%d-%m-%Y %H-%M-%S")

        # string to store the body of the mail
    message_template = read_template('Utilities/SendMail/LiveIMETemplate.txt')
    message_template2=read_template(convertedHTML)
    message = message_template.substitute(PERSON_NAME=namelist,PRODUCT_NAME=productName,LOCALE=locale)
    message2=message_template2.substitute(PERSON_NAME=namelist,PRODUCT_NAME=productName,LOCALE=locale)
        # attach the body with the msg instance
    combined=message+message2
    # msg.attach(MIMEText(message, 'html'))
    msg.attach(MIMEText(combined, 'html'))
    filelist=[filename]
    for filename in filelist:
        print(filename)
        attachment = open(filename, "rb")

            # instance of MIMEBase and named as p
        p = MIMEBase('application', 'octet-stream')

            # To change the payload into encoded form
        p.set_payload((attachment).read())

            # encode into base64
        encoders.encode_base64(p)

        p.add_header('Content-Disposition', "attachment; filename= %s" % "IME Framework Report.xlsx")
            # attach the instance 'p' to instance 'msg'
        msg.attach(p)

    s = smtplib.SMTP('Relay Server', 587)

    s.starttls()


    s.login(fromaddr, password)


    text = msg.as_string()


    s.sendmail(fromaddr, mix.split(','), text)


    s.quit()

    print("--Completed--")

def extract_text_after_https(link):
    if("https://" in link or "http://" in link):
        start_index = link.find("https://") + len("https://")
        end_index = link.find("/", start_index)
        if start_index != -1 and end_index != -1:
            return link[start_index:end_index]
        elif start_index != -1:
            return link[start_index:]
        else:
            return ""
    else:
        main=link.split('/')[0]
        return main
        
def newsendwebliveMail(productPath,notificationMail,locale,filename):
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.settimeout(1)   
    result = sock.connect_ex(('Relay Server',587))
    if result != 0:
        raise TimeoutError
    
    sock.close()
    
    productName=productPath
    productSubject=extract_text_after_https(productPath)
    password=getMailPassword()

    fromaddr=""
    toaddrlist=notificationMail
    namelist="All"

    cc = ""
    mix=toaddrlist+","+cc

    msg = MIMEMultipart()


    msg['From'] = fromaddr

    msg['To'] = toaddrlist
    msg['Cc'] = cc
    now=datetime.now()
    convertedHTML=filename+"-Convert.html"
    xlsx2html(filename, convertedHTML,sheet=1)
    msg['Subject'] = productSubject+" - Web IME Automation Report - " + "Locale - " + str(locale) + "-"+ now.strftime("%d-%m-%Y %H-%M-%S")

        # string to store the body of the mail
    message_template = read_template('Utilities/SendMail/LiveIMETemplate.txt')
    message_template2=read_template(convertedHTML)
    message = message_template.substitute(PERSON_NAME=namelist,PRODUCT_NAME=productName,LOCALE=locale)
    message2=message_template2.substitute(PERSON_NAME=namelist,PRODUCT_NAME=productName,LOCALE=locale)
    combined=message+message2

    msg.attach(MIMEText(combined, 'html'))
    filelist=[filename]
    for filename in filelist:
        print(filename)
        attachment = open(filename, "rb")


        p = MIMEBase('application', 'octet-stream')

    
        p.set_payload((attachment).read())


        encoders.encode_base64(p)

        p.add_header('Content-Disposition', "attachment; filename= %s" % "Web IME Framework Report.xlsx")

        msg.attach(p)

    s = smtplib.SMTP('Relay Server', 587)


    s.starttls()

        # Authentication
    s.login(fromaddr, password)

        # Converts the Multipart msg into a string
    text = msg.as_string()

    s.sendmail(fromaddr, mix.split(','), text)

        # terminating the session
    s.quit()

    print("--Completed--")