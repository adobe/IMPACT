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
def read_template(filename):
 with open(filename, 'r', encoding='utf-8') as template_file:
  template_file_content = template_file.read()
 return Template(template_file_content)
def sendDisplayOFFMail(productPath,notificationMail,locale):
    productName=productPath

    fromaddr=""
    toaddrlist=notificationMail
    namelist=notificationMail.split('@')[0]
    cc = ""

    mix=toaddrlist+","+cc

    msg = MIMEMultipart()

    msg['From'] = fromaddr
    msg['To'] = toaddrlist
    msg['Cc'] = cc
    now=datetime.now()

    msg['Subject'] = " System went in Sleep Mode IME Framework- " + "Locale - " + str(locale) + "-"+ now.strftime("%d-%m-%Y %H-%M-%S")

    message_template = read_template('Utilities/SendMail/DisplayOffTemplate.txt')
    message = message_template.substitute(PERSON_NAME=namelist,PRODUCT_NAME=productName,LOCALE=locale)

    msg.attach(MIMEText(message, 'html'))
    s = smtplib.SMTP('Relay Server', 587)

    s.starttls()


    s.login(fromaddr, "Password")

    text = msg.as_string()

    s.sendmail(fromaddr, mix.split(','), text)

    s.quit()

    print("--Completed of System Sleep Mail--")