import smtplib, ssl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.utils import COMMASPACE, formatdate
from email import encoders
from email.message import EmailMessage
import time
import datetime
import selenium
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains as actions
from selenium.webdriver.support.ui import Selec
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
import os
import re

def send_email(dokumenty_nazwa, *args, **kwargs):
        #Email Variables
        SMTP_SERVER = 'oneumbrella.pl' #Email Server (don't change!)
        SMTP_PORT = 587 #Server Port (don't change!)
        GMAIL_USERNAME = 'no-replay@oneumbrella.pl' #change this to match your gmail account
        GMAIL_PASSWORD = ''  #change this to match your gmail app-password
        
        class Emailer:
            def sendmail(self, recipient, subject, content):

                #Create Headers
                headers = ["From: " + GMAIL_USERNAME, "Subject: " + subject, "To: " + recipient,
                        "MIME-Version: 1.0", "Content-Type: text/html"]
                headers = "\r\n".join(headers)
                
                # Create the message
                message = EmailMessage()
                message['From'] = GMAIL_USERNAME
                message['To'] = recipient
                message['Subject'] = subject
                message.set_content(content)
                #message.add_attachment(attachment)

                #Connect to Gmail Server
                session = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
                session.ehlo()
                session.starttls()
                session.ehlo()

                #Login to Gmail
                session.login(GMAIL_USERNAME, GMAIL_PASSWORD)

                #Send Email & Exit
                session.send_message(message)
                session.quit

        sender = Emailer()

        sendTo = 'waldemar.lusiak@nksgroup.pl, remigiusz.zerbst@nksgroup.pl'
        emailSubject = "Automatically generated email from CERTUS WEB MANAGER"
        emailContent = "Units in: \n\n {}".format(dokumenty_nazwa)
        # ,michal.sniegocki@nksgroup.pl,remigiusz.zerbst@nksgroup.pl
        #Sends an email to the "sendTo" address with the specified "emailSubject" as the subject and "emailContent" as the email content.
        sender.sendmail(sendTo, emailSubject, emailContent)

login_certus = "NKSGroupSpzoo"
passw_certus = ""
customer_code = "163785"
url = "https://cloud.certus.software/"

#login_field = "/html/body/div[3]/div/div/div/div[1]/div[2]/form/ul/li[1]/input"
#password_field = "/html/body/div[3]/div/div/div/div[1]/div[2]/form/ul/li[2]/input"
#customer_code_field = "/html/body/div[3]/div/div/div/div[1]/div[2]/form/ul/li[3]/input"
#login_button = "/html/body/div[3]/div/div/div/div[1]/div[2]/form/ul/li[6]/a"
try:
    options = Options()
    options.BinaryLocation = r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
    driver_path = r"C:\\Users\\waldemar.lusiak\Downloads\\chromedriver.exe"
    driver = webdriver.Chrome(options=options, service=Service(driver_path))
    driver.get(url)
    time.sleep(1)
    
    #Login to certus site
    login_field = driver.find_element(By.XPATH,"/html/body/div[3]/div/div/div/div[1]/div[2]/form/ul/li[1]/input")
    login_field.send_keys(login_certus)
    time.sleep(1)

    password_field = driver.find_element(By.XPATH,"/html/body/div[3]/div/div/div/div[1]/div[2]/form/ul/li[2]/input")
    password_field.send_keys(passw_certus)
    time.sleep(1)

    customer_code_field = driver.find_element(By.XPATH,"/html/body/div[3]/div/div/div/div[1]/div[2]/form/ul/li[3]/input")
    customer_code_field.send_keys(customer_code)
    time.sleep(1)

    login_button = driver.find_element(By.XPATH,"/html/body/div[3]/div/div/div/div[1]/div[2]/form/ul/li[6]/a").click()
    time.sleep(2)

    #Erasure reports
    erasure_reports = driver.find_element(By.XPATH,"/html/body/div[4]/div/div[2]/div[2]/ul/li[3]/a/span[2]").click()
    time.sleep(3)

    select_report = driver.find_element(By.XPATH,"/html/body/div[4]/div/div[3]/div/div[3]/div/div[2]/div[1]/div[1]/div[2]/div/div/div[1]").click()
    time.sleep(3)

    export_button = driver.find_element(By.XPATH,"/html/body/div[4]/div/div[1]/div[3]/div/ul/li[1]/a/span").click()
    time.sleep(2)

    new_window = driver.find_element(By.XPATH, "/html/body/div[3]/div/div/div[2]/ul/li[1]/div/button[1]").click()
    time.sleep(2)

    export_button = driver.find_element(By.XPATH, "/html/body/div[3]/div/div/div[3]/a[2]").click()
    time.sleep(1)
    driver.quit()

    #path = "C:\\Users\\waldemar.lusiak\\Downloads"
try:    
    path = "C:\\Users\\walde\\Downloads"
    
    for filename in os.listdir(path):
        if re.match(r"certus_report.*\.csv", filename):
            file_path = os.path.join(path, filename)
            df = pd.read_csv(file_path, sep=".")
            #print("Full name of the file:", file_path)
            
            if "Erasure Status" in df.columns and "Erasure End Time" in df.columns and "Erasure Device Serial Number" in df.columns:
                # Wczytanie danych z pliku CSV do obiektu DataFrame

                # Usuwanie części dotyczącej strefy czasowej
                df['Erasure End Time'] = df['Erasure End Time'].str.replace('\s\(\+.*\)', '')

                # Konwersja kolumny Erasure End Time na format datetime, ignorując wartości N/A
                df['Erasure End Time'] = pd.to_datetime(df['Erasure End Time'], format='%Y-%m-%d %H:%M:%S', errors='coerce')

                # Konwersja kolumny Erasure End Time na format DD:MM:YYYY
                df['Erasure End Time'] = df['Erasure End Time'].dt.strftime('%d:%m:%Y')

                # Tworzenie pivot table
                pivot_table = df.pivot_table(index='Erasure End Time',
                                             columns='Erasure Status',
                                             values='Erasure Device Serial Number',
                                             aggfunc='count')

                # Wypisanie pivot table
                print(pivot_table)


except (Exception) as e:
    print(e)
