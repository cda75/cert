# -*- coding: utf-8 -*-

import pandas as pd
from datetime import datetime, timedelta
from login import sendmail
from selenium import webdriver
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from selenium.webdriver.support.ui import WebDriverWait
from ConfigParser import SafeConfigParser
from time import sleep
import os
from smtplib import SMTP
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText





cFile = 'config.cfg'
admin = 'd.chestnov@inlinegroup.ru'
reportFile = "cpapp_admin_cnt_xls_report_CertInd.xlsx"
outputFile = 'report.xlsx'
checkDays = 100
checkTime = datetime.now() + timedelta(checkDays)




def sendmail(msg_txt="\nCertification Expiry Warning!\n", recipients=admin):
	config = SafeConfigParser()
	config.read(cFile)
	sender = config.get('email', 'sender')
	subject = config.get('email', 'subject')
	msg = MIMEMultipart()
	msg['From'] = sender
	msg['To'] = recipients
	msg['Subject'] = subject
	msg.attach(MIMEText(msg_txt.encode('utf-8'),'plain'))
	try:
		server = SMTP('mail.inlinegroup.ru')
		server.sendmail(sender, recipients, msg.as_string())
		print '[+] E-mail successfully sent'
	except:
		print "[-] Error sending e-mail"
	finally:
		server.quit()


def login():
	config = SafeConfigParser()
	config.read(cFile)
	# Get User credentials
	user = config.get('auth', 'user')
	password = config.get('auth', 'password')
	# Get Login URL
	url = config.get('url', 'login')
	#Get Elements XPath
	user_input = config.get('xpath', 'user_input')
	password_input = config.get('xpath', 'password_input')
	button = config.get('xpath', 'login_button')
	#Start session
	profile = webdriver.FirefoxProfile()
	profile.set_preference('browser.download.folderList', 2) 
	profile.set_preference('browser.download.manager.showWhenStarting', False)
	profile.set_preference('browser.download.dir', os.getcwd())
	profile.set_preference('browser.helperApps.neverAsk.saveToDisk', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        driver = webdriver.Firefox(firefox_profile=profile)
	wait = WebDriverWait(driver, 10)
	#driver.maximize_window()
	driver.get(url)
	driver.find_element_by_xpath(user_input).send_keys(user)
	driver.find_element_by_xpath(button).click()
	sleep(3)
	driver.find_element_by_xpath(password_input).send_keys(password)
	driver.find_element_by_xpath(button).click()
	sleep(3)
	return driver


def get_report(drv):
	config = SafeConfigParser()
	config.read(cFile)
	url = config.get('url', 'report')
	button = config.get('xpath', 'geo_select')
	drv.get(url)
	sleep(3)
	drv.find_element_by_xpath("/html/body/div[3]/div/div/div[1]/div/img").click()
	sleep(3)
	drv.find_element_by_xpath("/html/body/div[3]/div/div/div[2]/div[2]/div[1]/div/ul/li[8]/a").click()
	sleep(3)
	drv.find_element_by_xpath(button).click()
	sleep(3)
	try:
		if os.path.isfile(reportFile):
			os.rename(reportFile, 'tmpFile')
		drv.find_element_by_xpath("/html/body/div[3]/div/div/div[2]/div[2]/div[2]/div[3]/div/form/table/tbody/tr[1]/td/b/a").click()
		result = True
	except:
		print "Oops! Error with file operation"
		result = False
	finally:
		drv.quit()
		return result



def report_to_dict(rFile = reportFile):
	#Read execl to Data Frame and convert object to datetime
	df = pd.read_excel(rFile)
	df["Expiry Date"] = pd.to_datetime(df["Expiry Date"])
	#Check datetime
	df1 = df[(df["Expiry Date"] < checkTime)]
	df2 = df1.loc[:,["First Name", "Last Name", "Email", "Certification", "Certification Description", "Expiry Date"]]
	#Write to Excel
	writer = pd.ExcelWriter(reportFile)
	df2.to_excel(writer, 'Sheet1', index_label=False, index=False, header=True)
	writer.save()
	#convert to Dictionary
	return df2.to_dict(orient='records')


drv = login()
get_report(drv)


msg_header = 'Следующие сертификационные статусы Cisco близки к окончанию срока действия!\n\n'.decode('utf-8')
admin_msg = ''
d = report_to_dict()
for i in d:
	msg = "%s\t%s\t%s\t%s\t%s\n" %(i['First Name'].ljust(12), i['Last Name'].ljust(13), i['Certification'].ljust(12), i['Expiry Date'].strftime('%d-%b-%Y').ljust(12), i['Certification Description'])
	recipient = i['Email']
	#sendmail(msg_header + msg, recipient)
	print msg
	admin_msg += msg


if admin_msg:
	sendmail(msg_header+admin_msg)
	print admin_msg
else:
	print "No messages"
