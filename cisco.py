# -*- coding: utf-8 -*-

import pandas as pd
from datetime import datetime, timedelta
from selenium import webdriver
from ConfigParser import SafeConfigParser
from time import sleep
import os
from smtplib import SMTP
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText



cFile = 'config.cfg'
admin = ['d.chestnov@inlinegroup.ru', 's.koptenkova@inlinegroup.ru', 'k.tikhanovski@inlinegroup.ru']
reportFile = "cpapp_admin_cnt_xls_report_CertInd.xlsx"
outputFile = 'report.xlsx'
checkDays = 100
checkTime = datetime.now() + timedelta(checkDays)


def sendmail(msg_txt="\nCertification Expiry Warning!\n", subject = 'Certification Status Warning!', recipients=admin):
	sender = 'info@inlinegroup.ru'
	msg = MIMEMultipart()
	msg['From'] = sender
	msg['To'] = ", ".join(recipients)
	msg['Subject'] = subject
	msg.attach(MIMEText(msg_txt.encode('utf-8'),'plain'))
	try:
		server = SMTP('10.8.50.75')
		server.sendmail(sender, recipients, msg.as_string())
		print '[+] E-mail to %s successfully sent' %recipients
	except:
		print "[-] Error sending e-mail"
	finally:
		server.quit()


def set_Firefox_profile():
	profile = webdriver.FirefoxProfile()
	profile.set_preference('browser.download.folderList', 2) 
	profile.set_preference('browser.download.manager.showWhenStarting', False)
	profile.set_preference('browser.download.dir', os.getcwd())
	profile.set_preference('browser.helperApps.neverAsk.saveToDisk', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
	return profile


def login():
	config = SafeConfigParser()
	config.read(cFile)
	user_input = '//*[@id="userInput"]'
	password_input = '//*[@id="passwordInput"]'
	button = '//*[@id="login-button"]'
	# Get User credentials
	user = config.get('auth', 'user')
	password = config.get('auth', 'password')
	# Get Login URL
	url = config.get('url', 'login')
	profile = set_Firefox_profile()
	driver = webdriver.Firefox(firefox_profile=profile)
	driver.get(url)
	driver.find_element_by_xpath(user_input).send_keys(user)
	driver.find_element_by_xpath(button).click()
	sleep(3)
	driver.find_element_by_xpath(password_input).send_keys(password)
	driver.find_element_by_xpath(button).click()
	sleep(3)
	return driver


def get_report(drv):
	url = "https://getlog.cloudapps.cisco.com/WWChannels/GETLOG/welcome.do#/reports"
	drv.get(url)
	sleep(3)
	drv.find_element_by_xpath("/html/body/div[3]/div/div/div[1]/div/img").click()
	sleep(3)
	drv.find_element_by_xpath("/html/body/div[3]/div/div/div[2]/div[2]/div[1]/div/ul/li[8]/a").click()
	sleep(3)
	drv.find_element_by_xpath("/html/body/div[5]/div/div/div[3]/div/server-directive/div[2]/table/tbody/tr/td/div/div[2]/button").click()
	sleep(3)
	if os.path.isfile('tmp'):
			os.remove('tmp')
	if os.path.isfile(reportFile):
			os.rename(reportFile, 'tmp')
	try:
		drv.find_element_by_xpath("/html/body/div[3]/div/div/div[2]/div[2]/div[2]/div[3]/div/form/table/tbody/tr[1]/td/b/a").click()
		os.remove('tmp')
	except Exception as e:
		print e
                if os.path.isfile('tmp'):
		    os.rename('tmp', reportFile)
	finally:
		drv.quit()




def report_to_dict(rFile = reportFile):
	#Read execl to Data Frame and convert object to datetime
	df = pd.read_excel(rFile)
	df["Expiry Date"] = pd.to_datetime(df["Expiry Date"])
	#Check datetime
	df1 = df[(df["Expiry Date"] < checkTime)]
	df2 = df1.loc[:,["First Name", "Last Name", "Email", "Certification", "Certification Description", "Expiry Date"]]
	#Write to Excel
	writer = pd.ExcelWriter(outputFile)
	df2.to_excel(writer, 'Sheet1', index_label=False, index=False, header=True)
	writer.save()
	#convert to Dictionary
	return df2.to_dict(orient='records')


drv = login()
get_report(drv)

print datetime.now().strftime("%d-%b-%Y %H:%M")

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

