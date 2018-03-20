from selenium import webdriver
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from selenium.webdriver.support.ui import WebDriverWait
from ConfigParser import SafeConfigParser
from time import sleep
import os
import glob
from smtplib import SMTP
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText



cFile = 'config.cfg'
admin = 'd.chestnov@inlinegroup.ru'
reportFile = "cpapp_admin_cnt_xls_report_CertInd.xlsx"


def sendmail(msg_txt="\nCertification Notification\n", recipients=admin):
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
	driver.maximize_window()
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
	drv.get("https://getlog.cloudapps.cisco.com/WWChannels/GETLOG/services/reports/downloadReports?cert_type=ALL")
	return drv
	

def download_report(drv):
	if os.path.isfile(reportFile):
		os.rename(reportFile, 'tmpFile')
	config = SafeConfigParser()
	config.read(cFile)
	url = config.get('url', 'report')
	drv.get(url)
	sleep(3)
	try:
		drv.get("https://getlog.cloudapps.cisco.com/WWChannels/GETLOG/services/reports/downloadReports?cert_type=ALL")
		os.rename(reportFile, 'cisco.xlsx')
		os.remove('tmpFile')
	except:
		print "Ops, Error with file operation"
	finally:
		drv.quit()



def main():
	drv = login()
	download_report(drv)



if __name__ == '__main__':
	main()




