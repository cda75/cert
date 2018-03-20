# -*- coding: utf-8 -*-

import pandas as pd
from datetime import datetime, timedelta
from login import sendmail



fullFile = 'cisco.xlsx'
reportFile = 'report.xlsx'
checkDays = 100
checkTime = datetime.now() + timedelta(checkDays)

#Read execl to Data Frame and convert object to datetime
df = pd.read_excel(fullFile)
df["Expiry Date"] = pd.to_datetime(df["Expiry Date"])

#Check datetime
df1 = df[(df["Expiry Date"] < checkTime)]
df2 = df1.loc[:,["First Name", "Last Name", "Email", "Certification", "Certification Description", "Expiry Date"]]
#print df2.head()

#Write to Excel
writer = pd.ExcelWriter(reportFile)
df2.to_excel(writer, 'Sheet1', index_label=False, index=False, header=True)
writer.save()

#convert to Dictionary
d = df2.to_dict(orient='records')
msg_header = 'Следующие сертификационные статусы Cisco близки к окончанию срока действия!\n\n'.decode('utf-8')
admin_msg = ''
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