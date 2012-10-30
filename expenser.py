#!/usr/bin/python
# Expenser
#
# A tool to generate PDFs from Email for easier expense reporting
# (c) 2012 Brian Harrington <bharrington@alticon.net>
#
# Requires: PyQt4 

#  Create a configuration file in ~/.expenser.cfg or the local directory with
#   the following configuration directives
#
#  [expenser]
#    hostname=imap.gmail.com
#    username=test_email@gmail.com
#    password=sample_pass
#    archive=0


import os, sys, datetime, re, base64
import imaplib
import ConfigParser
from  PyQt4.QtCore import *
from  PyQt4.QtGui import *
from  PyQt4.QtWebKit import *
import email
import mimetypes

# Read in values from the config file to know what to parse for the email 
#  settings
# Note:  This has the added value of not storing the information in the code
# to make checkin/out from source control easier

config = ConfigParser.SafeConfigParser(allow_no_value=True)
config.read(['expenser.cfg', os.path.expanduser('~/.expenser.cfg')])
hostname = config.get('expenser','hostname')
username = config.get('expenser','username')
password = config.get('expenser','password')
archive  = config.get('expenser','archive')

# Base path to store files in:
base_path = os.path.expanduser('~/Documents/redhat/expenses/')

# Setup PyQt4 parent objects & a web object for rendering (X)HTML
Appli=QApplication(sys.argv)
web=QWebView()

# Generate a generic printer object and set the formatting to US Letter
printer=QPrinter()
printer.setPageSize(QPrinter.Letter)
printer.setOutputFormat(QPrinter.PdfFormat)

# Perform function calls to log into email via SSL
imap = imaplib.IMAP4_SSL(hostname)
imap.login(username, password)
imap.select()

def getAttachment(msg,check):
  for part in msg.walk():
    if part.get_content_type() == 'application/octet-stream':
      if check(part.get_filename()):
        return part.get_payload(decode=1)


def embassy_suites():
	typ, national_data = imap.uid('search', None, '(HEADER FROM "@hilton.com")')
	for message in national_data[0].split():

		typ, fullmsg = imap.uid('fetch',message, '(RFC822)')
		msg = email.message_from_string(fullmsg[0][1])
		
		name_pat = re.compile('name=\".*\"')
		for part in msg.walk():
			if part.get_content_maintype() != 'application':
				continue

			file_type = part.get_content_type().split('/')[1]
			if not file_type:
				file_type = 'pdf'

			timestamp = datetime.datetime.strptime(str(msg['date'])[0:-12], '%a, %d %b %Y %H:%M:%S')
			timestamp += datetime.timedelta(hours=-5)
			filename = "national-%s.pdf" % timestamp.isoformat('_')[0:10]

			payload = part.get_payload(decode=True)

			if not os.path.isfile(filename) :
				# finally write the stuff
				fp = open(filename, 'wb')
				fp.write(part.get_payload(decode=True))
				fp.close()

# Retreival settings for National Rental Car - 
# http://www.nationalcar.com
def national():
	typ, national_data = imap.uid('search', None, '(HEADER FROM "Customerservice@nationalcar.com")')
	for message in national_data[0].split():

		typ, fullmsg = imap.uid('fetch',message, '(RFC822)')
		msg = email.message_from_string(fullmsg[0][1])
		
		name_pat = re.compile('name=\".*\"')
		for part in msg.walk():
			if part.get_content_maintype() != 'application':
				continue

			file_type = part.get_content_type().split('/')[1]
			if not file_type:
				file_type = 'pdf'

			timestamp = datetime.datetime.strptime(str(msg['date'])[0:-12], '%a, %d %b %Y %H:%M:%S')
			timestamp += datetime.timedelta(hours=-5)
			filename = "national-%s.pdf" % timestamp.isoformat('_')[0:10]

			payload = part.get_payload(decode=True)

			if not os.path.isfile(filename) :
				# finally write the stuff
				fp = open(filename, 'wb')
				fp.write(part.get_payload(decode=True))
				fp.close()

# Retreival settings for Clear Wireless- 
# http://www.clear.com
def clear_wireless():
	typ, clear_data = imap.search(None, 'FROM', '"noreply@Clear.com"')
	for message in clear_data[0].split():
		typ, fullmsg = imap.fetch(message, '(RFC822)')
		
		msg = email.message_from_string(fullmsg[0][1])
		for part in msg.walk():
			if part.get_content_maintype() == 'multipart':
				continue

			if part.get_content_subtype() != 'plain':
				continue

			payload = base64.b64decode(part.get_payload())

		
		header = "From: %s\nTo: %s\nSubject: %s\nDate: %s\n\n" % (msg['from'], msg['to'], msg['subject'], msg['date'])
		document = "<html><title>%s</title><body><font size='4'><pre>%s\n\n%s</pre></font></body></html>" % (msg['subject'], header, payload)
		web.setHtml(document)
		timestamp = datetime.datetime.strptime(str(msg['date'])[0:-6], '%a, %d %b %Y %H:%M:%S')
		timestamp += datetime.timedelta(hours=-5)
		filename = "clear-%s.pdf" % timestamp.isoformat('_')[0:10]
		printer.setOutputFileName(filename)
		web.print_(printer)


	
# Retreival settings for United Airlines- 
# http://www.united.com
def united():
	typ, united_data = imap.search(None, 'FROM', '"unitedairlines@united.com"')
	for message in united_data[0].split():
		typ, fullmsg = imap.fetch(message, '(RFC822)')
		
		msg = email.message_from_string(fullmsg[0][1])

		if not "receipt" in msg['subject'].lower():
			continue
		print msg['subject']
		for part in msg.walk():
                        print " +%s" % part.get_content_maintype()
                        print "   =%s" % part.get_content_subtype()

			if part.get_content_maintype() == 'multipart':
				continue

			if part.get_content_subtype() != 'html':
				continue

			payload = base64.b64decode(part.get_payload())

		filename = "united-%s.pdf" % msg['subject'].rstrip().replace(' ','_').replace(',','')

		web.setHtml(payload)
		printer.setOutputFileName(filename)
		web.print_(printer)

# Retreival settings for Uber car service-
# http://www.uber.com
def uber():
	typ, uber_data = imap.search(None, 'FROM', '"@uber.com"')
	for message in uber_data[0].split():
		typ, fullmsg = imap.fetch(message, '(RFC822)')
		
		msg = email.message_from_string(fullmsg[0][1])
		for part in msg.walk():
			#print part.get_content_maintype()
			#print part.get_content_subtype()
			if part.get_content_maintype() == 'multipart':
				continue

			if part.get_content_subtype() != 'html':
				continue

			payload = base64.b64decode(part.get_payload())

		timestamp = datetime.datetime.strptime(str(msg['date'])[0:-6], '%a, %d %b %Y %H:%M:%S')
		timestamp += datetime.timedelta(hours=-5)
		#filename = "uber-%s.pdf" % timestamp.isoformat('_')[0:10]
		filename = "uber-%s.pdf" % timestamp.isoformat('_')

		web.setHtml(payload)
		printer.setOutputFileName(filename)
		web.print_(printer)
		
# Retreival settings for Marriott Hotel Properties-
# http://www.marriott.com
def hilton():		
	typ, marriott_data = imap.search(None, 'FROM', '"@res.hilton.com"')
	for message in marriott_data[0].split():
		typ, fullmsg = imap.fetch(message, '(RFC822)')
		
		msg = email.message_from_string(fullmsg[0][1])
		for part in msg.walk():
			if part.get_content_maintype() == 'multipart':
				continue

			if part.get_content_subtype() != 'plain':
				continue

			payload = part.get_payload()

		msg_date = msg['subject'][5:31].rstrip('stay').strip(' ').replace(' ','_').replace(',','')
		web.setHtml(payload)
		filename = "hilton-%s.pdf" % msg_date
		printer.setOutputFileName(filename)
		web.print_(printer)
	

# Retreival settings for Marriott Hotel Properties-
# http://www.marriott.com
def marriott():		
	typ, marriott_data = imap.search(None, 'FROM', '"efolio@"')
	for message in marriott_data[0].split():
		typ, fullmsg = imap.fetch(message, '(RFC822)')
		
		msg = email.message_from_string(fullmsg[0][1])
		for part in msg.walk():
			if part.get_content_maintype() == 'multipart':
				continue

			if part.get_content_subtype() != 'html':
				continue

			payload = part.get_payload()

		msg_date = msg['subject'][5:31].rstrip('stay').strip(' ').replace(' ','_').replace(',','')
		web.setHtml(payload)
		filename = "marriott-%s.pdf" % msg_date
		printer.setOutputFileName(filename)
		web.print_(printer)
		


#clear_wireless()
#uber()
#marriott()
#united()
national()
#hilton()

imap.logout()



