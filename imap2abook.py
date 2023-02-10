#!/usr/bin/env python3
import imaplib
import email
import sys
import datetime
import time
import getopt
from getpass import getpass

def usage():
	print ("imap2abook - harvest destination addresses from an IMAP mailbox to create an address book")
	print ("Options:")
	print ("  -s, --server:    address of the IMAPS server")
	print ("  -p, --port:	   port of the IMAPS server")
	print ("  -u, --user:	   user name")
	print ("  -f, --folder:    IMAP folder")
	print ("  -p, --password:  optional password - will prompt for one if not specified")
	print ("  -a, --maxage:    maximum age (in days) of messages to harvest (optional; default=unlimited)")
	print ("  -v, --vcard:	   output in vCard format")
	print ("  -e, --email:	   for addresses with no name, use e-mail address as name; otherwise omitted")
	print ("  -F, --from:	   only consider messages sent From: this address (typically your own)")
	print ("  -h, --help:	   show this help info")

server = None
port = None
user = None
password = None
folder = None
vcard = False
emailasname = False
maxage = 0
from_address = None;

try:
	opts, args = getopt.getopt(sys.argv[1:], "hs:p:u:w:f:a:veF:", ["help", "server=", "port=", "user=", "password=", "folder=", "maxage=", "vcard", "email", "from=" ])
except getopt.GetoptError as err:
	print(err, file=sys.stderr)
	usage
	sys.exit(1)

for opt, arg in opts:
	if opt in ("-h", "--help"):
		usage()
		sys.exit(0)
	elif opt in ("-s", "--server"):
		server = arg
	elif opt in ("-p", "--port"):
		port = arg
	elif opt in ("-u", "--user"):
		user = arg
	elif opt in ("-w", "--password"):
		password = arg
	elif opt in ("-f", "--folder"):
		folder = arg
	elif opt in ("-a", "--maxage"):
		try:
			maxage = int(arg)
		except:
			pass
	elif opt in ("-v", "--vcard"):
		vcard = True
	elif opt in ("-e", "--email"):
		emailasname = True
	elif opt in ("-F", "--from"):
		from_address = arg;
	else:
		usage()
		sys.exit(1)
if not server or not port or not user or not folder:
	usage()
	sys.exit(1)

if not password:
	password = getpass()

if maxage:
	mindate = time.mktime((datetime.datetime.now() - datetime.timedelta(days=maxage)).timetuple())
else:
	mindate = 0

try:
	conn = imaplib.IMAP4_SSL(server, port)
	conn.login(user, password)
except imaplib.IMAP4.error as e:
	print("Cannot conect to IMAP server " + server + ":" + port + ": " + str(e), file=sys.stderr)
	exit(1)

try:
	conn.select(mailbox='"'+folder+'"', readonly=True)
except imaplib.IMAP4.error as e:
	print("Cannot select folder " + folder + ": " + str (e), file=sys.stderr)
	exit(1)

try:
	d = datetime.datetime.fromtimestamp(mindate).strftime("%d-%b-%Y")
	search = "SENTSINCE " + d
	if from_address:
		search += " FROM \"" + from_address + "\""
	typ, data = conn.search(None, "(" + search + ")")
	data = data[0].decode()
	count = len(data.split())
	num = data.replace(" ", ",")
	if not num:
		print("WARNING: no messages", file=sys.stderr)
		exit(0)
	print("Number of messages to download:", count, file=sys.stderr)

except imaplib.IMAP4.error as e:
       print("WARNING: cannot search messages, working with complete folder: " + str (e), file=sys.stderr)
       num="1:*"

try:
	typ, msg_data = conn.fetch(num, '(BODY.PEEK[HEADER.FIELDS (TO CC BCC DATE)])')
except imaplib.IMAP4.error as e:
	print("Cannot fetch headers: " + str(e), file=sys.stderr)
	exit(1)

abook=dict()


for m in msg_data:
	if not type(m) is tuple:
		continue
	try:
		msg = email.message_from_bytes(m[1])
		
		to = msg.get_all("To", [])
		cc = msg.get_all("Cc", [])
		bcc = msg.get_all("Bcc", [])
		date = msg.get_all("Date", [])
		
		try:
			date = email.utils.parsedate(msg['Date'])
			date = time.mktime(date)
			if date and date < mindate:
				continue
		except:
			pass

		for name, addr in email.utils.getaddresses(to + cc + bcc):
			addr = addr.lower()
			if addr:
				if name:
					name = str(email.header.make_header(email.header.decode_header(name)))
					abook[addr] = name
				else:
					if not addr in abook.keys():
						abook[addr] = None

	except:
		print("Cannot decode message:\n" + m[1].decode("utf-8"), file=sys.stderr)
		pass
		


if vcard:
	for adr, name in abook.items():
		if not name:
			if emailasname:
				name = adr
			else:
				continue

		print("BEGIN:VCARD")
		print("VERSION:2.1")
		print("N:" + name)
		print("EMAIL:" + adr)
		print("END:VCARD")
else:
	for adr, name in abook.items():
		if not name:
			if emailasname:
				name = adr
			else:
				continue
		print(adr, "\t" , name)
	
try:
    conn.logout()
    conn.close()
except:
     pass
