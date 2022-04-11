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
        print ("  -p, --port:      port of the IMAPS server")
        print ("  -u, --user:      user name")
        print ("  -f, --folder:    IMAP folder")
        print ("  -p, --password:  optional password - will prompt for one if not specified")
        print ("  -a, --maxage:    maximum age (in days) of messages to harvest (optional; default=unlimited)")
        print ("  -v, --vcard:     output in vCard format")
        print ("  -e, --email:     for addresses with no name, use e-mail address as name; otherwise omitted")
        print ("  -h, --help:      show this help info")

server = None
port = None
user = None
password = None
folder = None
vcard = False
emailasname = False
maxage = 0

try:
        opts, args = getopt.getopt(sys.argv[1:], "hs:p:u:w:f:a:ve", ["help", "server=", "port=", "user=", "password=", "folder=", "maxage=", "vcard", "email" ])
except getopt.GetoptError as err:
        print(err)
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
except:
        print("Cannot conect to IMAP server " + server + ":" + port, file=sys.stderr)
        exit(1)

try:
        conn.select(mailbox='"'+folder+'"', readonly=True)
except:
        print("Cannot select folder " + folder, file=sys.stderr)
        exit(1)

num="1:*"
try:
        typ, msg_data = conn.fetch(num, '(BODY.PEEK[HEADER.FIELDS (TO CC BCC DATE)])')
except:
        print("Cannot fetch headers")
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
