#!/bin/bash
# transforms the output of imap2abook into the VCARD format
# imap2abook can output vcard on its own, but sometimes you
# might want both formats and it's not ideal to run
# a slow imap2abook twice

while read a b; do
	echo "BEGIN:VCARD"
	echo "VERSION:2.1"
	echo "N: $a"
	echo "EMAIL: $adr"
	echo "END:VCARD"
done
