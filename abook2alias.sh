#!/bin/bash
# transforms the output of imap2abook into tha mail.alias format
# that can be used with Mutt

while read a b; 
	do if [[ $a == $b ]]; then 
		echo alias $a \<$a\>
	else 
		echo "alias $a \"$b\" <$a>"
	fi
done
