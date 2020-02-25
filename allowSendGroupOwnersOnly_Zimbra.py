#!/usr/bin/python2.7

###############################################################################
#Access to send emails  only to members of the Zimbra distribution list.
#Example:
#distribution list: "testprueba2@zimbradomain.com"
#             members: John Smith 
#			          Ann Gonzalez
#	These members can send/use this distribution list. Everybody else - no.

#Script:
# 1 - opens text file "distrList.txt" with distribution list
# 2 - goes throu it line by line
# 3 - removes trailing new line and other special characters
# 4 - adds 'allow' permissions to group members
# 5 - reboots Zimbra Milter server

#Note: you do not need to reboot any other Zimbra services 
###############################################################################

import subprocess

# Put your absolute path to the text file here
file = '/opt/zimbra/distrListACL/distrList.txt'

print("Starting script....\n")


with open(file) as f:
    for line in f.readlines():
	    line = line.strip()
		print('Working on ' + line)
    	subprocess.call(['zmprov', 'grr', 'dl', line, 'grp', line, 'sendToDistList'])
        print('Done')

print('Applying configuration to Zimbra Milter server....')
print('Please wait..')
subprocess.call(['zmmtactl', 'reload'])
print("Script completed")
