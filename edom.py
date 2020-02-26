##############################################################################################
#The script was created under "python2.7"
#This script is adding permissions to users of 2 domains to send emails 
#to internal Zimbra distributions list.  Everyone else will be prohibited.

#Requirements:
#- Install "termcolor module" with - 'pip install termcolor'.
#- Replace domain names for your own.
#- Create your own list of 'distribution list'.
#- Make sure that 'Milter Server' is running on your Zimbra server
#      You can enable it in Zimbra Admin Console - 
#      'Home > Configure > Servers > yourServer > MTA'
#      To start the 'milter server' use command - 'zmmilterctl start'
#          Check the status - 'zmmilterctl status'


#Execute script under 'zimbra' user account as - 'python edom.py'.

#To reverse script effect replace command for required domain with:
#>>>>  subprocess.call(['zmprov', 'rvr', 'dl', line, 'edom', domainCOM, 'sendToDistList'])

##############################################################################################


import subprocess
import os
from termcolor import colored

distrListFile = '/opt/zimbra/distrListACL/distrList.txt'
domainCL = 'yourdomain.cl'
domainCOM = 'yourdomain.com'
pathToScriptDir = '/opt/zimbra/distrListACL/'
scriptFileName = 'edom.py'

print colored("Starting script....\n", 'green')


try:
	print colored('Checking if file with distribution list exist', 'green')
	(os.path.isdir(pathToScriptDir)) and (os.path.isfile(distrListFile))
	try:
		with open(distrListFile) as f:
			for line in f.readlines():
				line = line.strip()
				subprocess.call(['zmprov', 'grr', 'dl', line, 'edom', domainCL, 'sendToDistList'])
				subprocess.call(['zmprov', 'grr', 'dl', line, 'edom', domainCOM, 'sendToDistList'])
				print colored(line + ' ->  Done.', 'green')
	except IOError as err:
		print colored(err, 'red')
except IOError as err1:
	print colored(err1, 'red')
else:
	print colored('Please insert code to create log file for this script', 'yellow')

subprocess.call(['zmmtactl', 'reload'])
print colored("Script completed", 'green')
