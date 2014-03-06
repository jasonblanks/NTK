#	NSF Validation Tool v .1
#	
#	This is the first attempt at using python to interact with com
#	and run various Lotus com functions and too validate deliveries.
#	
#	This is tested using IBM Notes 9 standalone client
#	This script temparly moves the user's default user.id file to user.id_temp
#	and runs through the load.csv file copying the id and nsf files as needed
#	from the base directory.
#	
#	Requiremnts:
#	IBM Notes 9 standalone client
#	Python 2.7
#
import os, win32com.client

#change to meet your needs
NSFPATH = "C:\\Users\\blanks\\Desktop\\TestNSF"
IDPATH = "C:\\Users\\blanks\\Desktop\\TestNSF\\IDFiles"
LOTUSDATAPATH = "C:\\Users\\blanks\\AppData\\Local\\IBM\\Notes\\Data"

#shouldn't need to change
LOADFILE = "load.csv"

TASKS = [line.strip().split(',') for line in open(LOADFILE)]

logfile = open('log.txt','w')
logfile.write("NSF File\tID File\tPASSWORD\tMSG Count\tSTATUS\n")
for task in TASKS:

	session = win32com.client.Dispatch("Lotus.NotesSession")
	try:
		session.Initialize(task[2])
		reg = session.CreateRegistration()
		reg.SwitchToID(os.path.join(IDPATH,task[1]),task[2])
		database = session.GetDatabase("", os.path.join(NSFPATH, task[0]))
		docs = database.AllDocuments
		logfile.write(str(task[0])+"\t"+str(task[1])+"\t"+str(task[2])+"\t"+str(docs.Count)+"\tGOOD\n")
	except:
		logfile.write(task[0]+"\t"+task[1]+"\t"+task[2]+"\tNA\tERROR: bad password or password id file mismatch\n")
		continue
		

