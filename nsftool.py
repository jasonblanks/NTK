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
import sys, os, win32com.client, shutil

#change to meet your needs
NSFPATH = "C:\\Users\\blanks\\Desktop\\TestNSF"
IDPATH = "C:\\Users\\blanks\\Desktop\\TestNSF"
IDFILEPATH = "C:\\Users\\blanks\\AppData\\Local\\IBM\\Notes\\Data"

#shouldn't change
LOADFILE = "load.csv"
IDFileDefault = "user.id"
IDTemp = os.path.join(IDFILEPATH,"user.id_temp")

TASKS = [line.strip().split(',') for line in open(LOADFILE)]

if os.path.isfile(os.path.join(IDFILEPATH,IDFileDefault)):
	os.rename(os.path.join(IDFILEPATH,IDFileDefault),IDTemp)
logfile = open('log.txt','w')
logfile.write("NSF File\tID File\tPASSWORD\tMSG Count\tSTATUS\n")
for task in TASKS:
	shutil.copyfile(os.path.join(IDPATH,task[1]),os.path.join(IDFILEPATH,"user.id"))

	session = win32com.client.Dispatch("Lotus.NotesSession")
	try:
		session.Initialize(task[2])
		database = session.GetDatabase("", os.path.join(NSFPATH, task[0]))
		docs = database.AllDocuments
	except:
		logfile.write(task[0]+"\t"+task[1]+"\t"+task[2]+"\tNA\tERROR: bad password or password id file mismatch\n")
		os.remove(os.path.join(IDFILEPATH,"user.id"))
		continue
		
	logfile.write(str(task[0])+"\t"+str(task[1])+"\t"+str(task[2])+"\t"+str(docs.Count)+"\tGOOD\n")
	database = session.GetDatabase("", os.path.join(NSFPATH,task[0]))
	docs = database.AllDocuments
	#print "NSFFile: "+task[0]+" IDFile: "+task[1]+" Password: "+task[2]+"MSG Count: "+str(docs.Count)
	os.remove(os.path.join(IDFILEPATH,"user.id"))

if os.path.isfile(IDTemp):
	os.rename(IDTemp,os.path.join(IDFILEPATH,IDFileDefault))
