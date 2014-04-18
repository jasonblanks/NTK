'''	NSF Validation Tool v .9
	This is the first attempt at using python to interact with com
	and run various Lotus com functions and too validate deliveries.

	This is tested using IBM Notes 9 standalone client
	This script temparly moves the user's default user.id file to user.id_temp
	and runs through the load.csv file copying the id and nsf files as needed
	from the base directory.

	Requirements:
	IBM Notes 9 standalone client
	NotesSQL 9.0 (Also called IBM ODBC Driver for Notes/Domino 9.0)
	Python 2.7
	win32com (http://sourceforge.net/projects/pywin32/files/pywin32/Build%20218/pywin32-218.win32-py2.7.exe/download))


	MAJOR TODO: Better error reporting and handling.


'''
import sys, os, win32com.client, shutil, getopt, subprocess, fileinput, pyodbc, re
import pythoncom
import multiprocessing
import time

#Environment Variables
##TODO: Fix so it accepts spaces in file path.
NSFPATH = r'Y:\NSF'
IDPATH = r'Y:\NSF\IDs'
##TODO:Should be able to look this path automatically.
LotusDataPATH = r'C:\Users\blanksj\AppData\Local\IBM\Notes\Data'
inifile = r'C:\Users\blanksj\AppData\Local\IBM\Notes\Data\notes.ini'
logpath = r'Y:\NSF'
##TODO: Add command line arg to specify this file path.
LOADFILE = r'Y:\NSF\load.txt'
DummyFile = r'C:\Users\blanksj\Desktop\test\dummy.id'
NotesSQLCFG = r'C:\NotesSQL\notessql.cfg'
workingDir = r'Y:\NSF'
BAD = []
GOOD = []

def NSFDecrypt(db, task, NSFPATH, logfile, DELETE, logpath):
		cloneFilename = task[0].split('.')
		dbclone = db.CreateFromTemplate("",os.path.join(NSFPATH, cloneFilename[0])+"--decrypt.nsf", False)
		#dbclone.Compact

		
		# We determianed you not only have to remove encryption by compact but clear out the ACL on that NSF
		# Below is a quick hack to allow everyone access, I would like to make this cleaner my creating a function
		# to clear out the whole ACL.
		dbclone.CompactWithOptions("L")
		dbclone.GrantAccess( "-Default-", "6" )
		dbclone.GrantAccess( "Anonymous", "6" )

		if DELETE:
			os.remove(os.path.join(NSFPATH, task[0]))

		OriginalDocCount = db.AllDocuments
		CloneDocCount = dbclone.AllDocuments

		if  CloneDocCount.Count == OriginalDocCount.Count:
			logfile = open(os.path.join(logpath,"decrypt_log.txt"),"a")
			logfile.write(str(task[0])+"\t"+str(task[1])+"\t"+str(task[2])+"\t"+str(CloneDocCount.Count)+"\tDecrypted\n")
			logfile.close()

		elif  CloneDocCount.Count != OriginalDocCount.Count:
			logfile = open(os.path.join(logpath,"decrypt_log.txt"),"a")
			logfile.write(str(task[0])+"\t"+str(task[1])+"\t"+str(task[2])+"\t"+str(CloneDocCount.Count)+"\tDecryption Failed\n")
			logfile.close()


#Useless function as all id/password combinations are set in the load file now.
def BruteForce(bad, TASKS, logfile, IDPATH, LotusDataPATH, NSFPATH, DELETE, logpath):
	for BadNSF in bad:
		for ID in TASKS:
			for PASSWORD in TASKS:
				shutil.copyfile(os.path.join(IDPATH,ID[1]),os.path.join(LotusDataPATH,"user.id"))
				session[count-1] = win32com.client.Dispatch("Lotus.NotesSession[count-1]")

				try:
					session[count-1].Initialize(PASSWORD[2])
					database = session[count-1].GetDatabase("", os.path.join(NSFPATH, BadNSF[0]))
					docs = database.AllDocuments
				except:
					logfile.write(BadNSF[0]+"\t"+ID[1]+"\t"+PASSWORD[2]+"\tNA\tERROR: BF Attempt Bad\n")
					os.remove(os.path.join(LotusDataPATH,"user.id"))
					continue
				logfile.write(str(BadNSF[0])+"\t"+str(ID[1])+"\t"+str(PASSWORD[2])+"\t"+str(docs.Count)+"\tGOOD\tBRUTEFORCE\n")
				#NSFDecrypt(database, task, NSFPATH, logfile)
				os.remove(os.path.join(LotusDataPATH,"user.id"))


def Validate(NSFPATH, IDPATH, LotusDataPATH, LOADFILE, decrypt, bruteForce, DELETE, inifile, GOOD, BAD, NotesSQLCFG, logpath, workingDir, DummyFile):
	#TODO: Add os check to see if LOADFILE exists. If not throw error and exit.
	TASKS = [line.strip().split(',') for line in open(LOADFILE)]
	logfile = open(os.path.join(logpath,"log.txt"),"w")
	logfile.write("NSF File\tID File\tPASSWORD\tMSG Count\tSTATUS\n")
	logfile.close()

	for line in fileinput.FileInput(inifile, inplace=1):
		if line.startswith("KeyFilename="):
			defaultID = line
			line = line.replace(line, "KeyFilename="+DummyFile)
		if line.startswith("KeyFileName_Owner"):
			defaultOwner = line
			line = line.replace(line, r'')
		if line.startswith("Location="):
			defaultOwner = line
			line = line.replace(line, "")
		if line.startswith("Directory="):
			defaultOwner = line
			line = line.replace(line, "Directory="+NSFPATH)
		print line.strip()

	try:
		session = win32com.client.Dispatch("Lotus.NotesSession")
		session.Initialize()
	except Exception as inst:
			print type(inst)
			print inst.args
			if type(inst) == AttributeError:
				print inst.args

	def buildfile(TASKS, NotesSQLCFG):
		CFGFile = open(NotesSQLCFG,"w")
		CFGFile.write("[Users]\n")
		CFGFile.close()
		for root, dirs, files in os.walk(workingDir):
			for dir in dirs:
				count = 0
				for custodian in TASKS:
					if custodian[0] == dir:
						CFGFile = open(NotesSQLCFG,"a")
						#Test full path
						CFGFile.write(custodian[0]+"/"+str(count)+"="+os.path.join(IDPATH,custodian[1])+"\n")
						#CFGFile.write(custodian[0]+"/"+str(count)+"="+custodian[1]+"\n")
						CFGFile.close()
						count = count + 1
	buildfile(TASKS, NotesSQLCFG)

	for root, dirs, files in os.walk(workingDir, topdown=False):
		for custodian in TASKS:
			count = 0
			for file in files:
				if custodian[0] in root.split('\\'):
					if file.endswith('nsf'):
						try:
							os.chdir(root)
							#Test moving id file to cwd.
							#print os.getcwd()
							#print os.path.join(IDPATH,custodian[1]) +"--"+custodian[1]
							##shutil.copy2(os.path.join(IDPATH,custodian[1]),".")
							#print "Driver={Lotus Notes SQL Driver (*.nsf)};UID="+custodian[0]+"/"+str(count)+";PWD="+custodian[2]+"; DATABASE="+os.path.join(root,file)+""
							connection=pyodbc.connect("Driver={Lotus Notes SQL Driver (*.nsf)};UID="+custodian[0]+"/"+str(count)+";PWD="+custodian[2]+"; DATABASE="+os.path.join(root,file)+"", autocommit=True)
							#connection=pyodbc.connect("Driver={Lotus Notes SQL Driver (*.nsf)};UID="+custodian[0]+"/"+str(count)+";PWD="+custodian[2]+"; DATABASE="+file+"", autocommit=True)
							#os.remove(os.path.join(IDPATH,custodian[1]))
							if connection:
								GOOD.append((os.path.join(root,file), custodian[1], custodian[2]))
								#print GOOD
						except Exception as inst:
										#print inst.args
										a,b = inst.args
										if re.search('Wrong Password',b):
											logfile = open(os.path.join(logpath,"log.txt"),"a")
											logfile.write(str(custodian[0])+"\t"+str(custodian[2])+"\t"+str(custodian[1])+"\t\tbad password/ID Combination\n")
											logfile.close()
											#print str(custodian[0])+"\t "+str(custodian[2])+"\t "+str(custodian[1])+"\t\tbad password/ID Combination\n"
										else:
											logfile = open(os.path.join(logpath,"log.txt"),"a")
											logfile.write(str(custodian[0])+"\t"+str(custodian[2])+"\t"+str(custodian[1])+"\t"+b+"\n")
											logfile.close()
	for task in GOOD:
		#print "working on second stage"
		#print task
		try:
			reg = session.createRegistration()
			reg.switchToID(os.path.join(IDPATH,task[1]),task[2])

		except Exception as inst:
				#print inst
				logfile = open(os.path.join(logpath,"log.txt"),"a")
				logfile.write(str(custodian[0])+"\t"+str(custodian[2])+"\t"+str(custodian[1])+"\t"+inst.args+"\n")
				logfile.close()
				#if type(inst) == AttributeError:
					#x, y ,u , i = inst.arg
					#print inst

		try:
			database = session.GetDatabase("", os.path.join(root, task[0]))
			docs = database.AllDocuments
			#print str(task[0])+"\t"+str(task[1])+"\t"+str(task[2])+"\t"+str(docs.Count)+"\tGOOD\n"
			logfile = open(os.path.join(logpath,"log.txt"),"a")
			logfile.write(str(task[0])+"\t"+str(task[1])+"\t"+str(task[2])+"\t"+str(docs.Count)+"\tWorks\n")
			logfile.close()
			if decrypt == 1:
				NSFDecrypt(database, task, NSFPATH, logfile, DELETE, logpath)
				#print str(task[0])+"\t"+str(task[1])+"\t"+str(task[2])+"\t"+str(docs.Count)+"\tworks\n"

			for line in fileinput.FileInput(inifile, inplace=1):
				if line.startswith("KeyFileName="):
					defaultID = line
					line = line.replace(line, "KeyFilename="+DummyFile)
				print line.strip()
			try:
				session = win32com.client.Dispatch("Lotus.NotesSession")
				session.Initialize()
			except Exception as inst:
					if type(inst) == AttributeError:
						logfile = open(os.path.join(logpath,"log.txt"),"a")
						logfile.write(str(custodian[0])+"\t"+str(custodian[2])+"\t"+str(custodian[1])+"\t"+inst.args+"\n")
						logfile.close()
						#print inst.args



		except Exception as inst:
			print inst.args
			x, y ,u , i = inst.args
			#print inst
			logfile = open(os.path.join(logpath,"log.txt"),"a")
			logfile.write(str(custodian[0])+"\t"+str(custodian[2])+"\t"+str(custodian[1])+"\t"+str(x)+str(y)+str(u)+str(i)+"\n")
			logfile.close()
			#if type(inst) == TypeError:
				#x, y ,u , i = inst.args
				#print inst.args

	#if decrypt == 1:
		#NSFDecrypt(database, task, NSFPATH, logfile, inifile,  DELETE, GOOD)
	#if bruteForce == 1:
			#BruteForce(bad, TASKS, logfile, IDPATH, LotusDataPATH, NSFPATH, DELETE, BAD)

#todo Finsh pass through function to identify none-secured NSF files
def CheckPwdProtected(NSFPATH, IDPATH, LotusDataPATH, LOADFILE, decrypt, bruteForce, DELETE, inifile, GOOD, BAD, NotesSQLCFG, logpath, workingDir, DummyFile):
	for line in fileinput.FileInput(inifile, inplace=1):
		if line.startswith("KeyFilename="):
			defaultID = line
			line = line.replace(line, "KeyFilename="+DummyFile)
		if line.startswith("KeyFileName_Owner"):
			defaultOwner = line
			line = line.replace(line, r'')
		if line.startswith("Location="):
			defaultOwner = line
			line = line.replace(line, "")
		if line.startswith("Directory="):
			defaultOwner = line
			line = line.replace(line, "Directory="+NSFPATH)
		print line.strip()

	try:
		session = win32com.client.Dispatch("Lotus.NotesSession")
		session.Initialize()
	except Exception as inst:
			logfile = open(os.path.join(logpath,"log.txt"),"a")
			logfile.write(str(custodian[0])+"\t"+str(custodian[2])+"\t"+str(custodian[1])+"\t"+inst+"\n")
			logfile.close()
			#print type(inst)
			#print inst.args
			#if type(inst) == AttributeError:
				#print inst.args

	for root, dirs, files in os.walk(workingDir, topdown=False):
		for file in files:
			try:
				session = win32com.client.Dispatch("Lotus.NotesSession")
				session.Initialize()
			except Exception as inst:
					logfile = open(os.path.join(logpath,"log.txt"),"a")
					logfile.write(str(custodian[0])+"\t"+str(custodian[2])+"\t"+str(custodian[1])+"\t"+inst+"\n")
					logfile.close()
					#print type(inst)
					#print inst.args
					#if type(inst) == AttributeError:
						#print inst.args
			if file.endswith('nsf'):
				try:
					os.chdir(root)
					connection=pyodbc.connect("Driver={Lotus Notes SQL Driver (*.nsf)};UID=dummy;DATABASE="+os.path.join(root,file)+"", autocommit=True)

				except Exception as inst:
					#print type(inst)
					#print inst.args
					logfile = open(os.path.join(logpath,"log.txt"),"a")
					logfile.write(str(custodian[0])+"\t"+str(custodian[2])+"\t"+str(custodian[1])+"\t"+inst+"\n")
					logfile.close()


def main(argv, GOOD, BAD, DummyFile, inifile):
	decrypt = 0
	bruteForce = 0
	DELETE = 0
	for line in fileinput.FileInput(inifile, inplace=1):
		if line.startswith("KeyFileName="):
			defaultID = line
			line = line.replace(line, "KeyFileName="+DummyFile)
		if line.startswith("KeyFileName_Owner"):
			defaultOwner = line
			line = line.replace(line, r'')
		if line.startswith("Location="):
			defaultOwner = line
			line = line.replace(line, "")
		if line.startswith("Directory="):
			defaultOwner = line
			line = line.replace(line, "Directory="+NSFPATH)
		print line.strip()

	try:
		opts, args = getopt.getopt(argv,"hdcs",)
		print opts
		print args
	except getopt.GetoptError:
		print 'test.py -h -c -d '
		sys.exit(2)

	#TODO: add arg to delete old file after decrypt.
	for opt, arg in opts:
		if opt == '-h':
			print 'test.py -d(decrypt valid files) -c(brute force with all given id and password combinations.)'
			sys.exit()
		elif opt == "-d":
			decrypt = True
			Validate(NSFPATH, IDPATH, LotusDataPATH, LOADFILE, decrypt, bruteForce, DELETE, inifile, GOOD, BAD, NotesSQLCFG, logpath, workingDir, DummyFile)
		elif opt == "-p":
			decrypt = True
			CheckPwdProtected(NSFPATH, IDPATH, LotusDataPATH, LOADFILE, decrypt, bruteForce, DELETE, inifile, GOOD, BAD, NotesSQLCFG)
		elif opt ==  "-c":
			bruteForce = True
			Validate(NSFPATH, IDPATH, LotusDataPATH, LOADFILE, IDDefault, IDTemp, decrypt, bruteForce, DELETE, logpath)
		elif opt == "-s":
			decrypt = True
			CheckPwdProtected(NSFPATH, IDPATH, LotusDataPATH, LOADFILE, decrypt, bruteForce, DELETE, inifile, GOOD, BAD, NotesSQLCFG, logpath, workingDir, DummyFile)
		else:
			Validate(NSFPATH, IDPATH, LotusDataPATH, LOADFILE, IDDefault, IDTemp, decrypt, bruteForce, DELETE, BAD, GOOD, logpath, workingDir)
	if not opts:
		Validate(NSFPATH, IDPATH, LotusDataPATH, LOADFILE, decrypt, bruteForce, DELETE, inifile, GOOD, BAD, NotesSQLCFG, logpath, workingDir, DummyFile)

if __name__ == "__main__":
	main(sys.argv[1:], GOOD, BAD, DummyFile, inifile)
