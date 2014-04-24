'''	NSF Validation Tool v .8.5
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
import sys, os, win32com.client, shutil, getopt, subprocess, fileinput, pyodbc, re, hashlib
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
logpath = r'C:\temp\test'
#logpath = r'Y:\NSF'
##TODO: Add command line arg to specify this file path.
LOADFILE = r'Y:\NSF\load.txt'
DummyFile = r'C:\Users\blanksj\Desktop\test\dummy.id'
NotesSQLCFG = r'C:\NotesSQL\notessql.cfg'
workingDir = r'C:\temp\test'
#workingDir = r'Y:\NSF'
BAD = []
GOOD = []

def NSFDecrypt(db, task, NSFPATH, logfile, DELETE, logpath):
		#TASKS is (root, file, username, IDFile, Password)
		cloneFilename = task[1].split('.')
		dbclone = db.CreateFromTemplate("",os.path.join(task[0], cloneFilename[0])+"--decrypt.nsf", False)
		#dbclone.Compact

		
		# We determianed you not only have to remove encryption by compact but clear out the ACL on that NSF
		# Below is a quick hack to allow everyone access, I would like to make this cleaner my creating a function
		# to clear out the whole ACL.
		dbclone.CompactWithOptions("L")
		dbclone.GrantAccess( "-Default-", "6" )
		dbclone.GrantAccess( "Anonymous", "6" )

		if DELETE:
			os.remove(os.path.join(task[0], task[1]))

		OriginalDocCount = db.AllDocuments
		CloneDocCount = dbclone.AllDocuments

		if  CloneDocCount.Count == OriginalDocCount.Count:
			logfile = open(os.path.join(logpath,"log.txt"),"a")
			logfile.write(str(os.path.join(task[0], task[1]))+"\t"+str(task[1])+"\t"+str(task[2])+"\tN\\A\t"+str(OriginalDocCount.Count)+"\t"+str(CloneDocCount.Count)+"\t"+str(os.path.getsize(os.path.join(task[0], cloneFilename[0])+"--decrypt.nsf"))+"\tDecrypted\n")
			logfile.close()

		elif  CloneDocCount.Count != OriginalDocCount.Count:
			logfile = open(os.path.join(logpath,"log.txt"),"a")
			logfile.write(str(os.path.join(task[0], task[1]))+"\t"+str(task[1])+"\t"+str(task[2])+"\tN\\A\t"+str(OriginalDocCount.Count)+"\t"+str(CloneDocCount.Count)+"\tDecryption Failed: Count Mismatch\n")
			logfile.close()


#Useless function as all id/password combinations are set in the load file now.
def BruteForce(bad, TASKS, logfile, IDPATH, LotusDataPATH, NSFPATH, DELETE, logpath):
	#TASKS is (root, file, username, IDFile, Password)
	for BadNSF in bad:
		for ID in TASKS:
			for PASSWORD in TASKS:
				shutil.copyfile(os.path.join(IDPATH,ID[3]),os.path.join(LotusDataPATH,"user.id"))
				session[count-1] = win32com.client.Dispatch("Lotus.NotesSession[count-1]")

				try:
					session[count-1].Initialize(PASSWORD[4])
					database = session[count-1].GetDatabase("", os.path.join(NSFPATH, BadNSF[0]))
					docs = database.AllDocuments
				except:
					logfile.write(BadNSF[0]+"\t"+BadNSF[1]+"\t"+ID[3]+"\t"+PASSWORD[4]+"\tNA\tERROR: BF Attempt Bad\n")
					os.remove(os.path.join(LotusDataPATH,"user.id"))
					continue
				logfile.write(BadNSF[0]+"\t"+BadNSF[1]+"\t"+ID[3]+"\t"+PASSWORD[4]+"\t"+str(docs.Count)+"\tGOOD\tBRUTEFORCE\n")
				#NSFDecrypt(database, task, NSFPATH, logfile)
				os.remove(os.path.join(LotusDataPATH,"user.id"))


def Validate(NSFPATH, IDPATH, LotusDataPATH, LOADFILE, decrypt, bruteForce, DELETE, inifile, GOOD, BAD, NotesSQLCFG, logpath, workingDir, DummyFile, filenameBlacklist, hashBlacklist):
	#TODO: Add os check to see if LOADFILE exists. If not throw error and exit.
	TASKS = [line.strip().split(',') for line in open(LOADFILE)]
	logfile = open(os.path.join(logpath,"log.txt"),"w")
	logfile.write("NSF File Path\tNSF File\tID File\tPASSWORD\tOriginal Count\tDecrypted Count\tFile Size\tSTATUS\n")
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
						CFGFile.write(custodian[0]+"/"+str(count)+"="+os.path.join(IDPATH,custodian[1])+"\n")
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
							if file in filenameBlacklist:
								logfile = open(os.path.join(logpath,"log.txt"),"a")
								logfile.write(str(root)+"\t"+str(file)+"\t"+str(custodian[1])+"\t"+str(custodian[2])+"\tN\\A\tN\\A\t"+str(os.path.getsize(os.path.join(root,file)))+"\tBlacklisted\n")
								logfile.close()
								continue

							with open(file) as file_to_check:
								# read contents of the file
								data = file_to_check.read(256)
								# pipe contents of the file through
								md5_returned = hashlib.md5(data).hexdigest()
								print "md5: "+md5_returned

							if md5_returned in hashBlacklist:
								logfile = open(os.path.join(logpath,"log.txt"),"a")
								logfile.write(str(root)+"\t"+str(file)+"\t"+str(custodian[1])+"\t"+str(custodian[2])+"\tN\\A\tN\\A\t"+str(os.path.getsize(os.path.join(root,file)))+"\tBlacklisted\n")
								logfile.close()
								continue
							else:
								#Test moving id file to cwd.
								#Custodian is (username, IDFile, Password)
								print "Driver={Lotus Notes SQL Driver (*.nsf)};UID="+custodian[0]+"/"+str(count)+";PWD="+custodian[2]+"; DATABASE="+os.path.join(root,file)+""
								connection=pyodbc.connect("Driver={Lotus Notes SQL Driver (*.nsf)};UID="+custodian[0]+"/"+str(count)+";PWD="+custodian[2]+"; DATABASE="+os.path.join(root,file)+"", autocommit=True)

								if connection:
									GOOD.append((root, file, custodian[0], custodian[1], custodian[2]))

						except MemoryError:
							print "MemoryError"
							continue
						except Exception as inst:
							print Exception
							print inst.args

							a, b = inst
							if re.search('Wrong Password',b):
								logfile = open(os.path.join(logpath,"log.txt"),"a")
								logfile.write(str(root)+"\t"+str(file)+"\t"+str(custodian[1])+"\t"+str(custodian[2])+"\tN\\A\tN\\A\t"+str(os.path.getsize(os.path.join(root,file)))+"\tERROR: bad password/ID Combination\n")
								logfile.close()
							elif re.search('08001',b):
								logfile = open(os.path.join(logpath,"log.txt"),"a")
								logfile.write(str(root)+"\t"+str(file)+"\t"+str(custodian[1])+"\t"+str(custodian[2])+"\tN\\A\tN\\A\t"+str(os.path.getsize(os.path.join(root,file)))+"\tERROR: unencrypted or could require additional certs\n")
								logfile.close()
								GOOD.append((root, file, custodian[0], custodian[1], custodian[2]))
							elif re.search('S1000',b):
								logfile = open(os.path.join(logpath,"log.txt"),"a")
								logfile.write(str(root)+"\t"+str(file)+"\t"+str(custodian[1])+"\t"+str(custodian[2])+"\tN\\A\tN\\A\t"+str(os.path.getsize(os.path.join(root,file)))+"\tERROR: possible corrupt file, please check\n")
								logfile.close()
							else:
								logfile = open(os.path.join(logpath,"log.txt"),"a")
								logfile.write(str(root)+"\t"+str(file)+"\t"+str(custodian[1])+"\t"+str(custodian[2])+"\tN\\A\tN\\A\t"+str(os.path.getsize(os.path.join(root,file)))+"\tERROR: "+b+"\n")
								logfile.close()

	for task in GOOD:
		#task is (root, file, username, IDFile, Password)
		try:
			reg = session.createRegistration()
			reg.switchToID(os.path.join(IDPATH,task[3]),task[4])

		except Exception as inst:
				print inst
				#x, y ,u , i = inst.args
				#logfile = open(os.path.join(logpath,"log.txt"),"a")
				#logfile.write(str(os.path.join(task[0], task[1]))+"\t"+str(task[1])+"\t"+str(task[3])+"\t"+str(custodian[4])+"\t"+str(x)+str(y)+str(u)+str(i)+"\n")
				#logfile.close()
				#if type(inst) == AttributeError:
					#x, y ,u , i = inst.arg
					#print inst

		try:
			database = session.GetDatabase("", os.path.join(task[0], task[1]))
			docs = database.AllDocuments
			logfile = open(os.path.join(logpath,"log.txt"),"a")
			logfile.write(str(os.path.join(task[0], task[1]))+"\t"+str(task[1])+"\t"+str(task[3])+"\t"+str(task[4])+"\t"+str(docs.Count)+"\t\N\\A\t"+str(os.path.getsize(os.path.join(task[0],task[1])))+"\tVerified\n")
			logfile.close()

			if decrypt == 1:
				NSFDecrypt(database, task, NSFPATH, logfile, DELETE, logpath)

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
						print custodian
						logfile.write(str(os.path.join(task[0], task[1]))+"\t"+str(task[1])+"\t"+str(task[3])+"\t"+str(task[4])+"\t"+inst.args+"\n")
						logfile.close()

		except Exception as inst:
			print inst
			#come back
			#x, y ,u , i = inst.args
			#logfile = open(os.path.join(logpath,"log.txt"),"a")
			#logfile.write(str(os.path.join(task[0], task[1]))+"\t"+str(task[1])+"\t"+str(task[3])+"\t"+str(task[4])+"\t\t"+str(x)+str(y)+str(u)+str(i)+"\n")
			#logfile.close()
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
			logfile.write("PP\t"+str(custodian[0])+"\t"+str(custodian[2])+"\t"+str(custodian[1])+"\t"+inst+"\n")
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
					logfile.write(str(root)+"\t"+str(file)+"\t"+str(custodian[1])+"\t"+str(custodian[2])+"\t"+inst+"\n")
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
					logfile = open(os.path.join(logpath,"log.txt"),"a")
					logfile.write(str(root)+"\t"+str(file)+"\t"+str(custodian[1])+"\t"+str(custodian[2])+"\t"+str(os.path.getsize(os.path.join(task[0],task[1])))+"\tERROR: "+inst+"\n")
					logfile.close()
def cli_progress_test(end_val, bar_length=20):
	for i in xrange(0, end_val):
		percent = float(i) / end_val
		hashes = '#' * int(round(percent * bar_length))
		spaces = ' ' * (bar_length - len(hashes))
		sys.stdout.write("\rPercent: [{0}] {1}%".format(hashes + spaces, int(round(percent * 100))))
		sys.stdout.flush()

def main(argv, GOOD, BAD, DummyFile, inifile):
	decrypt = 0
	bruteForce = 0
	DELETE = 0
	hashBlacklist = []
	filenameBlacklist = []
	fileblacklistIn=[line.strip() for line in open(os.path.join(workingDir, r'fblacklist.txt'))]
	hashblacklistIn=[line.strip() for line in open(os.path.join(workingDir, r'hblacklist.txt'))]
	
	#load filename blacklist
	for line in fileblacklistIn:
		filenameBlacklist.append(line)
	#load hash blacklist
	for line in hashblacklistIn:
		hashBlacklist.append(line)

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
			Validate(NSFPATH, IDPATH, LotusDataPATH, LOADFILE, decrypt, bruteForce, DELETE, inifile, GOOD, BAD, NotesSQLCFG, logpath, workingDir, DummyFile, filenameBlacklist, hashBlacklist)
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
		Validate(NSFPATH, IDPATH, LotusDataPATH, LOADFILE, decrypt, bruteForce, DELETE, inifile, GOOD, BAD, NotesSQLCFG, logpath, workingDir, DummyFile, filenameBlacklist, hashBlacklist)

if __name__ == "__main__":
	start_time = time.time()
	main(sys.argv[1:], GOOD, BAD, DummyFile, inifile)
	print time.time() - start_time, "seconds"
