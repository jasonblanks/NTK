'''	NSF Validation Tool v 1-alpha
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


	TODO's:
				* improve/test prerequisites checks
				* delay reporting on "possible corrupt file, please check" errors until after decryption has been attempted to cut down on over logging.
				* add new arg group "output" with output options (clone, clone directory structure), (inplace, output decrypted files in the same directory as the original), (delete, delete, deletes original files when completed
				* research unsupported NSF files and individually encrypted notes
				* display realtime processing status and statistics in the active window in a clean manner.
				* solidify command arg switches with group
				* discuss about cached settings/default config options
				* pull default paths/settings from Lotus registry settings
				* create a more verbose logging option
				* create more blacklist options ie. location switches, md5 and filename switches.
				**Propose distributed decrypting framework to speed up large deliveries.


'''
import sys, os, win32com.client, fileinput, pyodbc, re, hashlib, shutil, argparse, time
from _winreg import *

BAD = []
GOOD = []
REMOVE = []
MD5HashList = []
DummyFile = r'C:\Users\user\Desktop\test\dummy.id'
inifile = r'C:\Users\user\AppData\Local\IBM\Notes\Data\notes.ini'
#workingDir = r'Y:\NSF'
workingDir = r'C:\temp\test'

#not implemented yet...
def prerequisites():
	'''
	os
	if not os.path.exists(r'C:\NotesSQL'):
		print "it appears you do not have NoteSQL installed or installed to the Default location, would you like to proceed? [Y/N]: "
		choice = sys.input()
	if not os.path.exists(r'C:\Program Files (x86)\IBM\Notes') or os.path.exists(r'C:\Program Files\IBM\Notes'):
		print "it appears you do not have IBM Notes client installed or installed to the Default location, would you like to proceed? [Y/N]: "
		choice = sys.inpit()
	for file in os.listdir(r"C:\"+ PATH +r"\IBM\Notes\properties\version")
		if file.split('.').endswith(r'access-9.0.1.swtag')
			v9 = True
	if not v9:
		exit("You do not appear to be using the correct version of the notes client, please use 9.0.1")
	'''


	NotesClientExists = True
	NotesSQLExists = True
	NotesClientReg = ConnectRegistry(None,HKEY_CURRENT_USER)
	NotesSQLReg = ConnectRegistry(None,HKLM_LOCAL_MACHINE)
	try:
		#aKey = OpenKey(aReg, r"SOFTWARE\Clients\Mail\Lotus Notes", 0, KEY_WRITE)
		NotesClientKey = OpenKey(aReg, r"SOFTWARE\Clients\Mail\Lotus Notes")
	except WindowsError:
		NotesClientExists = False
	CloseKey(NotesClientKey)
	CloseKey(NotesClientReg)

	try:
		#aKey = OpenKey(aReg, r"SOFTWARE\Clients\Mail\Lotus Notes", 0, KEY_WRITE)
		NotesSQLKey = OpenKey(aReg, r"SOFTWARE\Wow6432Node\Lotus\Lotus Notes SQL Driver")
	except WindowsError:
		NotesSQLExists = False
	CloseKey(NotesSQLKey)
	CloseKey(NotesSQLReg)


	####
	try:
		if not NotesClientExists:
			print "ERROR: IBM Notes client must be installed."
			exit()
			#SetValueEx(aKey,registry_key_name,0, REG_SZ, r"" + folder + "\" + file_name)
	except EnvironmentError:
		print "Encountered problems writing into the Registry..."

	try:
		if not NotesSQLExists:
			print "ERROR: NoteSQL must be installed."
			exit()
			#SetValueEx(aKey,registry_key_name,0, REG_SZ, r"" + folder + "\" + file_name)
	except EnvironmentError:
		print "Encountered problems writing into the Registry..."





def NSFDecrypt(db, task, NSFPATH, logfile, DELETE, logpath):
		#TASKS is (root, file, username, IDFile, Password)
		cloneFilename = task[1].split('.')
		dbclone = db.CreateFromTemplate("", os.path.join(task[0], cloneFilename[0])+"--decrypt.nsf", False)


		
		# We determianed you not only have to remove encryption by compact but clear out the ACL on that NSF
		# Below is a quick hack to allow everyone access, I would like to make this cleaner my creating a function
		# to clear out the whole ACL.

		dbclone.GrantAccess( "-Default-", "6" )
		dbclone.GrantAccess( "Anonymous", "6" )
		dbclone.CompactWithOptions("L")
		if DELETE:
			os.remove(os.path.join(task[0], task[1]))

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

		try:
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

		except Exception as inst:
			print Exception
			print inst.args

			a, b = inst
			if re.search('encrypted',b):
				logfile = open(os.path.join(logpath,"log.txt"),"a")
				logfile.write(os.path.join(str(root),str(file))+"\t"+str(file)+"\t"+str(custodian[1])+"\t"+str(custodian[2])+"\tN\\A\tN\\A\t"+str(os.path.getsize(os.path.join(root,file)))+"\tERROR: Unsupported Encryption\n")
				logfile.close()
				REMOVE.append(os.path.join(task[0], cloneFilename[0])+"--decrypt.nsf")



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


def Validate(NSFPATH, IDPATH, LotusDataPATH, LOADFILE, decrypt, bruteForce, DELETE, inifile, GOOD, BAD, NotesSQLCFG, logpath, workingDir, DummyFile, filenameBlacklist, hashBlacklist, DedupeOption):
	#TODO: Add os check to see if LOADFILE exists. If not throw error and exit.
	TASKS = [line.strip().split(',') for line in open(LOADFILE)]
	logfile = open(os.path.join(logpath,"log.txt"),"w")
	logfile.write("NSF File Path\tNSF File\tID File\tPASSWORD\tOriginal Count\tDecrypted Count\tFile Size (Bytes)\tStatus\tHeader Text\n")
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

	#this will build the cfg file references for the NotesSQL two factor auth.
	def buildfile(TASKS, NotesSQLCFG):
		CFGFile = open(NotesSQLCFG,"w")
		CFGFile.write("[Users]\n")
		CFGFile.write("dummy="+os.path.join(IDPATH, DummyFile)+"\n")
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
	print "begining validation stage"
	for root, dirs, files in os.walk(workingDir, topdown=False):
		for custodian in TASKS:
			count = 0
			for file in files:
				if custodian[0] in root.split('\\'):
					if file.endswith('nsf'):
						try:
							os.chdir(root)
							#Unsupported file test
							f = open(file, "rb")
							i = 0
							unsupported = 0
							for line in f:
								if re.search('#2Notes90V1.3', line):
									print "unsupported file encrption: "+str(os.path.join(root, file))
									logfile = open(os.path.join(logpath,"log.txt"),"a")
									logfile.write(os.path.join(str(root),str(file))+"\t"+str(file)+"\t"+str(custodian[1])+"\t"+str(custodian[2])+"\tN\\A\tN\\A\t"+str(os.path.getsize(os.path.join(root,file)))+"\tERROR:  unsupported encryption\\file type\n")
									logfile.close()
									unsupported = 1
								i = i + 1
								if i == 3:
									break
							if unsupported == 1:
								continue
							#end of unsupported file test

							if file in filenameBlacklist:
								logfile = open(os.path.join(logpath,"log.txt"),"a")
								logfile.write(os.path.join(str(root),str(file))+"\t"+str(file)+"\t"+str(custodian[1])+"\t"+str(custodian[2])+"\tN\\A\tN\\A\t"+str(os.path.getsize(os.path.join(root,file)))+"\tBlacklisted\n")
								logfile.close()
								continue
							try:
								file_to_check = open(file, 'rb')
								# read contents of the file
								data = file_to_check.read(256)
								# pipe contents of the file through
								md5_returned = hashlib.md5(data).hexdigest()
								#print "md5: "+md5_returned
								file_to_check.close()

								#MD5 dedupe check
								if DedupeOption:
									if md5_returned in MD5HashList:
										logfile = open(os.path.join(logpath,"log.txt"),"a")
										logfile.write(os.path.join(str(root),str(file))+"\t"+str(file)+"\t"+str(custodian[1])+"\t"+str(custodian[2])+"\tN\\A\tN\\A\t"+str(os.path.getsize(os.path.join(root,file)))+"\tMD5 Deduped: "+str(md5_returned)+"\n")
										logfile.close()
										continue
									else:
										MD5HashList.append(md5_returned)

							except:
								print "unable to create md5 for: "+str(file)

							if md5_returned in hashBlacklist:
								logfile = open(os.path.join(logpath,"log.txt"),"a")
								logfile.write(os.path.join(str(root),str(file))+"\t"+str(file)+"\t"+str(custodian[1])+"\t"+str(custodian[2])+"\tN\\A\tN\\A\t"+str(os.path.getsize(os.path.join(root,file)))+"\tBlacklisted\n")
								logfile.close()
								continue
							else:
								#Custodian is (username, IDFile, Password)
								#print "Driver={Lotus Notes SQL Driver (*.nsf)};UID="+custodian[0]+"/"+str(count)+";PWD="+custodian[2]+"; DATABASE="+os.path.join(root,file)+""

								connection=pyodbc.connect("Driver={Lotus Notes SQL Driver (*.nsf)};UID="+custodian[0]+"/"+str(count)+";PWD="+custodian[2]+"; DATABASE="+os.path.join(root,file)+"", autocommit=True)
								if connection:
									GOOD.append((root, file, custodian[0], custodian[1], custodian[2]))

						except MemoryError:
							#before setting md5 chunksizes this was an issue, should not be anymore.
							#print "MemoryError"
							continue
						except Exception as inst:
							#print inst
							a, b = inst
							if re.search('Wrong Password',b):
								logfile = open(os.path.join(logpath,"log.txt"),"a")
								logfile.write(os.path.join(str(root),str(file))+"\t"+str(file)+"\t"+str(custodian[1])+"\t"+str(custodian[2])+"\tN\\A\tN\\A\t"+str(os.path.getsize(os.path.join(root,file)))+"\tERROR: bad password/ID Combination\n")
								logfile.close()
							elif re.search('08001',b):
								logfile = open(os.path.join(logpath,"log.txt"),"a")
								logfile.write(os.path.join(str(root),str(file))+"\t"+str(file)+"\t"+str(custodian[1])+"\t"+str(custodian[2])+"\tN\\A\tN\\A\t"+str(os.path.getsize(os.path.join(root,file)))+"\tERROR: unencrypted or could require additional certs\n")
								logfile.close()
								GOOD.append((root, file, custodian[0], custodian[1], custodian[2]))
							elif re.search('S1000',b):
								logfile = open(os.path.join(logpath,"log.txt"),"a")
								logfile.write(os.path.join(str(root),str(file))+"\t"+str(file)+"\t"+str(custodian[1])+"\t"+str(custodian[2])+"\tN\\A\tN\\A\t"+str(os.path.getsize(os.path.join(root,file)))+"\tERROR: possible corrupt file, please check\n")
								logfile.close()
							elif re.search('SECKFMDefaultPromptHandler',b):
								logfile = open(os.path.join(logpath,"log.txt"),"a")
								logfile.write(os.path.join(str(root),str(file))+"\t"+str(file)+"\t"+str(custodian[1])+"\t"+str(custodian[2])+"\tN\\A\tN\\A\t"+str(os.path.getsize(os.path.join(root,file)))+"\tERROR: Unsupported encryption\n")
								logfile.close()
							else:
								logfile = open(os.path.join(logpath,"log.txt"),"a")
								logfile.write(os.path.join(str(root),str(file))+"\t"+str(file)+"\t"+str(custodian[1])+"\t"+str(custodian[2])+"\tN\\A\tN\\A\t"+str(os.path.getsize(os.path.join(root,file)))+"\tERROR: "+b+"\n")
								logfile.close()
	NSFCount = len(GOOD)
	completedCount = NSFCount
	for task in GOOD:
		#Active window stats
		os.system('cls')
		print
		print "Current file("+str(completedCount)+"/"+str(NSFCount)+"): "+str(os.path.join(task[0], task[1]))
		completedCount = completedCount - 1
		#task is (root, file, username, IDFile, Password)
		try:
			reg = session.createRegistration()
			reg.switchToID(os.path.join(IDPATH,task[3]),task[4])

		except Exception as inst:
				print inst
				#x, y ,u , i = inst.args
				#logfile = open(os.path.join(logpath,"log.txt"),"a")
				#logfile.write(str(os.path.join(task[0], task[1]))+"\t"+str(task[1])+"\t"+str(task[3])+"\t"+str(custodian[4])+"\tNot Decrypted line 214\n")
				#logfile.close()
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
						print "242"
						if type(inst) == AttributeError:
							print inst.args
						continue
				continue

		try:
			database = session.GetDatabase("", os.path.join(task[0], task[1]))
			docs = database.AllDocuments

			if decrypt == 0:
				logfile = open(os.path.join(logpath,"log.txt"),"a")
				logfile.write(str(os.path.join(task[0], task[1]))+"\t"+str(task[1])+"\t"+str(task[3])+"\t"+str(task[4])+"\t"+str(docs.Count)+"\t\N\\A\t"+str(os.path.getsize(os.path.join(task[0],task[1])))+"\tVerified\n")
				logfile.close()

			if decrypt == 1:
				NSFDecrypt(database, task, NSFPATH, logfile, DELETE, logpath)



		except Exception as inst:
			print inst
			logfile = open(os.path.join(logpath,"log.txt"),"a")
			logfile.write(str(os.path.join(task[0], task[1]))+"\t"+str(task[1])+"\t"+str(task[3])+"\t"+str(task[4])+"\tUnknown Error\n")
			logfile.close()

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

	for root, dirs, files in os.walk(workingDir, topdown=False):
		for file in files:
			try:
				session = win32com.client.Dispatch("Lotus.NotesSession")
				session.Initialize()
			except Exception as inst:
					logfile = open(os.path.join(logpath,"log.txt"),"a")
					logfile.write(str(root)+"\t"+str(file)+"\t"+str(custodian[1])+"\t"+str(custodian[2])+"\t"+inst+"\n")
					logfile.close()

			if file.endswith('nsf'):
				try:
					os.chdir(root)
					connection=pyodbc.connect("Driver={Lotus Notes SQL Driver (*.nsf)};UID=dummy;DATABASE="+os.path.join(root,file)+"", autocommit=True)

				except Exception as inst:
					logfile = open(os.path.join(logpath,"log.txt"),"a")
					logfile.write(str(root)+"\t"+str(file)+"\t"+str(custodian[1])+"\t"+str(custodian[2])+"\t"+str(os.path.getsize(os.path.join(task[0],task[1])))+"\tERROR: "+inst+"\n")
					logfile.close()

#TODO: The below function is work in progress in creating active status display while decrypting etc.
def cli_progress_test(end_val, bar_length=20):
	for i in xrange(0, end_val):
		percent = float(i) / end_val
		hashes = '#' * int(round(percent * bar_length))
		spaces = ' ' * (bar_length - len(hashes))
		sys.stdout.write("\rPercent: [{0}] {1}%".format(hashes + spaces, int(round(percent * 100))))
		sys.stdout.flush()

def main(GOOD):

	decrypt = None
	bruteForce = None
	DELETE = None
	DedupeOption = None
	hashBlacklist = []
	NSFPATH = r'Y:\NSF'
	IDPATH = r'Y:\NSF\IDs'
	##TODO:Should be able to look this path automatically.
	LotusDataPATH = r'C:\Users\user\AppData\Local\IBM\Notes\Data'
	inifile = r'C:\Users\user\AppData\Local\IBM\Notes\Data\notes.ini'
	logpath = r'C:\temp\test'
	#logpath = r'Y:\NSF'
	##TODO: Add command line arg to specify this file path.
	#
	LOADFILE = r'Y:\NSF\load.txt'
	#LOADFILE = r'Y:\NSF\load.txt'
	DummyFile = r'C:\Users\user\Desktop\test\dummy.id'
	NotesSQLCFG = r'C:\NotesSQL\notessql.cfg'
	#workingDir = r'Y:\NSF'
	workingDir = r'C:\temp\test'
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


	parser = argparse.ArgumentParser()
	group = parser.add_mutually_exclusive_group()
	group.add_argument("-v", "--validate", help="This method provide a quick NSF counts and ID file checks without decrypting.",action="store_true")
	group.add_argument("-d", "--decrypt", help="this option both validates NSF File counts, ID/PASSWORD pares aswell as decrypts",action="store_true")


	#parser.add_argument("-h", "--help",help="This help")
	parser.add_argument("-id",help="full path to the id repository directory  (ex: -id C:\\ID_FILES)")
	parser.add_argument("-md5",help="this option tracks all md5's and ignores files that have already been added matching the same md5 value.",action="store_true")
	parser.add_argument("-l", "--loadfile",help="define which dummy file to use, full path muct be used. (ex: -l C:\\loadfile.txt)")
	parser.add_argument("-r", "--remove",help="remove original NSF files after decryption")
	parser.add_argument("-f", "--dummyfile",help="allows you to define which dummy file to use, full path muct be used. (ex: -f C:\\dummyfile.id)")



	#TODO:  add new arg group "output" with output options (clone, clone directory struct), (inplace, output decrypted files in the same directory
	#TODO:	    as the original), (delete, delete, deletes original files when completed


	args = parser.parse_args()

	if args.remove:
		DELETE = True
	elif args.dummyfile:
		DummyFile = args.dummyfile
	elif args.id:
		IDPATH = args.id
	elif args.loadfile:
		LOADFILE = args.loadfile
	elif args.decrypt:
		decrypt = True
	elif args.md5:
		DedupeOption = True
	#elif args.config:
	#	pass


	Validate(NSFPATH, IDPATH, LotusDataPATH, LOADFILE, decrypt, bruteForce, DELETE, inifile, GOOD, BAD, NotesSQLCFG, logpath, workingDir, DummyFile, filenameBlacklist, hashBlacklist, DedupeOption)

if __name__ == "__main__":
	start_time = time.time()
	main(GOOD)
	print time.time() - start_time, "seconds"

	print "cleaning up...."
	removefile = open(workingDir+"remove.bat", 'w')
	for r in REMOVE:
		removefile.write("REMOVE "+r+"\n")
	removefile.close()
