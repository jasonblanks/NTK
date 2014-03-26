'''	NSF Validation Tool v .2
	This is the first attempt at using python to interact with com
	and run various Lotus com functions and too validate deliveries.

	This is tested using IBM Notes 9 standalone client
	This script temparly moves the user's default user.id file to user.id_temp
	and runs through the load.csv file copying the id and nsf files as needed
	from the base directory.

	Requirements:
	IBM Notes 9 standalone client
	Python 2.7
	win32com (http://sourceforge.net/projects/pywin32/files/pywin32/Build%20218/pywin32-218.win32-py2.7.exe/download))


	MAJOR TODO: Better error reporting and handling.


'''
import sys, os, win32com.client, shutil, getopt, subprocess, fileinput
#import pywintypes
import pythoncom
import multiprocessing
import time

#Environment Variables
##TODO: Fix so it accepts spaces in file path.
NSFPATH = r'C:\\Users\user_profile\Desktop\test'
IDPATH = r'C:\\Users\user_profile\Desktop\test'
##TODO:Should be able to look this path automatically.
LotusDataPATH = r'C:\\Users\user_profile\AppData\Local\IBM\Notes\Data'
inifile = r'C:\\Users\user_profile\AppData\Local\IBM\Notes\Data\notes.ini'


##TODO: Add command line arg to specify this file path.
LOADFILE = r'C:\\Users\user_profile\Desktop\test\load.txt'
DummyFile = r'C:\\Users\user_profile\Desktop\test\dummy.id'
IDDefault = "user.id"
IDTemp = os.path.join(LotusDataPATH,"user.id_temp")
bad = []
global BAD
global GOOD
BAD = []
GOOD = []
timeout = 10

def test(NSFPATH, IDPATH, LotusDataPATH, LOADFILE, IDDefault, IDTemp, decrypt, bruteForce, DELETE, inifile, task, count, GOOD, BAD):
	password = None
	for line in fileinput.FileInput(inifile, inplace=1):
		if line.startswith(r'KeyFilename='):
			defaultID = line
			line = line.replace(line, r'KeyFilename='+DummyFile)
		if line.startswith("KeyFileName_Owner"):
			defaultOwner = line
			line = line.replace(line, r'')
		if line.startswith("Location="):
			defaultOwner = line
			line = line.replace(line, "")
		print line.strip()
	try:
		session = win32com.client.Dispatch("Lotus.NotesSession")
		session.Initialize()
	except Exception as inst:
			print type(inst)
			print inst.args
			if type(inst) == AttributeError:
				print inst.args
				#x, y ,u , i = insta.args
				#logfile = open("C:\\Users\\user_profile\\Desktop\\test\\log.txt","a")
				#logfile.write(task[0]+"\t"+task[1]+"\t"+"\tNA\t"+u[2]+"\n")
				#logfile.close()
	if task[2]:
		try:
			reg = session.createRegistration()
			reg.switchToID(os.path.join(IDPATH,task[1]),task[2])
		except Exception as inst:
				print type(inst)
				print inst.args
				if type(inst) == AttributeError:
					#x, y ,u , i = inst.arg
					print inst.arg
					#logfile = open("C:\\Users\\user_profile\\Desktop\\test\\log.txt","a")
					#logfile.write(task[0]+"\t"+task[1]+"\t"+"\tNA\t"+u[2]+"\n")
					#logfile.close()
					#task_tuple = (task[0],task[1],task[2])
					#BAD.append(task)

	elif not task[2]:
		password = None
		try:
			session.Initialize()
		except Exception as inst:
				if type(inst) == TypeError:
					x, y ,u , i = inst.args
					#logfile = open("C:\\Users\\user_profile\\Desktop\\test\\log.txt","a")
					#logfile.write(task[0]+"\t"+task[1]+"\t"+"\tNA\t"+u[2]+"\n")
					#logfile.close()

def NSFDecrypt(db, task, NSFPATH, logfile, DELETE):
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
			logfile = open("C:\\Users\\user_profile\\Desktop\\test\\log.txt","a")
			logfile.write(str(task[0])+"\t"+str(task[1])+"\t"+str(task[2])+"\t"+str(CloneDocCount.Count)+"\tDecrypted\n")
			logfile.close()

		elif  CloneDocCount.Count != OriginalDocCount.Count:
			logfile = open("C:\\Users\\user_profile\\Desktop\\test\\log.txt","a")
			logfile.write(str(task[0])+"\t"+str(task[1])+"\t"+str(task[2])+"\t"+str(CloneDocCount.Count)+"\tDecryption Failed\n")
			logfile.close()
def BruteForce(bad, TASKS, logfile, IDPATH, LotusDataPATH, NSFPATH, DELETE):
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

def Validate(NSFPATH, IDPATH, LotusDataPATH, LOADFILE, IDDefault, IDTemp, decrypt, bruteForce, DELETE, inifile, GOOD, BAD, timeout):
	#TODO: Add os check to see if LOADFILE exists. If not throw error and exit.
	TASKS = [line.strip().split(',') for line in open(LOADFILE)]
	logfile = open("C:\\Users\\user_profile\\Desktop\\test\\log.txt","w")
	logfile.write("NSF File\tID File\tPASSWORD\tMSG Count\tSTATUS\n")
	logfile.close()

	count = 0

	for task in TASKS:
		for line in fileinput.FileInput(inifile, inplace=1):
			if line.startswith(r'KeyFilename='):
				defaultID = line
				line = line.replace(line, r'KeyFilename='+DummyFile)
			if line.startswith("KeyFileName_Owner"):
				defaultOwner = line
				line = line.replace(line, r'')
			if line.startswith("Location="):
				defaultOwner = line
				line = line.replace(line, "")
			print line.strip()

		try:
			session = win32com.client.Dispatch("Lotus.NotesSession")
			session.Initialize()
		except Exception as inst:
				print type(inst)
				print inst.args
				if type(inst) == AttributeError:
					print inst.args
					#x, y ,u , i =
					#logfile = open("C:\\Users\\user_profile\\Desktop\\test\\log.txt","a")
					#logfile.write(task[0]+"\t"+task[1]+"\t"+"\tNA\t"+u[2]+"\n")
					#logfile.close()




		#os.system('taskkill /f /im SUService.exe')
		#os.system('taskkill /f /im ntmulti.exe')
		#os.system('taskkill /f /im nsd.exe')
		p = multiprocessing.Process(target=test, name="test", args=(NSFPATH, IDPATH, LotusDataPATH, LOADFILE, IDDefault, IDTemp, decrypt, bruteForce, DELETE, inifile, task, count, GOOD, BAD))
		count = count + 1
		p.start()
		p.join(timeout)
		if p.is_alive():
			p.terminate()
			BAD.append(task)
			logfile = open("C:\\Users\\user_profile\\Desktop\\test\\log.txt","a")
			logfile.write(task[0]+"\t"+task[1]+"\t"+"\tNA\tTIMEOUT\n")
			logfile.close()
			p.join()
		else:
			#print "good "+str(task)
			GOOD.append(task)

	print GOOD
	for line in fileinput.FileInput(inifile, inplace=1):
		if line.startswith(r'KeyFilename='):
			defaultID = line
			line = line.replace(line, r'KeyFilename='+DummyFile)
		if line.startswith("KeyFileName_Owner"):
			defaultOwner = line
			line = line.replace(line, r'')
		if line.startswith("Location="):
			defaultOwner = line
			line = line.replace(line, "")
		print line.strip()

	try:
		session = win32com.client.Dispatch("Lotus.NotesSession")
		session.Initialize()
	except Exception as inst:
			print type(inst)
			print inst.args
			if type(inst) == AttributeError:
				print inst.args
				#x, y ,u , i = insta.args
				#logfile = open("C:\\Users\\user_profile\\Desktop\\test\\log.txt","a")
				#logfile.write(task[0]+"\t"+task[1]+"\t"+"\tNA\t"+u[2]+"\n")
				#logfile.close()

	for task in GOOD:
		try:
			reg = session.createRegistration()
			reg.switchToID(os.path.join(IDPATH,task[1]),task[2])

		except Exception as inst:
				print type(inst)
				print inst.args
				if type(inst) == AttributeError:
					#x, y ,u , i = inst.arg
					print inst.arg
					#logfile = open("C:\\Users\\user_profile\\Desktop\\test\\log.txt","a")
					#logfile.write(task[0]+"\t"+task[1]+"\t"+"\tNA\t"+u[2]+"\n")
					#logfile.close()
					#task_tuple = (task[0],task[1],task[2])
					#BAD.append(task)

		try:
			database = session.GetDatabase("", os.path.join(NSFPATH, task[0]))
			docs = database.AllDocuments
			logfile = open("C:\\Users\\user_profile\\Desktop\\test\\log.txt","a")
			logfile.write(str(task[0])+"\t"+str(task[1])+"\t"+str(task[2])+"\t"+str(docs.Count)+"\tGOOD\n")
			logfile.close()
			if decrypt == 1:
				NSFDecrypt(database, task, NSFPATH, logfile, DELETE)
			print str(task[0])+"\t"+str(task[1])+"\t"+str(task[2])+"\t"+str(docs.Count)+"\tGOOD\n"
		except Exception as inst:
				if type(inst) == TypeError:
					x, y ,u , i = inst.args
					print u

	#if decrypt == 1:
		#NSFDecrypt(database, task, NSFPATH, logfile, inifile,  DELETE, GOOD)
	if bruteForce == 1:
			BruteForce(bad, TASKS, logfile, IDPATH, LotusDataPATH, NSFPATH, DELETE, BAD)


def main(argv, GOOD, BAD):
	decrypt = 0
	bruteForce = 0
	DELETE = 0
	try:
		opts, args = getopt.getopt(argv,"hdc",)
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
			Validate(NSFPATH, IDPATH, LotusDataPATH, LOADFILE, IDDefault, IDTemp, decrypt, bruteForce, DELETE, inifile, GOOD, BAD, timeout)
		elif opt ==  "-c":
			bruteForce = True
			Validate(NSFPATH, IDPATH, LotusDataPATH, LOADFILE, IDDefault, IDTemp, decrypt, bruteForce, DELETE)
		else:
			Validate(NSFPATH, IDPATH, LotusDataPATH, LOADFILE, IDDefault, IDTemp, decrypt, bruteForce, DELETE, BAD, GOOD)
	if not opts:
		Validate(NSFPATH, IDPATH, LotusDataPATH, LOADFILE, IDDefault, IDTemp, decrypt, bruteForce, DELETE, inifile, GOOD, BAD, timeout)



if __name__ == "__main__":
   main(sys.argv[1:], GOOD, BAD)
   session = None
