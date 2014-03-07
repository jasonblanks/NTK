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
import sys, os, win32com.client, shutil, getopt


#change to meet your needs
NSFPATH = "C:\\Users\\blanks\\Desktop\\TestNSF"
IDPATH = "C:\\Users\\blanks\\Desktop\\TestNSF\\IDFiles"
LotusDataPATH = "C:\\Users\\blanks\\AppData\\Local\\IBM\\Notes\\Data"

#shouldn't change
LOADFILE = "load.csv"
IDDefault = "user.id"
IDTemp = os.path.join(LotusDataPATH,"user.id_temp")
bad = []

def TempFileCheck():
	if os.path.isfile(IDTemp):
		os.rename(IDTemp,os.path.join(LotusDataPATH,IDDefault))
	elif os.path.isfile(os.path.join(LotusDataPATH,IDDefault)):
		os.rename(os.path.join(LotusDataPATH,IDDefault),IDTemp)


def NSFDecrypt(db, task, NSFPATH, logfile):
	dbclone = db.CreateFromTemplate("",os.path.join(NSFPATH, task[0])+"--decrypt", False)
	dbclone.Compact
	
	OriginalDocCount = db.AllDocuments
	CloneDocCount = dbclone.AllDocuments
	
	if  CloneDocCount.Count == OriginalDocCount.Count:
		logfile.write(str(task[0])+"\t"+str(task[1])+"\t"+str(task[2])+"\t"+str(CloneDocCount.Count)+"\tGOOD\tYES\n")
	
	elif  CloneDocCount.Count != OriginalDocCount.Count:
		logfile.write(str(task[0])+"\t"+str(task[1])+"\t"+str(task[2])+"\t"+str(CloneDocCount.Count)+"\tGOOD, but decypt did not match\tNO\n")

def BruteForce(bad, TASKS, logfile, IDPATH, LotusDataPATH, NSFPATH):
	for BadNSF in bad:
		for ID in TASKS:
			for PASSWORD in TASKS:
				shutil.copyfile(os.path.join(IDPATH,ID[1]),os.path.join(LotusDataPATH,"user.id"))
				session = win32com.client.Dispatch("Lotus.NotesSession")
				
				try:
					session.Initialize(PASSWORD[2])
					database = session.GetDatabase("", os.path.join(NSFPATH, BadNSF[0]))
					docs = database.AllDocuments
				except:
					logfile.write(BadNSF[0]+"\t"+ID[1]+"\t"+PASSWORD[2]+"\tNA\tERROR: BF Attempt Bad\n")
					os.remove(os.path.join(LotusDataPATH,"user.id"))
					continue
				logfile.write(str(BadNSF[0])+"\t"+str(ID[1])+"\t"+str(PASSWORD[2])+"\t"+str(docs.Count)+"\tGOOD\tBRUTEFORCE\n")
				#NSFDecrypt(database, task, NSFPATH, logfile)
				os.remove(os.path.join(LotusDataPATH,"user.id"))
				
def Validate(NSFPATH, IDPATH, LotusDataPATH, LOADFILE, IDDefault, IDTemp, bad, decrypt, bruteForce):
	TASKS = [line.strip().split(',') for line in open(LOADFILE)]
	TempFileCheck()
	logfile = open('log.txt','w')
	logfile.write("NSF File\tID File\tPASSWORD\tMSG Count\tSTATUS\n")

	for task in TASKS:
		shutil.copyfile(os.path.join(IDPATH,task[1]),os.path.join(LotusDataPATH,"user.id"))
		session = win32com.client.Dispatch("Lotus.NotesSession")
		
		try:
			session.Initialize(task[2])
			database = session.GetDatabase("", os.path.join(NSFPATH, task[0]))
			docs = database.AllDocuments
		except:
			logfile.write(task[0]+"\t"+task[1]+"\t"+task[2]+"\tNA\tERROR: bad password or password id file mismatch\n")
			os.remove(os.path.join(LotusDataPATH,"user.id"))
			bad.append(task)
			continue
				
		logfile.write(str(task[0])+"\t"+str(task[1])+"\t"+str(task[2])+"\t"+str(docs.Count)+"\tGOOD\n")
		if decrypt == 1:
			NSFDecrypt(database, task, NSFPATH, logfile)
		os.remove(os.path.join(LotusDataPATH,"user.id"))
	if bruteForce == 1:
		BruteForce(bad, TASKS, logfile, IDPATH, LotusDataPATH, NSFPATH)

	TempFileCheck()

def main(argv):
	decrypt = 0
	bruteForce = 0
	try:
		opts, args = getopt.getopt(argv,"hdc",)
		print opts
		print args
	except getopt.GetoptError:
		print 'test.py -h -c -d '
		sys.exit(2)

	for opt, arg in opts:
		#print "here"
		if opt == '-h':
			print 'test.py -d(decrypt valid files) -c(brute force with all given id and password combinations.)'
			sys.exit()
		#elif opt == "":
			#Validate(NSFPATH, IDPATH, LotusDataPATH, LOADFILE, IDDefault, IDTemp, bad, decrypt, bruteForce)
		elif opt == "-d":
			decrypt = True
			Validate(NSFPATH, IDPATH, LotusDataPATH, LOADFILE, IDDefault, IDTemp, bad, decrypt, bruteForce)
		elif opt ==  "-c":
			bruteForce = True
			Validate(NSFPATH, IDPATH, LotusDataPATH, LOADFILE, IDDefault, IDTemp, bad, decrypt, bruteForce)
		else:
			Validate(NSFPATH, IDPATH, LotusDataPATH, LOADFILE, IDDefault, IDTemp, bad, decrypt, bruteForce)
	if not opts:
		Validate(NSFPATH, IDPATH, LotusDataPATH, LOADFILE, IDDefault, IDTemp, bad, decrypt, bruteForce)
		print "here"




if __name__ == "__main__":
   main(sys.argv[1:])
