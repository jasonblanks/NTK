NSFTool
=======

NSFTool


Just some helpful links:

http://www-12.lotus.com/ldd/doc/domino_notes/6.5.1/help65_designer.nsf/f4b82fbb75e942a6852566ac0037f284/319c1e8e551143c185256e00004a54bd?OpenDocument

http://www-12.lotus.com/ldd/doc/domino_notes/rnext/help6_designer.nsf/f4b82fbb75e942a6852566ac0037f284/c4b8cb573d68e43c85256c54004c93ab?OpenDocument

http://publib.boulder.ibm.com/infocenter/domhelp/v8r0/index.jsp?topic=%2Fcom.ibm.designer.domino.main.doc%2FH_GRANTACCESS_METHOD.html

http://www-03.ibm.com/software/products/en/ibmnotes
https://www14.software.ibm.com/webapp/iwm/web/reg/download.do?source=ESD-NTSDOMTRL&S_TACT=109HD0PW&S_CMP=web_dw_rt_swd&lang=en_US&S_PKG=CRP4VEN&cp=UTF-8&dlmethod=http

http://www-12.lotus.com/ldd/doc/uafiles.nsf/docs/designer70/$file/prog2.pdf

---------------------------


This tool was designed to either quickly validate client deliveries by checking NSF accessablity and NSF Counts, second method is to massivly decrypt these NSF's using provided ID and Key pairs for ediscovery uses.


Directory Structure:

This script assumes that you have a large delivery involving multiple custodians aswell as ID files possibly even multiple key/id pairs to test.


What I've called the 'working directory' should at a minimum hold one 'custodian' directory and one 'IDs' Directory, though this will usually consist of multiple custodian directorys.

#typical working directory format
	root
	  |-IDs\
	  |-GomezAddams\
	  |-WednesdayAddams\
	  |-PugsleyAddams\
	  |-MorticiaAddams\
	  |-load.txt
	  |-fblacklist.txt
	  |-blacklist.txt

##
##Load File Breakdown
##

The load file is in the format of:

'custodian folder,id file,Password'

basicly this this breaks down and say for every file in said folder run and or test the following ID/Password combinations.  This allows you to attempt or use multiple id files and or test ID/Password combinations, one per line, coma seperated.

#typical loadfile.txt format
#start
	GomezAddams,gaddams.id,sd72dns8
	GomezAddams,gaddams_backup.id,calk72hx
	WednesdayAddams,waddams.id,password1
	WednesdayAddams,waddams.id,password2
	PugsleyAddams,user.id,sadsjwoc62
	...

  
## fblacklist.txt
fblacklist.txt is a filename blacklist file you may propigate.  It takes one file name per line.

#typical fblacklist.txt format
	help.nsf
	log.nsf
	logs.nsf
	...


## hblacklist.txt
hblacklist.txt is a MD5 Hash listing blacklist file you may propigate.  It takes one MD5 Hash per line.

#typical hblacklist.txt format
	adf32923e2c67d4798b8bf33f0312c41
	380a35234d5ca93f71eee06207cf7001
	3ac41a1dc73242048af3b8567d809af7
	...
##  
## Configuration options
##
  
  For now most options are hard coded in the main() function these variables are:
  
	workingDir = ''       #This is the location of your root processing directory
	LotusDataPATH = ''    #default: C:\Users\blanksj\AppData\Local\IBM\Notes\Data'
	IDPATH = ''           #default: ~workingDir\IDs
	inifile = ''          #default: ~LotusDataPATH\notes.ini
	logpath = ''          #default: ~workingDir\
	LOADFILE = ''         #default: ~workingDir\load.txt
	DummyFile = ''        #default: ~IDPATH\dummy.id'
	NotesSQLCFG =         #default:C:\NotesSQL\notessql.cfg, in the future it will be set automaticly.
	
  
