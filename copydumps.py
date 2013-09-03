import paramiko
import os, wmi

def getProcessList():
	try:
		c = wmi.WMI(moniker='//./root/cimv2:Win32_Process')
		procs = c.instances()
		procList = list()
		for i in procs:
			procList.append((i.ProcessId, i.Caption, i.CommandLine))
		return procList
	except:
		print "Exception while getting list of process"
		return None

def isDriverPrccessRunning():
	isDriverRunning = False
	try:
		procList = getProcessList()
		if procList == None:
			print "Unable to get the processess List"
		for proc in procList:
			procid = proc[0]
			procName = proc[1]
			procCmdLine = proc[2]
			if procName != None and procCmdLine != None:
				if procName.lower() == "python.exe":
					if "copydumps.py" in procCmdLine.lower():
						if str(os.getpid()) != str(procid):
							isDriverRunning = True
		return isDriverRunning
	except:
		print "Exception while checking if Driver is running"
		return None
def getParametersFile():
	dname = os.path.dirname(os.path.abspath(__file__))
	return os.path.join(dname, "Parameters.config")
	
def getParameters(filename = None):
	if filename is None:
		filename = getParametersFile()
		
	try:
		fin = open(filename)
	except IOError:
		print "Could not open file"
		fin.close()
		return
	
	parameters = dict()
	
	while 1:
		line = fin.readline()
		if not line: break
		words = line.split()
		
		if(len(words) !=2):
			continue
			
		parameters[words[0]] = words[1]
		
	fin.close()
	return parameters
		
if not isDriverPrccessRunning():

	parameters = getParameters()
	
	host = parameters["host"]
	port = int(parameters["port"])
	transport = paramiko.Transport((host, port))

	password = parameters["password"]
	username = parameters["username"]

	transport.connect(username = username, password = password)

	sftp = paramiko.SFTPClient.from_transport(transport)

	remotepath = parameters["remotepath"]
	localpath = parameters["localpath"]

	dirlist = sftp.listdir(remotepath)
	
	for file in dirlist:
		filepath = os.path.join(localpath, file)
		fileremotepath =os.path.join(remotepath, file)
	
		print 'Transferring the file %s'%file
		sftp.get(fileremotepath, filepath)
		
		print 'Removing the file %s'%file
		sftp.remove(fileremotepath)


	sftp.close()
	transport.close()
	print 'Transfer Finished.'
	
else: 
	print 'Driver is Running'
	