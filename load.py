import csv

class taskRank(object):
	def __init__(self, startTime,loadWithISR,load):
		self.startTime = startTime
		self.loadWithISR = loadWithISR
		self.load = load
	def __repr__(self):
		return "<startTime: %d, loadWithISR: %d, loadWithoutISR: %d>" % (self.startTime, self.loadWithISR, self.load)

class ISRRank(object):
	def __init__(self, startTime,load):
		self.startTime = startTime
		self.load = load
	def __repr__(self):
		return "<startTime: %d, load: %d>" % (self.startTime, self.load)

# Constants
TimeWindow = 5000000	#in ns
IdleTaskId = 0x1C		#IDLE Task ID
FuncNoneId = 0x0		#No Interrupt ID
numOfISRDirectFunc = 4
BCTaskID = [0x0, 0x2, 0xA, 0xB, 0xC, 0xD, 0xE, 0xF, 0x10, 0x11]
ResultsSlots = 3

# Variable for calculation IO
Tasks = [[0 for x in range(2)]]
ISRs = [[0 for x in range(2)]]
CanWakeUpInt = [[0 for x in range(2)]]
CanTxInt = [[0 for x in range(2)]]
CanRxInt = [[0 for x in range(2)]]
CanErrorInt = [[0 for x in range(2)]]

# Variable for ranking
initTaskRank = taskRank(0,0,0)
initISRRank = ISRRank(0,0)
rank_load = [initTaskRank]
rank_isr = [initISRRank]


def loadCSVFiles(fileName):
	global Tasks
	global ISRs
	global CanWakeUpInt
	global CanTxInt
	global CanRxInt
	global CanErrorInt

	with open(fileName,'rU') as csvfile:
		reader = csv.reader(csvfile)

		thisFile = list(reader)

		dataStartsHere = taskStartsHere = ISRStartsHere = functionStartsHere = CanWakeUpIntStartsHere = CanTxIntStartsHere = CanRxIntStartsHere = CanErrIntStartsHere = -1

		for i in range(len(thisFile)):
			if(thisFile[i] != []):
				if(thisFile[i][0]=="*********** History: Data ***************"):
					# If this line is not empty AND is the start of Data section
					dataStartsHere = i
					break

		# Error handling: if "History: Data" Section cannot be found in CSV file
		try:
			if(dataStartsHere<0):
				raise ValueError()
		except(ValueError):
			print("Error: Can't find 'History: Data' section in CSV file: " + fileName)
			return

		# Load Tasks, ISRs, CanWakeUpInterrupt0, CanTxInterrupt_0, CanRxInterrupt_0, CanErrorInterrupt_0

		#
		# Tasks
		#
		for i in range(dataStartsHere, len(thisFile)):
			if(thisFile[i] != []):
				if thisFile[i][0]=="Tasks":
					# If this line is not empty AND is the start of Tasks
					taskStartsHere = i + 2
					break

		# Error handling: if "Tasks" Section cannot be found in CSV file
		try:
			if(taskStartsHere<0):
				raise ValueError()
		except(ValueError):
			print("Error: Can't find 'Tasks' section in CSV file: " + fileName)
			return

		try:
			if(thisFile[taskStartsHere][0] == ""):
				raise ValueError()
		except(ValueError):
			print("Error: No valid data can be found in 'Tasks' Section: csv file: "+fileName)
			return
		else:
			Tasks[0][0] = int(thisFile[taskStartsHere][0])
			Tasks[0][1] = int(thisFile[taskStartsHere][1],16)
			j = 1

			while((taskStartsHere+j<len(thisFile)) and (thisFile[taskStartsHere+j]!=[]) and (thisFile[taskStartsHere+j][0]!="")):
				Tasks.append([0 for x in range(2)])
				Tasks[j][0] = int(thisFile[taskStartsHere+j][0],10)
				Tasks[j][1] = int(thisFile[taskStartsHere+j][1],16)
				j = j + 1

		#
		# ISRs
		#
		for i in range(taskStartsHere+j, len(thisFile)):
			if(thisFile[i] != []):
				if thisFile[i][0]=="ISRs2":
					# If this line is not empty AND is the start of ISRs
					ISRStartsHere = i + 2
					break

		# Error handling: if "ISRs" Section cannot be found in CSV file
		try:
			if(ISRStartsHere<0):
				raise ValueError()
		except(ValueError):
			print("Error: Can't find 'ISRs' section in CSV file: "+fileName)
			return

		try:
			if(thisFile[ISRStartsHere][0] == ""):
				raise ValueError()
		except(ValueError):
			print("Error: No valid data can be found in 'ISRs' Section, csv file: "+fileName)
			return
		else:
			ISRs[0][0] = int(thisFile[ISRStartsHere][0])
			ISRs[0][1] = int(thisFile[ISRStartsHere][1],16)
			j = 1
			while((ISRStartsHere+j<len(thisFile)) and (thisFile[ISRStartsHere+j]!=[]) and (thisFile[ISRStartsHere+j][0]!="")):
				ISRs.append([0 for x in range(2)])
				ISRs[j][0] = int(thisFile[ISRStartsHere+j][0])
				ISRs[j][1] = int(thisFile[ISRStartsHere+j][1],16)
				j = j + 1



		# Find "History: Functions"
		for i in range(ISRStartsHere+j, len(thisFile)):
			if(thisFile[i] != []):
				if(thisFile[i][0]=="********** History: Functions  **********"):
					# If this line is not empty AND is the start of Data section
					functionStartsHere = i
					break
		# Error handling: if "History: Data" Section cannot be found in CSV file
		try:
			if(functionStartsHere<0):
				raise ValueError()
		except(ValueError):
			print("Error: Can't find 'History: Function' section in CSV file: "+ fileName)
			return

		# Load CanWakeUpInterrupt0, CanTxInterrupt_0, CanRxInterrupt_0, CanErrorInterrupt_0
		#
		# CanWakeUpInterrupt0
		#
		# Error handling: if "CanWakeUpInterrupt0" Section cannot be found in CSV file
		for i in range(functionStartsHere, len(thisFile)):
			if(thisFile[i] != []):
				if thisFile[i][0]=="CanWakeUpInterrupt_0":
					# If this line is not empty AND is the start of CanWakeUpInterrupt0
					CanWakeUpIntStartsHere = i + 2
					break
		try:
			if(CanWakeUpIntStartsHere<0):
				raise ValueError()
		except(ValueError):
			print("Error: Can't find 'CanWakeUpInterrupt0' section in CSV file: "+ fileName)
			return

		try:
			if(thisFile[CanWakeUpIntStartsHere][0] == ""):
				raise ValueError()
		except(ValueError):
			print("Error: No valid data can be found in 'CanWakeUpInterrupt0' Section, csv file: "+fileName)
			return
		else:
			CanWakeUpInt[0][0] = int(thisFile[CanWakeUpIntStartsHere][0])
			CanWakeUpInt[0][1] = int(thisFile[CanWakeUpIntStartsHere][1],16)
			j = 1
			while((CanWakeUpIntStartsHere+j<len(thisFile)) and (thisFile[CanWakeUpIntStartsHere+j]!=[]) and (thisFile[CanWakeUpIntStartsHere+j][0]!="")):
				CanWakeUpInt.append([0 for x in range(2)])
				CanWakeUpInt[j][0] = int(thisFile[CanWakeUpIntStartsHere+j][0])
				CanWakeUpInt[j][1] = int(thisFile[CanWakeUpIntStartsHere+j][1],16)
				j = j + 1

		#
		# CanTxInterrupt_0
		#
		# Error handling: if "CanTxInterrupt_0" Section cannot be found in CSV file
		for i in range(functionStartsHere, len(thisFile)):
			if(thisFile[i] != []):
				if thisFile[i][0]=="CanTxInterrupt_0":
					# If this line is not empty AND is the start of CanTxInterrupt_0
					CanTxIntStartsHere = i + 2
					break

		try:
			if(CanTxIntStartsHere<0):
				raise ValueError()
		except(ValueError):
			print("Error: Can't find 'CanTxInterrupt_0' section in CSV file: "+fileName)
			return

		try:
			if(thisFile[CanTxIntStartsHere][0] == ""):
				raise ValueError()
		except(ValueError):
			print("Error: No valid data can be found in 'CanTxInterrupt_0' Section, csv file: "+fileName)
			return
		else:
			CanTxInt[0][0] = int(thisFile[CanTxIntStartsHere][0])
			CanTxInt[0][1] = int(thisFile[CanTxIntStartsHere][1],16)
			j = 1
			while((CanTxIntStartsHere+j<len(thisFile)) and (thisFile[CanTxIntStartsHere+j]!=[]) and (thisFile[CanTxIntStartsHere+j][0]!="")):
				CanTxInt.append([0 for x in range(2)])
				CanTxInt[j][0] = int(thisFile[CanTxIntStartsHere+j][0])
				CanTxInt[j][1] = int(thisFile[CanTxIntStartsHere+j][1],16)
				j = j + 1


		#
		# CanRxInterrupt_0
		#
		# Error handling: if "CanRxInterrupt_0" Section cannot be found in CSV file
		for i in range(functionStartsHere, len(thisFile)):
			if(thisFile[i] != []):
				if thisFile[i][0]=="CanRxInterrupt_0":
					# If this line is not empty AND is the start of CanRxInterrupt_0
					CanRxIntStartsHere = i + 2
					break

		try:
			if(CanRxIntStartsHere<0):
				raise ValueError()
		except(ValueError):
			print("Error: Can't find 'CanRxInterrupt_0' section in CSV file: "+fileName)
			return

		try:
			if(thisFile[CanRxIntStartsHere][0] == ""):
				raise ValueError()
		except(ValueError):
			print("Error: No valid data can be found in 'CanRxInterrupt_0' Section, csv file: "+fileName)
			return
		else:
			CanRxInt[0][0] = int(thisFile[CanRxIntStartsHere][0])
			CanRxInt[0][1] = int(thisFile[CanRxIntStartsHere][1],16)
			j = 1
			while((CanRxIntStartsHere+j<len(thisFile)) and (thisFile[CanRxIntStartsHere+j]!=[]) and (thisFile[CanRxIntStartsHere+j][0]!="")):
				CanRxInt.append([0 for x in range(2)])
				CanRxInt[j][0] = int(thisFile[CanRxIntStartsHere+j][0])
				CanRxInt[j][1] = int(thisFile[CanRxIntStartsHere+j][1],16)
				j = j + 1

		#
		# CanErrorInterrupt_0
		#
		# Error handling: if "CanErrorInterrupt_0" Section cannot be found in CSV file
		for i in range(functionStartsHere, len(thisFile)):
			if(thisFile[i] != []):
				if thisFile[i][0]=="CanErrorInterrupt_0":
					# If this line is not empty AND is the start of CanErrorInterrupt_0
					CanErrIntStartsHere = i + 2
					break

		try:
			if(CanErrIntStartsHere<0):
				raise ValueError()
		except(ValueError):
			print("Error: Can't find 'CanErrorInterrupt_0' section in CSV file: "+fileName)
			return

		try:
			if(thisFile[CanErrIntStartsHere][0] == ""):
				raise ValueError()
		except(ValueError):
			print("Error: No valid data can be found in 'CanErrorInterrupt_0' Section, csv file: "+fileName)
			return
		else:
			CanErrorInt[0][0] = int(thisFile[CanErrIntStartsHere][0])
			CanErrorInt[0][1] = int(thisFile[CanErrIntStartsHere][1],16)
			j = 1
			while((CanErrIntStartsHere+j<len(thisFile)) and (thisFile[CanErrIntStartsHere+j]!=[]) and (thisFile[CanErrIntStartsHere+j][0]!="")):
				CanErrorInt.append([0 for x in range(2)])
				CanErrorInt[j][0] = int(thisFile[CanErrIntStartsHere+j][0])
				CanErrorInt[j][1] = int(thisFile[CanErrIntStartsHere+j][1],16)
				j = j + 1


def getLoad(item):
	return item.load

def insertLoad(item, list):
	list.append(item)

	list.sort(key=getLoad,reverse=True)
	if(len(list) > ResultsSlots):
		del list[ResultsSlots]

#def ProcessISRfunctions(taskIdleStartTime, taskIdleEndTime, )

def load_calculation():
	print(len(Tasks))
	print(len(ISRs))
	print(len(CanWakeUpInt))
	print(len(CanTxInt))
	print(len(CanRxInt))
	print(len(CanErrorInt))

	# Data Initializations for CPU load calculation
	AccPercentage = 0
    AccPercentageNoISR = 0
    AccISRsPercentage = 0

	CurrentRow = 0
	TaskTimeError = 0
	while ((CurrentRow < len(Tasks)) and (TaskTimeError == 0)):
		StartTime = Task[CurrentRow][0]
		CurrentTime = StartTime
        # Init variables for CPU load calculation if current Window
        IdleTaskMeasuring = 0
        IdleTaskAcummulatedTime = 0
        FuncAllAccTime = 0
        TaskTimeError = 0

		while ((CurrentTime <= (StartTime + TimeWindow)) and (TaskTimeError = 0)):
			NoCPUusageTask = 0
			if(Task[CurrentRow][1] in BCTaskID):
				NoCPUusageTask = 1
			if()




	# insertLoad(taskRank(400,89,98),rank_load)

def cpu_load(fileName):
	print("Loading data from csv file: " + fileName + "...")
	print("")
	loadCSVFiles(fileName)
	print("Calculating CPU Load from csv file: " + fileName + "...")
	print("")
	load_calculation()
	print("Calculation completed for csv file: " + fileName + "...")
	print("")

	# insertLoad(taskRank(100,39,44),rank_load)
	# insertLoad(taskRank(200,49,56),rank_load)
	# insertLoad(taskRank(300,99,99),rank_load)
	# insertLoad(taskRank(400,89,98),rank_load)
	print(rank_load)



# Main functions: CPU Load calculation for each csv file generated
#cpu_load('test2.csv')
cpu_load('pccp_test.csv')
