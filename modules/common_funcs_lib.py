# THIS CONTAINS THE FUNCTIONS THAT'LL BE USED BY BOTH THE MEETING SCHEDULER AND MEETING CALLER

import os
import wmi
import sys
import time
import json
import winreg
import subprocess
from datetime import datetime

#function to display banner
def bannerDisp(heading):

    banLength = 60
    banChar = "-"
    headingLength = len(heading)
    
    print("{}".format(banChar * banLength))
    print("{}{}{}".format(" " * int((banLength / 2) - (headingLength / 2)), heading, " " * int((banLength / 2) - (headingLength / 2))))
    print("{}".format(banChar * banLength))

#function to return abs zoom path
def getZoomPath():

    try:
        #forming the abs path to zoom bin
        pathToAppData = os.getenv('APPDATA')
    except:
        #exit when appdata is not found
        print("\nAPPDATA location not found. Make sure you're running the script as administrator.")
        time.sleep(4)
        sys.exit()

    absPathToZoomBin = pathToAppData + "\\Zoom\\bin\\Zoom.exe"

    #checking if the zoom binary file exists
    if(os.path.isfile(absPathToZoomBin)):
        #return zoom.exe path when found
        return absPathToZoomBin
    else:
        #if Zoom is not found
        print("\nZoom.exe not found! Install Zoom Meetings App & restart the program.")
        time.sleep(4)
        sys.exit()

#funciton to add new meeting details to database
def addToDatabase(newMeetingDict):
    database = loadDatabase()
    database.append(newMeetingDict)
    saveDatabase(database)
    

#function to save to database
def saveDatabase(data):

    #the path where this file is kept (modules)
    cflPath = os.path.dirname(os.path.realpath(__file__))

    #variable to detect which script called it
        
    dataPath = ''
    try:
        if not os.path.exists(cflPath+'\\storage'):
            os.makedirs(cflPath+'\\storage')
    except Exception as e:
        print("Error creating database: {}\nExiting...".format(e))
        time.sleep(5)
        sys.exit()
    else:
        with open(cflPath+'\\storage\\database.json', "w") as outfile:  
            json.dump(data, outfile, default=str)

#function to read from database
def loadDatabase():
    #the path where this file is kept (modules)
    cflPath = os.path.dirname(os.path.realpath(__file__))

    try:
        fullPath = cflPath + '\\storage\\database.json'
            
        with open(fullPath) as infile: 
            database = json.load(infile)
    except Exception as e:
        #print("Database error: {}".format(e))
        database = []
        return database
    else:
        if(len(database)<1):
            return []
        else:
            #return database with datatime converted back from str to datetime again
            database = fixDatetime(database)
            return database

#function to convert json datetime strings back from str to datetime again
def fixDatetime(database):
    try:
        for i, meeting in enumerate(database):
            if(len(meeting["scheduled_at"]) > 0):
                meeting["scheduled_at"] = datetime.strptime(meeting["scheduled_at"], "%Y-%m-%d %H:%M:%S")
                
            if(len(meeting["stop_rec_time"]) > 0):
                meeting["stop_rec_time"] = datetime.strptime(meeting["stop_rec_time"], "%Y-%m-%d %H:%M:%S")

            if(len(meeting["end_at"]) > 0):
                meeting["end_at"] = datetime.strptime(meeting["end_at"], "%Y-%m-%d %H:%M:%S")
        return database
    except Exception as e:
        print("\nERROR CONVERTING DATETIME OBJECTS. DATABASE FAILURE. {}".format(e))
        time.sleep(3)
        sys.exit()



#function to find bandicam path
def findBandicamPath(findInstallPath = True):

    if(findInstallPath):
        #finding the installation path of bandicam through registry
        bandiPath = queryRegValue(r"SOFTWARE\BANDISOFT\BANDICAM", "ProgramPath")
        #checking if the bandicam binary file exists
        if(os.path.isfile(bandiPath)):
            #return zoom.exe path when found
            return bandiPath
    else:
        #to return the video saving path
        bandiPath = queryRegValue(r"SOFTWARE\BANDISOFT\BANDICAM\OPTION", "sOutputFolder")
        return bandiPath
    
    raise Exception("Sorry, Bandicam not found")

#function to registry query values
def queryRegValue(regPath, keyName):
    storedKey = winreg.OpenKey(winreg.HKEY_CURRENT_USER, regPath)
    keyValue =  winreg.QueryValueEx(storedKey, keyName)[0]

    return keyValue

#function to check running process
def checkProcRunning(procName):
    try:
        callTasklist = 'TASKLIST', '/FI', 'imagename eq %s' % procName
        result = subprocess.check_output(callTasklist, shell=True).decode()
        # checking end of string for proc name
        lastLine = result.strip().split('\r\n')[-1]
        # because Fail message could be translated
    except:
        print("Process checking failed. Recording might stop.")
    else:
        return lastLine.lower().startswith(procName.lower())

#function to execute command
def executeCommand(commandToExecute, shellBool, PopenBool = True):
    if(PopenBool):
        #supressing the output
        ONULL = open(os.devnull, 'w')
        #calling the command
        subprocess.Popen(commandToExecute, stdout=ONULL, stderr=ONULL, shell=shellBool)
    else:
        #supressing the output
        ONULL = open(os.devnull, 'w')
        #calling the command
        return subprocess.check_output(commandToExecute, shell=shellBool).decode().strip()

def terminateProcess(name):
    f = wmi.WMI()
    for process in f.Win32_Process():
        if process.name == name:
            process.Terminate()


#function to initialize bandicam
def initializeBandicamSetup(firstTime = False):
    #checking if SOFTWARE\\BANDISOFT\\BANDICAM\\OPTION exists
    try:
        storedKey = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"SOFTWARE\BANDISOFT\BANDICAM\OPTION")
    except:
        #directory doesn't exist - possible reason : bandicam was not executed even once after installation
        #possible fix - run bandicam once
        print("\nPrepearing Bandicam...")
        try:
            pathToBandicam = findBandicamPath()
        except:
            print("\nBandicam not found. Inatall Bandicam properly.")
        else:
            try:
                #execute bandicam once to get options directory in registry
                executeCommand(pathToBandicam, True)
                #wait for the execution
                time.sleep(5)
                #call the function again after the fix
                initializeBandicamSetup(True)
            except:
                print("\nError executing Bandicam. Manually open Bandicam & set recording option to full screen mode.")
    else:
        try:
            regPathToOptions = "SOFTWARE\\BANDISOFT\\BANDICAM\\OPTION"
            key1Name = "nTargetMode"
            key2Name = "nScreenRecordingSubMode"
            
            time.sleep(2)
            #check the values in current registry
            regKey1Val = queryRegValue(regPathToOptions, key1Name)
            regKey2Val = queryRegValue(regPathToOptions, key2Name)
            #change this reg keys: nScreenRecordingSubMode to 1, nTargetMode to 1 if not already correct
            print("\nChecking Bandicam configuarations...")
            if(regKey1Val != 1):
                #close the software and edit reg key
                if(firstTime):
                    try:
                        time.sleep(8)
                        if(checkProcRunning("bdcam.exe")):
                            terminateProcess("bdcam.exe")
                            firstTime = False
                    except:
                        print("\nManually open Bandicam & set recording option to full screen.")
                            
                time.sleep(2)
                executeCommand("REG ADD HKCU\\" + regPathToOptions + " /v " + key1Name + " /t REG_DWORD /d 1 /f", True)
                
            if(regKey2Val != 1):
                #close the software and edit reg key
                if(firstTime):
                    try:
                        time.sleep(8)
                        if(checkProcRunning("bdcam.exe")):
                            terminateProcess("bdcam.exe")
                            firstTime = False
                        else:
                            print("..")
                    except:
                        print("\nManually open Bandicam & set recording option to full screen.")
                                
                time.sleep(2)
                executeCommand("REG ADD HKCU\\" + regPathToOptions + " /v " + key2Name + " /t REG_DWORD /d 1 /f", True)
                
        except:
            print("Make sure bandicam is set to fullscreen mode.")
        else:
            #confirming the values have changed
            regKey1Val = queryRegValue(regPathToOptions, key1Name)
            regKey2Val = queryRegValue(regPathToOptions, key2Name)
            if(regKey1Val != 1 and regKey2Val != 1):
                print("First order configuration for recording failed, re-trying again.")
                initializeBandicamSetup()
            else:  
                print("\nBandicam is ready to record.")


#function to create a XML file for task import
def addTask(taskName, scheduleDateTime, pathToPythonExec, execModulePath, uniqueID):

    outXML = r'''<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <URI>\{taskName}</URI>
  </RegistrationInfo>
  <Triggers>
    <TimeTrigger>
      <StartBoundary>{scheduleDateTime}</StartBoundary>
      <Enabled>true</Enabled>
    </TimeTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>false</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>true</WakeToRun>
    <ExecutionTimeLimit>PT72H</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>"{pathToPythonExec}"</Command>
      <Arguments>"{execModulePath}" {uniqueID}</Arguments>
    </Exec>
  </Actions>
</Task>'''.format(taskName= taskName, scheduleDateTime= scheduleDateTime, pathToPythonExec= pathToPythonExec, execModulePath= execModulePath, uniqueID= uniqueID)

    #the path where this file is kept (modules)
    currentPath = os.path.dirname(os.path.realpath(__file__))
    #saving the task XML
    fullPath = currentPath + "\\taskSchedulerTemp.xml"
    f = open(fullPath, "w")
    f.write(outXML)
    f.close()

    commandToSchedule = 'SCHTASKS /CREATE /TN "'+ taskName +'" /XML "'+ fullPath +'" /F'

    executeCommand(commandToSchedule, True)


if os.name == 'nt':
    import ctypes
    from ctypes import windll, wintypes
    from uuid import UUID

    # ctypes GUID copied from MSDN sample code
    class GUID(ctypes.Structure):
        _fields_ = [
            ("Data1", wintypes.DWORD),
            ("Data2", wintypes.WORD),
            ("Data3", wintypes.WORD),
            ("Data4", wintypes.BYTE * 8)
        ] 

        def __init__(self, uuidstr):
            uuid = UUID(uuidstr)
            ctypes.Structure.__init__(self)
            self.Data1, self.Data2, self.Data3, \
                self.Data4[0], self.Data4[1], rest = uuid.fields
            for i in range(2, 8):
                self.Data4[i] = rest>>(8-i-1)*8 & 0xff

    SHGetKnownFolderPath = windll.shell32.SHGetKnownFolderPath
    SHGetKnownFolderPath.argtypes = [
        ctypes.POINTER(GUID), wintypes.DWORD,
        wintypes.HANDLE, ctypes.POINTER(ctypes.c_wchar_p)
    ]

    def _get_known_folder_path(uuidstr):
        pathptr = ctypes.c_wchar_p()
        guid = GUID(uuidstr)
        if SHGetKnownFolderPath(ctypes.byref(guid), 0, 0, ctypes.byref(pathptr)):
            raise ctypes.WinError()
        return pathptr.value

    FOLDERID_Download = '{374DE290-123F-4565-9164-39C4925E467B}'

    def getDownloadFolder():
        return _get_known_folder_path(FOLDERID_Download)
else:
    def getDownloadFolder():
        home = os.path.expanduser("~")
        return os.path.join(home, "Downloads")

    
