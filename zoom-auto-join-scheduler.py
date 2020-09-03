from datetime import datetime
from urllib import parse
import threading
import subprocess
import os
import winreg
import requests
import sys
import wmi
import time

from pyautogui import screenshot

database = []

def getZoomPath():

    try:
        #forming the abs path to zoom bin
        pathToAppData = os.getenv('APPDATA')
    except:
        #exit when appdata is not found
        sys.exit("\nAPPDATA location not found. Make sure you're running the script as administrator.")

    absPathToZoomBin = pathToAppData + "\\Zoom\\bin\\Zoom.exe"

    #checking if the zoom binary file exists
    if(os.path.isfile(absPathToZoomBin)):
        #return zoom.exe path when found
        return absPathToZoomBin
    else:
        #if Zoom is not found
        sys.exit("\nZoom.exe not found! Install Zoom Meetings App & restart the program.")

def bannerDisp(heading):

    banLength = 60
    banChar = "-"
    headingLength = len(heading)
    
    print("{}".format(banChar * banLength))
    print("{}{}{}".format(" " * int((banLength / 2) - (headingLength / 2)), heading, " " * int((banLength / 2) - (headingLength / 2))))
    print("{}".format(banChar * banLength))

def findBandicamPath():
    #finding the installation path of bandicam through registry
    bandiPath =  queryRegValue(r"SOFTWARE\BANDISOFT\BANDICAM", "ProgramPath")

    #checking if the bandicam binary file exists
    if(os.path.isfile(bandiPath)):
        #return zoom.exe path when found
        return bandiPath
    
    raise Exception("Sorry, Bandicam not found")

def queryRegValue(regPath, keyName):
    storedKey = winreg.OpenKey(winreg.HKEY_CURRENT_USER, regPath)
    keyValue =  winreg.QueryValueEx(storedKey, keyName)[0]

    return keyValue

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

def terminateProcess(name):
    f = wmi.WMI()
    for process in f.Win32_Process():
        if process.name == name:
            process.Terminate()

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

def downloadAndInstallBadicam():
    link = "https://dl.bandicam.com/bdcamsetup.exe"

    pathToDownloads = os.path.join(os.environ["HOMEPATH"], "Downloads")
    file_name = "bdcamsetup.exe"
    file_abs_path = pathToDownloads+"\\"+file_name

    
    
    with open(file_abs_path, "wb") as f:
            print("\nDownloading Bandicam")
            try:
                response = requests.get(link, stream=True)
                total_length = response.headers.get('content-length').strip()
            except:
                print("\nError downloading. Make sure you're connected to the internet.")
            
            if total_length is None: # no content length header
                f.write(response.content)
                
            try:
                fullSize = round(int(total_length) / (1024 * 1024), 1)
                print("\nSize: " + str(fullSize) + " MB")

            except:
                print("\nError downloading. Manually download and install Bandicam.")
                
            else:
                dl = 0
                total_length = int(total_length)
                for data in response.iter_content(chunk_size=2048):
                    dl += len(data)
                    f.write(data)
                    done = int(50 * dl / total_length)
                    sys.stdout.write("\r[{0}{1}] {2} MB / {3} MB".format(('=' * done), (' ' * (50-done)), round((dl / (1024 * 1024)), 1), fullSize))    
                    sys.stdout.flush()


    print("\nBandicam download complete!")
    print("\nInstalling...", end="")

    #installing bandicam
    argToInstallBandicam = " /S"
    try:
        executeCommand(file_abs_path + argToInstallBandicam, True)
    except:
        print("\nBandicam installation failed. Install it manually.")
    else:
        while 1:
            print(".", end ="")
            try:
                path = findBandicamPath()
            except:
                time.sleep(1)
            else:
                print("\nInstallation complete!")
                break
            
            time.sleep(3)

    initializeBandicamSetup()

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
            print("Failed to configure Bandicam recording options. Manually open Bandicam & set recording option to full screen.")
        else:
            #confirming the values have changed
            regKey1Val = queryRegValue(regPathToOptions, key1Name)
            regKey2Val = queryRegValue(regPathToOptions, key2Name)
            if(regKey1Val != 1 and regKey2Val != 1):
                print("First order configuration for recording failed, re-trying again.")
                initializeBandicamSetup()
            else:  
                print("\nBandicam is ready to record.")


def inputMeetingURL():
    while 1:
        #meeting URL input
        meetingURL = input("\nEnter the full zoom meeting URL here: ").strip()

        try:
            #parsing the URL into componenets
            parsed = parse.urlsplit(meetingURL)

            zoomServer = parsed.netloc
            
            #not a zoom meeting
            if "zoom.us" not in zoomServer:
                print("\nInvalid URL input. Your meeting URL should look like: https://us04web.zoom.us/j/XXXXXXXXXX?pwd=XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX. Try again.")
                continue

            
            #checking the validity of input url
            hashedMeetingPwd = parse.parse_qs(parsed.query)['pwd'][0]
            if(len(hashedMeetingPwd) != 32):
                print("\nInvalid pwd part of URL. Your meeting URL should look like: https://us04web.zoom.us/j/XXXXXXXXXX?pwd=XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX. Try again.")
                continue

            meetingID = parsed.path.split("/")[-1]
            if(not(meetingID.isdecimal())):
                print("\nInvalid ID part of URL. Your meeting URL should look like: https://us04web.zoom.us/j/XXXXXXXXXX?pwd=XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX. Try again.")
                continue
        except:
            print("\nURL parsing failed. Make sure you're providing the proper joining URL.")
            continue

        #return the url & they keys after input is processed & checked
        return (meetingURL, zoomServer, meetingID, hashedMeetingPwd)

def inputScheduledAt():
    while 1:
        #scheduling date-time input
        scheduledAt = input("\nEnter the date & time when the meeting is scheduled to start in DD-MM-YYYY HH:MM format (Ex: 28-08-2020 21:00): ").strip()

        #Parsing the date from str to datetime object & error-handling
        try: 
            scheduledAt = datetime.strptime(scheduledAt, '%d-%m-%Y %H:%M')
        except:
            print("\nInvalid date-time input. It should be in DD-MM-YYYY HH:MM format. Try again.")
            continue
        
        #checking the validity of input date-time
        if(datetime.now() > scheduledAt):
            print("\nInvalid Date & Time input. Scheduled date-time can't be lower than current date-time. Try again.")
            continue
        
        #return the scheduled time after input is processed & checked
        return scheduledAt

def inputUsername():
    #username Input
    joinAs = input("\nEnter a name to join the meeting as (Optional): ").strip()

    if(len(joinAs) < 1):
        joinAs = ''

    if(len(joinAs) > 40):
        joinAs = joinAs[:40]
        
    return joinAs

def inputEnableScreenshot():

    #input loop
    while 1:
        #screenshot option input
        enableScreenshot = input("\nDo you want to take a screenshot of your meeting? (y/n) ").strip().lower()
        
        if(enableScreenshot == 'y'):
            print("\nScreenshot will be taken 2 minutes after scheduled start-time & saved on your computer.")
            return True
        elif(enableScreenshot == 'n'):
            return False
        else:
            print("\nInvalid input. Press 'y' for 'Yes' or 'n' for 'No'.")
            continue
        
def inputAutoReConnect(scheduledAt):

    #input loop
    while 1:
        #auto-reconnect option input
        enableAutoReConnect = input("\nIf you have poor-network connection or if your meetings are executed part-wise in consecutive sessions, you can enable auto reconnect feature.\nThis will automatically connect you to the meeting if disconnected, removed or you leave. \nDo you want to auto re-connect to the meeting if it disconnects before a specified time? (y/n) ").strip().lower()

        if(enableAutoReConnect == 'y'):
            #if auto reconnect is enabled get the reconnect-till time
            while 1:
                #ending date-time input
                endAt = input("\nYou will be auto-reconnected to the meeting if it gets disconnected before the assumed ending time. \nEnter your assumed meeting ending time (in DD-MM-YYYY HH:MM format): ").strip()

                #parsing the date from str to datetime object & error-handling
                try: 
                    endAt = datetime.strptime(endAt, '%d-%m-%Y %H:%M')
                except:
                    print("\nInvalid date-time input. It should be in DD-MM-YYYY HH:MM format (Example: 28-08-2020 21:50). Try again.")
                    continue
                    
                #checking the validity of end input date-time
                if(scheduledAt > endAt):
                    print("\nInvalid Date & Time input. Meeting ending time can't be lower than the scheduled start time. Try again.")
                    continue

                return (True, endAt)
            
        elif(enableAutoReConnect == 'n'):
            return (False, '')
        
        else:
            print("\nInvalid input. Press 'y' for 'Yes' or 'n' for 'No'.")
            continue

        
def inputRecordingOptions(scheduledAt, endAt):
    #input loop
    while 1:
        #recording option input
        enableRecording = input("\nDo you want to record your meeting (Bandicam required)? (y/n) ").strip().lower()
        
        if(enableRecording == 'y'):

            #check if recording tool is installed
            try:
                progPath =  findBandicamPath()
            except:
                while 1:
                    downloadBool = input("\nBandicam is required for recording. Do you want to download & install Bandicam (y/n)?").strip().lower()
                    if(downloadBool == 'y'):
                        #download and install bandicam
                        downloadAndInstallBadicam()
                        break
                    elif(downloadBool == 'n'):
                        #disable recording feature when there is no bandicam and user does not want to download it
                        print("\nRecording feature can't be availed without installing Bandicam.")
                        return (False, '')
                    else:
                        #ask again, on ivalid input
                        print("\nInvalid input. Press 'y' for 'Yes' or 'n' for 'No'.")
                        continue

            needInputFlag = 0
            #if end time is specified, ask if user wants to stop recording at that time
            if(endAt != ''):
                stopRecTimeBool = input("\nDo you want to stop the recording at " + str(endAt) + "? (y/N) ").strip().lower()
                if (stopRecTimeBool == 'y'):
                    stopRecTime = endAt
                    return (True, stopRecTime)
                else:
                    needInputFlag = 1
                    
            if(endAt == '' or needInputFlag == 1):
                while 1:
                    stopRecTime = input("\nEnter the date-time when you want the recording to stop. Input in DD-MM-YYYY HH:MM format: ").strip()
                    #Parsing the date from str to datetime object & error-handling
                    try: 
                        stopRecTime = datetime.strptime(stopRecTime, '%d-%m-%Y %H:%M')
                    except:
                        print("\nInvalid date-time input. It should be in DD-MM-YYYY HH:MM format (Example: 28-08-2020 21:50). Try again.")
                        continue
                        
                    #checking the validity of end input date-time
                    if(scheduledAt > stopRecTime):
                        print("\nInvalid Date & Time input. Stop recodring time can't be lower than the scheduled start time. Try again.")
                        continue
                        
                    return (True, stopRecTime)
                    
        elif(enableRecording == 'n'):
            return (False, '')
        else:
            print("\nInvalid input. Press 'y' for 'Yes' or 'n' for 'No'.")
            continue    
        
def makeZoommtgURL(zoomServer, meetingID, hashedMeetingPwd, joinAs):
    #generating the new url which has zoommtg protocol (will be used to pass as argument to zoom.exe)
    zoommtgURL = "zoommtg://" + zoomServer + "/join?action=join&confno=" + meetingID + "&pwd=" + hashedMeetingPwd

    #adding username to zoommtgURL if available
    #check if username is empty
    if(len(joinAs) > 0):
        try:
            #url encoding the username
            joinAs = parse.quote(joinAs)
        except:
            print("Problem in URL encoding the username. Default username of your Zoom app will be used.")
        else:
            zoommtgURL = zoommtgURL + "&uname=" + joinAs

    return zoommtgURL

def dispAllMeetings():
    print("\nList of all scheduled meetings: ")
    print ("\n  {:<6} | {:<21} | {:<15} | {:<40}".format('SL.No', 'SCHEDULED_AT', 'MEETING_ID', 'USERNAME')) 

    global database

    #to prevent any accidental updation of global database, making a copy locally
    localDatabase = database.copy()

    for i, meeting in enumerate(localDatabase):
        print("  {:<70}".format( '-' * 63))
        print ("  {:<6} | {:<21} | {:<15} | {:<40}".format(i+1, str(meeting["scheduled_at"].strftime("%I:%M%p %d-%b-%Y")[:21]),
                                                                        str(meeting["meeting_id"][:15]), str(meeting["join_as"][:40])))
        print("  {:6} Screenshot Enabled: {}".format(" ", "Yes" if(meeting["enable_screenshot"]) else "No"))
        print("  {:6} Recording Enabled: {}".format(" ", "Yes" if(meeting["enable_recording"]) else "No"))
        if(meeting["enable_recording"]):
            print("  {:6} Record Till: {}".format(" ", str(meeting["stop_rec_time"].strftime("%I:%M%p %d-%b-%Y")) ))
        print("  {:6} Auto Reconnect Enabled: {}".format(" ", "Yes" if(meeting["enable_auto_reconnect"]) else "No"))
        if(meeting["enable_auto_reconnect"]):
            print("  {:6} Reconnect if disconnected before: {}".format(" ", str(meeting["end_at"].strftime("%I:%M%p %d-%b-%Y")) ))
        print("  {:6} Meeting URL: {}".format(" ", meeting["meeting_url"]))


    
def add_new_meeting():

    #clearing screen
    os.system('cls')

    bannerDisp("ADD NEW MEETING")
    print("\n")

    #get everything about meeting URL
    meetingURL, zoomServer, meetingID, hashedMeetingPwd = inputMeetingURL()

    #get the scheduled starting time
    scheduledAt = inputScheduledAt()

    #get the username
    joinAs = inputUsername()

    #get screenshot option
    enableScreenshot = inputEnableScreenshot()

    #get auto reconnect details
    enableAutoReConnect, endAt = inputAutoReConnect(scheduledAt)
    
    #get recording options
    enableRecording, stopRecTime = inputRecordingOptions(scheduledAt, endAt)

    #forming the zoommtg url
    zoommtgURL = makeZoommtgURL(zoomServer, meetingID, hashedMeetingPwd, joinAs)

    #structuring the data into a dictionary
    newMeeting = {
            "scheduled_at" : scheduledAt,
            "meeting_url" : meetingURL,
            "join_as" : joinAs,
            "enable_screenshot" : enableScreenshot,
            "enable_auto_reconnect" : enableAutoReConnect,
            "end_at" : endAt,
            "enable_recording" : enableRecording,
            "stop_rec_time" : stopRecTime,
            "zoom_server" : zoomServer,
            "meeting_id" : meetingID,
            "hashed_meeting_pwd" : hashedMeetingPwd,
            "zoommtg_url" : zoommtgURL
        }

    #adding new meeting details to global database
    global database
    database.append(newMeeting.copy())

    print("\nMeeting scheduled.")
    time.sleep(1)

    show_all_meetings()
    

def show_all_meetings():

    #clearing screen
    os.system('cls')

    bannerDisp("ALL MEETINGS")
    print("\n")

    dispAllMeetings()

    wait = input("\n\nPress any key to go back to Main Menu.")
    

def delete_meetings():
    #clearing screen
    os.system('cls')

    bannerDisp("DELETE MEETINGS")
    print("\n")

    dispAllMeetings()

    global database;

    while 1:
        print("\nTo go back to Main Menu enter 0. \nTo delete a meeting enter it's SL.No ")

        try:
            delMeetingIndex = int(input("Choose an option: ").strip())
        except:
            print("Invalid input. Try again.")
            continue
        else:
            if(delMeetingIndex > 0):
                #verify if there's a meeting with that sl. no.
                if(delMeetingIndex <= len(database)):
                    meetingID = database[delMeetingIndex-1]["meeting_id"]
                    scheduleStr = str(database[delMeetingIndex-1]["scheduled_at"].strftime("%I:%M%p %d-%b-%Y"))
                    confirmDel = input("\nAre you sure you want to delete meeting no {} (Meeting ID: {}, Scheduled At: {})? [y/N] ".format(delMeetingIndex, meetingID, scheduleStr)).strip().lower()
                    if(confirmDel == 'y'):
                        del database[delMeetingIndex-1]
                        print("Deleteing..")
                        time.sleep(1)
                        delete_meetings()
                        break
                    else:
                        print("Deletion aborted.")
                        continue
                else:
                    print("Invalid SL.No")
                    continue
            elif(delMeetingIndex == 0):
                #to go back to main menu
                break
            else:
                print("Invalid input")
                continue

def quit_script():
    #clearing screen
    os.system('cls')

    bannerDisp("Author : github.com/r4jdip")
    print("\n")

    time.sleep(3)
    sys.exit()


def main_menu():

    #clearing screen
    os.system('cls')

    bannerDisp("MAIN MENU")
    print("\n")
    
    #MAIN MENU
    mainMenu = {
                1:("Add a meeting schedule", add_new_meeting),
                2:("View all scheduled meetings", show_all_meetings),
                3:("Delete scheduled meetings", delete_meetings),
                4:("Quit", quit_script)
               }

    #print main menu
    for key in sorted(mainMenu.keys()):
         print("\t" + str(key) + ":" + mainMenu[key][0])

    #choose from main menu
    while 1:
        try:     
            option = int(input("\nChoose an option: ").strip())
        except:
            #when int conversion fails
            print("\nInvalid input. Please choose an option between 1 to " + str(len(mainMenu)))
            continue

        #if the input is out of range
        if(option < 1 or option > len(mainMenu)):
            print("\nInvalid input. Please choose an option between 1 to " + str(len(mainMenu)))
            continue
        break

    #select a function according to input
    mainMenu.get(option)[1]()


def initializeScreen():

    #clearing screen
    os.system('cls')
    
    print("\nChecking Zoom installation....")

    #checking if zoom is installed
    print("Detecting Zoom executable path: " + getZoomPath())

    print("Initializing...")

    time.sleep(1)


initializeScreen()
while 1:
    main_menu()
wait = input("PRESS ANY KEY TO EXIT")
