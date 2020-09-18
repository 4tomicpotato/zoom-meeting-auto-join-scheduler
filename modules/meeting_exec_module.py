import os
import sys
import threading
import time
from datetime import datetime
import common_funcs_lib as cfl
from pyautogui import screenshot

#fetch the index of the meeting provided as argument
try:
    uniqueID = int(sys.argv[1])
except:
    print("No valid meeting argument provided. Exiting...")
    time.sleep(3)
    sys.exit()

#load the database
database = cfl.loadDatabase()

def recordingFunction(enableRecording, stopRecTime):

    try:
        #start recorder if enabled
        if(enableRecording):
            #if badicam available
            try:
                pathToBandicam = cfl.findBandicamPath()
            except:
                print("\nRecording can't be enabled - Bandicam not found.")
                return False
            else:
                #start the initialize bandicam program to check and set reg keys
                cfl.initializeBandicamSetup()
                #start recording
                print("Initializing recording... \nMight take 10 to 15 seconds...")
                startRecording(pathToBandicam)
                print("Recording...")
                print("Your meeting will start in 20 seconds...")
                time.sleep(20)
                #start the keep recording function to keep recording even after 10mins
                print("\nNOTE: The recording will auto-start within 10 seconds if paused or stopped before {}".format(stopRecTime.strftime("%I:%M%p %d-%b-%Y")))
                print("To stop auto-starting of the recording please terminate this program OR close Bandicam from running apps & system tray.")
                keepRecording(stopRecTime, pathToBandicam)
                #the stop recrding function is integrated inside keep recording function
    except Exception as e:
        print("Recording failed! Error: {}".format(e))

#function to record the meeting using bandicam
def startRecording(pathToBandicam):

    #forming the full command to record using bandicam
    commandToStartRecording = '"' + pathToBandicam + '"' + " /record"
    try:
        cfl.executeCommand(commandToStartRecording, True)
    except:
        print("Recording failed. Please manually start recording.")
    

def keepRecording(stopRecTime, pathToBandicam):
    #check if bandicam is running and it's not past the stop time
    if((stopRecTime > datetime.now()) and cfl.checkProcRunning("bdcam.exe")):
        startRecording(pathToBandicam)
        #call itself with args after every inverval
        threading.Timer(10.0, keepRecording, [stopRecTime, pathToBandicam]).start()
    else:
        print("Stopping recording..")
        #call stop recording
        stopRecording()

def stopRecording():
    #finding the installation path of bandicam through registry
    try:
        pathToBandicam = cfl.findBandicamPath()
        outputPath =  cfl.findBandicamPath(False)
    except:
        print("\nBandicam not found. Please manually stop recording.")
    else:
        #stop recording
        commandToStopRecording = '"' + pathToBandicam + '"' + " /stop"
        try:
            cfl.executeCommand(commandToStopRecording, True)
            print("\nSaving video(s) to: {}".format(outputPath.strip()))
        except:
            print("Auto stop failed. Please manually stop recording.")
        else:
            print("Recorded video(s) saved.")
    
def commenceMeeting(zoommtgURL, meetingID, enableAutoReconnect, endAt):

    #check if zoom is available
    print("Finding Zoom on your computer...")
    zoomPath = cfl.getZoomPath()
    # form the command to join meeting
    argToPass = '--url="'+zoommtgURL+'"'
    commandToJoinMeeting = zoomPath+" "+argToPass

    #join the meeting
    print("Starting the meeting...")
    try:
        cfl.executeCommand(commandToJoinMeeting, False)
    except:
        print("Command failed. Meeting didn't start.")

    #if auto reconnect is enabled
    if(enableAutoReconnect):
        #if end at is empty stop auto reconnect feature
        if(endAt == ''):
            endAt = datetime.now()
        print("\nNOTE: You will be auto-reconnected to the meeting within 45 seconds, if disconnected before {} .".format(endAt.strftime("%I:%M%p %d-%b-%Y")))
        print("To stop auto reconnect feature before {}, terminate this program OR close Zoom meetings from running applications and system tray.".format(endAt.strftime("%I:%M%p %d-%b-%Y")))
        try:
            keepMeetingAlive(commandToJoinMeeting, endAt)
        except:
            print("Command failed. Auto-reconnect may not work.")

def keepMeetingAlive(commandToJoinMeeting, endAt):
    #check if zoom is running and it's not past the stop time
    if((endAt > datetime.now()) and cfl.checkProcRunning("Zoom.exe")):
        #start zoom meeting
        try:
            cfl.executeCommand(commandToJoinMeeting, False)
        except:
            pass
        #call itself with args after every inverval
        threading.Timer(45.0, keepMeetingAlive, [commandToJoinMeeting, endAt]).start()
    else:
        print("Zoom auto reconnection stopped.")

def screenShot(enableScreenshot):
    if(enableScreenshot):
        print("\nWill capture a screenshot after 2 mins")
        threading.Timer(120.0, takeScreenshot).start()

def takeScreenshot():
    try:
        meetingScreenShot = screenshot()
        pathToDekstop = os.path.join(os.environ["HOMEPATH"], "Desktop")
        imageName = "Meeting-"+ str(datetime.strftime(datetime.now(), '%d-%b-%Y-%H-%M-%S')) +".png"
        meetingScreenShot.save(pathToDekstop+"\\"+imageName)
    except:
        print("\nCapturing screenshot failed.")
    else:
        print("\nScreenshot saved to Desktop. Filename: "+imageName)

def handlePrevMeeting():
    #check if any meetings are running or not
    if(cfl.checkProcRunning("Zoom.exe")):
        print("\n\nWARNING: If you're already in a zoom meeting you will be disconnected in 20 seconds & your scheduled meeting will start.\nTo cancel the scheduled meeting terminate this program.")
        countDown(20)
        #terminate zoom if its running
        cfl.terminateProcess("Zoom.exe")
        time.sleep(2)
    


def countDown(sec):
    while sec: 
        mins, secs = divmod(sec, 60) 
        timer = '{:02d}:{:02d}'.format(mins, secs) 
        print(timer, end="\r") 
        time.sleep(1) 
        sec -= 1 
    print('Starting scheduled meeting') 

#function to start the schedule
def startMeeting(uniqueID):

    #clearing screen
    os.system('cls')
    cfl.bannerDisp("ZOOM AUTO SCHEDULER")
    
    global database

    try:

        foundFlag = 0
        #find the meeting with the given unique id
        for i, meeting in enumerate(database):
            if(meeting["unique_id"] == uniqueID):
                currentMeeting = database[i]
                foundFlag = 1
                break

    except Exception as e:
        print("Error: {}".format(e))
        print("Exiting in 3secs...")
        time.sleep(3)
        sys.exit()

    if(foundFlag == 0):
        print("Meeting with Unique ID: {} wasn't found.\nPossible reason: Meeting deleted.\nExiting...".format(uniqueID))
        time.sleep(4)
        sys.exit()


    #if there is any older scheduled meeting than the current time (except the current index) stop and delete it with logs - kill keep alive functions
    #handlePrevMeeting()
    
    #clearing screen
    os.system('cls')
    cfl.bannerDisp("MEETING IN PROGRESS")
    print("\n")

    #starting the recording if enabled
    recordingFunction(currentMeeting["enable_recording"], currentMeeting["stop_rec_time"])
    #start the meeting function
    commenceMeeting(currentMeeting["zoommtg_url"], currentMeeting["meeting_id"], currentMeeting["enable_auto_reconnect"], currentMeeting["end_at"])
    #take screenshot
    screenShot(currentMeeting["enable_screenshot"])

    #details about how to stop the function


startMeeting(uniqueID)
