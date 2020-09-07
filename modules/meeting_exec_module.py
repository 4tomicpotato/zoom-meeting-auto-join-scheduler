import os
import sys
import threading
import time
from datetime import datetime
import common_funcs_lib as cfl

#fetch the index of the meeting provided as argument
try:
    indexOfMeeting = int(sys.argv[1])
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
                startRecording(pathToBandicam)
                print("Initializing recording... \nMight take 10 to 15 seconds...")
                time.sleep(10)
                #start the keep recording function to keep recording even after 10mins
                keepRecording(stopRecTime, pathToBandicam)
                #the stop recrding function is integrated inside keep recording function
                print("Recording...")
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
    

#function to start the schedule
def startMeeting(indexOfMeeting):

    #clearing screen
    os.system('cls')

    cfl.bannerDisp("MEETING IN PROGRESS")
    print("\n")
    
    global database

    try:
        currentMeeting = database[indexOfMeeting]
    except Exception as e:
        print("Error: {}".format(e))
        print("Exiting in 3secs...")
        time.sleep(3)
        sys.exit()


    #if there is any older scheduled meeting than the current time (except the current index) stop and delete it with logs - kill keep alive functions


    #starting the recording if enabled
    recordingFunction(currentMeeting["enable_recording"], currentMeeting["stop_rec_time"])
    


    #check for zoom
        #start the meeting
        #if reconnect on
            #start the reconnecting feature and pass the end time

    #start the screenshot timer

    #details about how to stop the function


startMeeting(indexOfMeeting)
