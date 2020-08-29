from datetime import datetime
from urllib import parse
import threading
import subprocess
import os
import winreg
from pyautogui import screenshot


#function to join the meeting
def startMeeting():
    #to start recording
    if(recordOption == 'y' or recordOption == 'Y'):
        startRecording()
        keepRecording()
    #forming the abs path to zoom bin
    pathToAppData = os.getenv('APPDATA')
    absPathToZoomBin = pathToAppData + "\\Zoom\\bin\\Zoom.exe"
    #forming the arg to pass to zoom bin
    argToPass = '--url="'+zoommtgURL+'"'
    #merging the abs path and arg to form command
    commandToJoinMeeting = absPathToZoomBin+" "+argToPass
    #supressing the output
    ONULL = open(os.devnull, 'w')
    #calling the command
    subprocess.Popen(commandToJoinMeeting, stdout=ONULL, stderr=ONULL, shell=False)
    #screenshot will be take 60secs after joining meeting
    if(screenshotOption == 'y' or screenshotOption == 'Y'):
        threading.Timer(120.0, takeScreenshot).start()
    

#function to record the meeting using bandicam
def startRecording():
    #finding the installation path of bandicam through registry
    try:
        progPath =  findBandicamPath()
    except:
        print("\nBandicam is not installed in your computer. Please install Bandicam and re-try recording option.")
    #forming the full command to record using bandicam
    commandToStartRecording = progPath + " /record"
    #supressing the output
    ONULL = open(os.devnull, 'w')
    #calling the command to record
    subprocess.Popen(commandToStartRecording, stdout=ONULL, stderr=ONULL, shell=True)

def findBandicamPath():
    #finding the installation path of bandicam through registry
    storedKey = winreg.OpenKey(winreg.HKEY_CURRENT_USER, "SOFTWARE\\BANDISOFT\\BANDICAM")
    bandiPath =  winreg.QueryValueEx(storedKey, "ProgramPath")[0]
    return bandiPath

def keepRecording():
    if(checkProcRunning("bdcam.exe")):
        startRecording()
    threading.Timer(5.0, keepRecording).start()

def checkProcRunning(procName):
    try:
        callTasklist = 'TASKLIST', '/FI', 'imagename eq %s' % procName
        result = subprocess.check_output(callTasklist).decode()
        # checking end of string for proc name
        lastLine = result.strip().split('\r\n')[-1]
        # because Fail message could be translated
    except:
        print("Process checking failed. Recording might stop.")
    else:
        return lastLine.lower().startswith(procName.lower())
    

def takeScreenshot():
    try:
        meetingScreenShot = screenshot()
        pathToDekstop = os.path.join(os.environ["HOMEPATH"], "Desktop")
        imageName = "Meeting-"+ str(datetime.strftime(datetime.now(), '%d-%b-%Y-%H-%M-%S')) +".png"
        meetingScreenShot.save(pathToDekstop+"\\"+imageName)
    except:
        print("\nScreenshot failed.")
    else:
        print("\nScreenshot saved to Desktop. Filename: "+imageName)
        
        
while 1:
    #meeting URL input
    meetingURL = input("\nEnter the full zoom meeting URL here: ").strip()

    #parsing the URL into componenets
    parsed = parse.urlsplit(meetingURL)
    #not a zoom meeting
    if "zoom.us" not in parsed.netloc:
        print("\nInvalid URL input. Your meeting URL should look like: https://us04web.zoom.us/j/XXXXXXXXXX?pwd=XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX. Try again.")
        continue
    zoomServer = parsed.netloc
    
    #checking the validity of input url
    hashedMeetingPwd = parse.parse_qs(parsed.query)['pwd'][0]
    if(len(hashedMeetingPwd) != 32):
        print("\nInvalid pwd part of URL. Your meeting URL should look like: https://us04web.zoom.us/j/XXXXXXXXXX?pwd=XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX. Try again.")
        continue

    meetingID = parsed.path.split("/")[-1]
    if(not(meetingID.isdecimal())):
        print("\nInvalid ID part of URL. Your meeting URL should look like: https://us04web.zoom.us/j/XXXXXXXXXX?pwd=XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX. Try again.")
        continue
    break

while 1:
    #scheduling date-time input
    scheduledAt = input("\nEnter the date & time when the meeting is scheduled in DD-MM-YYYY HH:MM format (Ex: 28-08-2020 21:00): ").strip()

    #Parsing the date from str to datetime object & error-handling
    try: 
        scheduledAt = datetime.strptime(scheduledAt, '%d-%m-%Y %H:%M')
    except:
        print("\nInvalid date-time input. It should be in DD-MM-YYYY HH:MM format. Try again.")
        continue
    
    #checking the validity of input date-time
    now = datetime.now()
    if(now > scheduledAt):
        print("\nInvalid Date & Time input. Scheduled date-time can't be lower than current date-time. Try again.")
        continue   
    break

#generating the new url which has zoommtg protocol (will be used to pass as argument to zoom.exe)
zoommtgURL = "zoommtg://" + zoomServer + "/join?action=join&confno=" + meetingID + "&pwd=" + hashedMeetingPwd

#username Input
joinAs = input("\nEnter a name to join the meeting as (Optional - if you're already logged in Zoom): ").strip()
#if username is not empty
if(len(joinAs) > 0):
    #url encoding the username
    joinAs = parse.quote(joinAs)
    #adding it to zoommtgURL
    zoommtgURL = zoommtgURL + "&uname=" + joinAs


#screenshot option input
screenshotOption = input("\nDo you want to take a screenshot of your meeting? [y/N]").strip()
if(screenshotOption == 'y' or screenshotOption == 'Y'):
    print("\nScreenshot will be taken 2 minutes after joining the meeting and saved to Desktop")

while 1:    
    #recording option input
    recordOption = input("\nDo you want to record your meeting (Bandicam required) ? [y/N]").strip()
    if(recordOption == 'y' or recordOption == 'Y'):
        try:
            progPath =  findBandicamPath()
        except:
            print("\nBandicam is not installed in your computer. Please install Bandicam and re-try recording option.")
            continue
    break



#scheduling the meeting
delayInSecs = (scheduledAt - now).total_seconds()
threading.Timer(delayInSecs, startMeeting).start()
