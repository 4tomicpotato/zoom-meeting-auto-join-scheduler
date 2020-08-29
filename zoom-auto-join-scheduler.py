from datetime import datetime
from urllib import parse
import threading
import subprocess
import os
import winreg

#function to join the meeting
def startMeeting():
    if(recordOption == 'y' or recordOption == 'Y'):
        startRecording()
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
    #to start recording

#function to record the meeting using bandicam
def startRecording():
    #finding the installation path of bandicam through registry
    storedKey = winreg.OpenKey(winreg.HKEY_CURRENT_USER, "SOFTWARE\\BANDISOFT\\BANDICAM")
    progPath =  winreg.QueryValueEx(storedKey, "ProgramPath")[0]
    #forming the full command to record using bandicam
    commandToStartRecording = progPath + " /record"
    #supressing the output
    ONULL = open(os.devnull, 'w')
    #calling the command to record
    subprocess.Popen(commandToStartRecording, stdout=ONULL, stderr=ONULL, shell=True)
    
while 1:
    #user Input
    meetingURL = input("Enter the full zoom meeting URL here: ").strip()

    #parsing the URL into componenets
    parsed = parse.urlsplit(meetingURL)
    #not a zoom meeting
    if "zoom.us" not in parsed.netloc:
        print("Invalid URL input. Your meeting URL should look like: https://us04web.zoom.us/j/XXXXXXXXXX?pwd=XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX. Try again.")
        continue
    
    #checking the validity of input url
    hashedMeetingPwd = parse.parse_qs(parsed.query)['pwd'][0]
    if(len(hashedMeetingPwd) != 32):
        print("Invalid pwd part of URL. Your meeting URL should look like: https://us04web.zoom.us/j/XXXXXXXXXX?pwd=XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX. Try again.")
        continue

    meetingID = parsed.path.split("/")[-1]
    if(not(meetingID.isdecimal())):
        print("Invalid ID part of URL. Your meeting URL should look like: https://us04web.zoom.us/j/XXXXXXXXXX?pwd=XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX. Try again.")
        continue

    break

while 1:
    #user Input
    scheduledAt = input("Enter the date & time when the meeting is scheduled in DD-MM-YYYY HH:MM format (Ex: 28-08-2020 21:00): ").strip()

    #Parsing the date from str to datetime object & error-handling
    try: 
        scheduledAt = datetime.strptime(scheduledAt, '%d-%m-%Y %H:%M')
    except:
        print("Invalid date-time input. It should be in DD-MM-YYYY HH:MM format. Try again.")
        continue
    
    #checking the validity of input date-time
    now = datetime.now()
    if(now > scheduledAt):
        print("Invalid Date & Time input. Scheduled date-time can't be lower than current date-time. Try again.")
        continue
    break


#user Input
joinAs = input("Enter a name to join the meeting as: ").strip()
#user Input
recordOption = input("Do you want to record your meeting (Bandicam required) ? [y/N]").strip()

#generating the new url which has zoommtg protocol (will be used to pass as argument to zoom.exe)
zoommtgURL = "zoommtg://us04web.zoom.us/join?action=join&confno="+meetingID+"&pwd="+hashedMeetingPwd+"&uname="+joinAs

#scheduling the meeting
delayInSecs = (scheduledAt - now).total_seconds()
threading.Timer(delayInSecs, startMeeting).start()
