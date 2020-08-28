from datetime import datetime
from urllib import parse
import threading
import subprocess
import os


def startMeeting():
    pathToAppData = os.getenv('APPDATA')
    absPathToZoomBin = pathToAppData + "\\Zoom\\bin\\Zoom.exe"
    argToPass = '--url="'+zoommtgURL+'"'
    fullCommand = absPathToZoomBin+" "+argToPass
    ONULL = open(os.devnull, 'w')
    subprocess.call(fullCommand, stdout=ONULL, stderr=ONULL, shell=False)

#User Input
meetingURL = input("Enter the full meeting URL here: ").strip()
scheduledAt = input("Enter the date & time when the meeting is scheduled in DD-MM-YYYY HH:MM format (Ex: 28-08-2020 21:00): ").strip()
joinAs = input("Enter a name to join the meeting as: ").strip()
parsed = parse.urlsplit(meetingURL)

hashedMeetingPwd = parse.parse_qs(parsed.query)['pwd'][0]
meetingID = parsed.path.split("/")[-1]

#generating the new url which has zoommtg protocol (will be used to pass as argument to zoom.exe)
zoommtgURL = "zoommtg://us04web.zoom.us/join?action=join&confno="+meetingID+"&pwd="+hashedMeetingPwd+"&uname="+joinAs




#Parsing the date from str to datetime object
scheduledAt = datetime.strptime(scheduledAt, '%d-%m-%Y %H:%M')
#scheduling the meeting
now = datetime.now()
delayInSecs = (scheduledAt - now).total_seconds()
threading.Timer(delayInSecs, startMeeting).start()
