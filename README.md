# zoom-meeting-auto-join-scheduler

Join & record scheduled zoom meetings automatically.

Features:
  1. Auto join scheduled Zoom Meetings.
  2. Record the meetings.
  3. Take screenshot of the meeting.
  4. Auto re-join the meeting if you get diconnected before a spefied time (Useful if you have poor network connection or your meetings occur part-wise in consecutive sessions) 

Requirements:
  1. Only for Windows 10 & Windows 7
  2. Python 3.0 or higher should be installed in your computer.
  3. Zoom Meetings should be installed in your computer.

Installation instructions:
  1. Download and extract this repository
  2. Open Command Prompt as administrator.
  3. Change current directory to the extracted path (Using command: cd "path to direcotry in which you extracted the files")
  4. Install the python modules listed in requirements.txt (Using command: pip install -r requirements.txt )
  5. Run the zoom-auto-join-scheduler.py with Python (Using command: python zoom-auto-join-scheduler.py)

Note:
  1. To automatically join with audio during zoom meetings, enable "Automatically join audio by computer when joining a meeting" option in "Settings -> Audio" of Zoom Meetings app.
  2. To automatically mute your mic on joining, enable "Mute my microphone when joining a meeting" option in "Settings -> Audio" of Zoom Meetings app.
  3. To automatically disable your camera on joining, enable "Turn off my video when joining a meeting" option in  "Settings -> Video" of Zoom Meetings app.
  4. Closing "zoom-auto-join-scheduler.py" or rebooting your computer after scheduling a meeting won't effect the automation process. The scheduled meeting will still start if your computer is on during the scheduled time.
  5. To terminate a scheduled meeting, run "zoom-auto-join-scheduler.py" and Delete the scheduled meeting from the menu.

Precautions: 
  1. If UAC is enabled in your pc, make sure it isn't blocking the script.
  2. Make sure you're connected to the Internet when the meeting is scheduled to run.
  3. Do not move the script directory after you've scheduled a meeting, then the prviously scheduled meeting won't run. 
