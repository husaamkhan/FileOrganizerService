# FileOrganizerService

Author: Husaam Khan
Email: husaamkhan04@gmail.com
Date: 06/14/2024

## Purpose:
A windows service written in python that organizes your Desktop and Downloads files automatically.

## Dependencies:
OS: **Windows 11**
Libraries:
        - os
        - shutil
        - watchdog
        - pywin32 (win32serviceutil, win32service, win32event, win32evtlogutil, win32evtlog)
        - servicemanager
        - sys
        - wmi

## Installation:
- Ensure the listed dependencies are installed correctly.
- Open Windows Powershell as administrator and navigate to the directory script.py is in.
- To install the service, enter the command: ```python script.py install```

## Executing program:
- To start the service, enter the command: ```python script.py start```
- To stop the service, enter the command: ```python script.py stop```

## Setting up automatic startup:
- Install service
- Open Services and find FileOrganizerService in the list of services.
- Right click on FileOrganizerService and select Properties.
- In the General tab of the Properties menu, set the Startup type to Automatic.

## Files:
- README.md
- script.py

## Acknowledgements:
- [HaroldMills/Python-Windows-Service-Example](https://github.com/HaroldMills/Python-Windows-Service-Example)
