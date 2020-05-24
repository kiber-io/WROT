# WROT
Lightweight **Wrapper for Removing Office Tool**. Only for Windows. Requires VBScript support in system.

This only wrapper for Microsoft scripts: [Official Microsoft Repository (scripts sources)](https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/tree/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls).

This scripts completely remove ALL MS Office installed products and their files.

Very usefull when you cannot install new version of Office because of error **"You have already installed Office"**

![gif](https://github.com/kiber-io/WROT/blob/master/gif/wrot.gif?raw=true)

### Using WROT with python
```
python WROT.py
```
The script will request an elevate by using UAC and open in a separate window.

### Using WROT without python installed
You can use WROT without Python. Just one compile Python sources to *.exe file using [pyinstaller](https://www.pyinstaller.org/)
```
python -m pip install pyinstaller
pyinstaller -F WROT.py
```
Or just download already compiled WROT.exe from [RELEASES](https://github.com/kiber-io/WROT/releases/).

# Commands
![main_window](https://github.com/kiber-io/WROT/blob/master/screenshots/wrot_main.png?raw=true)
### del3, del7, del10, del13, del16, delc2r
Marks office products that match the commands for deletion.
  * del3 - Microsoft Office 2003
  * del7 - Microsoft Office 2007
  * del10 - Microsoft Office 2010
  * del13 - Microsoft Office 2013
  * del16 - Microsoft Office 2016
  * delc2r - Microsoft Office Click-To-Run (C2R)
  

### installed
 Try to find installed MS Office versions and show it. Also, if possible, tries to show the commands required to delete the founded products.
![wrot_installed](https://github.com/kiber-io/WROT/blob/master/screenshots/wrot_installed.png?raw=true)
### show
Show products what you marked for deletion.
![wrot_show](https://github.com/kiber-io/WROT/blob/master/screenshots/wrot_show.png?raw=true)
### clear (cls)
Delete from deletion list all products what you marked.
### force64 (force32)
Forcing script to use 64/32-bits cscript.exe
### start (run)
Start deletion process. 
![wrot_deletion](https://github.com/kiber-io/WROT/blob/master/screenshots/wrot_deletion.png?raw=true)
### start (run) parallel
Starts all deletion processes in parallel.

***ATTENTION!***
Due to the fact that the all deletion processes will be parallel, the load on the system increases and there may be lags and computer freezes.
![wrot_deletion_parallel](https://github.com/kiber-io/WROT/blob/master/screenshots/wrot_deletion_parallel.png?raw=true)
