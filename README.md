# The Scripts
A collection of Batch and PowerShell scripts used for various purposes.

## TechSetup.ps1
An in development and updated version of `ApplicationSetup.bat`
### Function
* Creates a logging form and shows on the screen.
* Verifies administration privileges.
    * If not running as administrator, log and exit
* Sets the Execution Policy.
* Configures, starts and installs Windows Updates (Check and list security and optional updates and force install all updates)
* Install Chocolatey, an application manager which installs TeamViewer, Adobe Reader, and Google Chrome.

<hr>

## DownloadApplicationPackage.bat
### Function
Lists all of the below packages that you can download individually or as a group/bundle together for efficiency (e.g. Download ApplicationSetup.bat and PersonaliseComputer.bat for running up a computer.)

<hr>

## ApplicationSetup.bat
### Function
A script used to automatically personalise a computer to deploy for clients.

### Features
* Install TeamViewer
* Install Adobe Reader
* Install Google Chrome

<hr>

## PersonaliseComputer.bat
### Function
A script used to automatically personalise a computer to deploy for clients.

### Features
* Hides Cortana Button.
* Hides Task View Button.
* Hides People Button.
* Removes News and Interests from Taskbar.
* Clear all Taskbar items.
* Adds This PC icon to Desktop.

### To Do
* Removes all Start Menu Icons/Tiles.
* Add User's Files to Desktop.

<hr>

## RestartPrintSpoolerService.bat
### Function
A script used to restart the Printer Spool Service if it gets stuck.