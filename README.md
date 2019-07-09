# Saverr
Powershell script with a front-end GUI to download media from Plex Servers.  
Allows you to search and save movies, tv shows, and music for offline/personal use.  
Expanded from the original non-gui version: [directplexDownloader](https://github.com/ninthwalker/directPlexDownloader).

## Details/Requirements
1. Windows 7/8/10
2. Powershell 3.0+
3 .Net Framework 3.5+ (Usually already on your windows computer)
3. Plex Server (Can be used on your own server or a remote friends server as well)
  
Supports:
* **Movies, TV Shows, and Music:** Download any and all media. Full TV Series or Albulms as well!
* **Searching:** Enter the name of what you are looking for to watch what you want, when you want.
* **Multiple Server support:** Quickly switch between multiple plex servers.
* **Size/ETA Estimates:** Shows current progress, size of download as well as an estimated time of finishing.
* **Pause/Resume:** Supports pausing and resuming of downloads.

## Launching Saverr

**Option One**
1. Copy the saverr.ink shortcut, as well as the saverr.ps1 to your computer. Place in the same directory.  
2. Double click the saverr shortcut to launch the app.  
3. Click the gear icon to configure initial settings. (See settings section below)  

**Option Two**
1. Copy the saverr.ps1 file to your desktop.
2. Open a powershell console and navigate to the folder you saved the saverr.ps1 file.
3. Enter the below command to temporarily set the execution policy:  
  `Set-ExecutionPolicy -Scope Process Bypass`  
  Alternatively, set the execution policy to permanently allow powershell scripts:  
  `Set-ExecutionPolicy -Scope Currentuser Unrestricted`  
4. Then enter the below command to launch the app:  
  `.\saverr.ps1`  

**Option Three**
1. If you trust me, I've compiled an .exe here as well.  
  source is just the saverr.ps1 file that is located here converted with the [PS2EXE](https://gallery.technet.microsoft.com/scriptcenter/PS2EXE-GUI-Convert-e7cb69d5) tool. You can do it yourself if you want using the saverr.ps1 file.
2. Save the saverr.exe to your computer and double click it to launch the app.  

## How to use  
1. Launch the script using one of the 3 methods listed above.  
2. Enter the name of the Movie, TV show or Music artist to search for.  
3. Select the desired result from the results box.  
  3a. If a Movie: Just click download.  
  3b. If a TV Show, select the season or episodes, then click download. (Can also select All seasons or All episodes)  
  3c. If Music, select the album or tack, then click download. (Can also select All albums or All tracks)  

## Config/Settings  
All config is done in the settings menu (Accessed by clicking the orange geer icon)  

TODO  

## Known Issues    

1. Some plex servers will not work or may act a bit funky.  
Not much I can do about this. Everyones plex servers are set up a bit differently. Some have firewalls, special routing configuration, or nginx proxy servers in front. However, from my own use using mutiple plex servers, this works the majority of the time and the servers that do not work are outliers.  

2. If trying to connect to a local Plex server, selecting it in the settings page may not work.  
This is because of routing/rebind protection settings at your house. To get around this issue, please modify the saverrSettings.xml file manually and enter in your local IPaddress:port in the server section. ie: 192.168.1.22:32400  

3. Minimizing the Saverr app window requires pausing of downloads.  
This is because of the powershell script download functions in the background. A pause is required to prevent hanging of the script. Minimizing will warn you that it will pause downloads. Click resume after maximizing the window again. You can also continue to do other things on your computer by opening other windows. The Saverr app will remain open behind these windows and can be re-selected again to use it without pausing of downloads.

4. Maximum number of downloads is 200.  
This is the default setting within windows for the BITS download function being used.  
More than 200 items will automatically be truncated to 200.  
If you need to increase the maximum download amount, please set this registry setting below to the value desired:  
  `HKLM\Software\Policies\Microsoft\Windows\BITS` `MaxFilesPerJob`  
Reference: [Bits](https://docs.microsoft.com/en-us/windows/desktop/bits/group-policies)


## Errors  
Some errors are self explanatory and output to the main app window, others are not. You can enable debugging in the settings menu. This will create a log file in the current saverr directory that will give more information on the error or issue.  

## Screenshots  

TODO  

