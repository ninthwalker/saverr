# Saverr
Powershell script with a frontend GUI to download media from a Plex Server.  
Allows you to search and save movies, tv shows, and music for offline/personal use.  
Expanded from the original non-gui version [directplexDownloader](https://github.com/ninthwalker/directPlexDownloader)

## Details/Requirements
Requires access to an existing plex server. Can be used on your own server or a remote friends server as well.  
  
Supports:
* **Movies, TV Shows, and Music:** Download any and all media. Full TV Series or Albulms as well!
* **Searching:** Enter the name of what you are looking for to watch what you want, when you want.
* **Multiple Server support:** Quickly switch between multiple plex servers.
* **Size/ETA Estimates:** Shows current progress, size of download as well as an estimated time of finishing.
* **Pause/Resume:** Supports pausing and resuming of downloads.

## How to use

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

## Config/Settings  
All config is done in the settings menu (Accessed by clicking the orange geer icon)  

TODO  


## How to use  
1. Launch the script using one of the 3 methods listed above in the setup section.  
2. Enter the name of the Movie, TV show or Music artist to search for.  
3. Select the desired result from the results box.  
  3a. If a Movie: Just click download.  
  3b. If a TV Show, select the season or episodes, then click download. (Can also select All seasons or All episodes)  
  3c. If Music, select the album or tack, then click download. (Can also select All albums or All tracks)  

## Known Issues    

TODO  

## Errors  
Some errors are self explanatory, others are not. You can enable debugging the settings menu.  
This will create a log file in the current saverr directory that will give more information on the error or issue.  

## Screenshots  

TODO  

