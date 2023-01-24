# Saverr

## Developers Note:  
This script/repo is no longer maintained.  
While it still works as far as i know, I'm not adding/changing to it any longer. I would suggest to check out a project called [PlexRipper](https://github.com/PlexRipper/PlexRipper) if you are interested in something like my Saverr, but with a lot better GUI and feature set. Thanks!


Powershell script with a front-end GUI to download media from Plex Servers.  
Allows you to search and save movies, tv shows, and music for offline/personal use.  
Expanded from the original non-gui version: [directplexDownloader](https://github.com/ninthwalker/directPlexDownloader).

## Details/Requirements
1. Windows 7/8/10
2. Powershell 3.0+
3. .Net Framework 3.5+ (Usually already on your windows computer)
4. Plex Server (Can be used on your own server or a remote friends server as well)
  
Supports:
* **Movies, TV Shows, and Music:** Download any and all media. Full TV Series or Albulms as well.
* **Searching:** Enter the name of what you are looking for to watch what you want, when you want.
* **Multiple Server support:** Quickly switch between multiple plex servers.
* **Size/ETA Estimates:** Shows current progress, size of download as well as an estimated time of finishing.
* **Pause/Resume:** Supports pausing and resuming of downloads.

## Launching Saverr

**Note:**
Depending on your download method/settings, you may need to 'unblock' the files after downloading them from Github. This is normal behavior for Microsoft Windows to do for files downloaded from the internet.  
`Right click > Properties > Check 'Unblock'`

**Option One**
1. Copy the Saverr.lnk shortcut, as well as the Saverr.ps1 file to your computer. Place in the same directory.
2. Double click the Saverr shortcut to launch the app.
3. Click the gear icon to configure initial settings. (See settings section below)  

**Option Two**
1. Copy the Saverr.ps1 file to your desktop.
2. Open a powershell console (Not ISE) and navigate to the folder you saved the Saverr.ps1 file.
3. Enter the below command to temporarily set the execution policy:  
  `Set-ExecutionPolicy -Scope Process Bypass`  
  Alternatively, set the execution policy to permanently allow powershell scripts:  
  `Set-ExecutionPolicy -Scope Currentuser Unrestricted`  
4. Then enter the below command to launch the app:  
  `.\Saverr.ps1`  

**Option Three**  
  
Not available anymore. Even though the .exe was just the powershell script compiled, windows would flag it as a virus so I removed this option. You can compile it yourself into an executable using this tool here if you would like:  
(https://gallery.technet.microsoft.com/scriptcenter/PS2EXE-GUI-Convert-e7cb69d5)
~~If you trust me, I've compiled an .exe here as well.~~
  ~~source is just the Saverr.ps1 file that is located here converted with the [PS2EXE] tool. You can do it yourself if you want using the Saverr.ps1 file.~~
~~Save the Saverr.exe to your computer and double click it to launch the app.~~

## How to use  
1. Launch the script using one of the methods listed above.  
2. Enter the name of the Movie, TV show or Music artist to search for.  
3. Select the desired result from the results box.  
  3a. If a Movie: Just click download.  
  3b. If a TV Show, select the season or episodes, then click download. (Can also select All seasons or All episodes)  
  3c. If Music, select the album or tack, then click download. (Can also select All albums or All tracks)  

## Config/Settings  
All config is done in the settings menu (Accessed by clicking the orange gear icon)  

* **Plex Username:** Plex Username. Can be an email or a username. (this is not saved)  
* **Plex Password:** Your Plex password. (this is not saved, only used to retrieve your plex token)  
Note for 2FA: If 2FA is enabled, please add your 2FA token to the end of your password. ie: MyP@ssword347821
* **Server:** After a token has been saved from entering a username/password you can select a Plex server to search from.  
* **Download Path:** Navigate and select a path to save downloads to.
* **SSL Required:** Select this if the server you are trying to connect to has 'Secure Connections: Required' enabled.
* **Debug:** Check to enable a log file for debugging issues. (Saves to saverrLog.txt)

## Known Issues    

1. Some plex servers will not work or may act a bit funky.  
Not much I can do about this. Everyones plex servers are set up a bit differently. Some have firewalls, special routing configuration, or nginx proxy servers in front. However, from my experience using mutiple plex servers, this works the majority of the time and the servers that do not work are outliers.

3. No mimimize button during an active download.  
This is intended. Because of the powershell script functions used in the background, a pause is required to prevent hanging of the script. The mimimize button becomes available after you pause a download. Click resume after maximizing the window again. You can also continue to do other things on your computer by opening other windows while the Saverr app downloads. Saverr will remain open behind these windows and can be re-selected again to use it without pausing of downloads.

4. Maximum number of downloads is 200.  
This is the default setting within windows for the BITS download function being used.  
More than 200 items will automatically be truncated to 200.  
If you need to increase the maximum download amount, please set this registry setting below to the value desired:  
  **Path:** `HKLM\Software\Policies\Microsoft\Windows\BITS`  
  **Dword:** `MaxFilesPerJob`  
  **Decimal Value:** `Dealers Choice`  
Reference: [Bits](https://docs.microsoft.com/en-us/windows/desktop/bits/group-policies)


## Errors  
Some errors are self explanatory and output to the main app window, others are not. You can enable debugging in the settings menu. This will create a log file (saverrLog.txt) in the current Saverr directory that will give more information on the error or issue.  

## Screenshots  

![](https://raw.githubusercontent.com/ninthwalker/saverr/master/screenshots/Saverr%20-%20Movie%20Search.png)  

![](https://raw.githubusercontent.com/ninthwalker/saverr/master/screenshots/Saverr%20-%20TV%20Search.png)  

![](https://raw.githubusercontent.com/ninthwalker/saverr/master/screenshots/Saverr%20-%20Music.png)  

![](https://raw.githubusercontent.com/ninthwalker/saverr/master/screenshots/Saverr%20-%20Downloading.png)  

![](https://raw.githubusercontent.com/ninthwalker/saverr/master/screenshots/Saverr%20-%20Settings.png)
