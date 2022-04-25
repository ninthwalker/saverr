#######################################
# Name:    Saverr                     #
# Desc:    d/l media from Plex        #
# Author:  Ninthwalker                #
# Date:    16NOV2021                  #
# Version: 1.1.1                      #
#######################################


###### NOTES FOR USER #######

# See Instruction online at: https://github.com/ninthwalker/saverr

# Execution policy may need to be set to run powershell scripts if not using the shortcut from Github
# ie: Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# Enforce TLS 1.1/1.2 if wanting. Uses HTTPS/SSL by default to retrieve plex tokens. May/may not break functionality depending on network setup.
#[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls

# If you need to increase the maximum downloads, please set this registry setting below to value desired.
# Bitstransfers default limit is 200. see: https://docs.microsoft.com/en-us/windows/desktop/bits/group-policies
# Path: HKLM\Software\Policies\Microsoft\Windows\BITS
# Dword: MaxFilesPerJob
# Decimal Value: Dealers Choice


#############################
####### DO NOT MODIFY #######
#############################

### setup environment ###

if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript") {
    $ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
}
else {
    $ScriptPath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0])
    if (!$ScriptPath) {
        $ScriptPath = "."
    }
}

# Set-Location $PSScriptRoot. Changed to above so it works with an .exe as well.
Set-Location $ScriptPath

#import Bitstransfer if not
if (!(Get-Module BitsTransfer)) {
    Import-Module BitsTransfer
}

# Set timeout value for how long to wait for download to start before giving up
$timeout = 30

# Get maximum BITS files value if set
$bitsRegistry =  'HKLM:\Software\Policies\Microsoft\Windows\BITS' 
$key = Get-Item -LiteralPath $bitsRegistry -ErrorAction SilentlyContinue
if ($key) {
    $limit = $key.GetValue("MaxFilesPerJob", 200)
}

# diff ways of d/l. invoke-restmethod seems to fail sometimes while webclient method does not.
# Set-Alias -Name plx -Value Invoke-RestMethod -Scope Script

# download function shortcut
function plx {
    
    Param([Parameter(Mandatory=$true)]
    [string]$url
    )

    # Will timeout after 20sec by default
    [xml](New-Object System.Net.WebClient).DownloadString($url)
}

# check invalid char's function
Function Remove-InvalidChars {

    Param([Parameter(Mandatory=$true)]
    [string]$name
    )

  $strip = $name -replace('[][]','')
  $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
  $re = "[{0}]" -f [RegEx]::Escape($invalidChars)
  return ($strip -replace $re)
}

# download path function
Function Get-SavePath($initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select a folder to save downloads to"
    $foldername.rootfolder = "MyComputer"

    if($foldername.ShowDialog() -eq "OK")
    {
        $folder += $foldername.SelectedPath
    }
    else {
        $folder = $false
    }
    return $folder
}

# logging function
function logIt {
    if ($debug) {
        $e = $_.Exception
        $line = $_.InvocationInfo.ScriptLineNumber
        $msg = $e.Message
	
        $eMSG = "$(Get-Date): caught exception: $e at $line. $msg"
        $eMSG | Out-File ".\saverrLog.txt" -Append
    }
}

# display size function
function byteSize($num)
{
    $suffix = "B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB"
    $index = 0
    while ($num -gt 1kb) 
    {
        $num = $num / 1kb
        $index++
    } 

    "{0:N1} {1}" -f $num, $suffix[$index]
}

# Import settings
if (Test-Path .\saverrSettings.xml) {
    $script:settings = Import-Clixml .\saverrSettings.xml
    if ((!($settings.name)) -or (!($settings.server)) -or (!($settings.userToken)) -or (!($settings.serverToken)) -or (!($settings.dlPath))) {
        $errorMsg = "Settings are not fully configured.`nPlease click the gear icon before searching."
    }
    else {
        $errorMsg = ""
    }
}
else {
    $errorMsg = "Settings file not detected. Please configure settings before searching."
}

# enable/disable debug
$debug = $settings.logging

# enable/disable ssl
$ssl = $settings.ssl

# The below is needed when the plex server has 'Secure connections: required' set. 
# when SSL is enforced, and the 'SSL Required' is checked on the Saverr settings page we will use HTTPS.
# However, because we have to access by IP, the cert will show as 'invalid' since the CN will not match the IP.
# These settings here, allow us to download from servers that enforce the SSL.

# also, later on we do a similar 'ignore cert errors' for the bitstransfer job
# using the below command/info:
# https://docs.microsoft.com/en-us/windows-server/administration/windows-commands/bitsadmin-setsecurityflags
# https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc753211(v=ws.10)?redirectedfrom=MSDN
# bitsadmin /SetSecurityFlags myJob 30

if ($ssl -eq $True) {
	$scheme = "https://"
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
}
else {
	$scheme = "http://"
}

# Plex signin for token url
$plexSignInUrl = "https://plex.tv/users/sign_in.xml"

# Plex servers list URL
$plexServersUrl = "https://plex.tv/pms/servers"

# init the cancel/pauseLoop variables
$script:cancelLoop = $false
$script:pauseLoop = $false


### Load required libraries ###

Add-Type -AssemblyName System.Windows.Forms, PresentationFramework, PresentationCore, WindowsBase, System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

################ Images ##################


# load these images inside the script so external calls arents
$plexImg = @'
    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
'@

$loadingImg = @'
    iVBORw0KGgoAAAANSUhEUgAAAHEAAACmCAYAAADtRWBHAAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwgAADsIBFShKgAAAABl0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMC4xMkM
    Ea+wAABC7SURBVHhe7Z0LlFTFncZ5DDCDAiIg4mOAYYZ5dE8/pqe7ZwbQAA4rIIxCQIYguCoqKL5FEZEEBBQTWRE0AopgwPg6RvNAjTGaXTdns48TyZo1iWvM6mbduBqjJhLpqt7v33
    RP/tVTPV0TWovb1HfO79z57nTfrqqv696qupehRzKZdHgc7U6Ht9DudHgL7U6Ht9DudHgL7U6Ht9DudHgL7U6Ht9DudHgL7U6Ht9DudHgL7U6Ht9DuzIeTfSl5cGOKk30peXBjipN9K
    XlwY4qTfSl5cGOKk30peXBjipN9KXlwY4qTfSl5cGOKk30peXBjipN9KXlwY4qTfSl5cGOKk30peXBjipN9KXlwY4qTfSl5cGOKk30peXBjipN9KXlwY4qTfSl5cGOKk30peXBjipN9
    KXlwY4qTfSl5cGOKk30peXBjipN9KXlwY4qTfSl5cGOKk30peXBjipN9KXlwY4qTfSl5cGOKk30peXBjipN9KXlwY4qTfSl5cGOKk30peXBjipN9KXlwY4qTfSl5cGOKk30peXBjipN
    9KXlwY4qTfSl5cGOKk30peXBjilNn+Xy+njU1NQOwDfv9/i/X1dW1V1VVDUr/uuBS8uDGlKNdCKoXGAhGgji4EfyD3+f7IBjwi5ZonQwF/BJh7sf+cQg3/c7CScmDG1OORiGQEvSu4R
    QKmA9/J7b7sf2YgpsQr5VfmjpGLm8/VT547dDEhFitwO/+jNdsQIh90ocpmJQ8uDHlaFF9fX0fhFADFiGQ1djuwvY18El9vV9ScEtnjRR3XnKieOym48V/bC0TH+3tLV+6bWAiHPQl8
    TrqjTtra2uPTx+yYFLy4MaUYhYavRfCG4ntxWA3eBn8Dnxaj1Ci4TqxYGpFYvtVw8SL6wfKN7aVyg/3lMhPHuktDzzSO0nbC9tGCRwjCSjEp9GDR6UPXzApeXBjSrEpHA73RmNXobGv
    x/af0fhvYPt7cBBICmTmxCqx4+phif2b+4u37+8nP9xbQqHJPz/aO8l5c3s/EYvUZkKk3vgKenBL+qMKJiUPbkzxqtCgdIqkbU8E1heNGwPbwG+w71P8LhVYNqGAT77+9VKRHVg2FOr
    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
    gpae6E4FPbBU7Oie+iuKEOk+Id3QPRmMKGJOAUuy6k4UR4hgdiAQKOo78giwF5iXVXfCkyF+LasSLsQCSMmDG1O6IxeiUnfCkyF+NasSLsQCSMmDG1O6IxeiUnfChegVuRCLQC7EIpA
    LsQjkQiwCFVWIqIgLUcWTId6RVQkXYgGk5MGNKd2RC1GpO+FC9IpciEWgYgtxY1YlXIgFkJIHN6Z0Ry5Epe6EC9ErciEWgVyIRaBiC/H2rEq4EAsgJQ9uTOmOXIhK3YniCTEUCvWaM2
    dOj2KFvqRFE6LP57stqxL03Ok6bKeDaUUM/dHe7C+wZ/+59wpNRegP2n6In4sWqh9Q/iwaoKffp6eb5rCk5MGNKd0RQqQ/BJv3j9geDaAdXgdD001zWFLy4MaU7giFp2vDFwD9Pxcm/
    0FJ0YF6f4rLyvPY+ugPNBVCSh7cmOJkX0oe3JjiZF9KHtyY4mRfSh7cmOJkX0oe3JjiZF9KHtyY4mRfSh7cmOJkX0oe3JjiZF9KHtyY4mRfSh7cmOJkX0oe3JjiZF9KHtyY4mRfSh7c
    mOJkX0oe3JjiZF9KHtyY4mRfSh7cmOJkX0oe3JjiZF9KHtyY4mRfSh7cmOJkX0oe3JjiZF9KHtyY4mRfSh7cmOJkX0oe3JjiZF9KHtyY4mRfSh7cmOJkX0oe3JjiZF9KHtyY4mRfSh7
    cmOJkX0oe3JjiZF9KHtyY4mRfSh7cmOJkX0oe3JjiZF9KHtyY4mRfSh7cmOJkX0oe3JjiZF9KHtyY4mRfSh7cOLyJdqfDW2h3OryFdqfDW2h3OryFdqfDW2h3OryFdqfDW2h3OryFdq
    fDW2h3OryFdqfDSyR7/D/88mPYQlLgrwAAAABJRU5ErkJggg==
'@


$iconImg = @'
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAQAAMMOAADDDgAAAAAAAAAAAAAtLS0ZLS0tyS0tLbQtLS2gLS0toS0tLaEtLS2hLS0toS0tLaEtLS2hLS0toS0
tLaEtLS2hLS0toC0tLbItLS3qLS0tIS0tLcgtLS0yLCwsAi0tLQUtLS0FLS0tBS0tLQUtLS0FLS0tBS0tLQUtLS0FLS0tBSwsLAItLS0xLS0t6C0tLSEtLS3ILS0tLi0tLQAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAtLS0ALS0tLS0tLegtLS0hLS0tyC0tLS4tLS0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALS0tAC0tLS0tLS3oLS0tI
S0tLcgtLS0uLS0tAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC0tLQAtLS0tLS0t6C0tLSEtLS3JLiopNzQVBggxIRoLLS0sCyMmKwsYHyoLGyEqCykrLAsuKigLMh0T
CzMbEAswIhsILS0tNi0tLektLSwgLS4u1yk5QdAoPUfCKjc9wy0sK8M2My7DQDkvwz03L8MuLi3DKy8ywyg8RsMoPEbDKzM3wi0sK84tLS36LSwsIC0tLtwfWHT/C5LS/w2Myf8iS1//RDw
w/5NtNf+oeTb/b1cy/zA1Nv8XbZX/CZfa/xGAtP8oPEb/LSws/y0tLSEtLCvcLC8w/xpni/QLk9TzEn6w9ig8RfdRRC/3mnE1951zNfdbSjH3KTlA9xR5qPYLktL2GWuS9ysyNfAuKylRKD
tD7yVFVPssMTTEKjY7eydBTk0pNz4/KSosQD04L0JHPjBCQzwwQi8rKUIrMzdCJ0JOQiZDUUMrMzY4LSsrzic+Sf8MkdD/EIO5/ic/S+wuKyq5LS0tbR8kLCYKFCgDERsqAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAC0tLFgsLzCkGWiN4wqW2P0aZor/MC4s/2RQMv5dSzHmOjUuqSwsLVovKCUbOwMAAjYRBQAAAAAAAAAAAAAAAAAvJyMAMCQeBC4pJyYkSVpqIk1htystL+1sVDL/
uYM2/4ZlNP8xLy37KTg+2Co2PJEvJiJCLiglDk0AAAAsLCwAAAAAAAAAAACEAAAAB0lIADUgHwkqKyw0PzkvfnhcM8iPajX1Qzot/yBRaP8QhLv/HV188ysyNcMtLCtALS0tAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAEhAAAAAAAVHSwQQDowRDYxLJAoO0TUDorF+Aee5v8jTGD+LikmYS0tLQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC0tLQAsLS4CLiooHiVFVGMZaI
y9IFNq5C4pJjQtLS0AAAAAAAAAAAAf/AAAH/wAAB/8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB/AAAADwAAgAMAAPABAAD8AQAA/4EAAA==
'@

$settingsImg = @'
iVBORw0KGgoAAAANSUhEUgAAABkAAAAZCAYAAADE6YVjAAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwgAADsIBFShKgAAAABl0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMC4xMkMEa+w
AAALtSURBVEhLtVVPSBRRGJ/V3ZnR3XVllS1rzf4RIWWWVFJ06FZ46Bh08ZS3oGMk1bW6RIHXwpz3ZmZ3dS3QEgLzVJQdoqAo8iIlRaQEIevO+17fPF+5jU93gvrBjx3e9/t+385733xPWw
/lIb2HUd1ixHy14Han5bIAjJ8wGDVfMFsfgpHNJ+VyeECh6TgjsSlOIpwTTZDRukE+3ReTEs1z4qO/Yj6B6pMlq75ThqsDbH0X0NqlSpNlo9hbZsXGGI1+CMZEPJ/pkBbhwJzkrMpoTTpxL
lPDAcYPZRmp/aE0W5MRXnKze6XFapSKbZ18ukvsN4zvNDyiP1IbrU88x0mea9sofN6dMxaLbVv9ZwHsoCnc7xlmRR/g75zKICzRa55R4wnQ6BwWvScKlGmyp7KL/jXBSRzFtzAsVTBIZpu3
IN9yDGyzFfIbDjOn4TqQyKoODJKRuqsas+ufqYKV9OxUn3jtADy3+XS1QtjWgxpMdMSZnRxQCXwyYoxKTyXKxHBVeZzW8LLddErKlsFI9KVKDG7qrJQowXOZPlUejpwBKVkBnk1BJfbcpjN
SooSX392rygOqnZeSFWBl5ajgtn5BSpTwiHlZlYdD9envWccnexs9J/nHsKskHuwsFHe0CnEAy5M4+l6V5xNjFp55Bs/CrNpd2L7PFwtbtktvARg+0IIf3mOVvpLMTb/xt+m2Khgks2q+YX
NQRiI3gNTeBVLzRaUL0rPMK9qSk+z+r198ftMR8er+HQFOfAbwYgJqfFaJwxLsOumjf8JvbEgU8MGLqa0f73fV+88LD9vTuC0jKoNqxIF4k+faE77PfHFfIww3t/jPSoCbTeM0/q4yWovYB
F95ThMFQuNvb0ZWyLyWqeEAww0HmRUprTKi5phnm/3gmBPBGN4fSziCtkmL6ig52T3MSTiVXefZDf0yLIBD9VplEUb0O4t2auUmDAueS3XhP7/EaGICD1SXywL+yPDnHd4vF0s0tV8uK6Bp
PwEAHUcTE+ClVQAAAABJRU5ErkJggg==
'@

# Image logo function
function DecodeBase64Image {
    param ([Parameter(Mandatory=$true)][String]$ImageBase64)
    $ObjBitmapImage = New-Object System.Windows.Media.Imaging.BitmapImage #Provides a specialized BitmapSource that is optimized for loading images using Extensible Application Markup Language (XAML).
    $ObjBitmapImage.BeginInit() #Signals the start of the BitmapImage initialization.
    $ObjBitmapImage.StreamSource = [System.IO.MemoryStream][System.Convert]::FromBase64String($ImageBase64) #Creates a stream whose backing store is memory.
    $ObjBitmapImage.EndInit() #Signals the end of the BitmapImage initialization.
    $ObjBitmapImage.Freeze() #Makes the current object unmodifiable and sets its IsFrozen property to true.
    $ObjBitmapImage
}

$plexImgDecoded = DecodeBase64Image -ImageBase64 $plexImg
$loadingImgDecoded = DecodeBase64Image -ImageBase64 $loadingImg
$settingsImgDecoded = DecodeBase64Image -ImageBase64 $settingsImg
$loading = [System.Drawing.Bitmap][System.Drawing.Image]::FromStream($loadingImgDecoded.StreamSource)
$logo = [System.Drawing.Bitmap][System.Drawing.Image]::FromStream($plexImgDecoded.StreamSource)
$gear = [System.Drawing.Bitmap][System.Drawing.Image]::FromStream($settingsImgDecoded.StreamSource)


# Icon
$iconBase64      = $iconImg
$iconBytes       = [Convert]::FromBase64String($IconBase64)
$stream          = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
$stream.Write($iconBytes, 0, $iconBytes.Length);

#################### FORMS ##############################

#Defaults
$label_mediaTitle_default_xy       = New-Object System.Drawing.Point(140,255)
$label_mediaRating_default_xy      = New-Object System.Drawing.Point(140,275)
$label_mediaScore_default_xy       = New-Object System.Drawing.Point(215,275)
$label_mediaSummary_default_xy     = New-Object System.Drawing.Point(140,295)
$label_mediaSummary_default_height = 110

# main form
$form                            = New-Object system.Windows.Forms.Form
$form.ClientSize                 = '550,500'
$form.text                       = "Saverr"
$form.BackColor                  = "#4a4a4a"
$form.TopMost                    = $false
$form.StartPosition              = 'CenterScreen'
$form.Icon                       = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())
$form.FormBorderStyle            = "FixedDialog"
$form.MaximizeBox                = $false

$label_title                     = New-Object system.Windows.Forms.Label
$label_title.text                = "Saverr"
$label_title.AutoSize            = $true
$label_title.width               = 30
$label_title.height              = 20
$label_title.location            = New-Object System.Drawing.Point(143,12)
$label_title.Font                = 'Microsoft Sans Serif,15,style=Bold'
$label_title.ForeColor           = "#f5a623"

$pictureBox_logo                 = New-Object system.Windows.Forms.PictureBox
$pictureBox_logo.width           = 75
$pictureBox_logo.height          = 75
$pictureBox_logo.location        = New-Object System.Drawing.Point(15,15)
$pictureBox_logo.image           = $logo
$pictureBox_logo.SizeMode        = [System.Windows.Forms.PictureBoxSizeMode]::normal

$pictureBox_thumb                = new-object Windows.Forms.PictureBox
$pictureBox_thumb.ImageLocation  = ""
$pictureBox_thumb.Visible        = $false
$pictureBox_thumb.Width          = 113 #680
$pictureBox_thumb.Height         = 166 #1000
$pictureBox_thumb.Anchor         = [System.Windows.Forms.AnchorStyles]::Top
$pictureBox_thumb.Location       = New-object System.Drawing.Size(15,250)
$pictureBox_thumb.image          = $loading
$pictureBox_thumb.SizeMode       = [System.Windows.Forms.PictureBoxSizeMode]::Zoom

$label_search                    = New-Object system.Windows.Forms.Label
$label_search.text               = "Search Movie:"
$label_search.AutoSize           = $true
$label_search.width              = 30
$label_search.height             = 15
$label_search.location           = New-Object System.Drawing.Point(15,105)
$label_search.Font               = 'Microsoft Sans Serif,8'
$label_search.ForeColor          = "#f5a623"

$textBox_search                  = New-Object system.Windows.Forms.TextBox
$textBox_search.multiline        = $false
$textBox_search.text             = ""
$textBox_search.width            = 300
$textBox_search.height           = 20
$textBox_search.location         = New-Object System.Drawing.Point(15,125)
$textBox_search.Font             = 'Microsoft Sans Serif,10'

$button_search                   = New-Object system.Windows.Forms.Button
$button_search.BackColor         = "#f5a623"
$button_search.text              = "Search"
$button_search.width             = 80
$button_search.height            = 25
$button_search.location          = New-Object System.Drawing.Point(345,124)
$button_search.Font              = 'Microsoft Sans Serif,9,style=Bold'
$button_search.FlatStyle         = "Flat"

$button_download                 = New-Object system.Windows.Forms.Button
$button_download.Enabled         = $false
$button_download.BackColor       = "#f5a623"
$button_download.text            = "Download"
$button_download.width           = 80
$button_download.height          = 25
$button_download.location        = New-Object System.Drawing.Point(345,174)
$button_download.Font            = 'Microsoft Sans Serif,9,style=Bold'
$button_download.FlatStyle       = "Flat"

$button_settings                 = New-Object system.Windows.Forms.Button
$button_settings.width           = 30
$button_settings.height          = 30
$button_settings.location        = New-Object System.Drawing.Point(508,8)
$button_settings.image           = $gear
$button_settings.FlatStyle       = "Flat"
$button_settings.BackColor       = "Transparent"
$button_settings.FlatAppearance.BorderSize = 0
$button_settings.FlatAppearance.MouseDownBackColor = "Transparent"
$button_settings.FlatAppearance.MouseOverBackColor = "#666666"

$groupBox_type                   = New-Object system.Windows.Forms.Groupbox
$groupBox_type.height            = 40
$groupBox_type.width             = 285
$groupBox_type.text              = "Select Media Type"
$groupBox_type.location          = New-Object System.Drawing.Point(140,46)

$RadioButton_movie               = New-Object system.Windows.Forms.RadioButton
$RadioButton_movie.text          = "Movies"
$RadioButton_movie.AutoSize      = $true
$RadioButton_movie.Checked       = $true
$RadioButton_movie.width         = 80
$RadioButton_movie.height        = 20
$RadioButton_movie.location      = New-Object System.Drawing.Point(15,16)
$RadioButton_movie.Font          = 'Microsoft Sans Serif,9'
$RadioButton_movie.ForeColor     = "#ffffff"

$RadioButton_tv                  = New-Object system.Windows.Forms.RadioButton
$RadioButton_tv.text             = "TV Shows"
$RadioButton_tv.AutoSize         = $true
$RadioButton_tv.width            = 80
$RadioButton_tv.height           = 20
$RadioButton_tv.location         = New-Object System.Drawing.Point(110,16)
$RadioButton_tv.Font             = 'Microsoft Sans Serif,9'
$RadioButton_tv.ForeColor        = "#ffffff"

$RadioButton_music               = New-Object system.Windows.Forms.RadioButton
$RadioButton_music.text          = "Artists"
$RadioButton_music.AutoSize      = $true
$RadioButton_music.width         = 80
$RadioButton_music.height        = 20
$RadioButton_music.location      = New-Object System.Drawing.Point(215,16)
$RadioButton_music.Font          = 'Microsoft Sans Serif,9'
$RadioButton_music.ForeColor     = "#ffffff"

$label_results                   = New-Object system.Windows.Forms.Label
$label_results.text              = "Results:"
$label_results.AutoSize          = $true
$label_results.width             = 30
$label_results.height            = 15
$label_results.location          = New-Object System.Drawing.Point(15,155)
$label_results.Font              = 'Microsoft Sans Serif,8'
$label_results.ForeColor         = "#f5a623"

$comboBox_results                = New-Object system.Windows.Forms.ComboBox
$comboBox_results.text           = ""
$comboBox_results.width          = 300
$comboBox_results.height         = 20
$comboBox_results.location       = New-Object System.Drawing.Point(15,175)
$comboBox_results.Font           = 'Microsoft Sans Serif,10'

$label_seasons                   = New-Object system.Windows.Forms.Label
$label_seasons.text              = ""
$label_seasons.AutoSize          = $true
$label_seasons.width             = 30
$label_seasons.height            = 15
$label_seasons.location          = New-Object System.Drawing.Point(15,225)
$label_seasons.Font              = 'Microsoft Sans Serif,8'
$label_seasons.ForeColor         = "#f5a623"

$comboBox_seasons                = New-Object system.Windows.Forms.ComboBox
$comboBox_seasons.Visible        = $false
$comboBox_seasons.text           = ""
$comboBox_seasons.width          = 130
$comboBox_seasons.height         = 20
$comboBox_seasons.location       = New-Object System.Drawing.Point(65,215)
$comboBox_seasons.Font           = 'Microsoft Sans Serif,10'

$label_episodes                  = New-Object system.Windows.Forms.Label
$label_episodes.text             = ""
$label_episodes.AutoSize         = $true
$label_episodes.width            = 30
$label_episodes.height           = 15
$label_episodes.location         = New-Object System.Drawing.Point(215,225)
$label_episodes.Font             = 'Microsoft Sans Serif,8'
$label_episodes.ForeColor        = "#f5a623"

$comboBox_episodes               = New-Object system.Windows.Forms.ComboBox
$comboBox_episodes.Visible       = $false
$comboBox_episodes.text          = ""
$comboBox_episodes.width         = 45
$comboBox_episodes.height        = 20
$comboBox_episodes.location      = New-Object System.Drawing.Point(265,215)
$comboBox_episodes.Font          = 'Microsoft Sans Serif,10'

$label_mediaTitle                = New-Object system.Windows.Forms.Label
$label_mediaTitle.text           = ""
$label_mediaTitle.AutoSize       = $false
$label_mediaTitle.AutoEllipsis   = $true
$label_mediaTitle.width          = 400
$label_mediaTitle.height         = 20
$label_mediaTitle.location       = $label_mediaTitle_default_xy
$label_mediaTitle.Font           = 'Microsoft Sans Serif,10,style=Bold'
$label_mediaTitle.ForeColor      = "#f5a623"

$label_mediaRating               = New-Object system.Windows.Forms.Label
$label_mediaRating.text          = ""
$label_mediaRating.AutoSize      = $true
$label_mediaRating.width         = 50
$label_mediaRating.height        = 20
$label_mediaRating.location      = $label_mediaRating_default_xy
$label_mediaRating.Font          = 'Microsoft Sans Serif,9,style=Bold'
$label_mediaRating.ForeColor     = "#ffffff"

$label_mediaScore                = New-Object system.Windows.Forms.Label
$label_mediaScore.text           = ""
$label_mediaScore.AutoSize       = $true
$label_mediaScore.width          = 50
$label_mediaScore.height         = 20
$label_mediaScore.location       = $label_mediaScore_default_xy
$label_mediaScore.Font           = 'Microsoft Sans Serif,9,style=Bold'
$label_mediaScore.ForeColor      = "#ffffff"

$label_mediaSummary              = New-Object system.Windows.Forms.Label
$label_mediaSummary.text         = $errorMsg
$label_mediaSummary.AutoSize     = $false
$label_mediaSummary.AutoEllipsis = $true
$label_mediaSummary.width        = 400
$label_mediaSummary.height       = $label_mediaSummary_default_height
$label_mediaSummary.location     = $label_mediaSummary_default_xy
$label_mediaSummary.Font         = 'Microsoft Sans Serif,9'
$label_mediaSummary.ForeColor    = "#ffffff"

$label_DLTitle                   = New-Object system.Windows.Forms.Label
$label_DLTitle.text              = ""
$label_DLTitle.AutoSize          = $false
$label_DLTitle.AutoEllipsis      = $true
$label_DLTitle.width             = 440
$label_DLTitle.height            = 20
$label_DLTitle.location          = New-Object System.Drawing.Point(15,430)
$label_DLTitle.Font              = 'Microsoft Sans Serif,9'
$label_DLTitle.ForeColor         = "#f5a623"

$label_DLProgress                = New-Object system.Windows.Forms.Label
$label_DLProgress.text           = ""
$label_DLProgress.AutoSize       = $false
$label_DLProgress.AutoEllipsis   = $true
$label_DLProgress.width          = 410
$label_DLProgress.height         = 20
$label_DLProgress.location       = New-Object System.Drawing.Point(15,455)
$label_DLProgress.Font           = 'Microsoft Sans Serif,9'
$label_DLProgress.ForeColor      = "#f5a623"

$button_cancel                   = New-Object system.Windows.Forms.Button
$button_cancel.Visible           = $false
$button_cancel.BackColor         = "#f5a623"
$button_cancel.text              = "Cancel"
$button_cancel.width             = 80
$button_cancel.height            = 25
$button_cancel.location          = New-Object System.Drawing.Point(456,470)
$button_cancel.Font              = 'Microsoft Sans Serif,9,style=Bold'
$button_cancel.FlatStyle         = "Flat"

$checkBoxButton_pause            = New-Object System.Windows.Forms.Checkbox 
$checkBoxButton_pause.location   = New-Object System.Drawing.Point(456,440)
$checkBoxButton_pause.Size       = New-Object System.Drawing.Size(80,25)
$checkBoxButton_pause.Appearance = [System.Windows.Forms.Appearance]::Button
$checkBoxButton_pause.Visible    = $false
$checkBoxButton_pause.Text       = "Pause"
$checkBoxButton_pause.FlatStyle  = "Flat"
$checkBoxButton_pause.BackColor  = "#f5a623"
$checkBoxButton_pause.width      = 80
$checkBoxButton_pause.height     = 25
$checkBoxButton_pause.TextAlign  = "MiddleCenter"
$checkBoxButton_pause.AllowDrop = $false
$checkBoxButton_pause.Font              = 'Microsoft Sans Serif,9,style=Bold'
$checkBoxButton_pause.FlatAppearance.BorderColor = "#000000"

$progressBar                     = New-Object system.Windows.Forms.ProgressBar
$progressBar.Name                = 'progressBar1'
$progressBar.BackColor           = "#f5a623"
$progressBar.ForeColor           = "#f5a623"
$progressBar.width               = 430
$progressBar.height              = 12
$progressBar.location            = New-Object System.Drawing.Point(15,480)
$progressBar.Value               = 0
$progressBar.Style               = "Continuous" #"Blocks"
$progressBar.Visible             = $false
$progressBar.Minimum             = 0
$progressBar.Maximum             = 110

# add tooltips
$toolTip                         = New-Object System.Windows.Forms.ToolTip
$toolTip.SetToolTip($label_search, "Searches by first Letter. Excluding 'The' and 'A'")

$form.controls.AddRange(@($groupbox_type,$label_title,$pictureBox_logo,$pictureBox_thumb,$label_search,$textBox_search,$progressBar,$button_search,$button_download,$button_settings,$label_mediaTitle,$label_mediaScore,$comboBox_results,$comboBox_seasons,$comboBox_episodes,$label_mediaRating,$label_mediaSummary,$label_results,$label_seasons,$label_episodes,$label_DLTitle,$label_DLProgress,$button_cancel,$checkBoxButton_pause))
$groupBox_type.controls.AddRange(@($RadioButton_movie,$RadioButton_tv,$RadioButton_music))


# settings form
$form2                           = New-Object system.Windows.Forms.Form
$form2.ClientSize                = '550,500'
$form2.text                      = "Saverr Settings"
$form2.BackColor                 = "#4a4a4a"
$form2.TopMost                   = $false
$form2.StartPosition             = 'CenterScreen'
$form2.Icon                      = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())
$form2.FormBorderStyle           = "FixedDialog"
$form2.MaximizeBox               = $false
$form2.StartPosition             = "CenterParent"

$label2_title                    = New-Object system.Windows.Forms.Label
$label2_title.text               = "Settings"
$label2_title.AutoSize           = $true
$label2_title.width              = 30
$label2_title.height             = 20
$label2_title.location           = New-Object System.Drawing.Point(143,12)
$label2_title.Font               = 'Microsoft Sans Serif,15,style=Bold'
$label2_title.ForeColor          = "#f5a623"

$label2_notice                   = New-Object system.Windows.Forms.Label
$label2_notice.text              = "Username/password is not saved.`nOnly used to retrieve Plex token"
$label2_notice.AutoSize          = $true
$label2_notice.width             = 70
$label2_notice.height            = 20
$label2_notice.location          = New-Object System.Drawing.Point(143,50)
$label2_notice.Font              = 'Microsoft Sans Serif,8'
$label2_notice.ForeColor         = "#ffffff"

$pictureBox2_logo                = New-Object system.Windows.Forms.PictureBox
$pictureBox2_logo.width          = 75
$pictureBox2_logo.height         = 75
$pictureBox2_logo.location       = New-Object System.Drawing.Point(15,15)
$pictureBox2_logo.image          = [System.Drawing.Bitmap][System.Drawing.Image]::FromStream($plexImgDecoded.StreamSource)
$pictureBox2_logo.SizeMode       = [System.Windows.Forms.PictureBoxSizeMode]::normal

$label2_username                 = New-Object system.Windows.Forms.Label
$label2_username.text            = "Plex Username:"
$label2_username.AutoSize        = $true
$label2_username.width           = 30
$label2_username.height          = 20
$label2_username.location        = New-Object System.Drawing.Point(15,121)
$label2_username.Font            = 'Microsoft Sans Serif,8'
$label2_username.ForeColor       = "#ffffff"

$textBox2_username               = New-Object system.Windows.Forms.TextBox
$textBox2_username.multiline     = $false
$textBox2_username.width         = 225
$textBox2_username.height        = 20
$textBox2_username.location      = New-Object System.Drawing.Point(110,110)
$textBox2_username.Font          = 'Microsoft Sans Serif,10'

$label2_password                 = New-Object system.Windows.Forms.Label
$label2_password.text            = "Plex Password:"
$label2_password.AutoSize        = $true
$label2_password.width           = 70
$label2_password.height          = 20
$label2_password.location        = New-Object System.Drawing.Point(15,166)
$label2_password.Font            = 'Microsoft Sans Serif,8'
$label2_password.ForeColor       = "#ffffff"

$textBox2_password               = New-Object system.Windows.Forms.TextBox
$textBox2_password.PasswordChar  = '*'
$textBox2_password.multiline     = $false
$textBox2_password.width         = 225
$textBox2_password.height        = 20
$textBox2_password.location      = New-Object System.Drawing.Point(110,155)
$textBox2_password.Font          = 'Microsoft Sans Serif,10'

$button2_getToken                = New-Object system.Windows.Forms.Button
$button2_getToken.BackColor      = "#f5a623"
$button2_getToken.text           = "Get Token"
$button2_getToken.width          = 95
$button2_getToken.height         = 25
$button2_getToken.location       = New-Object System.Drawing.Point(350,154)
$button2_getToken.Font           = 'Microsoft Sans Serif,9,style=Bold'
$button2_getToken.FlatStyle      = "Flat"

$label2_tokenStatus              = New-Object system.Windows.Forms.Label
$label2_tokenStatus.text         = ""
$label2_tokenStatus.AutoSize     = $true
$label2_tokenStatus.width        = 70
$label2_tokenStatus.height       = 20
$label2_tokenStatus.location     = New-Object System.Drawing.Point(350,180)
$label2_tokenStatus.Font         = 'Microsoft Sans Serif,8'
$label2_tokenStatus.ForeColor    = "#ffff00"

$label2_server                   = New-Object system.Windows.Forms.Label
$label2_server.text              = "Select Server:"
$label2_server.AutoSize          = $true
$label2_server.width             = 70
$label2_server.height            = 20
$label2_server.location          = New-Object System.Drawing.Point(15,211)
$label2_server.Font              = 'Microsoft Sans Serif,8'
$label2_server.ForeColor         = "#ffffff"

$comboBox2_servers               = New-Object system.Windows.Forms.ComboBox
$comboBox2_servers.text          = $settings.name
$comboBox2_servers.width         = 225
$comboBox2_servers.height        = 20
$comboBox2_servers.location      = New-Object System.Drawing.Point(110,200)
$comboBox2_servers.Font          = 'Microsoft Sans Serif,10'

$label2_dlPath                   = New-Object system.Windows.Forms.Label
$label2_dlPath.text              = "Download Path:"
$label2_dlPath.AutoSize          = $true
$label2_dlPath.width             = 70
$label2_dlPath.height            = 20
$label2_dlPath.location          = New-Object System.Drawing.Point(15,256)
$label2_dlPath.Font              = 'Microsoft Sans Serif,8'
$label2_dlPath.ForeColor         = "#ffffff"

$textBox2_dlPath                 = New-Object system.Windows.Forms.TextBox
$textBox2_dlPath.text            = $settings.dlPath
$textBox2_dlPath.ReadOnly        = $true
$textBox2_dlPath.multiline       = $false
$textBox2_dlPath.width           = 225
$textBox2_dlPath.height          = 20
$textBox2_dlPath.location        = New-Object System.Drawing.Point(110,245)
$textBox2_dlPath.Font            = 'Microsoft Sans Serif,10'
$textBox2_dlPath.Enabled         = $false

$button2_dlPath                  = New-Object system.Windows.Forms.Button
$button2_dlPath.BackColor        = "#f5a623"
$button2_dlPath.text             = "Select Path"
$button2_dlPath.width            = 95
$button2_dlPath.height           = 25
$button2_dlPath.location         = New-Object System.Drawing.Point(350,244)
$button2_dlPath.Font             = 'Microsoft Sans Serif,9,style=Bold'
$button2_dlPath.FlatStyle        = "Flat"

$label2_pathStatus               = New-Object system.Windows.Forms.Label
$label2_pathStatus.text          = ""
$label2_pathStatus.AutoSize      = $true
$label2_pathStatus.width         = 70
$label2_pathStatus.height        = 20
$label2_pathStatus.location      = New-Object System.Drawing.Point(350,270)
$label2_pathStatus.Font          = 'Microsoft Sans Serif,8'
$label2_pathStatus.ForeColor     = "#ffff00"

$label2_ssl                      = New-Object system.Windows.Forms.Label
$label2_ssl.text                 = "SSL Required:"
$label2_ssl.AutoSize             = $true
$label2_ssl.width                = 70
$label2_ssl.height               = 20
$label2_ssl.location             = New-Object System.Drawing.Point(15,290)
$label2_ssl.Font                 = 'Microsoft Sans Serif,8'
$label2_ssl.ForeColor            = "#ffffff"

$checkBox_ssl                    = New-Object System.Windows.Forms.Checkbox 
$checkBox_ssl.location           = New-Object System.Drawing.Point(110,285)
$checkBox_ssl.Size               = New-Object System.Drawing.Size(80,25)
$checkBox_ssl.width              = 80
$checkBox_ssl.height             = 25
$checkBox_ssl.checked            = $settings.ssl

$label2_debug                    = New-Object system.Windows.Forms.Label
$label2_debug.text               = "Debug Logging:"
$label2_debug.AutoSize           = $true
$label2_debug.width              = 70
$label2_debug.height             = 20
$label2_debug.location           = New-Object System.Drawing.Point(15,325)
$label2_debug.Font               = 'Microsoft Sans Serif,8'
$label2_debug.ForeColor          = "#ffffff"

$checkBox_debug                  = New-Object System.Windows.Forms.Checkbox 
$checkBox_debug.location         = New-Object System.Drawing.Point(110,320)
$checkBox_debug.Size             = New-Object System.Drawing.Size(80,25)
$checkBox_debug.width            = 80
$checkBox_debug.height           = 25
$checkBox_debug.checked          = $settings.logging

$label2_ssl_info                 = New-Object system.Windows.Forms.Label
$label2_ssl_info.text            = "[Restart Saverr after changing SSL or Debug options]"
$label2_ssl_info.AutoSize        = $true
$label2_ssl_info.width           = 70
$label2_ssl_info.height          = 20
$label2_ssl_info.location        = New-Object System.Drawing.Point(15,360)
$label2_ssl_info.Font            = 'Microsoft Sans Serif,8'
$label2_ssl_info.ForeColor       = "#ffffff"

$button2_servers                 = New-Object system.Windows.Forms.Button
$button2_servers.BackColor       = "#f5a623"
$button2_servers.text            = "List Servers"
$button2_servers.width           = 95
$button2_servers.height          = 25
$button2_servers.location        = New-Object System.Drawing.Point(350,199)
$button2_servers.Font            = 'Microsoft Sans Serif,9,style=Bold'
$button2_servers.FlatStyle       = "Flat"

$label2_serverStatus             = New-Object system.Windows.Forms.Label
$label2_serverStatus.text        = ""
$label2_serverStatus.AutoSize    = $true
$label2_serverStatus.width       = 70
$label2_serverStatus.height      = 20
$label2_serverStatus.location    = New-Object System.Drawing.Point(350,225)
$label2_serverStatus.Font        = 'Microsoft Sans Serif,8'
$label2_serverStatus.ForeColor   = "#ffff00"

$label2_notice2                  = New-Object system.Windows.Forms.Label
$label2_notice2.text             = ""
$label2_notice2.AutoSize         = $true
$label2_notice2.width            = 70
$label2_notice2.height           = 20
$label2_notice2.location         = New-Object System.Drawing.Point(15,395)
$label2_notice2.Font             = 'Microsoft Sans Serif,8'
$label2_notice2.ForeColor        = "#ffffff"

$label2_saveStatus               = New-Object system.Windows.Forms.Label
$label2_saveStatus.text          = ""
$label2_saveStatus.AutoSize      = $true
$label2_saveStatus.width         = 70
$label2_saveStatus.height        = 20
$label2_saveStatus.location      = New-Object System.Drawing.Point(110,275)
$label2_saveStatus.Font          = 'Microsoft Sans Serif,8'
$label2_saveStatus.ForeColor     = "#00ff00"

$label2_help                     = New-Object system.Windows.Forms.LinkLabel
$label2_help.text                = "Help"
$label2_help.AutoSize            = $true
$label2_help.width               = 70
$label2_help.height              = 20
$label2_help.location            = New-Object System.Drawing.Point(15,480)
$label2_help.Font                = 'Microsoft Sans Serif,9'
$label2_help.ForeColor           = "#00ff00"
$label2_help.LinkColor           = "#f5a623"
$label2_help.ActiveLinkColor     = "#f5a623"
$label2_help.add_Click({[system.Diagnostics.Process]::start("https://github.com/ninthwalker/saverr")})

$label2_version                  = New-Object system.Windows.Forms.Label
$label2_version.text             = "Ver. 1.1.1"
$label2_version.AutoSize         = $true
$label2_version.width            = 70
$label2_version.height           = 20
$label2_version.location         = New-Object System.Drawing.Point(480,480)
$label2_version.Font             = 'Microsoft Sans Serif,9'
$label2_version.ForeColor        = "#f5a623"

$form2.controls.AddRange(@($label2_title,$label2_username,$label2_password,$label2_dlPath,$label2_server,$label2_pathStatus,$label2_serverStatus,$label2_notice,$label2_notice2,$label2_saveStatus,$label2_tokenStatus,$label2_help,$label2_version,$textBox2_username,$textBox2_password,$textBox2_dlPath,$label2_debug,$checkBox_debug,$label2_ssl,$checkBox_ssl,$label2_ssl_info,$comboBox2_servers,$button2_servers,$pictureBox2_logo,$button2_getToken,$button2_dlPath))


############################## CODE ################################


# search server for media. Movies/tv are by firstletter. Music is by artist.
function search {

    Try {

        # clear old searches
        clearMediaInfo
        clearDLStatus
        $comboBox_results.Text = "Searching ..."

        # get sections
        $sections = $scheme + $settings.server + "/library/sections/" + "?X-Plex-Token=" + $settings.serverToken
        $xmlsearch = plx $sections

        #$sectionType = $groupBox_type.Controls | ? { $_.Checked } | Select-Object Text

            # GUI Selection of 'type' to search for. Remove spaces and leading 'the' and 'a' since plex removes those from search words
            $searchName = ($textBox_search.Text).TrimStart().TrimEnd()
            $searchName = $searchName -replace '^the |^a ', ''
            $sectionType = $groupBox_type.Controls | ? { $_.Checked -eq $true} | Select-Object Text

            Switch ($sectionType.Text)
            {
                'Movies'   {$script:type = "movie"
                            $subSection = "Video"; Break
                }
                'TV Shows' {$script:type = "show"
                            $subSection = "Directory"; Break
                }
                'Artists'  {$script:type = "artist"
                            $subSection = "Directory"; Break
                }
            }
            
            # get 'type' of key
            $sections2search = $xmlsearch.MediaContainer.Directory | ? {$_.type -eq $type} | select key
            $firstChar = $searchName.ToUpper()[0]

            # Search movies/tv
            if ($type -ne "artist") {
                $sectionsList = new-object collections.generic.list[object]

                foreach ($section in $sections2search) {
    
                    $sectionsUrl = $scheme + $settings.server + "/library/sections/$($section.key)/firstCharacter/$firstChar/" + "?X-Plex-Token=" + $settings.serverToken
                    $sectionsList.Add((plx $sectionsUrl))
                }

            # search through list for match
            $script:searchResults = $sectionsList.MediaContainer.$subSection | ? {$_.title -like "*$searchName*" -and $_.type -eq $type} | select title,type,key,tagline,summary,year,contentrating,thumb,rating

            }

            #search artists
            else {
                $artistList = new-object collections.generic.list[object]

                foreach ($section in $sections2search) {
    
                    $sectionsUrl = $scheme + $settings.server + "/library/sections/$($section.key)/all" + "?X-Plex-Token=" + $settings.serverToken
                    $artistList.Add((plx $sectionsUrl))
                     
                }

            $script:searchResults = $artistList.MediaContainer.$subSection | ? {$_.title -like "*$searchName*" -and $_.type -eq $type} | select title,type,key,thumb

            # search through list for match
            $trackList = new-object collections.generic.list[object]
            $artistPath = $artistsList.MediaContainer.$subSection | select key
                foreach ($artist in $artistPath) {
                    $artistURL = $scheme + $settings.server + "$($artist.key)" + "?X-Plex-Token=" + $settings.serverToken
                    $trackList.Add((plx $artistURl))
            }


            }

            #show results
            if ($searchresults) {
                $comboBox_results.Text = "$(@($searchResults).count) $($sectionType.Text) found!"
                $comboBox_results.Text = "$(@($searchResults).count) $($sectionType.Text) found!"
                foreach ($item in $searchresults) {
                    if ($type -ne "artist") {
                        if ($item.year) {
                            $comboBox_results.Items.Add($item.title + ' (' + $item.year + ')') 
                        }
                        else {
                            $comboBox_results.Items.Add($item.title)
                        }
                    }
                    else {
                        $comboBox_results.Items.Add($item.title)
                    }
                }
            }
            else {
                $ComboBox_results.Text = "No results found!"
            }
    }

    Catch {
        logit
        $comboBox_results.Items.Clear()
        $ComboBox_results.Text = "Error! Check settings/token/server status?"
    }
}

function mediaInfo {

    clearDLStatus
    $comboBox_index = $comboBox_results.SelectedIndex
    $script:info = $searchResults[$comboBox_index] | Select title,type,key,tagline,summary,year,contentrating,thumb,rating,size
    $thumb = $scheme + $settings.server + $info.thumb + "?X-Plex-Token=" + $settings.serverToken

    # enable download button if no other downloads in progress
    if (!(Get-BitsTransfer)) {
        $button_download.Enabled = $true
    }

    if ($info.thumb) {
        $pictureBox_thumb.ImageLocation = "$thumb"
        $pictureBox_thumb.Visible = $true
    }
    else {
        $pictureBox_thumb.ImageLocation = ""
        $pictureBox_thumb.Image = $loading
        $pictureBox_thumb.Visible = $true
    }
    if ($info.year) {
        $label_mediaTitle.Text = "$($info.title) ($($info.year))"
    }
    else {
        $label_mediaTitle.Text = "$($info.title)"
    }
    if ($info.contentrating) {
        $label_mediaRating.Text = "$($info.contentrating)"
    }
    elseif ($type -ne "artist") {
        $label_mediaRating.Text = "No Rating"
    }
    if ($info.rating) {
        $label_mediaScore.Text = "$($info.rating)"
    }
    if ($info.summary) {
        $label_mediaSummary.Text = "$($info.summary)"
    }

    # show season/ep boxes
    if ($type -eq "show") {

        #get seasons
        $comboBox_seasons.Items.Clear()
        $seasonPath = $scheme + $settings.server + "$($info.key)" + "?X-Plex-Token=" + $settings.serverToken
        $xmlSeason = plx $seasonPath
            $script:seasons = $xmlSeason.MediaContainer.directory | select title,key,index
        foreach ($season in $seasons) {
            $comboBox_seasons.Items.Add($season.title)
        }
        if (($comboBox_seasons.Items).count -ne "1" ) {
            $comboBox_seasons.Text = $comboBox_seasons.items[1]
        }
        else {
            $comboBox_seasons.Text = $comboBox_seasons.items[0]
        }
    }

    if ($type -eq "artist") {

        #get albums
        $comboBox_seasons.Items.Clear()
        $seasonPath = $scheme + $settings.server + "$($info.key)" + "?X-Plex-Token=" + $settings.serverToken
        $xmlSeason = plx $seasonPath
            $script:seasons = $xmlSeason.MediaContainer.directory | select title,key,index,year

        foreach ($season in $seasons) {
            $comboBox_seasons.Items.Add($season.title)
        }

        if (($comboBox_seasons.Items).count -ge 2) {
            $comboBox_seasons.Items.Add("All Albums")
        }

        $comboBox_seasons.Text = $comboBox_seasons.items[0]
    }

}

#get episodes/tracks
function episodeSelection {
    
    clearDLStatus
    # Movies and tv shows
    if ($type -ne "artist") {
        $comboBox_episodes.Items.Clear()
        $script:comboBox_seasons_index = $comboBox_seasons.SelectedIndex

        if ($comboBox_seasons.Text -ne "All Episodes") {
            $episodePath = $scheme + $settings.server + "$($seasons[$comboBox_seasons_index].key)" + "?X-Plex-Token=" + $settings.serverToken
            $script:xmlEpisode = plx $episodePath
                $script:episodes = $xmlEpisode.MediaContainer.video | select title,key,contentrating,summary,rating,year,thumb,originallyAvailableAt,index,duration
        
            foreach ($episode in $episodes) {
                $comboBox_episodes.Items.Add($episode.index)
            }
            
            if ($comboBox_episodes.Items -ge 2) {
                $comboBox_episodes.Items.Add("All")
            }

            $comboBox_episodes.Text = $comboBox_episodes.items[0]
        }
        else {
            $comboBox_episodes.Text = "All"
            $label_mediaTitle.Text = "Download All Episodes from All Seasons"
            $label_mediaRating.Text = ""
            $label_mediaScore.Text = ""
            $label_mediaSummary.Text = "Notice: This may take a very long time depending on number of seasons/episodes."
        }
    }

    # music
    else {
        $comboBox_episodes.Items.Clear()
        $comboBox_seasons_index = $comboBox_seasons.SelectedIndex
        $episodePath = $scheme + $settings.server + "$($seasons[$comboBox_seasons_index].key)" + "?X-Plex-Token=" + $settings.serverToken
        $script:xmlEpisode = plx $episodePath
        $script:episodes = $xmlEpisode.MediaContainer.track | select title,key,index,duration,summary,parentYear,thumb,grandparentthumb,addedAt

        foreach ($episode in $episodes) {
            $comboBox_episodes.Items.Add($episode.title)
        }
        if (($comboBox_episodes.Items).count -ge 2) {
            $comboBox_episodes.Items.Add("All Tracks")
        }

        if ($comboBox_seasons.text -ne "All Albums") {
            $comboBox_episodes.Text = $comboBox_episodes.items[0]
        }
        else {
            $comboBox_episodes.Text = "All Tracks"
            $label_mediaTitle.Text = "Download All Albums"
            $label_mediaRating.Text = ""
            $label_mediaScore.Text = ""
            $label_mediaSummary.Text = "Notice: This may take a very long time depending on number of albums/tracks."

        }

    }
}



function mediaEpInfo {

    clearDLStatus
    $comboBox_episode_index = $comboBox_episodes.SelectedIndex
    if ($comboBox_episode_index -ne $null -and $episodes -ne $null) {
        $script:infoEp = $episodes[$comboBox_episode_index] | Select title,type,key,tagline,summary,year,contentrating,thumb,rating
    }
    
    if ($comboBox_episodes.Text -eq "All" -and $comboBox_seasons.Text -ne "All episodes") {
        $label_mediaTitle.Text = "Download All Episodes from $($comboBox_seasons.Text)"
        $label_mediaRating.Text = ""
        $label_mediaScore.Text = ""
        $label_mediaSummary.Text = "Notice: This may take a very long time depending on number of episodes."
    }
    elseif ($comboBox_episodes.Text -eq "All Tracks" -and $comboBox_seasons.Text -ne "All Albums") {
        $label_mediaTitle.Text = "Download All Tracks from $($comboBox_seasons.Text)"
        $label_mediaRating.Text = ""
        $label_mediaScore.Text = ""
        $label_mediaSummary.Text = "Notice: This may take a very long time depending on number of tracks."
    }
    else {
        if ($infoEp.year) {
            $label_mediaTitle.Text = "$($infoEp.title) ($($infoEp.year))"
        }
        else {
            $label_mediaTitle.Text = "$($infoEp.title)"
        }
        if ($infoEp.contentrating) {
            $label_mediaRating.Text = "$($infoEp.contentrating)"
        }
        elseif ($type -ne "artist") {
            $label_mediaRating.Text = "No Rating"
        }
        if ($infoEp.rating) {
            $label_mediaScore.Text = "$($infoEp.rating)"
        }
        if ($infoEp.summary) {
            $label_mediaSummary.Text = "$($infoEp.summary)"
        }
        else {
            $label_mediaSummary.Text = ""
        }
    }
}

# retrieve token. Does not store username/password, only the token
function getToken {
    try { 
        $label2_tokenStatus.ForeColor = "#ffff00"
        $label2_tokenStatus.Text = "Retrieving ..."
        $username = $textBox2_username.Text
        $password = $textBox2_password.Text

        # Use this method for now for more powershell version backwards compatability instead of -credential
        $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $username,$password)))

        $headers = @{
            "X-Plex-Version" = "1.1.1"
            "X-Plex-Product" = "Saverr"
            "X-Plex-Client-Identifier" = "271938"
            "Content-Type" = "application/xml"
            "Authorization" = ("Basic {0}" -f $base64AuthInfo)
        }

        $data = Invoke-RestMethod -Method POST -Uri $plexSignInUrl -Headers $headers
        $script:userToken = $data.user.authToken

        # update settings file
        if (Test-Path .\saverrSettings.xml) {
            $script:settings = Import-Clixml .\saverrSettings.xml
            Add-Member -InputObject $settings -MemberType NoteProperty -Name 'userToken' -Value $userToken -force
            $settings | Export-Clixml .\saverrSettings.xml
        }
        else {
            $script:settings = [pscustomobject] @{
                userToken = $userToken
            }
            $settings | Export-Clixml .\saverrSettings.xml
        }
        $settings = Import-Clixml .\saverrSettings.xml
        $label2_tokenStatus.ForeColor = "#00ff00"
        $label2_tokenStatus.Text = "Token Saved!"
    }
    catch {
        logit
        $label2_tokenStatus.ForeColor = "#ff0000"
        $label2_tokenStatus.Text = "Error! User/Pass?"
    }
}

function getServers {
    Try {
        $label2_serverStatus.text = "Searching ..."
        $label2_notice2.text = ""
        # full server url
        $serversUrl = $plexServersUrl + "?X-Plex-Token=" + $settings.userToken

        # get servers
        $serversXml = plx $serversUrl
        $script:serverList =  $serversxml.MediaContainer.Server | select name,host,port,accessToken,localAddresses,owned

        #output servers
        if ($serverList) {
            $comboBox2_servers.Items.Clear();
            $comboBox2_servers.Text = "$(@($serverList).count) Plex Servers found!"
            foreach ($server in $serverList) {
                $comboBox2_servers.Items.Add($server.Name)
            }
        }
        else {
            $comboBox2_servers.Items.Clear()
            $ComboBox2_servers.Text = "No Plex servers found! Got Token?"
        }
        $label2_serverStatus.text = ""
    }
    Catch {
        logit
        $label2_serverStatus.text = ""
        $comboBox2_servers.Items.Clear()
        $ComboBox2_servers.Text = "Error! Check token?"
    }
}


# save all setttings to file
function saveServer {

    Try {
        $comboBox2_index = $comboBox2_servers.SelectedIndex
        $selectedServer = $serverList[$comboBox2_index]
                
        if ($selectedServer.owned -ne "1") {
            $serverUrl = $selectedServer.Host + ":" + $selectedServer.Port
        }
        else {
            $ipCheck = Invoke-RestMethod http://ipinfo.io/json | Select -exp ip
            if ($ipcheck -eq $selectedServer.Host) {

                if ( (($selectedServer.localaddresses).GetType().name) -eq "String" ) {
                    $serverUrl = $selectedServer.localAddresses + ":" + "32400"
                }
                elseif ( (($selectedServer.localaddresses).GetType().name) -eq "Array" ) {
                    $serverUrl = $selectedServer.localAddresses[0] + ":" + "32400"
                }
            }
            else {
                $serverUrl = $selectedServer.Host + ":" + $selectedServer.Port
            }
        }

        $serverToken = $($selectedServer.accessToken)
        $serverName = $($selectedServer.name)

        [PsCustomObject] @{
        name = "$serverName"
        server = "$serverUrl"
        dlPath = "$($textBox2_dlPath.Text)"
        userToken = "$($settings.userToken)"
        serverToken = "$serverToken"
		ssl = "$ssl"
        logging = "$debug"
        } | Export-Clixml .\saverrSettings.xml

        $label2_serverStatus.ForeColor = "#00ff00"
        $label2_serverStatus.text = "Server Saved!"
    }
    Catch {
        logit
        $label2_notice2.ForeColor = "#ff0000"
        $label2_notice2.text = "Error saving! Got token? selected a server?"
    }
}

function clearStatusSave {

    # save path
    # $dlPath = "$($textBox2_dlPath.Text)" # old way before dialog box

    if (Test-Path .\saverrSettings.xml) {
        $script:settings = Import-Clixml .\saverrSettings.xml
        if ($dlPath) {
            Add-Member -InputObject $settings -MemberType NoteProperty -Name 'dlPath' -Value $dlPath -force
            $settings | Export-Clixml .\saverrSettings.xml
        }
    }
    else {
        if ($dlPath) {
            $script:settings = [pscustomobject] @{
                dlPath = $dlPath
            }
        }
    }
    
    # clear status
    clearMediaInfo
    clearDLStatus
    $errorMsg = ""
    $label2_saveStatus.text   = ""
    $label2_serverStatus.text = ""
    $label2_tokenStatus.Text  = ""
    $label2_pathStatus.Text = ""
    $label2_notice2.Text = ""
    $textBox_search.Text = ""
    $textBox2_dlPath.Text = $settings.dlPath
    $textBox2_username.Text = ""
    $textBox2_password.Text = ""
    $comboBox2_servers.Items.Clear()
    $comboBox2_servers.Text = $settings.name

    Try {
        if ($dlPath) {
            New-Item -ItemType Directory -Force -Path $dlPath
        }
        $label_mediaTitle.ForeColor = "#ffffff"
        $label_mediaTitle.Text = ""
        $label_mediaSummary.Text = ""
    }
    Catch {
        logit
        $label_mediaTitle.ForeColor = "#ff0000"
        $label_mediaTitle.Text = "Error creating download Path"
        $label_mediaSummary.Text = "Could Not validate download directory:`n$dlPath.`n`nCheck path name or system permissions maybe?"
    }
    
    Try {
        if ($settings) {
            $settings | Export-Clixml .\saverrSettings.xml
        }
        $label_mediaTitle.ForeColor = "#ffffff"
        $label_mediaTitle.Text = ""
        $label_mediaSummary.Text = ""
    }
    Catch {
        logit
        $label_mediaTitle.ForeColor = "#ff0000"
        $label_mediaTitle.Text = "Error saving settings"
        $label_mediaSummary.Text = "Could Not Create settings file at:`n$PSScriptRoot.`n`nCheck path or system permissions maybe?"
    }

    if ((!($settings.name)) -or (!($settings.server)) -or (!($settings.userToken)) -or (!($settings.serverToken)) -or (!($settings.dlPath))) {
        $label_mediaSummary.text = "Settings are not fully configured.`nPlease click the gear icon before searching."
    }
    else {
        $label_mediaSummary.text = ""
    }

}

function clearMediaInfo {

    $comboBox_results.Items.Clear()
    $comboBox_results.Text = ""
    $pictureBox_thumb.Visible = $false
    $pictureBox_thumb.image = $loading
    $pictureBox_thumb.ImageLocation = ""
    $label_mediaTitle.Text = ""
    $label_mediaRating.Text = ""
    $label_mediaScore.Text = ""
    $label_mediaSummary.Text = ""
    $comboBox_seasons.Text = ""
    $comboBox_episodes.Text = ""
    $comboBox_seasons.Items.Clear()
    $comboBox_episodes.Items.Clear()
    $button_download.Enabled = $false

}

function clearDLStatus {

    if ($label_DLTitle.Text -match "^Download Completed|^Download Failed|^Download Cancelled|^There was an error|^Error") {
            $label_DLTitle.Text = ""
    }
    if ($label_DLProgress.Text -ne "Download Paused!") {
        $label_DLProgress.Text = ""
        $label_DLProgress.ForeColor = "#f5a623"
    }

}

function cancelJob {
    $progressBar.Value = 0
    $label_DLTitle.ForeColor = "#ff0000"
    $label_DLTitle.Text = "Download Cancelled!"
    $label_DLProgress.Text = ""
    $progressBar.Visible = $false
    if ($comboBox_results.Items) {
        $button_download.Enabled = $true
    }
    $button_cancel.Visible = $false
    $button_cancel.Enabled = $false
    $CheckBoxButton_pause.Visible = $false
    $CheckboxButton_pause.Enabled = $false
    $script:pauseLoop = $false
    $checkBoxButton_pause.Text = "Pause"
    Get-BitsTransfer | Complete-BitsTransfer

    # clean up any empty folders created.
    # This can throw an error in the console sometimes if the path is deleted too fast and then it doesn't exist. Not worth it to remove error, doesn't stop the app.
    if ($dlType -eq "allEp" -or $dlType -eq "allSeasons") {
        if (Test-Path $allSeasonPath) {
            Get-ChildItem $allSeasonPath -Directory -recurse | where {-NOT $_.GetFiles("*","AllDirectories")} | del -recurse -ErrorAction SilentlyContinue
            if (Test-Path $allSeasonPath) {
                if ((Get-ChildItem $allSeasonPath | Measure-Object).Count -eq 0) {
                    del $allSeasonPath -ErrorAction SilentlyContinue
                }
            }
        }
    }

    # enable minimize again
    $form.MinimizeBox = $true

}

########### OPERATIONS ###############

# Buttons
$button_download.Add_Click({

    Try {
        if ((Get-BitsTransfer).JobState -ne "Suspended") {
            # tell them what's going on, reset status
            $button_download.Enabled = $false
            $label_DLTitle.ForeColor = "#f5a623"
            $label_DLProgress.ForeColor = "#f5a623"
            $label_DLProgress.Text = ""
            $label_DLTitle.Text = "Processing download request"
            $label_DLTitle.Refresh()
            $button_download.Text = "Download"
            $checkBoxButton_pause.Checked = $false
            $status = "success"

            # movie dl links
            if ($type -eq "movie") {
                $mediaURL = $scheme + $settings.server + $info.key + "?X-Plex-Token=" + $settings.serverToken
                $mediaPath = plx $mediaURL
                $mediaInfo = $mediaPath.MediaContainer.Video.Media.Part | select key,file -First 1
                $dlURL = $scheme + $settings.server + $mediaInfo.key + "?download=1" + "&X-Plex-Token=" + $settings.serverToken
                $script:dlName = Split-Path $mediaInfo.file -Leaf
                $script:dlType = "one"
            }

            # tv show dl links
            if ($type -eq "show") {

                # if all episodes from one season is selected
                if ($comboBox_episodes.Text -eq "All" -and $comboBox_seasons.Text -ne "All episodes") {
                    $script:allSeasonPath = "$($settings.dlPath)\$(Remove-InvalidChars $info.title)"
                    $allEpPath = "$allSeasonPath\$(Remove-InvalidChars $comboBox_seasons.Text)"
                    New-Item -ItemType Directory -Force -Path $allEpPath
                    $allEp = $xmlepisode.MediaContainer.Video.Media.Part | select @{n="Source";e={$scheme + $settings.server + $_.key + "?X-Plex-Token=" + $settings.serverToken}},@{n="Destination";e={$allEpPath + "\" + (Split-Path $_.file -Leaf)}}
                    $script:dlType = "allEp"

                    # remove links that have already been downloaded
                    $allEpData = @()
                    For ($I=0; $I -lt $allep.count; $I++) {

                        if (!(Test-Path $allEp.destination[$I])) {
                            $allEpData += [pscustomobject] @{
                                Source  = $allep.source[$I]
                                Destination = $allEp.destination[$I]
                            }
                        }

                    }

                }
                # if all seasons and all episodes is selected
                elseif ($comboBox_episodes.Text -eq "All" -and $comboBox_seasons.Text -eq "All episodes") {
                    $script:allSeasonPath = "$($settings.dlPath)\$(Remove-InvalidChars $info.title)"
                    $mediaURL = $scheme + $settings.server + $seasons.key[0] + "?X-Plex-Token=" + $settings.serverToken
                    $mediaPath = plx $mediaURL
                    $seasonNumber = $mediaPath.MediaContainer.Video | select parenttitle
                    $seasonClean = $seasonNumber.parenttitle | % {Remove-InvalidChars $_}
                    $allEp = $mediaPath.MediaContainer.Video.Media.Part | select @{n="Source";e={$scheme + $settings.server + $_.key + "?X-Plex-Token=" + $settings.serverToken}},@{n="Destination";e={(Split-Path $_.file -Leaf)}}
                    $allEpClean = $allEp.destination | % {Remove-InvalidChars $_}
                    $script:dlType = "allSeasons"

                    # combine source/destination/season data for Bitstransfer import
                    # remove links that have already been downloaded
                    $allEpData = @()
                    For ($I=0; $I -lt $allep.count; $I++) {

                        $finalDestination = $allSeasonPath + "\" + $seasonClean[$I] + "\" + $allEpClean[$I]

                        if (!(Test-Path $finalDestination)) {
                            $allEpData += [pscustomobject] @{
                                Source  = $allep.source[$I]
                                Destination = $finalDestination
                            }
                        }
                    }


                    # Bitstransfer defaults to 200 max files per job. Truncate download to 200 unless registry value is set to something else.
                    if ($allEpData.length -ge $limit) {
                        $setLimit = $limit - 1
                        $allEpData = $allEpData[0..$setLimit]
                        $noLimitStatus = "(Max limit set to $limit)"
                    }    

                    # pre-create directories for seasons
                    $seasonClean | select -Unique | % {New-Item -ItemType Directory -Force -Path "$allSeasonPath\$_"}

                }

                # if just one episode or a movie is selected
                else {
                    $mediaURL = $scheme + $settings.server + $infoEp.key + "?X-Plex-Token=" + $settings.serverToken
                    $mediaPath = plx $mediaURL
                    $mediaInfo = $mediaPath.MediaContainer.Video.Media.Part | select key,file -First 1
                    $dlURL = $scheme + $settings.server + $mediaInfo.key + "?download=1" + "&X-Plex-Token=" + $settings.serverToken
                    $script:dlName = Split-Path $mediaInfo.file -Leaf | % {Remove-InvalidChars $_}
                    $script:dlType = "one"
                }
             
            }
  
            # music dl links
            if ($type -eq "artist") {

                # if all tracks from one album is selected
                if ($comboBox_episodes.Text -eq "All Tracks" -and $comboBox_seasons.Text -ne "All Albums") {
                    $script:allSeasonPath = "$($settings.dlPath)\$(Remove-InvalidChars $info.title)"
                    $allEpPath = "$allSeasonPath\$(Remove-InvalidChars $comboBox_seasons.Text)"
                    New-Item -ItemType Directory -Force -Path $allEpPath
                    $allEp = $xmlepisode.MediaContainer.Track.Media.Part | select @{n="Source";e={$scheme + $settings.server + $_.key + "?X-Plex-Token=" + $settings.serverToken}},@{n="Destination";e={$allEpPath + "\" + (Split-Path $_.file -Leaf)}}
                    $script:dlType = "allTracks"

                    # remove links that have already been downloaded
                    $allEpData = @()
                    For ($I=0; $I -lt $allep.count; $I++) {

                        if (!(Test-Path $allEp.destination[$I])) {
                            $allEpData += [pscustomobject] @{
                                Source  = $allep.source[$I]
                                Destination = $allEp.destination[$I]
                            }
                        }

                    }

                }

                # if all tracks and all albums is selected
                elseif ($comboBox_episodes.Text -eq "All Tracks" -and $comboBox_seasons.Text -eq "All Albums") {
                    $script:allSeasonPath = "$($settings.dlPath)\$(Remove-InvalidChars $info.title)"

                    # collect all album metadata paths
                    $mediaURL = @()
                    $seasons | % { $mediaURL += $scheme + $settings.server + $_.key + "?X-Plex-Token=" + $settings.serverToken }

                    # get all tracks
                    $all = $mediaURL | % {(plx $_).MediaContainer.Track}
                    $allEp = $all.media.part | select @{n="Source";e={$scheme + $settings.server + $_.key + "?X-Plex-Token=" + $settings.serverToken}},@{n="Destination";e={(Split-Path $_.file -Leaf)}}  
                    $allClean = $all.parenttitle | % {Remove-InvalidChars $_}
                    $allSeasonClean = $all.parenttitle | % {Remove-InvalidChars $_}
                    $allEpClean = $allEp.destination | % {Remove-InvalidChars $_}
                    $script:dlType = "allAlbums"

                    # combine source/destination/season data for Bitstransfer import
                    $allEpData = @()
                    For ($I=0; $I -lt $allEp.count; $I++) {

                        $finalDestination = $allSeasonPath + "\" + $allSeasonClean[$I] + "\" + $allEpClean[$I]

                        if (!(Test-Path $finalDestination)) {
                            $allEpData += [pscustomobject] @{
                                Source  = $allEp.source[$I]
                                Destination = $finalDestination
                            }
                        }

                    }

                    # Bitstransfer defaults to 200 max files per job. Truncate download to 200 unless registry value is set to something else.
                    if ($allEpData.length -ge $limit) {
                        $setLimit = $limit - 1
                        $allEpData = $allEpData[0..$setLimit]
                        $noLimitStatus = "(Max limit set to $limit)"
                    }

                    # pre-create directories for seasons
                    $allClean | select -Unique | % {New-Item -ItemType Directory -Force -Path "$allSeasonPath\$_"}

                }

                # if just one music track is selected
                else {
                    $mediaURL = $scheme + $settings.server + $infoEp.key + "?X-Plex-Token=" + $settings.serverToken
                    $mediaPath = plx $mediaURL
                    $mediaInfo = $mediaPath.MediaContainer.track.media.part | select key,file -First 1
                    $mediaInfo2 = $mediaPath.MediaContainer.track | select grandparentTitle,parentTitle,title -First 1
                    $dlURL = $scheme + $settings.server + $mediaInfo.key + "?download=1" + "&X-Plex-Token=" + $settings.serverToken
                    $script:dlName = Split-Path $mediaInfo.file -Leaf | % {Remove-InvalidChars $_}
                    $script:dlType = "one"
                }

            }        
        
            # Cancelling all old Bits jobs
            Get-BitsTransfer | Remove-BitsTransfer

            # get starting time
            $startTime = Get-Date

            # disable minimize since it causes issues during downloads
            $form.MinimizeBox = $false

            # download all episodes from a season or album
            if ($dlType -eq "allEp" -or $dlType -eq "allTracks") {
                $script:myjob = Start-BitsTransfer -source "$($allEpData.Source[0])" -Destination "$($allEpData.Destination[0])" -DisplayName "Downloading ..." -Description "All Episodes" -Asynchronous -Suspended
                $allEpData[1..($allEpData.Length -1)] | Add-BitsFile $myjob
                if ($ssl -eq $True) {bitsadmin /SetSecurityFlags $myjob.displayname 30}
                Resume-BitsTransfer $myjob -Asynchronous
            }

            # download all seasons or all albums
            elseif ($dltype -eq "allSeasons" -or $dlType -eq "allAlbums") {
                $script:myjob = Start-BitsTransfer -source "$($allEpData.Source[0])" -Destination "$($allEpData.Destination[0])" -DisplayName "Downloading ..." -Description "All Episodes" -Asynchronous -Suspended
                $allEpData[1..($allEpData.Length -1)] | Add-BitsFile $myjob
                if ($ssl -eq $True) {bitsadmin /SetSecurityFlags $myjob.displayname 30}
                Resume-BitsTransfer $myjob -Asynchronous
            }

            # download a movie or one episode or one song
            else {
                $script:myjob = Start-BitsTransfer -Source $dlURL -Destination "$($settings.dlPath)\$dlName" -DisplayName "Downloading ..." -Description $dlName -Asynchronous -Suspended
                if ($ssl -eq $True) {bitsadmin /SetSecurityFlags $myjob.displayname 30}
				Resume-BitsTransfer $myjob -Asynchronous
            }

        }
        elseif ((Get-BitsTransfer).JobState -eq "Suspended") {
            # resume if they paused the download
            $button_download.Enabled = $false
            $button_download.Text = "Download"
            $label_DLTitle.ForeColor = "#f5a623"
            $label_DLProgress.ForeColor = "#f5a623"
            $label_DLTitle.Text = "Paused download detected. Resuming progress ..."
            $checkBoxButton_pause.Text = "Pause"
            $label_DLProgress.Text = ""
            $label_DLProgress.ForeColor = "#f5a623"
            Resume-BitsTransfer $myjob -Asynchronous
            $label_DLTitle.Refresh()
            $status = "success"
        }
        # disable minimize since it causes issues during downloads
        $form.MinimizeBox = $false

        # Pause to let it start before checking progress. Timeout after 30sec
        $count    = 0
        $noDot    = (0,4,8,12,16,20,24,28)
        $oneDot   = (1,5,9,13,17,21,25,29)
        $twoDot   = (2,6,10,14,18,22,26,30)
        $threeDot = (3,7,11,15,19,23,27)

        # timeout
        :check while ($count -lt $timeout) {

            if (((Get-BitsTransfer | ? { $_.JobState -eq "Transferring" }).Count -gt 0) -or (Get-BitsTransfer | ? { $_.JobState -eq "Transferred" }) -or (Get-BitsTransfer | ? { $_.JobState -eq "Error" })) {   
                # exit check
                break check
            }

            if ($noDot -contains $count) {
                $label_DLTitle.Text = "Processing download request"
                $label_DLTitle.Refresh()
            }
            elseif ($oneDot -contains $count) {
                $label_DLTitle.Text = "Processing download request ."
                $label_DLTitle.Refresh()
            }
            elseif ($twoDot -contains $count) {
                $label_DLTitle.Text = "Processing download request . ."
                $label_DLTitle.Refresh()
            }
            elseif ($threeDot -contains $count) {
                $label_DLTitle.Text = "Processing download request . . ."
                $label_DLTitle.Refresh()
            }

            # increase counter
            Start-Sleep -Seconds 1
            $count++

        }

        if ($count -ge $timeout) {
            $status = "failed"

            if (Get-BitsTransfer | ? { $_.JobState -like "*Error*" }) {
                $bitsError = (Get-BitsTransfer | select ErrorDescription).ErrorDescription
                if ($debug) {
                    $eMSG = "$(Get-Date): Connection Timed out after $timeout seconds. $bitsError"
                    $eMSG | Out-File ".\saverrLog.txt" -Append
                }
            }

            # remove bitstransfer jobs and clean up empty directories created
            Get-BitsTransfer | Remove-BitsTransfer
            if ($dlType -like "all*") {
                if (Test-Path $allSeasonPath) {
                    Get-ChildItem $allSeasonPath -Directory -recurse | where {-NOT $_.GetFiles("*","AllDirectories")} | del -recurse -ErrorAction SilentlyContinue
                    if (Test-Path $allSeasonPath) {
                        if ((Get-ChildItem $allSeasonPath | Measure-Object).Count -eq 0) {
                            del $allSeasonPath -ErrorAction SilentlyContinue
                        }
                    }
                }
            }

            $label_DLTitle.ForeColor = "#ff0000"
            $label_DLTitle.Text = "Download Failed! Timed out.`n Error: $bitsError"
            $button_download.Enabled = $true
            $label_DLProgress.Text = ""
        }

        if ((Get-BitsTransfer | select ErrorDescription).ErrorDescription -like "*404*") {
            $status = "failed"

            # remove bitstransfer jobs and clean up empty directories created
            Get-BitsTransfer | Remove-BitsTransfer
            if ($dlType -like "all*") {
                if (Test-Path $allSeasonPath) {
                    Get-ChildItem $allSeasonPath -Directory -recurse | where {-NOT $_.GetFiles("*","AllDirectories")} | del -recurse -ErrorAction SilentlyContinue
                    if (Test-Path $allSeasonPath) {
                        if ((Get-ChildItem $allSeasonPath | Measure-Object).Count -eq 0) {
                            del $allSeasonPath -ErrorAction SilentlyContinue
                        }
                    }
                }
            }

            $label_DLTitle.ForeColor = "#ff0000"
            $label_DLTitle.Text = "Download Failed! File not found. Check server"
            $button_download.Enabled = $true
            $label_DLProgress.Text = ""

        }

        if ($status -eq "success") {
            # Show progress
            $progressBar.Visible = $true
            $progressBar.Value = 0
            $label_DLProgress.Text = ""
            $label_DLProgress.ForeColor = "#f5a623"

            # Init CancelLoop
            $script:cancelLoop = $false
            $button_cancel.Enabled = $true
            $button_cancel.Visible = $true
            $checkBoxButton_pause.Enabled = $true
            $checkBoxButton_pause.Visible = $true

            if (Get-BitsTransfer | ? { $_.JobState -ne "Transferred" }) {

                :xfer while ((Get-BitsTransfer | ? { $_.JobState -eq "Transferring" }).Count -gt 0) { 
                    $totalbytes=0;    
                    $bytestransferred=0; 
                    $timeTaken = 0;
                    foreach ($job in (Get-BitsTransfer | ? { $_.JobState -eq "Transferring" } | Sort-Object CreationTime)) {
                             
                        $totalbytes += [math]::round($job.BytesTotal /1MB);
                        $totalSize = byteSize $($job.BytesTotal)         
                        $bytestransferred += [math]::round($job.bytestransferred /1MB)
                        $transferSize = byteSize $($job.bytestransferred)   
                        if ($timeTaken -eq 0) { 
                            #Get the time of the oldest transfer aka the one that started first
                            $timeTaken = ((Get-Date) - $job.CreationTime).TotalMinutes 
                        }
                    }    
                    #TimeRemaining = (TotalFileSize - BytesDownloaded) * TimeElapsed/BytesDownloaded
                    [System.Windows.Forms.Application]::DoEvents()

                    # cancel download if asked
                    if ($script:cancelLoop -eq $true) {
                        cancelJob
                        # exit loop
                        break xfer
                    }

                    # pause download if asked
                    if ($script:pauseLoop -eq $true) {
                        $label_DLProgress.Text = "Download Paused!"
                        $label_DLProgress.ForeColor = "#ffff00" # yellow
                        Get-BitsTransfer | Suspend-BitsTransfer

                        # allow minimize while paused
                        $form.MinimizeBox = $true

                        # exit loop
                        break xfer
                    }

                    if ($totalbytes -gt 0 -and $bytestransferred -gt 0 -and $timetaken -gt 0) {        
                        [int]$timeLeft = ($totalBytes - $bytestransferred) * ($timeTaken / $bytestransferred)
                        [int]$pctComplete = $(($bytestransferred*100)/$totalbytes);     
                        $label_DLProgress.Text = "$transferSize of $totalSize ($pctComplete%) - Approx. $timeLeft minutes remaining"
                        $progressBar.Value = $pctComplete
                        #$label_DLProgress.Refresh()

                        Switch ($dlType)
                            {
                                'one'           { $label_DLTitle.Text = "Downloading: $dlName"; Break }
                                'allEp'         { $label_DLTitle.Text = "Completed: $([int]$myjob.FilesTransferred) of $([int]$myjob.FilesTotal) Episodes $noLimitStatus"; Break }
                                'allTracks'     { $label_DLTitle.Text = "Completed: $([int]$myjob.FilesTransferred) of $([int]$myjob.FilesTotal) Tracks $noLimitStatus"; Break }
                                'allSeasons'    { $label_DLTitle.Text = "Completed: $([int]$myjob.FilesTransferred) of $([int]$myjob.FilesTotal) Episodes $noLimitStatus"; Break }
                                'allAlbums'     { $label_DLTitle.Text = "Completed: $([int]$myjob.FilesTransferred) of $([int]$myjob.FilesTotal) Tracks $noLimitStatus"; Break }
                            }

                        $progressBar.PerformStep()
                        Start-Sleep -Seconds 1
                    }
                }
            }

            # download went too fast. Pause to show title downloaded
            else {
                Switch ($dlType)
                    {
                        'one'           { $label_DLTitle.Text = "Downloading: $dlName"; Break }
                        'allEp'         { $label_DLTitle.Text = "Completed: $([int]$myjob.FilesTransferred) of $([int]$myjob.FilesTotal) Episodes"; Break }
                        'allTracks'     { $label_DLTitle.Text = "Completed: $([int]$myjob.FilesTransferred) of $([int]$myjob.FilesTotal) Tracks"; Break }
                        'allSeasons'    { $label_DLTitle.Text = "Completed: $([int]$myjob.FilesTransferred) of $([int]$myjob.FilesTotal) Episodes"; Break }
                        'allAlbums'     { $label_DLTitle.Text = "Completed: $([int]$myjob.FilesTransferred) of $([int]$myjob.FilesTotal) Tracks"; Break }
                    }

                Start-Sleep -Seconds 2

            }

            # Finish and close Bitstransfer
            if ($script:cancelLoop -eq $false -and $script:pauseLoop -eq $false) {

                if (Get-BitsTransfer | ? { $_.JobState -like "*Error*" }) {
                    $bitsError = (Get-BitsTransfer | select ErrorDescription).ErrorDescription
                    if ($debug) {
                        $eMSG = "$(Get-Date): Download Error. $bitsError"
                        $eMSG | Out-File ".\saverrLog.txt" -Append
                    }
                    $label_DLTitle.ForeColor = "#ff0000"
                    $label_DLProgress.Text = ""
                    $label_DLTitle.Text = "Error: $bitsError"
                    $progressBar.Visible = $false
                    $progressBar.Value = 0
                    $button_download.Enabled = $true
                    $button_cancel.Visible = $false
                    $button_cancel.Enabled = $false
                    $checkBoxButton_pause.Visible = $false
                    $CheckBoxButton_pause.Enabled = $false
                    
                }
                else {
                    if ($startTime -ne $null) {
                        $howLong = (get-date).Subtract($startTime)
                        if ($howLong.Minutes -eq "0") {
                            $dlTime = "$($howLong.Seconds) seconds"
                        }
                        else {
                            $dlTime = "$($howLong.Minutes) minutes"
                        }
                    }
                    else {
                        $dlTime = "Unknown time"
                    }

                    $label_DLProgress.Text = ""
                    $label_DLTitle.ForeColor = "#00ff00"
                    $label_DLTitle.Text = "Download Completed in: $dlTime"
                    $button_download.Enabled = $true
                    $button_cancel.Visible = $false
                    $button_cancel.Enabled = $false
                    $checkBoxButton_pause.Visible = $false
                    $CheckBoxButton_pause.Enabled = $false
                    $progressBar.Visible = $false
                    $progressBar.Value = 0
                }

                Get-BitsTransfer | Complete-BitsTransfer

                # remove any empty folder created
                if ($dlType -like "all*") {
                    if (Test-Path $allSeasonPath) {
                        Get-ChildItem $allSeasonPath -Directory -recurse | where {-NOT $_.GetFiles("*","AllDirectories")} | del -recurse -ErrorAction SilentlyContinue
                        if (Test-Path $allSeasonPath) {
                            if ((Get-ChildItem $allSeasonPath | Measure-Object).Count -eq 0) {
                                del $allSeasonPath -ErrorAction SilentlyContinue
                            }
                        }
                    }
                }

                # allow minimize again
                $form.MinimizeBox = $true

            }
        }

    }

    Catch {
        logit
        Get-BitsTransfer | Remove-BitsTransfer

        # clean up any empty folders
        if ($dlType -like "all*") {
            if (Test-Path $allSeasonPath) {
                Get-ChildItem $allSeasonPath -Directory -recurse | where {-NOT $_.GetFiles("*","AllDirectories")} | del -recurse -ErrorAction SilentlyContinue
                if (Test-Path $allSeasonPath) {
                    if ((Get-ChildItem $allSeasonPath | Measure-Object).Count -eq 0) {
                        del $allSeasonPath -ErrorAction SilentlyContinue
                    }
                }
            }
        }

        $label_DLTitle.ForeColor = "#ff0000"
        $label_DLProgress.Text = ""
        $label_DLTitle.Text = "There was an error with the download."
        $progressBar.Visible = $false
        $progressBar.Value = 0
        $button_download.Enabled = $true
        $button_cancel.Visible = $false
        $button_cancel.Enabled = $false

        # enable minimize again
        $form.MinimizeBox = $true
    }

})


$button_search.Add_Click({search})

$button_settings.Add_Click({[void]$form2.ShowDialog()})

$button_cancel.Add_Click({
    $cancelMsg = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to cancel the download?`nAll non-completed files will be deleted!",'Cancel Download','YesNo','Question')

    if ($cancelMsg -eq "Yes" -and $script:pauseLoop -ne $true) {
        $script:cancelLoop = $true
    }
    elseif ($cancelMsg -eq "Yes" -and $script:pauseLoop -eq $true) {
        $script:cancelLoop = $true
        cancelJob
    }
})

$button2_getToken.Add_Click({getToken})

$button2_servers.Add_Click({getServers})

$button2_dlPath.Add_Click({
    try {
        $script:dlPath = Get-SavePath
        if ($dlPath) {
            $textBox2_dlPath.Text = $dlPath
            $label2_pathStatus.ForeColor = "#00ff00"
            $label2_pathStatus.Text = "Path Saved!"
        }
    }
    catch {
        logit
        $label2_pathStatus.ForeColor = "#ff0000"
        $label2_pathStatus.Text = "Error! Check log"
    }

})

# save token on enter key
$textBox2_password.Add_KeyUp({
    if ($_.KeyCode -eq "Enter") {
        getToken
    }
})

# search on enter key
$textBox_search.Add_KeyUp({
    if ($_.KeyCode -eq "Enter") {search}
})

 $checked_type ={
             if ($RadioButton_movie.Checked){
                   clearMediaInfo
                   clearDLStatus
                   $textBox_search.Text = ""
                   $label_seasons.Text = ""
                   $label_episodes.Text = ""
                   $label_search.text = "Search Movie:"
                   $combobox_seasons.Visible = $false
                   $combobox_episodes.Visible = $false
                   $label_mediaTitle.location = $label_mediaTitle_default_xy
                   $label_mediaRating.location = $label_mediaRating_default_xy
                   $label_mediaScore.location = $label_mediaScore_default_xy
                   $label_mediaSummary.location = $label_mediaSummary_default_xy
                   $label_mediaSummary.height = $label_mediaSummary_default_height
                   $toolTip.SetToolTip($label_search, "Searches by First Letter. Excluding 'The' and 'A'")
                   }
            if ($RadioButton_tv.Checked){
                   clearMediaInfo
                   clearDLStatus
                   $textBox_search.Text = ""
                   $label_search.text = "Search TV Show:"
                   $label_seasons.Text = "Season:"
                   $label_episodes.text = "Episode:"
                   $combobox_seasons.Visible = $true
                   $combobox_episodes.Visible = $true
                   $label_episodes.location = New-Object System.Drawing.Point(215,225)
                   $comboBox_episodes.location = New-Object System.Drawing.Point(270,215)
                   $label_mediaTitle.location = $label_mediaTitle_default_xy
                   $label_mediaRating.location = $label_mediaRating_default_xy
                   $label_mediaScore.location = $label_mediaScore_default_xy
                   $label_mediaSummary.location = $label_mediaSummary_default_xy
                   $label_mediaSummary.height = $label_mediaSummary_default_height
                   $combobox_seasons.width = 130
                   $combobox_episodes.width = 45
                   $toolTip.SetToolTip($label_search, "Searches by first Letter. Excluding 'The' and 'A'")
                   }
            elseif ($radiobutton_music.Checked) {
                   clearMediaInfo
                   clearDLStatus
                   $textBox_search.Text = ""
                   $label_search.text = "Search Artist:"
                   $label_seasons.Text = "Album:"
                   $label_episodes.Text = "Track:"
                   $label_episodes.location = New-Object System.Drawing.Point(280,225)
                   $comboBox_episodes.location = New-Object System.Drawing.Point(325,215)
                   $combobox_seasons.Visible = $true
                   $combobox_episodes.Visible = $true
                   $combobox_seasons.width = 195
                   $combobox_episodes.width = 210
                   $label_mediaTitle.location = New-Object System.Drawing.Point(140,280)
                   $label_mediaRating.location = New-Object System.Drawing.Point(140,300)
                   $label_mediaScore.location = New-Object System.Drawing.Point(140,300)
                   $label_mediaSummary.location = New-Object System.Drawing.Point(140,320)
                   $label_mediaSummary.height = 85
                   $toolTip.SetToolTip($label_search, "Searches by Artist Name")
                }
}

$checkBoxButton_pause.Add_CheckedChanged({
    if ($checkBoxButton_pause.Checked -eq $true){
        $script:pauseLoop = $true
        $checkBoxButton_pause.Text = "Resume"
    }
    else {
        if ((Get-BitsTransfer).JobState -eq "Suspended") {
            $script:pauseLoop = $false
            $checkBoxButton_pause.Text = "Pause"
            $button_download.Enabled = $true
            $button_download.PerformClick()
        }
    }
})

$checkBox_debug.Add_CheckedChanged({
    if ($checkBox_debug.Checked -eq $true){
        $setDebug = $true
    }
    else {
        $setDebug = $false
    }
    # update settings file
    if (Test-Path .\saverrSettings.xml) {
        $script:settings = Import-Clixml .\saverrSettings.xml
        Add-Member -InputObject $settings -MemberType NoteProperty -Name 'logging' -Value $setDebug -force
        $settings | Export-Clixml .\saverrSettings.xml
    }
    else {
        $script:settings = [pscustomobject] @{
            logging = $setDebug
        }
        $settings | Export-Clixml .\saverrSettings.xml
    }
    $settings = Import-Clixml .\saverrSettings.xml
    $debug = $settings.logging
})

$checkBox_ssl.Add_CheckedChanged({
    if ($checkBox_ssl.Checked -eq $true){
        $setSSL = $true
    }
    else {
        $setSSL = $false
    }
    # update settings file
    if (Test-Path .\saverrSettings.xml) {
        $script:settings = Import-Clixml .\saverrSettings.xml
        Add-Member -InputObject $settings -MemberType NoteProperty -Name 'ssl' -Value $setSSL -force
        $settings | Export-Clixml .\saverrSettings.xml
    }
    else {
        $script:settings = [pscustomobject] @{
            ssl = $setSSL
        }
        $settings | Export-Clixml .\saverrSettings.xml
    }
    $settings = Import-Clixml .\saverrSettings.xml
    $ssl = $settings.ssl
})


# show extra season/artist fields
$RadioButton_movie.Add_CheckedChanged($checked_type)
$RadioButton_tv.Add_CheckedChanged($checked_type)
$RadioButton_music.Add_CheckedChanged($checked_type)

# Show season media info on selection
$comboBox_results.Add_SelectedIndexChanged({mediaInfo})

# show episodes after season selection
$comboBox_seasons.Add_SelectedIndexChanged({episodeSelection})

# show ep media info after season selection
$comboBox_episodes.Add_SelectedIndexChanged({mediaEpInfo})

# save server on selection
$comboBox2_servers.Add_SelectedValueChanged({saveServer})

$form2.add_FormClosing({clearStatusSave})

# confirm closing. clear any downloads on close.
$form.add_FormClosing({
    if (Get-BitsTransfer) {
        $question = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to Exit?`nAll non-completed files will be deleted!", 'Exit Saverr', 'YesNo', 'Question')

        if ($question -eq 'Yes') {
            $script:cancelLoop = $true
            Get-BitsTransfer | Complete-BitsTransfer

            # remove empty folders created
            if (($dlType -like "all*") -and ($script:pauseLoop -eq $true)) {
                if (Test-Path $allSeasonPath) {
                    Get-ChildItem $allSeasonPath -Directory -recurse | where {-NOT $_.GetFiles("*","AllDirectories")} | del -recurse -ErrorAction SilentlyContinue
                    if (Test-Path $allSeasonPath) {
                        if ((Get-ChildItem $allSeasonPath | Measure-Object).Count -eq 0) {
                            del $allSeasonPath -ErrorAction SilentlyContinue
                        }
                    }
                }
            }

        }
        else {
            $_.Cancel = $true
        }
    }
})

# show the form
[void]$form.ShowDialog()

# close the forms
$form.Dispose()
$form2.Dispose()
# end
