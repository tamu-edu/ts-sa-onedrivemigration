<#
.SYNOPSIS
  OneDrive File Migration Script
.DESCRIPTION
  Script checks for necessary permissions and staging requirements before guiding user through migration progress
.PARAMETER <Parameter_Name>
  No paramters required, adjust script settings in "[Initializations]" section
.INPUTS
  None
.OUTPUTS
  Transcript stored in "C:\logs\OneDriveMigration\Transcript_$($env:USERNAME)_$scriptStartTime.log"
  Missing Files list stored in "C:\logs\OneDriveMigration\MissingFiles_$(Get-Date -Format "yyyyMMdd_HHmmss")_$user.log"
.NOTES
  Version:        0.9
  Author:         Callan Christensen callan@tamu.edu
  Creation Date:  3/22/2024
  Purpose/Change: First general release
  
.EXAMPLE
  .\ComputerLogon-OneDriveMigration.ps1
#>

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#NoModulesRequired

#Prereq for popup user info windows
Add-Type -AssemblyName System.Windows.Forms


#---------------------------------------------------------[Initializations]--------------------------------------------------------

# Transcript the entire script and log it with date/time
$scriptStartTime = Get-Date -Format "yyyyMMddHHmmss"
$transcriptLogPath = "C:\logs\OneDriveMigration\OneDriveMigration_$($env:USERNAME)_$scriptStartTime.log"
Start-Transcript -Path $transcriptLogPath

#Grab the currently logged on user's domain and username
$domain, $user = ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name -split '\\')

#If you have a location that your users can write to if you want logs to be placed in a network location
#$transcriptLogPath2 = "\\mySecure\Log\location\OneDriveMigration_$($env:USERNAME)_$scriptStartTime.log"

#Define the user's home folder location using the currently logged on user's name
$userHomeDrive = "\\dsa.reldom.tamu.edu\student affairs\Departments\Information Technology\Staff\cctest"

#Define the local user folder path (C:\users\$username)
$userFolderPath = [System.Environment]::GetFolderPath('UserProfile')

# Define the location inside of the OneDrive folder that the files will migrate to
$migrationDest = Join-Path -Path $userFolderPath -ChildPath ("OneDrive - Texas A&M University\Documents\" + ($user.Substring(0,1).ToUpper() + $user.Substring(1) + " OneDrive Migration Files " + (Get-Date -Format "yyyy-MM-dd HH_mm_ss")))


# Check if migration has already been completed previously, flag is set in User's Home Drive location

$flagPattern = "$userHomeDrive\MigrationCompleted_*.flag"
if (Test-Path $flagPattern) {
    Write-Warning "User $User has already completed migration, the .flag marker has been detected in: $userHomeDrive"
    exit
}

# Test the path to $userHomeDrive 
if (-not (Test-Path $userHomeDrive)) {
    Write-Warning "User $User network home drive not acessible, it either doesn't exist or there is a user permission missing."
    Stop-Transcript
    exit
}

<#Use this section for if you need to load the Home Directory or other information from an user's AD object attributes
# Define the AD/LDAP server and domain for user lookup
$ldapServer = "LDAP://auth.tamu.edu"  # Replace with AD that has user's HomeDrive info
$ldapDomain = "DC=auth,DC=tamu,DC=edu"  # Replace with  AD that has user's HomeDrive info

# LDAP Query to search for user based on "city" attribute ("l" for location)
$queryString = "(&(objectCategory=user)(l=$user))"

#Create an [adsisearcher] object with the query
$searcher = New-Object System.DirectoryServices.DirectorySearcher
# Set the SearchRoot to the specified domain and server
$searcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("$ldapServer/$ldapDomain")
# Set the Filter
$searcher.Filter = $queryString

# Specify the properties you want to retrieve 
$searcher.PropertiesToLoad.Add("sAMAccountName") > $null
#This is a useful attribute to store info, doesn't exist until you define it
$searcher.PropertiesToLoad.Add("info") > $null
#Would change this setting Once the migration is completed
$searcher.PropertiesToLoad.Add("description") > $null  # Retrieve the description attribute
#>

#-----------------------------------------------------------[Functions]------------------------------------------------------------
# Function to test user read permission on a path
function Test-UserReadPermission {
    param (
        [string]$path,
        [string]$user
    )
    $fullUserName = "$env:USERDOMAIN\$user"
    $acl = Get-Acl -Path $path
    # If you are writing the .flag file to the folder when you are done, more than READ permission is necessary
    #$readRights = [System.Security.AccessControl.FileSystemRights]::Read
    $modifyRights = [System.Security.AccessControl.FileSystemRights]::Modify
    $fullControlRights = [System.Security.AccessControl.FileSystemRights]::FullControl

    foreach ($accessRule in $acl.Access) {
        if ($accessRule.IdentityReference.Value -eq $fullUserName) {
            $rights = $accessRule.FileSystemRights

            if (($rights -band $fullControlRights) -eq $fullControlRights) {
                return $true
            }
            if (($rights -band $modifyRights) -eq $modifyRights) {
                return $true
            }
            if (($rights -band $readRights) -eq $readRights) {
                return $true
            }
        }
    }
    return $false
}


# Function to get a list of files excluding "desktop.ini"
function Get-FileList {
    param (
        [string]$folderPath
    )
    Get-ChildItem -Path $folderPath -Recurse -File | Where-Object { $_.Name -ne "desktop.ini" } | ForEach-Object { $_.FullName.Replace($folderPath, '') }
}

# Function to check if OneDrive is installed and user is signed in
function CheckOneDrive {

    $onedriveExePaths = @("C:\Program Files\Microsoft OneDrive\OneDrive.exe", "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe", "C:\Program Files (x86)\Microsoft OneDrive\OneDrive.exe")
    $onedriveExePath = $onedriveExePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    
    if (-not $onedriveExePath) {
        Write-Warning "OneDrive is not installed or not found. Please contact IT."
        return $false
    }

    $oneDriveFolderPath = "$env:USERPROFILE\OneDrive - Texas A&M University"
    if (-not (Test-Path $oneDriveFolderPath)) {
        # Open OneDrive and hopefully user will sign in and press retry
        try {
            Start-Process -FilePath $onedriveExePath -ErrorAction Stop
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to start OneDrive. Please contact the Student Affairs Helpdesk for assistance.", "DSA Home Folder Migration", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            Write-Warning "One Drive failed to open or it is already opened and logged in."
        }
        return $false
    }

    #OneDrive is installed and user is logged in
    return $true
}

# Watchdog script to minimize MS Teams from poping over the migration prompt
$TeamswatchdogScript = {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    public class User32 {
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    }
"@
    
    #An unforunate solution for keeping teams minimized during migration
    $startTime = Get-Date
    $foundAndMinimized = $false
    while ((New-TimeSpan -Start $startTime -End (Get-Date)).TotalMinutes -lt 1 -and !$foundAndMinimized) {
        $hWnd = [User32]::FindWindow([NullString]::Value, "Microsoft Teams")
        if ($hWnd -ne [IntPtr]::Zero) {
            [User32]::ShowWindow($hWnd, 2)  # 2 = SW_MINIMIZE
            $foundAndMinimized = $true
        }
        Start-Sleep -Seconds 1
    }
}

#OneDrive watchdog makes sure the OneDrive process is running, only after 3 exits will it stop trying or after 45 minutes.

$OneDriveRelauncherWatchdog = {
    $restartAttempts = 0
    $maxRestartAttempts = 3
    $timeLimit = 2700
    $isOneDriveRestarted = $false

    while ($true) {
        $oneDriveProc = Get-Process -Name 'OneDrive' -ErrorAction SilentlyContinue

        if ($null -ne $oneDriveProc) {
            if ($isOneDriveRestarted) {
                $elapsedTime = (Get-Date) - $oneDriveProc.StartTime
                if ($elapsedTime.TotalSeconds -ge $timeLimit) {
                    Write-Host "OneDrive has been running for $timeLimit seconds post-restart. Exiting watchdog."
                    return
                }
            }
        }
        else {
            if ($restartAttempts -ge $maxRestartAttempts) {
                Write-Host "Reached maximum restart attempts. Exiting watchdog."
                return
            }

            #Add your additional OneDrive paths here
            $possiblePaths = @(
                'C:\Program Files\Microsoft OneDrive\OneDrive.exe',
                'C:\Program Files (x86)\Microsoft OneDrive\OneDrive.exe',
                "$env:APPDATA\Microsoft\OneDrive\OneDrive.exe"
            )

            foreach ($path in $possiblePaths) {
                if (Test-Path $path) {
                    Start-Process -FilePath $path
                    Write-Host "Launched OneDrive from $path using Start-Process."
                    break
                }
            }

            $isOneDriveRestarted = $true
            $restartAttempts++
            Write-Host "Attempt $restartAttempts to restart OneDrive."
        }

        Start-Sleep -Seconds 1
    }
}

<#Watchdog Script to quit and reopen OneDrive when a specific sign in screen is detected. This "catch and release"
of the OneDrive user sign-in exprience is a workaround that enforces group policy settings that aren't applied at this stage of the login.#>

$OneDrivePromptWatchdog = { Add-Type -AssemblyName "UIAutomationClient"

    # Function to recursively find a UI element by its name
    function Find-UIElementByName {
        param(
            [System.Windows.Automation.AutomationElement]$Root,
            [string]$Name
        )

        $condition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, $Name)
        return $Root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $condition)
    }

    # Function to launch OneDrive from the correct path
    function Launch-OneDrive {
        $paths = @(
            "C:\Program Files\Microsoft OneDrive\OneDrive.exe",
            "C:\Program Files (x86)\Microsoft OneDrive\OneDrive.exe",
            "$env:USERPROFILE\AppData\Local\Microsoft\OneDrive\OneDrive.exe"
        )
    
        foreach ($path in $paths) {
            if (Test-Path $path) {
                start-sleep -Seconds 2
                Start-Process -FilePath $path
                break
            }
        }
    }

    # Main watchdog script
    while ($true) {
        Start-Sleep -Milliseconds 650 # Check every 250ms

        # Get the main OneDrive window
        $condition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, 'Microsoft OneDrive')
        $oneDriveWindow = [System.Windows.Automation.AutomationElement]::RootElement.FindFirst([System.Windows.Automation.TreeScope]::Children, $condition)

        if ($oneDriveWindow) {
            # Check for unique toggle/button/text
            $uniqueElement = Find-UIElementByName -Root $oneDriveWindow -Name 'Ready to back up'

            if ($uniqueElement) {
                # If the unique element is found, kill OneDrive
                Stop-Process -Name "OneDrive" -Force

                # Wait until the OneDrive process has fully stopped
                do {
                    Start-Sleep -Milliseconds 50
                } until (-not (Get-Process "OneDrive" -ErrorAction SilentlyContinue))
                

                Start-Sleep -Milliseconds 500
                # Launch OneDrive from the correct path
                Launch-OneDrive
            }
        }
    }

    #End of Script block
}



#-----------------------------------------------------------[Execution]------------------------------------------------------------


#Watchdog Script Launching

# Convert the Teams watchdog to a Base64 encoded string to avoid character limits
$Teamsbytes = [System.Text.Encoding]::Unicode.GetBytes($TeamswatchdogScript.ToString())
$TeamswatchdogScriptEncodedCommand = [Convert]::ToBase64String($Teamsbytes)

# Start Teams watchdog in a separate hidden PowerShell process
Start-Process PowerShell.exe -ArgumentList "-NoProfile -EncodedCommand $TeamswatchdogScriptEncodedCommand" -WindowStyle Hidden

# Convert the OneDrive watchdog to a Base64 encoded string to avoid character limits
$OneDrive1bytes = [System.Text.Encoding]::Unicode.GetBytes($OneDriveRelauncherWatchdog.ToString())
$OneDriveWatchdogEncodedCommand = [Convert]::ToBase64String($OneDrive1bytes)

# Start the OneDriveWatchdog in a separate hidden PowerShell process
Start-Process PowerShell.exe -ArgumentList "-NoProfile -EncodedCommand $OneDriveWatchdogEncodedCommand" -WindowStyle Hidden

# Convert the Teams watchdog to a Base64 encoded string to avoid character limits
$OneDrive2bytes = [System.Text.Encoding]::Unicode.GetBytes($OneDrivePromptWatchdog.ToString())
$OneDrivePromptWatchdogEncodedCommand = [Convert]::ToBase64String($OneDrive2bytes)

#Start the OneDrive "Back Up Folder on this PC" user prompt catch/exit watchdog
Start-Process PowerShell.exe -ArgumentList "-NoProfile -EncodedCommand $OneDrivePromptWatchdogEncodedCommand" -WindowStyle Hidden

# File Migration Start

$hasPermission = Test-UserReadPermission -path $userHomeDrive -user $user -domain $domain

# If user doesn't have permission to their home directory, show a popup and exit
if (-not $hasPermission) {
    $message = "You do not have the necessary permissions to access all the required folders. Please contact IT support."
    $caption = "Permission Error"
    [System.Windows.Forms.MessageBox]::Show($message, $caption, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    
    # Log the error
    $logPath = "C:\logs\OneDriveMigration\PermissionError_$(Get-Date -Format "yyyyMMdd_HHmmss")_$user.log"
    $errorMessage = "User $user does not have read permission to $userfolderpath. Please contact IT support."
    Add-Content -Path $logPath -Value $errorMessage
    
    Write-Warning $errorMessage
    Stop-Transcript
    exit
}

#Runs CheckOneDrive Function and will set as $true if OneDrive check passes
$onedriveStatus = CheckOneDrive

#
while (-not $onedriveStatus) {
    #Open OneDrive and hopefully user will sign in and press retry
    #Start-Process -FilePath $onedriveExePath -ErrorAction Stop
    $result = [System.Windows.Forms.MessageBox]::Show('Please sign into TAMU OneDrive with your NetID and click retry to retrieve your user files. Press Cancel to Exit.', 'OneDrive Migration', 'RetryCancel', 'Warning')
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
        Write-Warning "User exited migration at OneDrive Sign-On Prompt."
        exit
    }
    #Loop this function until it becomes true
    $onedriveStatus = CheckOneDrive
}

$ReadyToBeginMigrationPrompt = [System.Windows.Forms.MessageBox]::Show('To begin your Home Folder Migration. Please press OK or press Cancel to exit.', 'OneDrive Migration', [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Information)

if ($ReadyToBeginMigrationPrompt -eq [System.Windows.Forms.DialogResult]::Cancel) {
    # User clicked Cancel, exit script
    Write-Warning "User exited after signing into OneDrive but before running the migration."
    Stop-Transcript
    exit
}

$logTimeStamp = Get-Date -Format "yyyyMMddHHmmss"
$robocopyLogPath = "C:\logs\OneDriveMigration\Robocopy_$user`_$logTimeStamp.log"

if(-not (Test-path $migrationDest)){new-item -ItemType Directory -Path $migrationDest}

if ((Test-Path $userHomeDrive) -and (test-path $migrationDest)) {
    
    robocopy $userHomeDrive $migrationDest /E /Z /LOG+:$robocopyLogPath /NP /R:3 /W:3 /TEE /COPY:DAT /XO /XF desktop.ini
    
    <#Robocopy Mirror Purge Option (This option uses the "mirror" flag which can copy if the files are locked)
    robocopy $desktopSource $desktopDest /MIR /COPY:DAT /V /LOG+:$robocopyLogPath /XF "desktop.ini"
    #>
    # Removes potential desktop.ini from destination since robocopy doesn't seem to always respect file exclusions
    
    $desktopIniPath = Join-Path -Path $migrationDest -ChildPath "desktop.ini"
    if (Test-Path $desktopIniPath) {
        Remove-Item -Path $desktopIniPath -Force
    }
}

$sourceFilesCompare = Get-FileList -folderPath $userHomeDrive
$destFilesCompare = Get-FileList -folderPath $migrationDest

$missingFiles = $sourceFilesCompare | Where-Object { $_ -notin $destFilesCompare}


if ($missingFiles.Count -gt 0) {
    $logPath = "C:\logs\OneDriveMigration\MissingFiles_$(Get-Date -Format "yyyyMMdd_HHmmss")_$user.log"
    $missingFiles | Out-File -filepath $logPath
    [System.Windows.Forms.MessageBox]::Show("Some files may not have copied successfully. Press `"OK`" to open the file copy log. Please contact IT Support if an important file is missing.", 'OneDrive Migration', 'OK', 'Warning')
    # Open the log in Notepad
    Start-Process -FilePath "notepad.exe" -ArgumentList $logPath
}

# Create migration completion flag and placing it in their Home Drive 
# Be sure your users can write to the root of the $userHomeDrive, or change the path to somewhere they can
$flagFileName = "MigrationCompleted_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".flag"
$flagPath = Join-Path -Path $userHomeDrive -ChildPath $flagFileName
New-Item -Path $flagPath -ItemType File


# Create a form
$form = New-Object Windows.Forms.Form
$form.TopMost = $true

# Create a message box with an OK button
[System.Windows.Forms.MessageBox]::Show($form, 'OneDrive Migration was completed. Press OK to view the migrated content.', 'OneDrive Migration Completed', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

write-output $migrationDest

# After OK is pressed, open the windows for all the content

Start-Process "explorer.exe" -ArgumentList $migrationDest


Stop-Transcript
