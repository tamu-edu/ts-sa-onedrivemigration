<#
.SYNOPSIS
  Staff Folder Size Reporting
.DESCRIPTION
  Reads directory of Home Directories/Staff Folders and presets a list of all users with more than 1GB in decending order
.PARAMETER <Parameter_Name>
  No paramters required, adjust script settings in "[Initializations]" section
.INPUTS
  None
.OUTPUTS
  Outputs to console window and prompts to copy results to clipboard
.NOTES
  Version:        0.9
  Author:         Callan Christensen callan@tamu.edu
  Creation Date:  3/27/2024
  Purpose/Change: First general release
  
.EXAMPLE
  .\ComputerLogon-OneDriveMigration.ps1
#>

#----------------------------------------------------------[Declarations]----------------------------------------------------------

# NoModulesRequired

#---------------------------------------------------------[Initializations]--------------------------------------------------------

# Defining the location of home directory or staff folder 
$staffFolderLocation = get-childitem "\\dsa\student affairs\Departments\Information Technology\Staff" -Directory | Sort-Object

# Defining the Array for results collection
$allResults = @()
# Setting Percentage to 0 as a starting value
$progress = 1

#-----------------------------------------------------------[Execution]------------------------------------------------------------

# Looping throught
foreach ($folder in $staffFolderLocation) {
    # Math for percentage at top of screen
    $percent = [Math]::Round((($progress / $staffFolderLocation.length) * 100), 0)
    # Updates progress report
    Write-Progress -Activity "Calculating Staff Folder Sizes" -Status "Percent Complete: $percent%" -PercentComplete $percent

    # Create Folder Object
    $objFSO = New-Object -com  Scripting.FileSystemObject
    $Name = [string] $folder.Name
    $SizeGB = [Math]::Round((($objFSO.GetFolder($folder.FullName).Size) / 1GB), 2)

    $Result = New-Object PSObject
    Add-Member -InputObject $Result -MemberType NoteProperty -Name Name -Value $Name
    Add-Member -InputObject $Result -MemberType NoteProperty -Name Size-GB -Value $SizeGB

    # Add the results from this single loop to the other results
    $allResults += $Result

    # Increment progress counter
    $progress++
}

Write-Progress -Activity "Calculating Staff Folder Sizes" -Completed -Status "Complete"

# Sorts folders in decending size order and then filters out objects smaller than 1GB
$finalOutput = $allResults | Sort-Object Size-GB -Descending | where-object Size-GB -gt "1"

# Outputs the results to the screen
$finalOutput | ft -AutoSize

# Asks user if they would like output copied to clipboard
$clipboardout = Read-Host "Do you want to copy the results to your clipboard? [y/n]"
if ($clipboardout -eq 'y') {
    $finalOutput | ft -AutoSize | out-string | Set-Clipboard
    Write-output "Clipboard Set with Results."
}