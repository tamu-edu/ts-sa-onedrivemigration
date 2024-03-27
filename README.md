# TS-SA-OneDriveMigration  

These scripts are provided to TAMU Technology Services staff for migrating from a legacy home directory to OneDrive. They are designed for a specific migration scenario but may be adapted for similar uses. Refinements and updates will be made to this repository over time.  

## Contributions and Support  
- To suggest refinements, updates, or to report issues, please submit an issue or create a pull request with your correction.  
- For support or questions, contact callan@tamu.edu.  

## Disclaimer  
- Use these scripts at your own risk. Always test scripts in a non-production environment before widespread use.  
- Ensure you meet all prerequisites, such as specific PowerShell version requirements.  
- These scripts contains no capability or code for deleting items.  
- If switching to an updated script, please be sure to redefine all required variables.  

## Acknowledgments  
- Thanks to the TAMU IT community for your interest in these scripts.  
 

Script Notes:  

**ComputerLogon-OneDriveMigration.ps1**  
**- Variables to set:** $userHomeDrive  
**- Use:** GPO User Logon Script  

**Get-StaffFolderSizes.ps1**  
**- Variables to set:** $staffFolderLocation  
**- Use:** Admin reporting tool, outputs a descending list of staff folder sizes  