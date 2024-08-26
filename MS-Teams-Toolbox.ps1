<#
Script Written By Eric Marsi | www.ericmarsi.com | https://www.ericmarsi.com/2023/01/27/microsoft-teams-user-account-provisioning-utility/

ChangeLog-----------------------
v2301.1
   -Initial Release

v2302.1
   -CHANGE - Script Minimum Teams PS Module updated to 4.9.3 from 4.9.1
   -FEATURE - Added PhoneNumberType to the Import CSV & Single User Mode. The script can then be used for Direct Routing, Operator Connect, and Calling Plans customers
    -Supported Values are DirectRouting, CallingPlan, and OperatorConnect
   -FEATURE - Updated Text on PhoneNumber Provisioning to Support the move away from LineURI to PhoneNumber
   -FEATURE - Added a Script GitHub Updater function. If this fails (Firewall Blocking, etc.), existing version continues working

v2310.1_BETA
   -BUG - Updated misc script descriptors and other text objects for accuracy
   -CHANGE - Script Minimum Teams PS Module updated to 5.7.1 from 4.9.3
   -FEATURE - Added support for LocationID in Set-CsPhoneNumberAssignment. This field is optional for DR, required for CP/OC,and requires a new Template CSV
   -FEATURE - Added support for assigning Caller ID Policies (CallingLineIdentity) to users

v2405.1
   -CHANGE - Script Minimum Teams PS Module updated to 6.1.0 from 5.7.1

v2408.1
   -BUG - Fixed Issue with Disconnect Teams PS Function not updating main menu
   -CHANGE - Script Renamed from "Microsoft Teams User Account Provisioning Utility" to "MS Teams Account Provisioning Utility"
   -CHANGE - Script Minimum Teams PS Module updated to 6.4.0 from 6.1.0
   -CHANGE - Reorganized the order of policy assignment to be alphabetical based on PowerShell cmdlet
   -FEATURE - Remove Single User Provisioning Mode, Not Needed/Clumbersome to manage
   -FEATURE - Added support to assign a Call Park, Calling Policy, Voice Application Policy, Voicemail Policy, Shared Calling Policy, and/or a IP Phone Policy to a user
    -Supported Policies: CsCallingLineIdentity, CsOnlineAudioConferencingRoutingPolicy, CsOnlineVoicemailPolicy, CsOnlineVoiceRoutingPolicy, CsTeamsCallingPolicy, CsTeamsCallParkPolicy, CsTeamsEmergencyCallingPolicy, CsTeamsEmergencyCallRoutingPolicy, CsTeamsIPPhonePolicy, CsTeamsSharedCallingRoutingPolicy, CsTeamsVoiceApplicationsPolicy, and CsTenantDialPlan

v2408.2
   -CHANGE - Script Minimum Teams PS Module updated to 6.5.0 from 6.4.0
   -CHANGE - Script Renamed from "MS Teams Account Provisioning Utility" to "MS Teams Toolbox"
   -Change - Change the Script Updater to get the name of largest file in the latest release. This allows for freedom renaming the tool in future releases
    -A Delta updater was released with the old script filename. This allows for all old versions of the script to auto-update going forward.
   -CHANGE - Optimized Provsioning Functions with a new EM-PolicyAssignment function
   -CHANGE - StatusFlag Bits are now set on user enablement for error checking. To be made granular in future feature enhancements
   -CHANGE - General Code Optimizations & code preparation for upcoming features with Exchange Online
   -CHANGE - Cleaned up the script to have a sub function for policy assignement. Makes adding new policies in the future easier and minifies the script.
   -FEATURE - Added a $Script:ConsoleDebugEnable flag to the code header that is enabled by default. This shows or hides the skipped policies and items when provisioning users. Items are still written to log.
   -FEATURE - Added the ability to switch clouds from Commercial Cloud to GCCH, DOD, and China
   -FEATURE - Added the ability to set a tenant id from a customer's verified domain name or a static tenant id. This is used to connect to the tenant as a guest user from another tenant or as a Microsoft Partner
   -FEATURE - Added support for assigning Private Lines to Users
   -FEATURE - Added support for assigning the Survivable Branch Policy (CsTeamsSurvivableBranchAppliancePolicy) to a user
   -FEATURE - Added the ability to Enterprise Voice Enable a User with No PhoneNumber, PrivateLineNumber, PhoneNumberType, or LocationID set.
   -FEATURE - Added the ability to auto-normalize Numbers not stating with a + to hopefully E.164 format - Format is not validated to be proper e.164 to prevent issues
   -FEATURE - Added Support for Importing Excel User Files for all options. This allows users to edit the user data in Excel form and then use it directly in the script
    -Had to fix a bug under PhoneNumber and PrivateLineNumber Assignment that acted different due to the import-excel function in LocationID


**Future Release Things to Add/Change/Fix**
   -BUG - Not Working on a Mac - IsAdmin and File Import Dialog - Have to migrate to PS 7.2 to support this
   -CHANGE - Update Script to Require PowerShell 7.2 for all functions due to Teams PS Module 6.3.0 now supporting the newer release.
   -FEATURE - Add a function to validate that users are ready to be provisioned for CP/OC/DR. Maybe Add a SFB User Prep too but TBD on that.
   -FEATURE - Rewrite line uri assignment/EV Enable under a sub functon (EM-SetCsUserPhoneNumberAssignment)
   -FEATURE - Write-Log of UPN in Separate Column and a Data Column. Maybe a Separate function just for ease of fixing the issue in the future.
   -FEATURE - Provision Teams Rooms Accounts from CSV and Rebrand the script
#>

#Base Script Variables--------------------------------------------------------------------------------------------------------------------------------
    $Script:Name = "MS Teams Toolbox By Eric Marsi"
    $Script:BuildVersion = "2408.2"
    $Script:LogPath = "C:\_Logs\EM-MSTeamsToolbox\"
    $Script:LogFileName = "ScriptLog"
    $Script:TeamsPSMinVer = "6.5.0"
    $Script:ImportExcelPSMinVer = "7.8.9"
    $Script:ConsoleDebugEnable = $True #Variable to enable or disable showing skipped policy assignments in the console log
    $Script:ScriptUpdaterEnabled = $True #Variable to enable or disable the Script GitHub Updater function.
    $Script:ScriptUpdaterGithubRepo = "EricMarsi/MS-Teams-Toolbox"
    #Dont Change:
    $Script:M365EnvironmentNameID = "Commercial Cloud (CC) & Government Cloud (GCC)"
    $Script:TeamsEnvironmentNameID = "TeamsCC-GCC"
    $Script:ExchangeEnvironmentNameID = "O365Default"
    $Script:TeamsSession = $False
    $Script:BetaFlightsEnabled = $False #Variable to Enable Beta Features, Do not change here, activate with activation code from main menu
    $Script:ReqTenantID = "<Not Specified>"
    $Script:ReqTenantDomain = "<Not Specified>"
    $Script:TenantDomain = "<Not Connected>"
    $Script:TenantID = "<Not Connected>"
    $Script:M365Admin = "<Not Connected>"

Clear-Host
$DT = Get-Date -Format "MM/dd/yyyy HH:mm:ss:ffff"
Write-Host "$($Script:Name) v$($Script:BuildVersion) Started at: $($DT)`n" -ForegroundColor Green

#Script Tests-----------------------------------------------------------------------------------------------------------------------------------------
#Verify that the Script is executing as an Administrator
Write-Host "Verifying that the script is executing as an Administrator"
function Test-IsAdmin {
    ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    }
    if (!(Test-IsAdmin)){
        throw "Please run this script as an Administrator!"
    }
    else {
        Write-Host "Pass: The script is executing as an Administrator `n" -ForegroundColor Green
    }

#Create Logging Directory & Determine Log File Name
Write-Host "Verifying Log File Directory Exists"
if(Test-Path "$($Script:LogPath)")
    {
        Write-Host "Pass: Log Folder Already Exists`n" -ForegroundColor Green
    }
else
    {
        try
            {
                New-Item -ItemType Directory -Force -Path $Script:LogPath -ErrorAction Stop | out-null
            }
        catch
            {
                Write-Host "An Unexpected Error occured! The exception caught was $_ " -ForegroundColor Red
                Write-Error "Script Terminating due to error when attempting to create a Log File Directory! " -ErrorAction Stop
            }
        Write-Host "Pass: Log Folder did not exist, but was created`n" -ForegroundColor Green
    }

#Logging Function
$DT2 = $DT -replace ("/","") -replace (":","") -replace (" ","_")
$Script:LogFilePath = $($Script:LogPath) + $($Script:LogFileName) + "_" + $($DT2) + ".csv"
function Write-Log {
        [CmdletBinding()]
        param(
            [Parameter()]
            [ValidateNotNullOrEmpty()]
            [string]$Message,
 
            [Parameter()]
            [ValidateNotNullOrEmpty()]
            [ValidateSet('INFO','WARN','ERR','BLANK')]
            [string]$Severity = 'Info'
        )
        $Severity = $Severity.ToUpper()
        [pscustomobject]@{
            Time = (Get-Date -Format "MM/dd/yyyy HH:mm:ss:ffff")
            Severity = $Severity
            Message = $Message
        } | Export-Csv -Path "$($Script:LogFilePath)" -Append -NoTypeInformation
    }
Write-Log -Severity Info -Message "$($Script:Name) v$($Script:BuildVersion) Started at: $($DT)"
Write-Log -Severity Info -Message "Script is Running as an Admin"
Write-Log -Severity Info -Message "Pass: Log File Directory Check"
Write-Log -Severity Info -Message "Pass: Logging Function Imported"

#Verify that at least PowerShell 5.1 is Installed
Write-Host "Verifying that at least PowerShell 5.1 is Installed"
Write-Log -Severity Info -Message "Verifying that at least PowerShell 5.1 is Installed"
    if([Version]'5.1.00000.000' -GT $PSVersionTable.PSVersion)
    {
        Write-Log -Severity ERR -Message "PowerShell 5.1 is Not Installed!"
        Write-Error "The host must be upgraded to at least PowerShell 5.1! Please Refer to: https://www.ericmarsi.com/2021/02/27/installing-the-microsoft-teams-powershell-module/" -ErrorAction Stop
    }else {
        Write-Log -Severity Info -Message "Pass: At Least PowerShell 5.1 is Installed"
        Write-Host "Pass: The host has at least PowerShell 5.1 Installed" -ForegroundColor Green
    }
Write-Host "***NOTE: This is the last version of this script that will continue to run on PowerShell 5.1. This script WILL require 7.2 in future updates!***`n" -ForegroundColor Yellow

#Verify that the script is executing in the PowerShell Console and not the ISE
Write-Host "Verifying that the script is executing in the PowerShell Console and not the ISE"
Write-Log -Severity Info -Message "Verifying that the script is executing in the PowerShell Console and not the ISE"
    if((Get-Host).Name -eq "ConsoleHost")
    {
        Write-Log -Severity Info -Message "Pass: The script is executing in the PowerShell Console"
        Write-Host "Pass: The script is executing in the PowerShell Console`n" -ForegroundColor Green
    }else {
        Write-Log -Severity ERR -Message "The script is not executing in the PowerShell Console!"
        Write-Error "The script is not executing in the PowerShell Console!" -ErrorAction Stop
    }

function EM-GetLatestGitHubRelease
    {
        $ProgressPreference = 'SilentlyContinue' #Speed Up Invoke-WebRequest
        if ($Script:ScriptUpdaterEnabled = $True)
            {
                Write-Host "Script Updater Enabled, Checking the server for any avaliable updates. Please Standby..."
                Write-Log -Severity Info -Message "Script Updater Enabled, Checking the server for any avaliable updates. Please Standby..."
                $ReleasesURL = "https://api.github.com/repos/$($Script:ScriptUpdaterGithubRepo)/releases"
                try
                    {
                        $ServerVersion = (Invoke-WebRequest $ReleasesURL -ErrorAction Stop | ConvertFrom-Json)[0].tag_name
                        Write-Log -Severity Info -Message "Obtained Latest Script Version from the Server"
                        $GetServerVersion = $True
                    }
                catch
                    {
                        Write-Host "Failed to get latest version from the server. Continuing with the currently installed version. The Error was: $_.`n" -ForegroundColor Yellow
                        Write-Log -Severity ERR -Message "Failed to get latest version from the server. Continuing with the currently installed version. The Error was: $_" 
                        $GetServerVersion = $False
                    }
            
                if ($GetServerVersion -eq $True )
                    {
                        if ([Version]$Script:BuildVersion -lt [Version]$ServerVersion)
                            {
                                Write-Host "This script has an update avaliable. Some features may not work unless you upgrade to the latest version.`n" -ForegroundColor Yellow
                                Write-Log -Severity Info -Message "This script has an update avaliable. Script Version: v$($Script:BuildVersion). Server Version: v$($ServerVersion)"
                                Write-Host "Script Version: v$($Script:BuildVersion)"
                                Write-Host "Server Version: v$($ServerVersion)`n"

                                $UpdateResponse = Read-Host "Would you like to upgrade to the latest version of this script? (Y/N)"

                                if ($UpdateResponse -eq "Y")
                                    {
                                        Write-Host "Downloading & Replacing Script with Server Version"
                                        Write-Log -Severity Info -Message "User Accepted the Update. Downloading & Replacing Script with Server Version"

                                        if ($PSScriptRoot -ne "") #If Function isnt running in a script, throw the download in C:\
                                            {
                                                $DownloadPath = $PSScriptRoot + "\"
                                            }
                                        else
                                            {
                                                $DownloadPath = "C:\"
                                            }
										
										#Get Largest File that is a .ps1 file from the latest release
										$GHAssets = (Invoke-WebRequest $ReleasesURL -ErrorAction Stop | ConvertFrom-Json)[0].assets
										[System.Collections.ArrayList]$GHValues = @()
										foreach ($File in $GHAssets)
											{
												if ($File.Name -match "^*.ps1$")
													{
														$GHValuesOut = New-Object PSCustomObject
														$GHValuesOut | Add-Member -NotePropertyName Name -NotePropertyValue $File.Name
														$GHValuesOut | Add-Member -NotePropertyName Size -NotePropertyValue $File.Size
														$GHValuesOut | Add-Member -NotePropertyName DownloadURL -NotePropertyValue $File.browser_download_url
														$GHValues += $GHValuesOut
													}
											}
										#Parse URL and Filename
										$ScriptDownloadURL = ($GHValues | Sort-Object -Property Size -Descending | Select-Object -First 1).DownloadURL
                                        $TargetScriptName = ($GHValues | Sort-Object -Property Size -Descending | Select-Object -First 1).Name
                                        $NewScriptPath = "$($DownloadPath)$($TargetScriptName)"
                                        
                                        try
                                            {
                                                Invoke-WebRequest -URI $ScriptDownloadURL -Out "$($NewScriptPath)" -ErrorAction Stop
                                                Write-Host "Obtained the latest script version from the server. Relaunching with the updated script. Please Standby...`n" -ForegroundColor Green
                                                Write-Log -Severity Info -Message "Obtained the latest script version from the server. Relaunching with the updated script. Please Standby..."
                                                Start-Sleep 3
                                                $DT = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
                                                Write-Host "$($Script:Name) v$($Script:BuildVersion) Stopped at: $($DT)`n" -ForegroundColor Green
                                                Write-Log -Severity Info -Message "$($Script:Name) v$($Script:BuildVersion) Stopped at: $($DT)"
                                                . $NewScriptPath
                                                Exit
                                            }
                                        catch
                                            {
                                                Write-Host "Failed to get latest script version from the server. Continuing with the currently installed version. The Error was: $_.`n" -ForegroundColor Yellow
                                                Write-Log -Severity ERR -Message "Failed to get latest script version from the server. Continuing with the currently installed version. The Error was: $_."
                                            }
                                    }
                                else
                                    {
                                        Write-Host "User declined the avaliable update. Continuing with the currently installed version.`n" -ForegroundColor Yellow
                                        Write-Log -Severity WARN -Message "User declined the avaliable update. Continuing with the currently installed version."
                                    }
                            }
                        elseif ([Version]$Script:BuildVersion -eq [Version]$ServerVersion)
                            {
                                Write-Host "Pass: The latest version of this script (v$($Script:BuildVersion)) is already installed. No Update Required.`n" -ForegroundColor Green
                                Write-Log -Severity Info -Message "Pass: The latest version of this script (v$($Script:BuildVersion)) is already installed. No Update Required."
                            }
                        else
                            {
                                Write-Host "Pass: The script version (v$($Script:BuildVersion)) is a higher than that on the server (v$($ServerVersion)). No Update Required.`n" -ForegroundColor Green
                                Write-Log -Severity Info -Message "Pass: The script version (v$($Script:BuildVersion)) is a higher than that on the server (v$($ServerVersion)). No Update Required."
                            }
                    }    
            }
        else   
            {
                Write-Host "Script Updater Disabled. Continuing without checking for any avaliable updates`n"
                Write-Log -Severity Info -Message "Script Updater Disabled. Continuing without checking for any avaliable updates"
            }
    }

#Verify that the latest version of the script is installed. If not, ask the user if they would like to install it
EM-GetLatestGitHubRelease

function EM-ValidatePSModule {
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$DisplayName,
        [string]$ModuleName,
        [string]$MinimumVersion
    )
    
    Write-Host "Verifying that at least the $($DisplayName) $($MinimumVersion) PowerShell Module is Installed"
    Write-Log -Severity Info -Message "Verifying that at least the $($DisplayName) $($MinimumVersion) PowerShell Module is Installed"
    try {
            if(Get-Module -ListAvailable $($ModuleName))
            {
                if(!([Version]((Get-Module -ListAvailable $($ModuleName))[0].Version) -ge [Version]$($MinimumVersion)))
                    {
                        Write-Host "The $($DisplayName) PS Module is Out of Date and Needs to Be Updated. Attempting Upgrade" -ForegroundColor Yellow
                        Write-Log -Severity Warn -Message "The $($DisplayName) PS Module is Out of Date and Needs to Be Updated. Attempting Upgrade"
                        try
                            {
                                Update-Module $($ModuleName) -Force -Confirm:$True -ErrorAction Stop
                                if(!([Version]((Get-Module -ListAvailable $($ModuleName))[0].Version) -ge [Version]$($MinimumVersion)))
                                    {
                                        Write-Log -Severity ERR -Message "Script Terminating due to error after upgrading to the latest $($DisplayName) PS Module!"
                                        Write-Error "Script Terminating due to error after upgrading to the latest $($DisplayName) PS Module!" -ErrorAction Stop
                                    }
                            }
                        catch
                            {
                                Write-Host "An Unexpected Error occured! The exception caught was $_ " -ForegroundColor Red
                                Write-Log -Severity ERR -Message "Script Terminating due to error during the $($DisplayName) PS Module Upgrade Test!"
                                Write-Error "Script Terminating due to error during the $($DisplayName) PS Module Upgrade Test!" -ErrorAction Stop
                            }
                    }
                Write-Host "Pass: The Required $($DisplayName) PS Module is Installed`n" -ForegroundColor Green
                Write-Log -Severity Info -Message "Pass: The Required $($DisplayName) PS Module is Installed"
            }else {
                Write-Host "The $($DisplayName) PS Module is not installed. Attempting to Install the $($DisplayName) PS Module"
                Write-Log -Severity Warn -Message "The $($DisplayName) PS Module is not installed. Attempting to Install the $($DisplayName) PS Module"
                Install-Module $($ModuleName) -ErrorAction Stop
                Write-Host ""
                if(!([Version]((Get-Module -ListAvailable $($ModuleName))[0].Version) -ge [Version]$($MinimumVersion)))
                    {
                        Write-Log -Severity ERR -Message "Script Terminating due to error after installing the latest $($DisplayName) PS Module!"
                        Write-Error "Script Terminating due to error after installing the latest $($DisplayName) PS Module!" -ErrorAction Stop
                    }
            }   
        }
    catch {
        Write-Host "An Unexpected Error occured! The exception caught was $_ " -ForegroundColor Red
        Write-Log -Severity ERR -Message "Script Terminating due to error during the $($DisplayName) PS Module Test!"
        Write-Error "Script Terminating due to error during the $($DisplayName) PS Module Test! " -ErrorAction Stop
        }
}

#Verify that the Required PowerShell Modules are installed. If not installed, attempt to install or update
EM-ValidatePSModule -DisplayName "Teams" -ModuleName "MicrosoftTeams" -MinimumVersion $($Script:TeamsPSMinVer)
EM-ValidatePSModule -DisplayName "Import-Excel" -ModuleName "ImportExcel" -MinimumVersion $($Script:ImportExcelPSMinVer) #Special Thanks to @DougCharlesFinke
Import-Module ImportExcel

pause

#ScriptFunctions-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- 
function EM-ConnectTeamsPS
    {
        [CmdletBinding()]
        param(
            [Parameter()]
            [ValidateNotNullOrEmpty()]
            [string]$TenantID,
 
            [Parameter()]
            [ValidateNotNullOrEmpty()]
            [ValidateSet('TeamsCC-GCC','TeamsGCCH','TeamsDOD','TeamsChina')]
            [string]$TeamsEnvironment  = 'TeamsCC-GCC'
        )
                
        Write-Log -Severity Info -Message "Running the EM-ConnectTeamsPS Function"
        try
            {
                Import-Module MicrosoftTeams
                Write-Log -Severity Info -Message "Teams Module Imported"

                if ($Script:TeamsEnvironmentNameID -eq "TeamsCC-GCC" -and $Script:ReqTenantID -eq "<Not Specified>")
                    {
                        $Script:TeamsConnection = Connect-MicrosoftTeams -ErrorAction Stop
                    }
                elseif ($Script:TeamsEnvironmentNameID -eq "TeamsCC-GCC" -and $Script:ReqTenantID -match "^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$")
                    {
                        $Script:TeamsConnection = Connect-MicrosoftTeams -TenantID $Script:ReqTenantID -ErrorAction Stop
                    }
                elseif ($Script:TeamsEnvironmentNameID -ne "TeamsCC-GCC" -and $Script:ReqTenantID -eq "<Not Specified>")
                    {
                        $Script:TeamsConnection = Connect-MicrosoftTeams -TeamsEnvironmentName $Script:TeamsEnvironmentNameID -ErrorAction Stop
                    }
                else #GCCH/DOD/China and Specific Tenant ID
                    {
                        $Script:TeamsConnection = Connect-MicrosoftTeams -TeamsEnvironmentName $Script:TeamsEnvironmentNameID -TenantID $Script:ReqTenantID -ErrorAction Stop
                    }
                                
                #Set Envrionment Information
                try
                    {
                        $Script:TenantDomain = (Get-CsTenant -ErrorAction Stop).SipDomain[0]
                        $Script:M365Admin = $($Script:TeamsConnection).Account
                        $Script:TenantID = $($Script:TeamsConnection).TenantID
                        $Script:TeamsSession = $True
                        Write-Host "Successfully Connected to Microsoft Teams PowerShell`n" -ForegroundColor Green
                        Write-Log -Severity Info -Message "Successfully Connected to Microsoft Teams PowerShell"
                    }
                catch
                    {
                        Write-Error "An error occured while trying to run Get-CsTenant. The Error was: $_`n" -ForegroundColor Red
                        Write-Log -Severity ERR -Message "An error occured while trying to run Get-CsTenant. The Error was: $_"
                        EM-DisconnectTeamsPS
                    }
            }
        catch
            {
                Write-Log -Severity ERR -Message "An Unexpected Error occured when Connecting to Microsoft Teams PowerShell. The Error was: $_`n"
                Write-Host "An Unexpected Error occured when Connecting to Microsoft Teams PowerShell. The Error was: $_" -ForegroundColor Red
            }
    }

function EM-DisconnectTeamsPS
    {
        Write-Log -Severity Info -Message "Running the EM-DisconnectTeamsPS Function"
        if ($Script:TeamsSession -eq $True)
            {
                try{
                    Disconnect-MicrosoftTeams -ErrorAction Stop
                    $Script:TenantDomain = "<Not Connected>"
                    $Script:M365Admin = "<Not Connected>"
                    $Script:TenantID = "<Not Connected>"
                    Write-Log -Severity Info -Message "Disconnected from Microsoft Teams PowerShell"
                    $Script:TeamsSession = $False
                }catch{
                    Write-Log -Severity ERR -Message "An Unexpected Error occured when Disconnecting from Microsoft Teams PowerShell. The Error was: $_"
                    Write-Host "An Unexpected Error occured when Disconnecting from Microsoft Teams PowerShell. The Error was: $_" -ForegroundColor Red
                }    
            }
        else
            {
                Write-Log -Severity Info -Message "The EM-DisconnectTeamsPS Function Found No Active Teams PowerShell Sessions, Continuing..."
            }
    }

function EM-GetDataFile
    {
        Write-Log -Severity Info -Message "Running the EM-GetDataFile Function"
        Write-Host "Please Select the XLSX/XLS/CSV file you wish to import"
        Add-Type -AssemblyName System.Windows.Forms
        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
        $FileBrowser.filter = "Excel or CSV Files (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv|All files (*.*)|*.*"
        [void]$FileBrowser.ShowDialog()
		#Parse for XLSX and CSV
			if (($FileBrowser.FileName -like "*.xlsx") -or ($FileBrowser.FileName -like "*.xls"))
				{
					Write-Log -Severity Info -Message "EM-GetDataFile - XLSX/XLS File Selected"
                    Write-Host "XLSX/XLS File Selected`n" -ForegroundColor Green
					Import-Excel -Path $FileBrowser.FileName
				}
			elseif ($FileBrowser.FileName -like "*.csv") 
				{
					Write-Log -Severity Info -Message "EM-GetDataFile - CSV File Selected"
                    Write-Host "CSV File Selected`n" -ForegroundColor Green
					Import-Csv -Path $FileBrowser.FileName
				}
			elseif ($FileBrowser.FileName -eq "") 
				{
					Write-Log -Severity Info "EM-GetDataFile - No XLSX/XLS/CSV File Was Selected"
                    Write-Error "No XLSX/XLS/CSV File Was Selected!" -ErrorAction Stop
				}
			else
				{
					Write-Log -Severity Info "EM-GetDataFile - Invalid File Type Selected"
                    Write-Error "Invalid File Type Selected!" -ErrorAction Stop
				}
            $Script:CurrentUserCsvPath = $FileBrowser.FileName
            Write-Log -Severity Info "Current User CSV File Path Set to $($Script:CurrentUserCsvPath)"
    }

function EM-PolicyAssignment #Used Inside #EM-ProvisionUsers
    {
        [CmdletBinding()]
        param(
            [Parameter()]
            [ValidateNotNullOrEmpty()]
            [string]$TeamsCmdlet,
            [string]$TeamsCmdletDescription,
            [string]$Identity,
            [string]$PolicyName,
            [string]$StatusFlagBit
        )     
        #Clear Input Attributes
        $CMD = $null

        if (($PolicyName-eq "") -or ($PolicyName -eq "null") -or ($PolicyName -eq $null) -or ($PolicyName -eq "N/A"))
            {
                if ($Script:ConsoleDebugEnable -eq $True ){Write-Host "- Skipping the Assignment of the $($TeamsCmdletDescription) as the Value Provided is NULL" -ForegroundColor Yellow}
                Write-Log -Severity Info -Message "Skipping the Assignment of the $($TeamsCmdletDescription) to $($Identity) as the Value Provided is NULL"  
            }
        else
            {
                try
                    {
                        $CMD = "Grant-$($TeamsCmdlet) -Identity $($Identity) -PolicyName $($PolicyName) -ErrorAction Stop"
                        Invoke-Expression $CMD -ErrorAction Stop
                        Write-Host "- Assigned the $($PolicyName) $($TeamsCmdletDescription) Successfully" -ForegroundColor Green
                        Write-Log -Severity Info -Message "Assigned $($Identity) the $($PolicyName) $($TeamsCmdletDescription) Successfully"
                    }
                catch
                    {
                        Write-Host "- FAILED to Assign the $($PolicyName) $($TeamsCmdletDescription). The Error Was: $_" -ForegroundColor Red
                        Write-Log -Severity ERR -Message "FAILED to Assign $($Identity) the $($PolicyName) $($TeamsCmdletDescription). The Error Was: $_"
                        $Script:ErrorCommands += $CMD
                        $Script:StatusFlags += $StatusFlagBit
                    }
            }
    }

function EM-ProvisionUsers
    {
        Write-Log -Severity Info -Message "Running the EM-ProvisionUsers Function"
        Write-Host "Provisioning $($Script:Count) User(s) for Microsoft Teams Voice. Please Standby...`n"
        Write-Log -Severity Info -Message "Provisioning $($Script:Count) User(s) for Microsoft Teams Voice. Please Standby..."
        $Script:ErrorCommands = $null
        [System.Collections.ArrayList]$Script:ErrorCommands = @()

        foreach ($User in $Script:Users)
            {
                Write-Host "-----------------------------------------------------------------------------------------------"
                Write-Log -Severity Info -Message "-----------------------------------------------------------------------------------------------"
                Write-Host "Provisioning $($User.UserPrincipalName) for Microsoft Teams Voice"
                Write-Log -Severity Info -Message "Provisioning $($User.UserPrincipalName) for Microsoft Teams Voice"
                
                #Null Status Flags Inbetween Each User
                $Script:StatusFlags = 0x0

                #Parse PhoneNumber field to start with a +
                    if ($User.PhoneNumber -match "^\+?(.*)")
                        {
                            $User.PhoneNumber | Select-String -pattern "^\+?(.*)" | foreach-object {$_.line -match "^\+?(.*)" > $nul}
                            $UserPhoneNumberToAssign = "+$($matches[1])"
                        }
                    #Really don't know what the format is but still allow the script user to assign it ¯\_(ツ)_/¯
                    else 
                        {
                            $UserPhoneNumberToAssign = $User.PhoneNumber
                        }

                #Assign a Phone Number to the User
                if (($User.PhoneNumber -eq "") -or ($User.PhoneNumber -eq "null") -or ($User.PhoneNumber -eq $null) -or ($User.PhoneNumber -eq "N/A") -or ($User.PhoneNumberType -eq "") -or ($User.PhoneNumberType -eq "null") -or ($User.PhoneNumberType -eq $null) -or ($User.PhoneNumberType -eq "N/A"))
                    {
                        if ($Script:ConsoleDebugEnable -eq $True ){Write-Host "- Skipping the Assignment of a Phone Number as the Value Provided for PhoneNumber and/or PhoneNumberType is NULL" -ForegroundColor Yellow}
                        Write-Log -Severity Info -Message "Skipping the Assignment of a Phone Number to $($User.UserPrincipalName) as the Value Provided for PhoneNumber and/or PhoneNumberType is NULL"  
                    }
                else
                    {
                        try
                            {
                                if (($User.LocationID -eq "") -or ($User.LocationID -eq $null))
                                    {
                                        Set-CsPhoneNumberAssignment -Identity $User.UserPrincipalName -PhoneNumberType $User.PhoneNumberType -PhoneNumber $UserPhoneNumberToAssign -ErrorAction Stop
                                        Write-Host "- Assigned the $($UserPhoneNumberToAssign) PhoneNumber with a PhoneNumberType of $($User.PhoneNumberType) Successfully" -ForegroundColor Green
                                        Write-Log -Severity Info -Message "Assigned $($User.UserPrincipalName) the $($UserPhoneNumberToAssign) PhoneNumber with a PhoneNumberType of $($User.PhoneNumberType) Successfully"
                                    }
                                else
                                    {
                                        Set-CsPhoneNumberAssignment -Identity $User.UserPrincipalName -PhoneNumberType $User.PhoneNumberType -PhoneNumber $UserPhoneNumberToAssign -LocationID $User.LocationID -ErrorAction Stop
                                        Write-Host "- Assigned the $($UserPhoneNumberToAssign) PhoneNumber with a PhoneNumberType of $($User.PhoneNumberType) and LocationID of $($User.LocationID) Successfully" -ForegroundColor Green
                                        Write-Log -Severity Info -Message "Assigned $($User.UserPrincipalName) the $($UserPhoneNumberToAssign) PhoneNumber with a PhoneNumberType of $($User.PhoneNumberType) and LocationID of $($User.LocationID) Successfully"
                                    }
                            }
                        catch
                            {
                                if (($User.LocationID -eq "") -or ($User.LocationID -eq $null))
                                    {
                                        Write-Host "- FAILED to Assign the $($UserPhoneNumberToAssign) PhoneNumber with a PhoneNumberType of $($User.PhoneNumberType). The Error Was: $_" -ForegroundColor Red
                                        Write-Log -Severity ERR -Message "FAILED to Assign $($User.UserPrincipalName) the $($UserPhoneNumberToAssign) PhoneNumber with a PhoneNumberType of $($User.PhoneNumberType). The Error Was: $_"
                                        $Script:ErrorCommands += "Set-CsPhoneNumberAssignment -Identity $($User.UserPrincipalName) -PhoneNumberType $($User.PhoneNumberType) -PhoneNumber $($UserPhoneNumberToAssign) -ErrorAction Stop"
                                    }
                                else
                                    {
                                        Write-Host "- FAILED to Assign the $($UserPhoneNumberToAssign) PhoneNumber with a PhoneNumberType of $($User.PhoneNumberType) and LocationID of $($User.LocationID). The Error Was: $_" -ForegroundColor Red
                                        Write-Log -Severity ERR -Message "FAILED to Assign $($User.UserPrincipalName) the $($UserPhoneNumberToAssign) PhoneNumber with a PhoneNumberType of $($User.PhoneNumberType) and LocationID of $($User.LocationID). The Error Was: $_"
                                        $Script:ErrorCommands += "Set-CsPhoneNumberAssignment -Identity $($User.UserPrincipalName) -PhoneNumberType $($User.PhoneNumberType) -PhoneNumber $($UserPhoneNumberToAssign) -LocationID $($User.LocationID) -ErrorAction Stop"
                                    }
                                $Script:StatusFlags += 0x1
                            }
                    }
                
                #Parse PrivateLineNumber field to start with a +
                    if ($User.PrivateLineNumber -match "^\+?(.*)")
                        {
                            $User.PrivateLineNumber | Select-String -pattern "^\+?(.*)" | foreach-object {$_.line -match "^\+?(.*)" > $nul}
                            $UserPrivateLineNumberToAssign = "+$($matches[1])"
                        }
                    #Really don't know what the format is but still allow the script user to assign it ¯\_(ツ)_/¯
                    else 
                        {
                            $UserPrivateLineNumberToAssign = $User.PrivateLineNumber
                        }


                #Assign a Private Line to the User
                if (($User.PrivateLineNumber -eq "") -or ($User.PrivateLineNumber -eq "null") -or ($User.PrivateLineNumber -eq $null) -or ($User.PrivateLineNumber -eq "N/A") -or ($User.PhoneNumberType -eq "") -or ($User.PhoneNumberType -eq "null") -or ($User.PhoneNumberType -eq $null) -or ($User.PhoneNumberType -eq "N/A"))
                    {
                        if ($Script:ConsoleDebugEnable -eq $True ){Write-Host "- Skipping the Assignment of a Private Line as the Value Provided for PrivateLineNumber and/or PhoneNumberType is NULL" -ForegroundColor Yellow}
                        Write-Log -Severity Info -Message "Skipping the Assignment of a Private Line to $($User.UserPrincipalName) as the Value Provided for PrivateLineNumber and/or PhoneNumberType is NULL"  
                    }
                else
                    {
                        try
                            {
                                if (($User.LocationID -eq "") -or ($User.LocationID -eq $null))
                                    {
                                        Set-CsPhoneNumberAssignment -Identity $User.UserPrincipalName -PhoneNumberType $User.PhoneNumberType -PhoneNumber $UserPrivateLineNumberToAssign -AssignmentCategory Private -ErrorAction Stop
                                        Write-Host "- Assigned the $($UserPrivateLineNumberToAssign) PrivateLineNumber with a PhoneNumberType of $($User.PhoneNumberType) Successfully" -ForegroundColor Green
                                        Write-Log -Severity Info -Message "Assigned $($User.UserPrincipalName) the $($UserPrivateLineNumberToAssign) PrivateLineNumber with a PhoneNumberType of $($User.PhoneNumberType) Successfully"
                                    }
                                else
                                    {
                                        Set-CsPhoneNumberAssignment -Identity $User.UserPrincipalName -PhoneNumberType $User.PhoneNumberType -PhoneNumber $UserPrivateLineNumberToAssign -AssignmentCategory Private -LocationID $User.LocationID -ErrorAction Stop
                                        Write-Host "- Assigned the $($UserPrivateLineNumberToAssign) PrivateLineNumber with a PhoneNumberType of $($User.PhoneNumberType) and LocationID of $($User.LocationID) Successfully" -ForegroundColor Green
                                        Write-Log -Severity Info -Message "Assigned $($User.UserPrincipalName) the $($UserPrivateLineNumberToAssign) PrivateLineNumber with a PhoneNumberType of $($User.PhoneNumberType) and LocationID of $($User.LocationID) Successfully"
                                    }
                            }
                        catch
                            {
                                if (($User.LocationID -eq "") -or ($User.LocationID -eq $null))
                                    {
                                        Write-Host "- FAILED to Assign the $($UserPrivateLineNumberToAssign) PrivateLineNumber with a PhoneNumberType of $($User.PhoneNumberType). The Error Was: $_" -ForegroundColor Red
                                        Write-Log -Severity ERR -Message "FAILED to Assign $($User.UserPrincipalName) the $($UserPrivateLineNumberToAssign) PrivateLineNumber with a PhoneNumberType of $($User.PhoneNumberType). The Error Was: $_"
                                        $Script:ErrorCommands += "Set-CsPhoneNumberAssignment -Identity $($User.UserPrincipalName) -PhoneNumberType $($User.PhoneNumberType) -PhoneNumber $($UserPrivateLineNumberToAssign) -AssignmentCategory Private -ErrorAction Stop"
                                    }
                                else
                                    {
                                        Write-Host "- FAILED to Assign the $($UserPrivateLineNumberToAssign) PrivateLineNumber with a PhoneNumberType of $($User.PhoneNumberType) and LocationID of $($User.LocationID). The Error Was: $_" -ForegroundColor Red
                                        Write-Log -Severity ERR -Message "FAILED to Assign $($User.UserPrincipalName) the $($UserPrivateLineNumberToAssign) PrivateLineNumber with a PhoneNumberType of $($User.PhoneNumberType) and LocationID of $($User.LocationID). The Error Was: $_"
                                        $Script:ErrorCommands += "Set-CsPhoneNumberAssignment -Identity $($User.UserPrincipalName) -PhoneNumberType $($User.PhoneNumberType) -PhoneNumber $($UserPrivateLineNumberToAssign) -AssignmentCategory Private -LocationID $($User.LocationID) -ErrorAction Stop"
                                    }
                                $Script:StatusFlags += 0x2
                            }
                    }

                #Enterprise Voice Enable Only a User - Used when a user only wants to be EV Enabled, but no DID assigned
                if ((($User.EnterpriseVoiceEnabled -eq "True") -or ($User.EnterpriseVoiceEnabled -eq $True ) -or ($User.EnterpriseVoiceEnabled -eq "Yes")) -and (($User.PhoneNumber -eq "") -or ($User.PhoneNumber -eq "null") -or ($User.PhoneNumber -eq $null) -or ($User.PhoneNumber -eq "N/A")) -and (($User.PrivateLineNumber -eq "") -or ($User.PrivateLineNumber -eq "null") -or ($User.PrivateLineNumber -eq $null) -or ($User.PrivateLineNumber -eq "N/A")) -and (($User.PhoneNumberType -eq "") -or ($User.PhoneNumberType -eq "null") -or ($User.PhoneNumberType -eq $null) -or ($User.PhoneNumberType -eq "N/A")))
                    {
                        try
                            {
                                Set-CsPhoneNumberAssignment -Identity $User.UserPrincipalName -EnterpriseVoiceEnabled $True -ErrorAction Stop
                                Write-Host "- Set EnterpriseVoiceEnabled to TRUE Successfully" -ForegroundColor Green
                                Write-Log -Severity Info -Message "Set EnterpriseVoiceEnabled to TRUE for $($User.UserPrincipalName) Successfully"
                            }
                        catch
                            {
                                Write-Host "- FAILED to set EnterpriseVoiceEnabled to TRUE. The Error Was: $_" -ForegroundColor Red
                                Write-Log -Severity ERR -Message "FAILED to set EnterpriseVoiceEnabled to TRUE for $($User.UserPrincipalName). The Error Was: $_"
                                $Script:ErrorCommands += "Set-CsPhoneNumberAssignment -Identity $($User.UserPrincipalName) -EnterpriseVoiceEnabled $True -ErrorAction Stop"
                                $Script:StatusFlags += 0x4
                            }
                    }
                #User has an assigned DID to either PhoneNumber or PrivateLineNumber and it provisioned successfully. If not successful, follow else statement
                elseif (($Script:StatusFlags -eq 0x0) -and (($User.PhoneNumber -ne "" ) -or ($User.PrivateLineNumber -ne "")) -and (($Script:User.EnterpriseVoiceEnabled -eq "TRUE") -or ($Script:User.EnterpriseVoiceEnabled -eq $True)))
                    {
                            Write-Host "- Set EnterpriseVoiceEnabled to TRUE Successfully" -ForegroundColor Green
                            Write-Log -Severity Info -Message "Set EnterpriseVoiceEnabled to TRUE for $($User.UserPrincipalName) Successfully"
                    }
                #EVDisable Code - Not adding in to the codebase, but keeping here for reference as you should use the mass-disable mode
                #elseif ((($User.EnterpriseVoiceEnabled -eq "False") -or ($User.EnterpriseVoiceEnabled -eq $False ) -or ($User.EnterpriseVoiceEnabled -eq "No")) -and (($User.PhoneNumber -eq "") -or ($User.PhoneNumber -eq "null") -or ($User.PhoneNumber -eq $null) -or ($User.PhoneNumber -eq "N/A")) -and (($User.PrivateLineNumber -eq "") -or ($User.PrivateLineNumber -eq "null") -or ($User.PrivateLineNumber -eq $null) -or ($User.PrivateLineNumber -eq "N/A")) -and (($User.PhoneNumberType -eq "") -or ($User.PhoneNumberType -eq "null") -or ($User.PhoneNumberType -eq $null) -or ($User.PhoneNumberType -eq "N/A")))
                else
                    {

                        if ($Script:ConsoleDebugEnable -eq $True ){Write-Host "- Skipping Enterprise Voice ONLY Enablement as either EnterpriseVoiceEnabled is not TRUE and/or PhoneNumber/Type fields are not NULL." -ForegroundColor Yellow}
                        Write-Log -Severity Info -Message "Skipping the Enterprise Voice ONLY Enablement of $($User.UserPrincipalName) as either EnterpriseVoiceEnabled is not TRUE and/or PhoneNumber, PrivateLineNumber, and/or PhoneNumberType is not NULL."  
                    }

                #Policy Assignment
                EM-PolicyAssignment -TeamsCmdlet "CsCallingLineIdentity" -TeamsCmdletDescription "Caller ID Policy" -Identity $User.UserPrincipalName -PolicyName $User.CsCallingLineIdentity -StatusFlagBit 0x8
                EM-PolicyAssignment -TeamsCmdlet "CsOnlineAudioConferencingRoutingPolicy" -TeamsCmdletDescription "Audio Conferencing Routing Policy" -Identity $User.UserPrincipalName -PolicyName $User.CsOnlineAudioConferencingRoutingPolicy -StatusFlagBit 0x10
                EM-PolicyAssignment -TeamsCmdlet "CsOnlineVoicemailPolicy" -TeamsCmdletDescription "Voicemail Policy" -Identity $User.UserPrincipalName -PolicyName $User.CsOnlineVoicemailPolicy -StatusFlagBit 0x20
                EM-PolicyAssignment -TeamsCmdlet "CsOnlineVoiceRoutingPolicy" -TeamsCmdletDescription "Online Voice Routing Policy" -Identity $User.UserPrincipalName -PolicyName $User.CsOnlineVoiceRoutingPolicy -StatusFlagBit 0x40
                EM-PolicyAssignment -TeamsCmdlet "CsTeamsCallingPolicy" -TeamsCmdletDescription "Calling Policy" -Identity $User.UserPrincipalName -PolicyName $User.CsTeamsCallingPolicy -StatusFlagBit 0x80
                EM-PolicyAssignment -TeamsCmdlet "CsTeamsCallParkPolicy" -TeamsCmdletDescription "Call Park Policy" -Identity $User.UserPrincipalName -PolicyName $User.CsTeamsCallParkPolicy -StatusFlagBit 0x100
                EM-PolicyAssignment -TeamsCmdlet "CsTeamsEmergencyCallingPolicy" -TeamsCmdletDescription "Emergency Calling Policy" -Identity $User.UserPrincipalName -PolicyName $User.CsTeamsEmergencyCallingPolicy -StatusFlagBit 0x200
                EM-PolicyAssignment -TeamsCmdlet "CsTeamsEmergencyCallRoutingPolicy" -TeamsCmdletDescription "Emergency Call Routing Policy" -Identity $User.UserPrincipalName -PolicyName $User.CsTeamsEmergencyCallRoutingPolicy -StatusFlagBit 0x400
                EM-PolicyAssignment -TeamsCmdlet "CsTeamsIPPhonePolicy" -TeamsCmdletDescription "IP Phone Policy" -Identity $User.UserPrincipalName -PolicyName $User.CsTeamsIPPhonePolicy -StatusFlagBit 0x800
                EM-PolicyAssignment -TeamsCmdlet "CsTeamsSharedCallingRoutingPolicy" -TeamsCmdletDescription "Shared Calling Routing Policy" -Identity $User.UserPrincipalName -PolicyName $User.CsTeamsSharedCallingRoutingPolicy -StatusFlagBit 0x1000
                EM-PolicyAssignment -TeamsCmdlet "CsTeamsSurvivableBranchAppliancePolicy" -TeamsCmdletDescription "Survivable Branch Appliance Policy" -Identity $User.UserPrincipalName -PolicyName $User.CsTeamsSurvivableBranchAppliancePolicy -StatusFlagBit 0x2000
                EM-PolicyAssignment -TeamsCmdlet "CsTeamsVoiceApplicationsPolicy" -TeamsCmdletDescription "Voice Applications Policy" -Identity $User.UserPrincipalName -PolicyName $User.CsTeamsVoiceApplicationsPolicy -StatusFlagBit 0x4000
                EM-PolicyAssignment -TeamsCmdlet "CsTenantDialPlan" -TeamsCmdletDescription "Tenant Dial Plan" -Identity $User.UserPrincipalName -PolicyName $User.CsTenantDialPlan -StatusFlagBit 0x8000


                $Script:Count = $Script:Count - 1 #Decrease remaining users count by 1

                if ($StatusFlags -eq 0x0)
                    {
                        Write-Host ""
                        Write-Host "Provisioned $($User.UserPrincipalName) Successfully! $($Script:Count) of $($Script:CountInitial) User(s) Remain...`n" -ForegroundColor Green
                        Write-Log -Severity Info -Message "Provisioned $($User.UserPrincipalName) Successfully! $($Script:Count) of $($Script:CountInitial) User(s) Remain..."
                    }
                else
                    {
                        Write-Host ""
                        Write-Host "One or More Errors Caused Provisioning to Fail for $($User.UserPrincipalName). $($Script:Count) of $($Script:CountInitial) User(s) Remain...`n" -ForegroundColor Red
                        Write-Log -Severity Info -Message "One or More Errors Caused Provisioning to Fail for $($User.UserPrincipalName). $($Script:Count) of $($Script:CountInitial) User(s) Remain..."
                    }
                
            }
    }

    function EM-RemoveAllCsPhoneNumberAssignments
    {
        Write-Log -Severity Info -Message "Running the EM-RemoveAllCsPhoneNumberAssignments Function"
        Write-Host "Removing ALL Phone Numbers from $($Script:Count) Microsoft Teams Voice User(s). Please Standby...`n"
        Write-Log -Severity Info -Message "Removing ALL Phone Numbers from $($Script:Count) Microsoft Teams Voice User(s). Please Standby..."
        $Script:ErrorCommands = $null
        [System.Collections.ArrayList]$Script:ErrorCommands = @()

        foreach ($User in $Script:Users)
            {
                Write-Host "-----------------------------------------------------------------------------------------------"
                Write-Log -Severity Info -Message "-----------------------------------------------------------------------------------------------"
                Write-Host "Attempting to Remove ALL Phone Number(s) from $($User.UserPrincipalName)"
                Write-Log -Severity Info -Message "Attempting to Remove ALL Phone Number(s) from $($User.UserPrincipalName)"
                
                #Null Status Flags Inbetween Each User
                $Script:StatusFlags = 0x0

                #Enterprise Voice Enable Only a User - Used when a user only wants to be EV Enabled, but no DID assigned
                try
					{
						Remove-CsPhoneNumberAssignment -Identity $User.UserPrincipalName -RemoveAll -ErrorAction Stop
						Write-Host "- Successfully Removed ALL Phone Number(s) Assigned" -ForegroundColor Green
						Write-Log -Severity Info -Message "Successfully Removed ALL Phone Number(s) Assigned to $($User.UserPrincipalName)"
					}
				catch
					{
						Write-Host "- FAILED to Remove ALL Phone Number(s) Assigned. The Error Was: $_" -ForegroundColor Red
						Write-Log -Severity ERR -Message "FAILED to Remove ALL Phone Number(s) Assigned to $($User.UserPrincipalName). The Error Was: $_"
						$Script:ErrorCommands += "Remove-CsPhoneNumberAssignment -Identity $($User.UserPrincipalName) -RemoveAll -ErrorAction Stop"
						$Script:StatusFlags += 0x1
					}

                if ($StatusFlags -eq 0x0)
                    {
                        Write-Host ""
                        Write-Host "Provisioned $($User.UserPrincipalName) Successfully! $($Script:Count) of $($Script:CountInitial) User(s) Remain...`n" -ForegroundColor Green
                        Write-Log -Severity Info -Message "Provisioned $($User.UserPrincipalName) Successfully! $($Script:Count) of $($Script:CountInitial) User(s) Remain..."
                    }
                else
                    {
                        Write-Host ""
                        Write-Host "One or More Errors Caused Provisioning to Fail for $($User.UserPrincipalName). $($Script:Count) of $($Script:CountInitial) User(s) Remain...`n" -ForegroundColor Red
                        Write-Log -Severity Info -Message "One or More Errors Caused Provisioning to Fail for $($User.UserPrincipalName). $($Script:Count) of $($Script:CountInitial) User(s) Remain..."
                    }
            }
    }

function EM-RetryProvisioningErrors
    {
        Write-Log -Severity Info -Message "Running the EM-RetryProvisioningErrors Function"
        if (($Script:ErrorCommands).Count -ne 0)
            {
                $Script:ErrorCount = ($Script:ErrorCommands).Count
                $Script:ErrorCountInitial = ($Script:ErrorCommands).Count  
                $Retry = Read-Host "There were $(($Script:ErrorCommands).Count) Command(s) that failed to run properly. Would you like to retry these command(s)? (Y/N)"

                if ($Retry -eq "Y")
                    {
                        Write-Log -Severity Info -Message "EM-RetryProvisioningErrors: User Selected to Retry $(($Script:ErrorCommands).Count) Command(s) that failed to run properly."
                        Write-Host "Retrying $(($Script:ErrorCommands).Count) Command(s) that failed to run properly. Please Standby...`n"
                        Write-Log -Severity Info -Message "Retrying $(($Script:ErrorCommands).Count) Command(s) that failed to run properly. Please Standby..."
                        Write-Host "-----------------------------------------------------------------------------------------------"
                        Write-Log -Severity Info -Message "-----------------------------------------------------------------------------------------------"
                        foreach ($CMD in $Script:ErrorCommands)
                            {
                                $Script:ErrorCount = $Script:ErrorCount - 1
                                try
                                    {
                                        Invoke-Expression $CMD -ErrorAction Stop
                                        Write-Host "The Command: $($CMD) | Completed Successfully. $($Script:ErrorCount) of $($Script:ErrorCountInitial) Command(s) Remain..." -ForegroundColor Green
                                        Write-Log -Severity Info -Message "The Command: $($CMD) | Completed Successfully. $($Script:ErrorCount) of $($Script:ErrorCountInitial) Command(s) Remain..."
                                    }
                                catch
                                    {
                                        Write-Host "The Command: $($CMD) | Failed Again. The Error Was: $_. $($Script:ErrorCount) of $($Script:ErrorCountInitial) Command(s) Remain..." -ForegroundColor Red
                                        Write-Log -Severity ERR -Message "The Command: $($CMD) | Failed Again. The Error Was: $_. $($Script:ErrorCount) of $($Script:ErrorCountInitial) Command(s) Remain..."
                                    }
                            }
                    }
            }
    }
function Get-TenantID
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory)]
        [String]$Domain
    )

    try {
            (Invoke-WebRequest https://login.windows.net/$($Domain)/.well-known/openid-configuration|ConvertFrom-Json).token_endpoint.Split('/')[3]
        }
    catch
        {
            Write-Error "Failed to get the Tenant ID. The exception caught was $_" -ErrorAction Stop
        }
}

function EM-MainMenu
    {
        clear-host
        Write-Log -Severity Info -Message "Presenting the User Main Menu Options"
        Write-Host "$($Script:Name) v$($Script:BuildVersion)`n"
        Write-Host "Environment Information---------------------------------------------------------------------------"
        Write-Host "Microsoft 365 Environment     : "-ForegroundColor Green -NoNewLine
        Write-Host "$($Script:M365EnvironmentNameID)" -ForegroundColor Yellow
        Write-Host "Requested Tenant Domain       : "-ForegroundColor Green -NoNewLine
        Write-Host "$($Script:ReqTenantDomain)" -ForegroundColor Yellow
        Write-Host "Requested Tenant ID           : "-ForegroundColor Green -NoNewLine
        Write-Host "$($Script:ReqTenantID)" -ForegroundColor Yellow
        Write-Host "Connected Tenant Domain       : "-ForegroundColor Green -NoNewLine
        Write-Host "$($Script:TenantDomain)" -ForegroundColor Yellow
        Write-Host "Connected Tenant ID           : "-ForegroundColor Green -NoNewLine
        Write-Host "$($Script:TenantID)" -ForegroundColor Yellow
        Write-Host "M365 Admin Credentials        : "-ForegroundColor Green -NoNewLine
        Write-Host "$($Script:M365Admin)" -ForegroundColor Yellow
        Write-Host "Teams PS Session Active?      : "-ForegroundColor Green -NoNewLine
        Write-Host "$($Script:TeamsSession)" -ForegroundColor Yellow
        Write-Host "Beta Flights Enabled          : "-ForegroundColor Green -NoNewLine
        Write-Host "$($Script:BetaFlightsEnabled)" -ForegroundColor Yellow
        Write-Host "Script Console Debug Enabled  : "-ForegroundColor Green -NoNewLine
        Write-Host "$($Script:ConsoleDebugEnable)" -ForegroundColor Yellow
        Write-Host "Script GitHub Updater Enabled : "-ForegroundColor Green -NoNewLine
        Write-Host "$($Script:ScriptUpdaterEnabled)" -ForegroundColor Yellow
        Write-Host "Script Log File Path          : "-ForegroundColor Green -NoNewLine
        Write-Host "$($Script:LogFilePath)`n`n" -ForegroundColor Yellow
        Write-Host "Admin Connections---------------------------------------------------------------------------------"
        Write-Host " Option 1: Required - Connect to Teams PowerShell" -ForegroundColor Green
        Write-Host " Option 2: Optional - Connect to Exchange PowerShell (Required only for Rooms)" -ForegroundColor Green
        Write-Host " Option 3: Optional - Specify Tenant ID (Guest Access & Microsoft Partners)" -ForegroundColor Green
        Write-Host " Option 4: Optional - Change Microsoft 365 Cloud Environments" -ForegroundColor Green
        Write-Host " Option 9: Disconnect All Admin Connections`n" -ForegroundColor Green

        Write-Host "Script Modes--------------------------------------------------------------------------------------"
        Write-Host " Option 10: Deprecated" -ForegroundColor Green
        Write-Host " Option 11: Provision Multiple User Accounts (XLSX & CSV Import)" -ForegroundColor Green
        Write-Host " Option 12: Bulk Remove ALL CsPhoneNumberAssignments from Users (XLSX & CSV Import)" -ForegroundColor Green
        Write-Host " Option 13: Export User Calling Settings (CSV Import ONLY)" -ForegroundColor Green
        if ($Script:BetaFlightsEnabled -eq $True)
            {
                Write-Host " Option 14: (Beta) Validate Teams Only Users for Readiness (XLSX & CSV Import)`n" -ForegroundColor Green
            }
        else
            {
                Write-Host ""
            }
        Write-Host "Option 99: Terminate this Script`n"-ForegroundColor Red

        #Write Current Environment Variables to the Log File
        [string]$Script:EnvInfo = @()
        $Script:EnvInfo += "Environment Information--------------------------------------------`n"
        $Script:EnvInfo += "Microsoft 365 Environment : $($Script:M365EnvironmentNameID)`n"
        $Script:EnvInfo += "Requested Tenant Domain : $($Script:ReqTenantDomain)`n"
        $Script:EnvInfo += "Requested Tenant ID : $($Script:ReqTenantID)`n"
        $Script:EnvInfo += "Connected Tenant Domain : $($Script:TenantDomain)`n"
        $Script:EnvInfo += "Connected Tenant ID : $($Script:TenantID)`n"
        $Script:EnvInfo += "M365 Admin Credentials : $($Script:M365Admin)`n"
        $Script:EnvInfo += "Teams PS Session Active? : $($Script:TeamsSession)`n"
        $Script:EnvInfo += "Script Beta Flights Enabled : $($Script:BetaFlightsEnabled)`n"
        $Script:EnvInfo += "Script Console Debug Enabled : $($Script:ConsoleDebugEnable)`n"
        $Script:EnvInfo += "Script GitHub Updater Enabled : $($Script:ScriptUpdaterEnabled)`n"
        $Script:EnvInfo += "Script Log File Path : $($Script:LogFilePath)"
        Write-Log -Severity Info -Message $($Script:EnvInfo)
    }

#Main Menu--------------------------------------------------------------------------------------------------------------------------------------------
do{
    EM-MainMenu
    $Confirm1 = Read-Host "Of the above options, what mode would you like to run this script in? (Enter the Option Number)"
    Clear-Host

if ($Confirm1 -eq "1")
    {
        Write-Host "Option 1: Required - Connect to Teams PowerShell Selected. Setting Up Connections...`n"
        Write-Log -Severity Info -Message "Option 1: Required - Connect to Teams PowerShell Selected. Setting Up Connections..."
        Write-Host "Connecting to Microsoft Teams PowerShell`n"
        EM-ConnectTeamsPS
        pause
        Write-Log -Severity Info -Message "Option 1: Required - Connect to Teams PowerShell Complete, Returning to the Main Menu"
    }

elseif ($Confirm1 -eq "2")
    {
        Write-Host "Option 2: Optional - Connect to Exchange PowerShell (Required only for Rooms) Selected. Setting Up Connections...`n"
        Write-Log -Severity Info -Message "Option 2: Optional - Connect to Exchange PowerShell (Required only for Rooms) Selected. Setting Up Connections..."
        Write-Host "Connecting to Microsoft Exchange PowerShell`n"
        Write-Host "Feature Coming Soon -EM"
        ########
        pause
        Write-Log -Severity Info -Message "Option 2: Optional - Connect to Exchange PowerShell (Required only for Rooms) Complete, Returning to the Main Menu"
    }

elseif ($Confirm1 -eq "3")
    {
        Write-Host "Option 3: Optional - Specify Tenant ID (Guest Access & Microsoft Partners)`n"
        Write-Log -Severity Info -Message "Option 3: Optional - Specify Tenant ID (Guest Access & Microsoft Partners)"
        Write-Host "Please specify a Tenant ID or a verified domain in the Microsoft 365 Tenant that you are wanting to"
        Write-Host "connect to, else leave this field blank to reset to connection defaults`n"
        $Script:ReqTenantDomain = Read-Host "Tenant ID or Domain"
        if ($Script:ReqTenantDomain -eq "")
            {
                $Script:ReqTenantID = "<Not Specified>"
                $Script:ReqTenantDomain = "<Not Specified>"
                Write-Host ""
                Write-Host "Set the Requested Tenant ID and Requested Tenant Domain to Defaults`n" -ForegroundColor Green
            }
        elseif ($Script:ReqTenantDomain -match "^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$")
            {
                #Static Assignment with an already known tenant ID
                $Script:ReqTenantID = $Script:ReqTenantDomain
                $Script:ReqTenantDomain = "<Not Specified>"
                Write-Host ""
                Write-Host "User Provided a Tenant ID of $($Script:ReqTenantID)`n" -ForegroundColor Green
            }
        else
            {
                
                try
                    {
                        $Script:ReqTenantID = Get-TenantID -Domain $Script:ReqTenantDomain -ErrorAction Stop
                        Write-Host ""
                        Write-Host "Retrived the Tenant ID of $($Script:ReqTenantID) for $($Script:ReqTenantDomain)`n" -ForegroundColor Green
                    }
                catch
                    {
                        Write-Host "FAILED to retrieve the Tenant ID for $($Script:ReqTenantDomain). The Error was $_`n" -ForegroundColor Red
                        $Script:ReqTenantID = "<Not Specified>"
                        $Script:ReqTenantDomain = "<Not Specified>"
                        Write-Host ""
                        Write-Host "Set the Requested Tenant ID and Requested Tenant Domain to Defaults`n" -ForegroundColor Green
                    }
            }
        EM-DisconnectTeamsPS

        pause
        Write-Log -Severity Info -Message "Option 3: Optional - Specify Tenant ID (Guest Access & Microsoft Partners) Complete, Returning to the Main Menu"  
    }

elseif ($Confirm1 -eq "4")
    {
        Write-Host "Option 4: Optional - Change Microsoft 365 Cloud Environments Selected`n"
        Write-Log -Severity Info -Message "Option 4: Optional - Change Microsoft 365 Cloud Environments Selected"

        Write-Host "Please select the number from the list below that is Microsoft 365 Cloud Environment you want to connect to:`n"
        Write-Host "1: "-ForegroundColor Green -NoNewLine
        Write-Host "CC-GCC - Commercial Cloud (CC) & Government Cloud (GCC)"
        Write-Host "2: "-ForegroundColor Green -NoNewLine
        Write-Host "CC-Germany - Commercial Cloud (CC) Teams & O365Germany Exchange"
        Write-Host "3: "-ForegroundColor Green -NoNewLine
        Write-Host "GCCH - US Goverment High Cloud (GCCH)"
        Write-Host "4: "-ForegroundColor Green -NoNewLine
        Write-Host "DOD - US Department of Defense (DOD)"
        Write-Host "5: "-ForegroundColor Green -NoNewLine
        Write-Host "China-21Vianet - Microsoft China Operated By 21Vianet`n"

        $Script:M365EnvironmentRaw = Read-Host "M365 Environment"
        Write-Host""

        if ($Script:M365EnvironmentRaw -eq 1)
            {
                $Script:M365EnvironmentNameID = "Commercial Cloud (CC) & Government Cloud (GCC)"
                $Script:TeamsEnvironmentNameID = "TeamsCC-GCC"
                $Script:ExchangeEnvironmentNameID = "O365Default"
            }
        elseif ($Script:M365EnvironmentRaw -eq 2)
            {
                $Script:M365EnvironmentNameID = "Commercial Cloud (CC) Teams & O365Germany Exchange"
                $Script:TeamsEnvironmentNameID = "TeamsCC-GCC"
                $Script:ExchangeEnvironmentNameID = "O365GermanyCloud"
            }
        elseif ($Script:M365EnvironmentRaw -eq 3)
            {
                $Script:M365EnvironmentNameID = "US Goverment High Cloud (GCCH)"
                $Script:TeamsEnvironmentNameID = "TeamsGCCH"
                $Script:ExchangeEnvironmentNameID = "O365USGovGCCHigh"
            }
        elseif ($Script:M365EnvironmentRaw -eq 4)
            {
                $Script:M365EnvironmentNameID = "US Department of Defense (DOD)"
                $Script:TeamsEnvironmentNameID = "TeamsDOD"
                $Script:ExchangeEnvironmentNameID = "O365USGovDoD"
            }
        elseif ($Script:M365EnvironmentRaw -eq 5)
            {
                $Script:M365EnvironmentNameID = "Microsoft China Operated By 21Vianet"
                $Script:TeamsEnvironmentNameID = "TeamsChina"
                $Script:ExchangeEnvironmentNameID = "O365China"
            }
        else
            {
                Write-Host "Invalid Selection`n" -ForegroundColor Red
            }
        Write-Host "Microsoft 365 Cloud Environment is set to $($Script:M365EnvironmentNameID)`n" -ForegroundColor Green
        EM-DisconnectTeamsPS

        pause
        Write-Log -Severity Info -Message "Option 4: Optional - Change Microsoft 365 Cloud Environments Complete, Returning to the Main Menu"  
    }

elseif ($Confirm1 -eq "9")
    {
        Write-Host "Option 9: Disconnect All Admin Connections Selected. Closing Connections...`n"
        Write-Log -Severity Info -Message "Option 9: Disconnect All Admin Connections Selected. Closing Connections..."
        EM-DisconnectTeamsPS
        Write-Log -Severity Info -Message "Clearing all Admin Connection Variables"
        $Script:TeamsConnection = $null
        $Script:TenantDomain = "<Not Connected>"
        $Script:TenantID = "<Not Connected>"
        $Script:M365Admin = "<Not Connected>"
        Write-Log -Severity Info -Message "All Admin Connection Variables Cleared"
        Write-Log -Severity Info -Message "Option 9: Disconnect All Admin Connections Complete, Returning to the Main Menu"
    }

elseif ($Confirm1 -eq "10")
    {
        Write-Host "Option 10: Deprecated Selected`n"
        Write-Log -Severity Info -Message "Option 10: Deprecated Selected"

        Write-Host "This feature has been deprecated. Please use Option 11: Provision Multiple User Accounts (XLSX & CSV Import)`n"

        pause
        Write-Log -Severity Info -Message "Option 10: Deprecated Complete, Returning to the Main Menu"
    }

elseif ($Confirm1 -eq "11")
    {
        Write-Host "Option 11: Provision Multiple User Accounts (XLSX & CSV Import) Selected`n"
        Write-Log -Severity Info -Message "Option 11: Provision Multiple User Accounts (XLSX & CSV Import) Selected"

        #Ensure Teams PS Admin Connection is Setup
        if ($TeamsSession -ne $True)
            {
                Write-Host "Teams PowerShell Session is Not Active. Setting Up the Needed Admin Connection`n" -ForegroundColor Yellow
                Write-Log -Severity Info -Message "Teams PowerShell Session is Not Active. Setting Up the Needed Admin Connection"
                EM-ConnectTeamsPS
            }

        #Ensure Input Variables are Null
        $Confirmation = $null
        $Script:Users = $null
        [System.Collections.ArrayList]$Script:Users = @(EM-GetDataFile)
        $Script:Count = $null
        $Script:CountInitial = $null
        $Script:Count = $Script:Users.Count
        $Script:CountInitial = $Script:Users.Count

        $Confirmation = Read-Host "Are you sure that you want to provision $Script:Count User(s) for Microsoft Teams Voice? (Y/N)"

        if ($Confirmation -eq "Y")
            {
                EM-ProvisionUsers
                Write-Host "-----------------------------------------------------------------------------------------------"
                Write-Log -Severity Info -Message "-----------------------------------------------------------------------------------------------"
                EM-RetryProvisioningErrors
                
            }
        else
            {
                Write-Host "Operator Canceled the User Provisioning Operation" -ForegroundColor Yellow
                Write-Log -Severity WARN -Message "Operator Canceled the User Provisioning Operation"
            }

        pause
        Write-Log -Severity Info -Message "Option 10: Provision Multiple User Accounts (XLSX & CSV Import) Complete, Returning to the Main Menu"
    }

elseif ($Confirm1 -eq "12")
    {
        Write-Host "Option 12: Bulk Remove ALL CsPhoneNumberAssignments from Users (XLSX & CSV Import) Selected`n"
        Write-Log -Severity Info -Message "Option 12: Bulk Remove ALL CsPhoneNumberAssignments from Users (XLSX & CSV Import) Selected"

        #Ensure Teams PS Admin Connection is Setup
        if ($TeamsSession -ne $True)
            {
                Write-Host "Teams PowerShell Session is Not Active. Setting Up the Needed Admin Connection`n" -ForegroundColor Yellow
                Write-Log -Severity Info -Message "Teams PowerShell Session is Not Active. Setting Up the Needed Admin Connection"
                EM-ConnectTeamsPS
            }

        #Ensure Input Variables are Null
        $Confirmation = $null
        $Script:Users = $null
        [System.Collections.ArrayList]$Script:Users = @(EM-GetDataFile)
        $Script:Count = $null
        $Script:CountInitial = $null
        $Script:Count = $Script:Users.Count
        $Script:CountInitial = $Script:Users.Count

        $Confirmation = Read-Host "Are you sure that you want to remove ALL phone numbers from $Script:Count User(s) for Microsoft Teams Voice? (Y/N)"

        if ($Confirmation -eq "Y")
            {
                EM-RemoveAllCsPhoneNumberAssignments
                Write-Host "-----------------------------------------------------------------------------------------------"
                Write-Log -Severity Info -Message "-----------------------------------------------------------------------------------------------"
                EM-RetryProvisioningErrors
                
            }
        else
            {
                Write-Host "Operator Canceled the User Provisioning Operation" -ForegroundColor Yellow
                Write-Log -Severity WARN -Message "Operator Canceled the User Provisioning Operation"
            }

        pause
        Write-Log -Severity Info -Message "Option 12: Bulk Remove ALL CsPhoneNumberAssignments from Users (XLSX & CSV Import) Complete, Returning to the Main Menu"
    }

elseif ($Confirm1 -eq "13")
    {
        Write-Host "Option 13: Export User Calling Settings (XLSX & CSV Import) Selected`n"
        Write-Log -Severity Info -Message "Option 13: Export User Calling Settings (XLSX & CSV Import) Selected"
    
        #Ensure Teams PS Admin Connection is Setup
        if ($TeamsSession -ne $True)
            {
                Write-Host "Teams PowerShell Session is Not Active. Setting Up the Needed Admin Connection`n" -ForegroundColor Yellow
                Write-Log -Severity Info -Message "Teams PowerShell Session is Not Active. Setting Up the Needed Admin Connection"
                EM-ConnectTeamsPS
            }
    
        #Ensure Input Variables are Null
        $Confirmation = $null
        $Script:Users = $null
        $output = $null

        #Get Data
        [System.Collections.ArrayList]$Script:Users = @(EM-GetDataFile)
        $Script:CountInitial = $Script:Users.Count
        $Script:Count = $Script:Users.Count
        
        #Process Data
        Write-Host "Gathering Teams Calling Settings for $($Script:CountInitial) Users. Please Standby...`n"
        Write-Log -Severity Info -Message "Gathering Teams Calling Settings for $($Script:CountInitial) Users"
        [System.Collections.ArrayList]$Output = @()
        foreach ($User in $Script:Users )
            {
                try
                    {
                        $Output += Get-CsUserCallingSettings -Identity $User.UserPrincipalName -ErrorAction Stop | Select-Object SipUri,IsForwardingEnabled,ForwardingTarget,ForwardingTargetType,ForwardingType,IsUnansweredEnabled,UnansweredDelay,UnansweredTarget,UnansweredTargetType
                        $Script:Count -= 1
                        Write-Host "Gathered User Calling Settings for $($User.UserPrincipalName). $($Script:Count) of $($Script:CountInitial) Users Remaining..." -ForegroundColor Green
                        Write-Log -Severity Info -Message "Gathered User Calling Settings for $($User.UserPrincipalName). $($Script:Count) of $($Script:CountInitial) Users Remaining..."
                    }
                catch
                    {
                        $Script:Count -= 1
                        Write-Host "FAILED to Gathered User Calling Settings for $($User.UserPrincipalName). The Error was: $_. $($Script:Count) of $($Script:CountInitial) Users Remaining..." -ForegroundColor Red
                        Write-Log -Severity ERR -Message "FAILED to Gathered User Calling Settings for $($User.UserPrincipalName). The Error was: $_. $($Script:Count) of $($Script:CountInitial) Users Remaining..."
                    }  
            }

        #Export Collected Data
        Write-Host ""
        Write-Host "Please select the folder where the list of Teams User Calling Settings (CSV) should be saved to"
        Add-Type -AssemblyName System.Windows.Forms
        $FileBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        [void]$FileBrowser.ShowDialog()
        $Path = $FileBrowser.SelectedPath 

        if ($Path -ne "")
            {
                $DT3 = Get-Date -Format "MMddyyyy_HHmmssffff"
                $ExportPath = $($Path) + "\TeamsUserCallingSettings_" + $($DT3) + ".csv"
                Write-Log -Severity Info -Message "Saving Exported Data to: $($ExportPath)"
                try
                    {
                        $Output | Export-Csv -NoTypeInformation -Path $ExportPath -ErrorAction Stop
                        Write-Host "Saved Exported Data to: $($ExportPath)" -ForegroundColor Green
                        Write-Log -Severity Info -Message "Saved Exported Data to: $($ExportPath)"
                    }
                catch
                    {
                        Write-Host "FAILED to save Exported Data to: $($ExportPath). The Error Was: $_" -ForegroundColor Green
                        Write-Log -Severity Info -Message "FAILED to save Exported Data to: $($ExportPath). The Error Was: $_" 
                    }
            }
        else
            {
                Write-Host "No Valid Path Selected. Returning to the Main Menu" -ForegroundColor Yellow
                Write-Log -Severity WARN -Message "No Valid Path Selected. Returning to the Main Menu" -ForegroundColor Yellow
            }

        pause
        Write-Log -Severity Info -Message "Option 13: Export User Calling Settings (XLSX & CSV Import) Complete, Returning to the Main Menu"
    }

elseif ($Confirm1 -eq "14")
    {
        if ($Script:BetaFlightsEnabled -eq $True)
            {
                Write-Host "Option 14: (Beta) Validate Teams Only Users for Readiness (XLSX & CSV Import) Selected"
                Write-Log -Severity Info -Message "Option 14: (Beta) Validate Teams Only Users for Readiness (XLSX & CSV Import) Selected"
            
                #Ensure Teams PS Admin Connection is Setup
                if ($TeamsSession -ne $True)
                    {
                        Write-Host "Teams PowerShell Session is Not Active. Setting Up the Needed Admin Connection`n" -ForegroundColor Yellow
                        Write-Log -Severity Info -Message "Teams PowerShell Session is Not Active. Setting Up the Needed Admin Connection"
                        EM-ConnectTeamsPS
                    }
            
                #Ensure Input Variables are Null
                $Confirmation = $null
                $Script:Users = $null
                $output = $null

                #Get Data
                [System.Collections.ArrayList]$Script:Users = @(EM-GetDataFile)
                $Script:CountInitial = $Script:Users.Count
                $Script:Count = $Script:Users.Count
                
                #Process Data
                Write-Host "Work In Progress for Future Release"

                pause
                Write-Log -Severity Info -Message "Option 14: (Beta) Validate Teams Only Users for Readiness (XLSX & CSV Import) Complete, Returning to the Main Menu"
            }
        else
            {
                Write-Host "User not Authorized for this task!!!" -ForegroundColor Red
                Write-Log -Severity WARN -Message "User not Authorized for this task!!!"
                pause
                Write-Log -Severity Info -Message "Option 14: Returning to the Main Menu"
            }
    }

elseif ($Confirm1 -eq "933")
    {
        $Script:BetaFlightsEnabled = $True
        Write-Host "Beta Flights Enabled for This Session!" -ForegroundColor Green
        Write-Log -Severity Info -Message "Beta Flights Enabled for This Session!"
        pause
        Write-Log -Severity Info -Message "Option 933 Easter Egg: Enable Beta Flights Complete, Returning to the Main Menu"
    }

else
    {
        if ($Confirm1 -eq "99")
            {
                Write-Host "Script Terminated by User" -ForegroundColor Yellow
                Write-Log -Severity Info -Message "Script Terminated by User"
            }
        elseif ($Confirm1 -ne "")
            {
                Write-Host "Invalid Mode Selected" -ForegroundColor Yellow
                Write-Log -Severity Info -Message "Invalid Mode Selected"
                $Confirm1 = "99"
            }
        else
            {
                Write-Host "No Mode Selected" -ForegroundColor Yellow
                Write-Log -Severity Info -Message "No Mode Selected"
                #Disabling Invalid Mode Due to Keyboard Inputs potentially causing issues. Keeping code here
                #$Confirm1 = "99"
            }
    }
}
while ($Confirm1 -ne "99") 

EM-DisconnectTeamsPS


$DT = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
Write-Host "$($Script:Name) v$($Script:BuildVersion) Stopped at: $($DT)`n" -ForegroundColor Green
Write-Log -Severity Info -Message "$($Script:Name) v$($Script:BuildVersion) Stopped at: $($DT)"