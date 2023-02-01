#Script Written By Eric Marsi | www.ericmarsi.com
#Base Script Variables--------------------------------------------------------------------------------------------------------------------------------
    $Script:Name = "Microsoft Teams User Account Provisioning Utility By Eric Marsi"
    $Script:BuildVersion = "v2301.1"
    $Script:LogPath = "C:\_Logs\EM-MSTeamsUserAccountProvUtil\"
    $Script:LogFileName = "ScriptLog"
    $Script:TeamsPSMinVer = "4.9.1"
    $Script:TeamsSession = $False
    $Script:BetaFlightsEnabled = $False #Variable to Enable Beta Features, Do not change here, activate with activation code from main menu
    $Script:TeamsConnection = "<Not Set>"
    $Script:TenantDomain = "<Not Set>"
    $Script:TenantID = "<Not Set>"
    $Script:M365Admin = "<Not Set>"

Clear-Host
$DT = Get-Date -Format "MM/dd/yyyy HH:mm:ss:ffff"
Write-Host "$($Script:Name) $($Script:BuildVersion) Started at: $($DT)`n" -ForegroundColor Green

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
Write-Log -Severity Info -Message "$($Script:Name) $($Script:BuildVersion) Started at: $($DT)"
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
        Write-Host "Pass: The host has at least PowerShell 5.1 Installed`n" -ForegroundColor Green
    }

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

pause

#ScriptFunctions--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function EM-MainMenu
    {
        clear-host
        Write-Log -Severity Info -Message "Presenting the User Main Menu Options"
        Write-Host "$($Script:Name) $($Script:BuildVersion)`n"
        Write-Host "Environment Information---------------------------------------------------------------------------"
        Write-Host "Tenant Domain               : "-ForegroundColor Green -NoNewLine
        Write-Host "$($Script:TenantDomain)" -ForegroundColor Yellow
        Write-Host "Tenant ID                   : "-ForegroundColor Green -NoNewLine
        Write-Host "$($Script:TenantID)" -ForegroundColor Yellow
        Write-Host "M365 Admin Credentials      : "-ForegroundColor Green -NoNewLine
        Write-Host "$($Script:M365Admin)" -ForegroundColor Yellow
        Write-Host "Teams PS Session Active?    : "-ForegroundColor Green -NoNewLine
        Write-Host "$($Script:TeamsSession)" -ForegroundColor Yellow
        Write-Host "Beta Flights Enabled        : "-ForegroundColor Green -NoNewLine
        Write-Host "$($Script:BetaFlightsEnabled)" -ForegroundColor Yellow
        Write-Host "Script Log File Path        : "-ForegroundColor Green -NoNewLine
        Write-Host "$($Script:LogFilePath)`n" -ForegroundColor Yellow
        Write-Host "Admin Connections (Required)----------------------------------------------------------------------"
        Write-Host "Option 1: Setup Admin Connections" -ForegroundColor Green
        Write-Host "Option 2: Close all Admin Connections`n" -ForegroundColor Green
        Write-Host "Script Modes--------------------------------------------------------------------------------------"
        Write-Host "Option 10: Provision a Single User Account" -ForegroundColor Green
        Write-Host "Option 11: Provision Multiple User Accounts (CSV Import)" -ForegroundColor Green
        Write-Host "Option 12: Export User Calling Settings (CSV Import)" -ForegroundColor Green
        if ($Script:BetaFlightsEnabled -eq $True)
            {
                Write-Host "Option 13: (Beta) Validate Teams Only Users for Readiness (CSV Import)`n" -ForegroundColor Green
            }
        else
            {
                Write-Host ""
            }
        Write-Host "Option 99: Terminate this Script`n"-ForegroundColor Red

        #Write Current Environment Variables to the Log File
        [string]$Script:EnvInfo = @()
        $Script:EnvInfo += "Environment Information--------------------------------------------`n"
        $Script:EnvInfo += "Tenant Domain : $($Script:TenantDomain)`n"
        $Script:EnvInfo += "Tenant ID : $($Script:TenantID)`n"
        $Script:EnvInfo += "M365 Admin Credentials : $($Script:M365Admin)`n"
        $Script:EnvInfo += "Teams PS Session Active? : $($Script:TeamsSession)`n"
        $Script:EnvInfo += "Beta Flights Enabled : $($Script:BetaFlightsEnabled)`n"
        $Script:EnvInfo += "Script Log File Path : $($Script:LogFilePath)"
        Write-Log -Severity Info -Message $($Script:EnvInfo)
    }
 
function EM-ConnectTeamsPS
    {
        Write-Log -Severity Info -Message "Running the EM-ConnectTeamsPS Function"
            try{
                Import-Module MicrosoftTeams
                Write-Log -Severity Info -Message "Teams Module Imported"
                $Script:TeamsConnection = Connect-MicrosoftTeams -ErrorAction Stop
                Write-Log -Severity Info -Message "Connected to Microsoft Teams PowerShell"
                $Script:TenantDomain = $(($Script:TeamsConnection).Account -split "@")[1]
                $Script:M365Admin = $($Script:TeamsConnection).Account
                $Script:TenantID = $($Script:TeamsConnection).TenantID
                $Script:TeamsSession = $True
            }catch{
                Write-Log -Severity ERR -Message "An Unexpected Error occured! The exception caught was $_"
                Write-Log -Severity ERR -Message "An Unexpected Error occured when Connecting to Microsoft Teams PowerShell!"
                Write-Output "An Unexpected Error occured! The exception caught was $_"
                Write-Error "An Unexpected Error occured when Connecting to Microsoft Teams PowerShell!" -ErrorAction Stop
            }
    }

function EM-DisconnectTeamsPS
    {
        Write-Log -Severity Info -Message "Running the EM-DisconnectTeamsPS Function"
        if ($Script:TeamsSession -eq $True)
            {
                try{
                    Disconnect-MicrosoftTeams -ErrorAction Stop
                    Write-Log -Severity Info -Message "Disconnected from Microsoft Teams PowerShell"
                    $Script:TeamsSession = $False
                }catch{
                    Write-Log -Severity ERR -Message "An Unexpected Error occured! The exception caught was $_"
                    Write-Log -Severity ERR -Message "An Unexpected Error occured when disconnecting from Microsoft Teams PowerShell!"
                    Write-Output "An Unexpected Error occured! The exception caught was $_"
                    Write-Host "An Unexpected Error occured when disconnecting from Microsoft Teams PowerShell!" -ForegroundColor Red
                }    
            }
        else
            {
                Write-Log -Severity Info -Message "The EM-DisconnectTeamsPS Function Found No Active Teams PowerShell Sessions, Continuing..."
            }
    }

function EM-GetUsersCsv
    {
        Write-Log -Severity Info -Message "Running the EM-GetUsersCsv Function"
        Write-Host "Please Select the CSV containing a list of users"
        Add-Type -AssemblyName System.Windows.Forms
        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
        $FileBrowser.filter = "Csv (*.csv)| *.csv"
        [void]$FileBrowser.ShowDialog()
            if ($FileBrowser.FileName -eq "")
                {
                    Write-Log -Severity ERR -Message "No CSV File Was Selected!"
                    Write-Error "No CSV File Was Selected!" -ErrorAction Stop
                }
            else
                {
                    Write-Log -Severity Info -Message "Pass: CSV File Selected"
                    Write-Host "Pass: CSV File Selected`n" -ForegroundColor Green
                }
            [System.Collections.ArrayList]$Script:Users = @(Import-Csv -Path $FileBrowser.FileName)
            $Script:CurrentUserCsvPath = $FileBrowser.FileName
            Write-Log -Severity Info "Current User CSV File Path Set to $($Script:CurrentUserCsvPath)"
    }

function EM-ProvisionUsers
    {
        Write-Log -Severity Info -Message "Running the EM-ProvisionUsers Function"
        Write-Host "Provisioning $($Script:Count) User(s) for Microsoft Teams Voice. Please Standby...`n"
        Write-Log -Severity Info -Message "Provisioning $($Script:Count) User(s) for Microsoft Teams Voice. Please Standby..."
        [System.Collections.ArrayList]$Script:ErrorCommands = @()

        foreach ($User in $Script:Users)
            {
                Write-Host "-----------------------------------------------------------------------------------------------"
                Write-Log -Severity Info -Message "-----------------------------------------------------------------------------------------------"
                Write-Host "Provisioning $($User.UserPrincipalName) for Microsoft Teams Voice"
                Write-Log -Severity Info -Message "Provisioning $($User.UserPrincipalName) for Microsoft Teams Voice"

                #Assign the Phone Number to the User
                if (($User.PhoneNumber -eq "") -or ($User.PhoneNumber -eq "null") -or ($User.PhoneNumber -eq $null) -or ($User.PhoneNumber -eq "N/A"))
                    {
                        Write-Host "- Skipping the Assignment of a Phone Number as the Value Provided is NULL" -ForegroundColor Yellow
                        Write-Log -Severity Info -Message "Skipping the Assignment of a Phone Number to $($User.UserPrincipalName) as the Value Provided is NULL"  
                        $UserLineURISuccess = $True
                    }
                else
                    {
                        try
                            {
                                Set-CsPhoneNumberAssignment -Identity $User.UserPrincipalName -PhoneNumberType DirectRouting -PhoneNumber $User.PhoneNumber -ErrorAction Stop
                                Write-Host "- Assigned the tel:$($User.PhoneNumber) LineURI Successfully" -ForegroundColor Green
                                Write-Log -Severity Info -Message "Assigned $($User.UserPrincipalName) the tel:$($User.PhoneNumber) LineURI Successfully"
                                $UserLineURISuccess = $True
                            }
                        catch
                            {
                                Write-Host "- FAILED to Assign the tel:$($User.PhoneNumber) LineURI. The Error Was: $_" -ForegroundColor Red
                                Write-Log -Severity ERR -Message "FAILED to Assign $($User.UserPrincipalName) the tel:$($User.PhoneNumber) LineURI. The Error Was: $_"
                                $Script:ErrorCommands += "Set-CsPhoneNumberAssignment -Identity $($User.UserPrincipalName) -PhoneNumberType DirectRouting -PhoneNumber $($User.PhoneNumber) -ErrorAction Stop"
                                $UserLineURISuccess = $False
                            }
                    }

                #Assign the OVRP to the User
                if (($User.OnlineVoiceRoutingPolicy -eq "") -or ($User.OnlineVoiceRoutingPolicy -eq "null") -or ($User.OnlineVoiceRoutingPolicy -eq $null) -or ($User.OnlineVoiceRoutingPolicy -eq "N/A"))
                    {
                        Write-Host "- Skipping the Assignment of a Online Voice Routing Policy as the Value Provided is NULL" -ForegroundColor Yellow
                        Write-Log -Severity Info -Message "Skipping the Assignment of a Online Voice Routing Policy to $($User.UserPrincipalName) as the Value Provided is NULL"  
                        $UserOVRPSuccess = $True
                    }
                else
                    {
                        try
                            {
                                Grant-CsOnlineVoiceRoutingPolicy -Identity $User.UserPrincipalName -PolicyName $User.OnlineVoiceRoutingPolicy -ErrorAction Stop
                                Write-Host "- Assigned the $($User.OnlineVoiceRoutingPolicy) Online Voice Routing Policy Successfully" -ForegroundColor Green
                                Write-Log -Severity Info -Message "Assigned $($User.UserPrincipalName) the $($User.OnlineVoiceRoutingPolicy) Voice Routing Policy Successfully"
                                $UserOVRPSuccess = $True
                            }
                        catch
                            {
                                Write-Host "- FAILED to Assign the $($User.OnlineVoiceRoutingPolicy) Voice Routing Policy. The Error Was $_" -ForegroundColor Red
                                Write-Log -Severity ERR -Message "FAILED to Assign $($User.UserPrincipalName) the $($User.OnlineVoiceRoutingPolicy) Voice Routing Policy. The Error Was $_"
                                $Script:ErrorCommands += "Grant-CsOnlineVoiceRoutingPolicy -Identity $($User.UserPrincipalName) -PolicyName $($User.OnlineVoiceRoutingPolicy) -ErrorAction Stop"
                                $UserOVRPSuccess = $False
                            }
                    }
                
                #Assign the OACRP to the User
                if (($User.OnlineAudioConferencingRoutingPolicy -eq "") -or ($User.OnlineAudioConferencingRoutingPolicy -eq "null") -or ($User.OnlineAudioConferencingRoutingPolicy -eq $null) -or ($User.OnlineAudioConferencingRoutingPolicy -eq "N/A"))
                    {
                        Write-Host "- Skipping the Assignment of a Online Audio Conferencing Routing Policy as the Value Provided is NULL" -ForegroundColor Yellow
                        Write-Log -Severity Info -Message "Skipping the Assignment of a Online Audio Conferencing Routing Policy to $($User.UserPrincipalName) as the Value Provided is NULL"  
                        $UserOACRPSuccess = $True
                    }
                else
                    {
                        try
                            {
                                Grant-CsOnlineAudioConferencingRoutingPolicy -Identity $User.UserPrincipalName -PolicyName $User.OnlineAudioConferencingRoutingPolicy -ErrorAction Stop
                                Write-Host "- Assigned the $($User.OnlineAudioConferencingRoutingPolicy) Online Audio Conferencing Routing Policy Successfully" -ForegroundColor Green
                                Write-Log -Severity Info -Message "Assigned $($User.UserPrincipalName) the $($User.OnlineAudioConferencingRoutingPolicy) Online Audio Conferencing Routing Policy Successfully"
                                $UserOACRPSuccess = $True
                            }
                        catch
                            {
                                Write-Host "- FAILED to Assign the $($User.OnlineAudioConferencingRoutingPolicy) Online Audio Conferencing Routing Policy. The Error Was: $_" -ForegroundColor Red
                                Write-Log -Severity ERR -Message "FAILED to Assign $($User.UserPrincipalName) the $($User.OnlineAudioConferencingRoutingPolicy) Online Audio Conferencing Routing Policy. The Error Was: $_"
                                $Script:ErrorCommands += "Grant-CsOnlineAudioConferencingRoutingPolicy -Identity $($User.UserPrincipalName) -PolicyName $($User.OnlineAudioConferencingRoutingPolicy) -ErrorAction Stop"
                                $UserOACRPSuccess = $False
                            }
                    }
                
                #Assign the Dial Plan to the User
                if (($User.TenantDialPlan -eq "") -or ($User.TenantDialPlan -eq "null") -or ($User.TenantDialPlan -eq $null) -or ($User.TenantDialPlan -eq "N/A"))
                    {
                        Write-Host "- Skipping the Assignment of a Tenant Dial Plan as the Value Provided is NULL" -ForegroundColor Yellow
                        Write-Log -Severity Info -Message "Skipping the Assignment of a Tenant Dial Plan to $($User.UserPrincipalName) as the Value Provided is NULL"  
                        $UserDPSuccess = $True
                    }
                else
                    {
                        try
                            {
                                Grant-CsTenantDialPlan -Identity $User.UserPrincipalName -PolicyName $User.TenantDialPlan -ErrorAction Stop
                                Write-Host "- Assigned the $($User.TenantDialPlan) Dial Plan Successfully" -ForegroundColor Green
                                Write-Log -Severity Info -Message "Assigned $($User.UserPrincipalName) the $($User.TenantDialPlan) Dial Plan Successfully"
                                $UserDPSuccess = $True
                            }
                        catch
                            {
                                Write-Host "- FAILED to Assign the $($User.TenantDialPlan) Dial Plan. The Error Was: $_" -ForegroundColor Red
                                Write-Log -Severity ERR -Message "FAILED to Assign $($User.UserPrincipalName) the $($User.TenantDialPlan) Dial Plan. The Error Was: $_"
                                $Script:ErrorCommands += "Grant-CsTenantDialPlan -Identity $($User.UserPrincipalName) -PolicyName $($User.TenantDialPlan) -ErrorAction Stop"
                                $UserDPSuccess = $False
                            }
                    }

                #Assign the Emergency Calling Policy to the User
                if (($User.TeamsEmergencyCallingPolicy -eq "") -or ($User.TeamsEmergencyCallingPolicy -eq "null") -or ($User.TeamsEmergencyCallingPolicy -eq $null) -or ($User.TeamsEmergencyCallingPolicy -eq "N/A"))
                    {
                        Write-Host "- Skipping the Assignment of a Emergency Calling Policy as the Value Provided is NULL" -ForegroundColor Yellow
                        Write-Log -Severity Info -Message "Skipping the Assignment of a Emergency Calling Policy to $($User.UserPrincipalName) as the Value Provided is NULL"  
                        $UserECPSuccess = $True
                    }
                else
                    {
                        try
                            {
                                Grant-CsTeamsEmergencyCallingPolicy -Identity $User.UserPrincipalName -PolicyName $User.TeamsEmergencyCallingPolicy -ErrorAction Stop
                                Write-Host "- Assigned the $($User.TeamsEmergencyCallingPolicy) Emergency Calling Policy Successfully" -ForegroundColor Green
                                Write-Log -Severity Info -Message "Assigned $($User.UserPrincipalName) the $($User.TeamsEmergencyCallingPolicy) Emergency Calling Policy Successfully"
                                $UserECPSuccess = $True
                            }
                        catch
                            {
                                Write-Host "- FAILED to Assign the $($User.TeamsEmergencyCallingPolicy) Emergency Calling Policy. The Error Was: $_" -ForegroundColor Red
                                Write-Log -Severity ERR -Message "FAILED to Assign $($User.UserPrincipalName) the $($User.TeamsEmergencyCallingPolicy) Emergency Calling Policy. The Error Was: $_"
                                $Script:ErrorCommands += "Grant-CsTeamsEmergencyCallingPolicy -Identity $($User.UserPrincipalName) -PolicyName $($User.TeamsEmergencyCallingPolicy) -ErrorAction Stop"
                                $UserECPSuccess = $False
                            }
                    }

                #Assign the Emergency Call Routing Policy to the User
                if (($User.TeamsEmergencyCallRoutingPolicy -eq "") -or ($User.TeamsEmergencyCallRoutingPolicy -eq "null") -or ($User.TeamsEmergencyCallRoutingPolicy -eq $null) -or ($User.TeamsEmergencyCallRoutingPolicy -eq "N/A"))
                    {
                        Write-Host "- Skipping the Assignment of a Emergency Call Routing Policy as the Value Provided is NULL" -ForegroundColor Yellow
                        Write-Log -Severity Info -Message "Skipping the Assignment of a Emergency Call Routing Policy to $($User.UserPrincipalName) as the Value Provided is NULL"  
                        $UserECRPSuccess = $True
                    }
                else
                    {
                        try
                            {
                                Grant-CsTeamsEmergencyCallRoutingPolicy -Identity $User.UserPrincipalName -PolicyName $User.TeamsEmergencyCallRoutingPolicy -ErrorAction Stop
                                Write-Host "- Assigned the $($User.TeamsEmergencyCallRoutingPolicy) Emergency Call Routing Policy Successfully" -ForegroundColor Green
                                Write-Log -Severity Info -Message "Assigned $($User.UserPrincipalName) the $($User.TeamsEmergencyCallRoutingPolicy) Emergency Call Routing Policy Successfully"
                                $UserECRPSuccess = $True
                            }
                        catch
                            {
                                Write-Host "- FAILED to Assign the $($User.TeamsEmergencyCallRoutingPolicy) Emergency Call Routing Policy. The Error Was: $_" -ForegroundColor Red
                                Write-Log -Severity ERR -Message "FAILED to Assign $($User.UserPrincipalName) the $($User.TeamsEmergencyCallRoutingPolicy) Emergency Call Routing Policy. The Error Was: $_"
                                $Script:ErrorCommands += "Grant-CsTeamsEmergencyCallRoutingPolicy -Identity $($User.UserPrincipalName) -PolicyName $($User.TeamsEmergencyCallRoutingPolicy) -ErrorAction Stop"
                                $UserECRPSuccess = $False
                            }
                    }

                $Script:Count = $Script:Count - 1 #Decrease remaining users count by 1

                if (($UserLineURISuccess -eq $True) -and ($UserOVRPSuccess -eq $True) -and ($UserOACRPSuccess -eq $True) -and ($UserDPSuccess -eq $True) -and ($UserECPSuccess -eq $True) -and ($UserECRPSuccess -eq $True))
                    {
                        Write-Host ""
                        Write-Host "Provisioned $($User.UserPrincipalName) Successfully! $($Script:Count) of $($Script:CountInitial) User(s) Remain...`n" -ForegroundColor Green
                        Write-Log -Severity Info -Message "Provisioned $($User.UserPrincipalName) Successfully! $($Script:Count) of $($Script:CountInitial) User(s) Remain..."
                    }
                else
                    {
                        Write-Host ""
                        Write-Host "One or More Errors Caused Provisioning to Fail for $($User.UserPrincipalName). $($Script:Count) of $($Script:CountInitial) User(s) Remain...`n" -ForegroundColor Red
                        Write-Log -Severity ERR -Message "One or More Errors Caused Provisioning to Fail for $($User.UserPrincipalName). $($Script:Count) of $($Script:CountInitial) User(s) Remain..."
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
                                        Write-Log -Severity Info -Message "The Command: $($CMD) | Failed Again. The Error Was: $_. $($Script:ErrorCount) of $($Script:ErrorCountInitial) Command(s) Remain..."
                                    }
                            }
                    }
            }
    }

#Main Menu--------------------------------------------------------------------------------------------------------------------------------------------
do{
    EM-MainMenu
    $Confirm1 = Read-Host "Of the above options, what mode would you like to run this script in? (Enter the Option Number)"
    Clear-Host

if ($Confirm1 -eq "1")
    {
        Write-Host "Option 1: Setup Admin Connections Selected. Setting Up Connections..."
        Write-Log -Severity Info -Message "Option 1: Setup Admin Connections Selected. Setting Up Connections..."
        Write-Host "Connecting to Microsoft Teams PowerShell"
        EM-ConnectTeamsPS
        Write-Log -Severity Info -Message "Option 1: Setup Admin Connections Complete, Returning to the Main Menu"
    }

elseif ($Confirm1 -eq "2")
    {
        Write-Host "Option 2: Close all Admin Connections Selected. Closing Connections..."
        Write-Log -Severity Info -Message "Option 2: Close all Admin Connections Selected. Closing Connections..."
        EM-DisconnectTeamsPS
        Write-Log -Severity Info -Message "Clearing all Admin Connection Variables"
        $Script:TeamsConnection = "<Not Set>"
        $Script:TenantDomain = "<Not Set>"
        $Script:TenantID = "<Not Set>"
        $Script:M365Admin = "<Not Set>"
        Write-Log -Severity Info -Message "All Admin Connection Variables Cleared"
        Write-Log -Severity Info -Message "Option 2: Close all Admin Connections Complete, Returning to the Main Menu"
    }

elseif ($Confirm1 -eq "10")
    {
        Write-Host "Option 10: Provision a Single User Account Selected"
        Write-Log -Severity Info -Message "Option 10: Provision a Single User Account Selected"

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
        $Users = $null
        $UserUPN = $null
        $UserPN = $null
        $UserOVRP = $null
        $UserOACRP = $null
        $UserDP = $null
        $UserECP = $null
        $UserECRP = $null
        $Script:Count = 1

        #Get User Information
        Write-Host "For any of the below values, if you would like no value set, please leave the field blank or enter null for none`n"-ForegroundColor Yellow
        $UserUPN = Read-Host "Please enter the UPN for the User you wish to provision (Ex:User@domain.com)"
        $UserPN = Read-Host "Please enter the phone number to assign (Ex: +13305550001)"
        $UserOVRP = Read-Host "Please enter the name of the Online Voice Routing Policy to assign"
        $UserOACRP = Read-Host "Please enter the name of the Online Audio Conferencing Routing Policy to assign"
        $UserDP = Read-Host "Please enter the name of the Dial Plan to assign"
        $UserECP = Read-Host "Please enter the name of the Emergency Calling Policy to assign"
        $UserECRP = Read-Host "Please enter the name of the Emergency Call Routing Policy to assign"

        [System.Collections.ArrayList]$Script:Users = @()
        $Users = New-Object PSCustomObject
        $Users | Add-Member -NotePropertyName UserPrincipalName -NotePropertyValue $UserUPN
        $Users | Add-Member -NotePropertyName PhoneNumber -NotePropertyValue $UserPN
        $Users | Add-Member -NotePropertyName OnlineVoiceRoutingPolicy -NotePropertyValue $UserOVRP
        $Users | Add-Member -NotePropertyName OnlineAudioConferencingRoutingPolicy -NotePropertyValue $UserOACRP
        $Users | Add-Member -NotePropertyName TenantDialPlan -NotePropertyValue $UserDP
        $Users | Add-Member -NotePropertyName TeamsEmergencyCallingPolicy -NotePropertyValue $UserECP
        $Users | Add-Member -NotePropertyName TeamsEmergencyCallRoutingPolicy -NotePropertyValue $UserECRP
        $Script:Users = $Users

        Write-Host ""
        $Confirmation = Read-Host "Are you sure that you want to provision 1 User for Microsoft Teams Voice? (Y/N)"

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
        Write-Log -Severity Info -Message "Option 10: Provision a Single User Account Complete, Returning to the Main Menu"  
    }

elseif ($Confirm1 -eq "11")
    {
        Write-Host "Option 11: Provision Multiple User Accounts (CSV Import) Selected"
        Write-Log -Severity Info -Message "Option 11: Provision Multiple User Accounts (CSV Import) Selected"

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
        EM-GetUsersCsv
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
        Write-Log -Severity Info -Message "Option 11: Provision Multiple User Accounts (CSV Import) Complete, Returning to the Main Menu"
    }

elseif ($Confirm1 -eq "12")
    {
        Write-Host "Option 12: Export User Calling Settings (CSV Import) Selected"
        Write-Log -Severity Info -Message "Option 12: Export User Calling Settings (CSV Import) Selected"
    
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
        EM-GetUsersCsv
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
                        Write-Host "FAILED to Gathered User Calling Settings for $($User.UserPrincipalName). The Error was $_. $($Script:Count) of $($Script:CountInitial) Users Remaining..." -ForegroundColor Red
                        Write-Log -Severity ERR -Message "FAILED to Gathered User Calling Settings for $($User.UserPrincipalName). The Error was $_. $($Script:Count) of $($Script:CountInitial) Users Remaining..."
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
                        Write-Host "FAILED to save Exported Data to: $($ExportPath). The Error Was $_" -ForegroundColor Green
                        Write-Log -Severity Info -Message "FAILED to save Exported Data to: $($ExportPath). The Error Was $_" 
                    }
            }
        else
            {
                Write-Host "No Valid Path Selected. Returning to the Main Menu" -ForegroundColor Yellow
                Write-Log -Severity WARN -Message "No Valid Path Selected. Returning to the Main Menu" -ForegroundColor Yellow
            }

        pause
        Write-Log -Severity Info -Message "Option 12: Export User Calling Settings (CSV Import) Complete, Returning to the Main Menu"
    }

elseif ($Confirm1 -eq "13")
    {
        if ($Script:BetaFlightsEnabled -eq $True)
            {
                Write-Host "Option 13: (Beta) Validate Teams Only Users for Readiness (CSV Import) Selected"
                Write-Log -Severity Info -Message "Option 13: (Beta) Validate Teams Only Users for Readiness (CSV Import) Selected"
            
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
                EM-GetUsersCsv
                $Script:CountInitial = $Script:Users.Count
                $Script:Count = $Script:Users.Count
                
                #Process Data
                Write-Host "Work In Progress for Future Release"

                pause
                Write-Log -Severity Info -Message "Option 13: (Beta) Validate Teams Only Users for Readiness (CSV Import) Complete, Returning to the Main Menu"
            }
        else
            {
                Write-Host "User not Authorized for this task!!!" -ForegroundColor Red
                Write-Log -Severity WARN -Message "User not Authorized for this task!!!"
                pause
                Write-Log -Severity Info -Message "Option 13: Returning to the Main Menu"
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
                $Confirm1 = "99"
            }
    }
}
while ($Confirm1 -ne "99") 

EM-DisconnectTeamsPS

$DT = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
Write-Host "$($Script:Name) $($Script:BuildVersion) Stopped at: $($DT)`n" -ForegroundColor Green
Write-Log -Severity Info -Message "$($Script:Name) $($Script:BuildVersion) Stopped at: $($DT)"
