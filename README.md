# MicrosoftTeamsUserAccountProvisioningUtility
This PowerShell Script Provisions Users for Microsoft Teams as either a single user or bulk import. It also has some validation functions for users, more to come soon!

Features:
•	Option 10: Provision a Single Use
•	Option 11: Provision Multiple Users (CSV Import) - Template Provided
•	Option 12: Get User Calling Settings for Validation


Script Requirements:
•	Run as an Admin (For Logging)
•	Teams PS 4.9.1 Minimum - Script will auto update if this is not installed.


Script Operation:
•	From the Main Menu, you can either directly run a function or first setup admin connections. If you attempt to run a mode and the Teams PS Session is not active, it will have you connect.
•	All results are written to a log file located at C:\_Logs\


Option 10:
•	The script will ask for all the values to provision a single user. Includes: UserPrincipalName,PhoneNumber,OnlineVoiceRoutingPolicy,OnlineAudioConferencingRoutingPolicy,TenantDialPlan,TeamsEmergencyCallingPolicy,TeamsEmergencyCallRoutingPolicy
•	If you don’t want to set any of these values, leave the field completely blank.
•	The script will then provision the users, see the details of this in Option 11


Option 11:
•	The script will ask for a CSV containing the users that are to be provisioned. A template is provided in this repository. Any blank, Null or N/A values will be skipped.
•	A prompt asking to confirm the provisioning of X number of users will be presented.
•	Once accepted, the script will attempt the following logic:
o	If the value is Null, Blank, or N/A it will be skipped.
o	Else the cmd for the policy or value will be ran.
	If successful, it will be written to log file and the console.
	If failed, it will be written to log file and the console, but also will be stored in memory to be retried at the end of the run of users.
o	If a cmdlet failed, it will be retried at the end. This is useful for telephone numbers especially as they will occasionally fail due to the API.
 
Option 12:
•	Exports CsUserCallingSettings for all Users in the above CSV
