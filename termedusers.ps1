#Requires -Version 3
#Requires -Modules AzureADPreview
#Requires -Modules Microsoft.Online.SharePoint.PowerShell

<#
.SYNOPSIS
    Used to term users from the company network and Office 365

.Description
    Takes employeeid value as parameter, finds account by employeeid, and does the following
   
    - hides from GAL
    - renames mailbox to indicate its termed and enddate for deletion
    - converts to shared mailbox
    - delegates mailbox access to manager
    - sets OoO 

#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory = $True, Position = 1)]
    [string]$employeeid,

    [Parameter(Mandatory = $false, Position = 2)]
    [string]$ManagerEmail,

    # Toggle OOF
    [Parameter(Mandatory = $false)]
    [bool]$EnableOOO = $true,

    # Set transcript path
    [Parameter(Mandatory = $false)]
    [string]$transcriptpath = "C:\windows\temp\transcript\transcript.txt"
)

$TermDateText = Get-Date -format "yyMMdd-HHmm"
$TermDate = Get-Date
[string]$transcriptpath = "C:\windows\temp\transcript\transcript-$TermDateText.txt"
Start-Transcript -Path $transcriptpath -Append

write-information -MessageData "Processing EmployeeID - $employeeid"
Set-Variable -name O365Credentials -Value (Get-StoredCredential -target O365) -Scope Global
Set-Variable -name OnPremCredentials -Value (Get-StoredCredential -target OnPrem) -Scope Global


write-information -MessageData "Validating User Enabled - Started"
$USR = $null
$USR = get-aduser -Credential $global:OnPremCredentials -Properties samaccountname,enabled,employeeid -Filter {(employeeid -eq $employeeid) -and (enabled -eq $false)}
If(!$USR){exit}
write-information -MessageData "Validating User Enabled - Completed"

write-information -MessageData "Validating mailbox - Started"
$mailbox = $null
$mailbox = get-aduser -Credential $global:OnPremCredentials -Properties samaccountname,employeeid -Filter {employeeid -eq $employeeid} | select -ExpandProperty samaccountname
If(!$mailbox){exit}
write-information -MessageData "Validating mailbox - Completed"

write-information -MessageData "Clearing PSSessions - Started"
Get-PSSession | Remove-PSSession
write-information -MessageData "Clearing PSSessions - Completed"

write-information -MessageData "Creating PSSessions - Started"
$PremSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchangeserver/PowerShell/ -Authentication Kerberos -Credential $global:OnPremCredentials
Import-PSSession $PremSession -Prefix 'Prem' -ErrorAction Stop -allowclobber
$0365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Authentication basic -Credential $global:O365Credentials -AllowRedirection
Import-PSSession $0365Session -DisableNameChecking -allowclobber
write-information -MessageData "Creating PSSessions - Completed"


write-information -MessageData "Instantiate user object - Started"
[psobject]$CurrentUser = get-aduser -Credential $global:OnPremCredentials -Properties samaccountname,employeeid -Filter {employeeid -eq $employeeid}
write-information -MessageData "Instantiate user object - Completed"

write-information -MessageData "Verifying user info - Started"
$username = $CurrentUser.givenname + "_" + $CurrentUser.surname
$name = $CurrentUser.givenname + " " + $CurrentUser.surname
write-information -MessageData "Verifying user info - Completed"

write-information -MessageData "Verifying manager email - Started"
$ManagerEmail = get-aduser $mailbox -Credential $global:OnPremCredentials -Properties Manager | select -expandproperty Manager | get-aduser -Properties emailaddress | select -expandproperty emailaddress
write-information -MessageData "Verifying manager email - Completed"

write-information -MessageData "Instantiate MGR mailbox object - Started"
[psobject]$Manager = Get-PremRemoteMailbox -Identity $ManagerEmail
write-information -MessageData "Instantiate MGR mailbox object - Completed"

[string]$ManagerName = $Manager.DisplayName
[string]$OoO = "$name is no longer with Company Name. Please contact $ManagerName by email at $ManagerEmail. Thank you!"
write-information -MessageData "Setting Exch attribute data - Started"
set-ADUser -Identity $CurrentUser.SamAccountName  -Replace @{msExchRemoteRecipientType = '100'; msExchRecipientTypeDetails = '34359738368'} -Credential $global:OnPremCredentials -Server ADServerName
write-information -MessageData "Setting Exch attribute data - Completed"

write-information -MessageData "Hiding from GAL - Started"
Set-PremRemoteMailbox -HiddenFromAddressListsEnabled $true -Identity $Mailbox
write-information -MessageData "Hiding from GAL - Completed"

write-information -MessageData "Setting account name with Enddate - Started"
[string]$EndDate = $termdate.adddays(30).ToShortDateString()
Set-PremRemoteMailbox -Identity $Mailbox -Name "Term - $name - $EndDate"
write-information -MessageData "Setting account name with Enddate - Completed"

write-information -MessageData "Converting mailbox to shared - Started"
Set-Mailbox -Identity $Mailbox -Type Shared
write-information -MessageData "Converting mailbox to shared - Completed"

write-information -MessageData "Adding MGR FullAccess to mailbox - Started"
Add-MailboxPermission -Identity $Mailbox -User $ManagerEmail -AccessRights FullAccess -InheritanceType All
write-information -MessageData "Adding MGR FullAccess to mailbox - Completed"

write-information -MessageData "Setting OOO - Started"
Set-MailboxAutoReplyConfiguration -Identity $Mailbox -AutoReplyState Enabled -InternalMessage $OoO -ExternalMessage $OoO
write-information -MessageData "Setting OOO - Completed"
    
write-information -MessageData "Delegating Onedrive - Started"
$TenantUrl = "https://tenant-admin.sharepoint.com/"
write-information -MessageData "Delegating Onedrive - URL Set"
write-information -MessageData "Delegating Onedrive - Connecting SPOService"
Connect-SPOService -Url $TenantUrl -Credential $global:O365Credentials
write-information -MessageData "Delegating Onedrive - Connected SPOService"
write-information -MessageData "Delegating Onedrive - Identifying Onedrive object for user"
$useronedrive = Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like '$username'" 
write-information -MessageData "Delegating Onedrive - Identified Onedrive object for user"
write-information -MessageData "Delegating Onedrive - Delegating to MGR"
Set-SPOUser -Site $useronedrive.url -LoginName $manageremail -IsSiteCollectionAdmin $true
write-information -MessageData "Delegating Onedrive - Delegated to MGR"
write-information -MessageData "Delegating Onedrive - Completed"

write-information -MessageData "**********************************"
write-information -MessageData "**********************************"
write-information -MessageData "USR Name - $name"
write-information -MessageData "USR employeeid - $employeeid"
write-information -MessageData "MGR Name - $ManagerName"
write-information -MessageData "OoO - $OoO"
write-information -MessageData 'Onedrive - $useronedrive.url'
write-information -MessageData $useronedrive.url
write-information -MessageData "**********************************"
write-information -MessageData "**********************************"
Stop-Transcript
