<#
.SYNOPSIS
Manages Azure AD user lifecycle operations including onboarding, updating, and offboarding,
for both single users and bulk operations via CSV import.

.DESCRIPTION
This script provides a centralized way to perform various Azure AD user lifecycle tasks.
It supports:
- Onboarding a single new user with specified attributes.
- Updating attributes for a single existing user.
- Offboarding (disabling or deleting) a single existing user.
- Bulk onboarding of new users from a CSV file.
- Bulk updating of existing users from a CSV file.
- Bulk offboarding of existing users from a CSV file.

All operations are logged to a specified log file and a PowerShell transcript is created.
The script relies on the Az.Accounts and Az.Resources (or Az.MicrosoftGraph) PowerShell modules.

.PARAMETER LifecycleAction
Specifies the lifecycle action to perform.
Valid values: SingleOnboard, SingleUpdate, SingleOffboard, BulkOnboard, BulkUpdate, BulkOffboard.

.PARAMETER UserPrincipalName
User Principal Name (UPN) for single-user operations.

.PARAMETER ObjectId
Object ID (GUID) for single-user operations (alternative to UserPrincipalName for Update/Offboard).

.PARAMETER CSVPath
Path to the CSV file for bulk operations.

.PARAMETER LogFilePath
Path to the main log file. Defaults to ".\ScriptLogs\AzureADLifecycle_yyyyMMdd_HHmmss.log".
Transcript will be saved in the same directory.

.PARAMETER DisplayName
DisplayName for the user (used in SingleOnboard, SingleUpdate).

.PARAMETER Department
Department for the user (used in SingleOnboard, SingleUpdate).

.PARAMETER JobTitle
JobTitle for the user (used in SingleOnboard, SingleUpdate).

.PARAMETER OfficeLocation
OfficeLocation for the user (used in SingleUpdate).

.PARAMETER StreetAddress
StreetAddress for the user (used in SingleUpdate).

.PARAMETER City
City for the user (used in SingleUpdate).

.PARAMETER State
State for the user (used in SingleUpdate).

.PARAMETER PostalCode
PostalCode for the user (used in SingleUpdate).

.PARAMETER Country
Country for the user (used in SingleUpdate).

.PARAMETER MobilePhone
MobilePhone for the user (used in SingleUpdate).

.PARAMETER OfficePhone
OfficePhone for the user (used in SingleUpdate).

.PARAMETER Password
Password for the user (SecureString) for SingleOnboard. If not provided, a random one is generated.

.PARAMETER ForceChangePasswordNextLogin
Boolean. $true to force password change on next login (used in SingleOnboard). Defaults to $true.

.PARAMETER InitialGroups
Array of Group ObjectIDs or DisplayNames to add the user to (used in SingleOnboard, SingleUpdate -GroupsToAdd).

.PARAMETER LicenseSKUs
Array of License SKU IDs (GUIDs) to assign (used in SingleOnboard, SingleUpdate -LicensesToAssign).

.PARAMETER ManagerUPN
UPN of the manager (used in SingleOnboard, SingleUpdate). For SingleUpdate, an empty string removes the manager.

.PARAMETER GroupsToAdd
Array of Group ObjectIDs or DisplayNames to add the user to (used in SingleUpdate).

.PARAMETER GroupsToRemove
Array of Group ObjectIDs or DisplayNames to remove the user from (used in SingleUpdate).

.PARAMETER LicensesToAssign
Array of License SKU IDs (GUIDs) to assign (used in SingleUpdate).

.PARAMETER LicensesToRemove
Array of License SKU IDs (GUIDs) to remove (used in SingleUpdate).

.PARAMETER OffboardAction
Action for SingleOffboard: 'Disable' or 'Delete'. Defaults to 'Disable'.

.PARAMETER RevokeSignInSessions
Boolean. For SingleOffboard. Defaults to $true.

.PARAMETER RemoveAllLicenses
Boolean. For SingleOffboard. Defaults to $true.

.PARAMETER RemoveFromAllGroups
Boolean. For SingleOffboard. Defaults to $false.

.EXAMPLE
# Single User Onboarding (minimal)
.\Manage-AzureADUserLifecycle.ps1 -LifecycleAction SingleOnboard -UserPrincipalName "test.user@contoso.com" -DisplayName "Test User"

.EXAMPLE
# Single User Onboarding (with more attributes)
$pwd = ConvertTo-SecureString "P@sswOrd123!" -AsPlainText -Force
.\Manage-AzureADUserLifecycle.ps1 -LifecycleAction SingleOnboard -UserPrincipalName "new.hire@contoso.com" -DisplayName "New Hire" -Password $pwd -Department "IT" -JobTitle "Support Tech" -InitialGroups "IT Support Users" -LicenseSKUs "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

.EXAMPLE
# Single User Update
.\Manage-AzureADUserLifecycle.ps1 -LifecycleAction SingleUpdate -UserPrincipalName "test.user@contoso.com" -Department "Marketing" -JobTitle "Marketing Specialist"

.EXAMPLE
# Single User Offboarding (Disable)
.\Manage-AzureADUserLifecycle.ps1 -LifecycleAction SingleOffboard -UserPrincipalName "leaving.user@contoso.com"

.EXAMPLE
# Single User Offboarding (Delete with specific options)
.\Manage-AzureADUserLifecycle.ps1 -LifecycleAction SingleOffboard -UserPrincipalName "former.employee@contoso.com" -OffboardAction Delete -RemoveAllLicenses $false

.EXAMPLE
# Bulk User Onboarding
.\Manage-AzureADUserLifecycle.ps1 -LifecycleAction BulkOnboard -CSVPath .\onboard_users.csv -LogFilePath .\Logs\BulkOnboard.log

.EXAMPLE
# Bulk User Update
.\Manage-AzureADUserLifecycle.ps1 -LifecycleAction BulkUpdate -CSVPath .\update_users.csv

.EXAMPLE
# Bulk User Offboarding
.\Manage-AzureADUserLifecycle.ps1 -LifecycleAction BulkOffboard -CSVPath .\offboard_users.csv -Verbose

.NOTES
Author: AI Agent (Software Engineering Specialization)
Version: 1.1
Date: $(Get-Date -Format yyyy-MM-dd)

Requires Az.Accounts and Az.Resources (or Az.MicrosoftGraph) modules.
Ensure the script file is unblocked if downloaded from the internet (Unblock-File -Path .\Manage-AzureADUserLifecycle.ps1).
The script defines and uses several helper functions for single and bulk operations.
CSV formats for bulk operations are critical. Refer to README.md for details.
The default LogFilePath includes a timestamp to prevent overwriting logs from previous runs if not specified.
The script will attempt to create the directory for LogFilePath and TranscriptPath if it doesn't exist.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $true, HelpMessage = "Specifies the lifecycle action to perform.")]
    [ValidateSet("SingleOnboard", "SingleUpdate", "SingleOffboard", "BulkOnboard", "BulkUpdate", "BulkOffboard")]
    [string]$LifecycleAction,

    [Parameter(ParameterSetName = "SingleUser", Mandatory = $false, HelpMessage = "User Principal Name for single-user operations.")]
    [Parameter(ParameterSetName = "SingleOnboard", Mandatory = $true)]
    [Parameter(ParameterSetName = "SingleUpdateUPN", Mandatory = $true)]
    [Parameter(ParameterSetName = "SingleOffboardUPN", Mandatory = $true)]
    [string]$UserPrincipalName,

    [Parameter(ParameterSetName = "SingleUpdateObjectId", Mandatory = $true, HelpMessage = "Object ID for single-user update operations.")]
    [Parameter(ParameterSetName = "SingleOffboardObjectId", Mandatory = $true, HelpMessage = "Object ID for single-user offboard operations.")]
    [string]$ObjectId,

    [Parameter(ParameterSetName = "BulkOperation", Mandatory = $true, HelpMessage = "Path to the CSV file for bulk operations.")]
    [string]$CSVPath,

    [Parameter(Mandatory = $false, HelpMessage = "Path to the main log file. Transcript will be in the same directory.")]
    [string]$LogFilePath, # Default value handled in Begin block

    # Parameters for SingleOnboard / SingleUpdate
    [Parameter(ParameterSetName = "SingleOnboard", Mandatory = $true)]
    [Parameter(ParameterSetName = "SingleUpdateUPN")]
    [Parameter(ParameterSetName = "SingleUpdateObjectId")]
    [string]$DisplayName,

    [Parameter(ParameterSetName = "SingleOnboard")]
    [Parameter(ParameterSetName = "SingleUpdateUPN")]
    [Parameter(ParameterSetName = "SingleUpdateObjectId")]
    [string]$Department,

    [Parameter(ParameterSetName = "SingleOnboard")]
    [Parameter(ParameterSetName = "SingleUpdateUPN")]
    [Parameter(ParameterSetName = "SingleUpdateObjectId")]
    [string]$JobTitle,
    
    [Parameter(ParameterSetName = "SingleUpdateUPN")]
    [Parameter(ParameterSetName = "SingleUpdateObjectId")]
    [string]$OfficeLocation,

    [Parameter(ParameterSetName = "SingleUpdateUPN")]
    [Parameter(ParameterSetName = "SingleUpdateObjectId")]
    [string]$StreetAddress,

    [Parameter(ParameterSetName = "SingleUpdateUPN")]
    [Parameter(ParameterSetName = "SingleUpdateObjectId")]
    [string]$City,

    [Parameter(ParameterSetName = "SingleUpdateUPN")]
    [Parameter(ParameterSetName = "SingleUpdateObjectId")]
    [string]$State,

    [Parameter(ParameterSetName = "SingleUpdateUPN")]
    [Parameter(ParameterSetName = "SingleUpdateObjectId")]
    [string]$PostalCode,

    [Parameter(ParameterSetName = "SingleUpdateUPN")]
    [Parameter(ParameterSetName = "SingleUpdateObjectId")]
    [string]$Country,
    
    [Parameter(ParameterSetName = "SingleUpdateUPN")]
    [Parameter(ParameterSetName = "SingleUpdateObjectId")]
    [string]$MobilePhone,

    [Parameter(ParameterSetName = "SingleUpdateUPN")]
    [Parameter(ParameterSetName = "SingleUpdateObjectId")]
    [string]$OfficePhone,

    # SingleOnboard specific
    [Parameter(ParameterSetName = "SingleOnboard")]
    [System.Security.SecureString]$Password,

    [Parameter(ParameterSetName = "SingleOnboard")]
    [bool]$ForceChangePasswordNextLogin = $true,

    [Parameter(ParameterSetName = "SingleOnboard")]
    [string[]]$InitialGroups,

    [Parameter(ParameterSetName = "SingleOnboard")]
    [string[]]$LicenseSKUs,
    
    # SingleOnboard & SingleUpdate common (Manager, also Groups/Licenses for Update)
    [Parameter(ParameterSetName = "SingleOnboard")]
    [Parameter(ParameterSetName = "SingleUpdateUPN")]
    [Parameter(ParameterSetName = "SingleUpdateObjectId")]
    [string]$ManagerUPN,

    [Parameter(ParameterSetName = "SingleUpdateUPN")]
    [Parameter(ParameterSetName = "SingleUpdateObjectId")]
    [string[]]$GroupsToAdd,

    [Parameter(ParameterSetName = "SingleUpdateUPN")]
    [Parameter(ParameterSetName = "SingleUpdateObjectId")]
    [string[]]$GroupsToRemove,

    [Parameter(ParameterSetName = "SingleUpdateUPN")]
    [Parameter(ParameterSetName = "SingleUpdateObjectId")]
    [string[]]$LicensesToAssign,

    [Parameter(ParameterSetName = "SingleUpdateUPN")]
    [Parameter(ParameterSetName = "SingleUpdateObjectId")]
    [string[]]$LicensesToRemove,

    # SingleOffboard specific
    [Parameter(ParameterSetName = "SingleOffboardUPN")]
    [Parameter(ParameterSetName = "SingleOffboardObjectId")]
    [ValidateSet("Disable", "Delete")]
    [string]$OffboardAction = "Disable",

    [Parameter(ParameterSetName = "SingleOffboardUPN")]
    [Parameter(ParameterSetName = "SingleOffboardObjectId")]
    [bool]$RevokeSignInSessions = $true,

    [Parameter(ParameterSetName = "SingleOffboardUPN")]
    [Parameter(ParameterSetName = "SingleOffboardObjectId")]
    [bool]$RemoveAllLicenses = $true,

    [Parameter(ParameterSetName = "SingleOffboardUPN")]
    [Parameter(ParameterSetName = "SingleOffboardObjectId")]
    [bool]$RemoveFromAllGroups = $false
)

#region Function Definitions

#--------------------------------------------------------------------------------
# Logging Function (Write-ScriptLog)
#--------------------------------------------------------------------------------
function Write-ScriptLog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [ValidateSet('INFO', 'WARN', 'ERROR', 'VERBOSE', 'DEBUG')]
        [string]$Level = 'INFO',
        [Parameter(Mandatory = $true)]
        [string]$LogFilePathIn
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp [$Level] - $Message"
    try {
        if (-not (Test-Path -Path (Split-Path -Path $LogFilePathIn -Parent) -PathType Container)) {
            New-Item -ItemType Directory -Path (Split-Path -Path $LogFilePathIn -Parent) -Force -ErrorAction Stop | Out-Null
        }
        Add-Content -Path $LogFilePathIn -Value $logEntry -ErrorAction Stop
    }
    catch {
        Write-Error "FATAL: Failed to write to log file '$LogFilePathIn'. Error: $($_.Exception.Message)"
    }
    switch ($Level) {
        'INFO'    { Write-Host $logEntry }
        'WARN'    { Write-Warning $logEntry }
        'ERROR'   { Write-Error $logEntry }
        'VERBOSE' { Write-Verbose $logEntry }
        'DEBUG'   { Write-Debug $logEntry }
        default   { Write-Host $logEntry }
    }
}

#--------------------------------------------------------------------------------
# Single-User Onboarding Function (Invoke-AADUserOnboarding)
#--------------------------------------------------------------------------------
function Invoke-AADUserOnboarding {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true)] [string]$UserPrincipalName,
        [Parameter(Mandatory = $true)] [string]$DisplayName,
        [Parameter(Mandatory = $false)] [System.Security.SecureString]$Password,
        [Parameter(Mandatory = $false)] [bool]$ForceChangePasswordNextLogin = $true,
        [Parameter(Mandatory = $false)] [string[]]$InitialGroups,
        [Parameter(Mandatory = $false)] [string[]]$LicenseSKUs,
        [Parameter(Mandatory = $false)] [string]$Department,
        [Parameter(Mandatory = $false)] [string]$JobTitle,
        [Parameter(Mandatory = $false)] [string]$ManagerUPN,
        [Parameter(Mandatory = $true)] [string]$LogFilePath # Renamed from LogFilePathIn for clarity
    )
    function _EnsureWriteScriptLog { if (-not (Get-Command Write-ScriptLog -ErrorAction SilentlyContinue)) { New-Alias -Name Write-ScriptLog -Value Write-Host -Scope Script -Force } }
    _EnsureWriteScriptLog
    begin {
        Write-ScriptLog -Message "Invoke-AADUserOnboarding: Starting for UPN '$UserPrincipalName'." -Level VERBOSE -LogFilePath $LogFilePath
        try {
            $azContext = Get-AzContext -ErrorAction Stop
            Write-ScriptLog -Message "Invoke-AADUserOnboarding: Azure Context Tenant: $($azContext.Tenant.Id) for UPN '$UserPrincipalName'." -Level DEBUG -LogFilePath $LogFilePath
        } catch {
            Write-ScriptLog -Message "Invoke-AADUserOnboarding: Not connected to Azure. Aborting for UPN '$UserPrincipalName'." -Level ERROR -LogFilePath $LogFilePath
            return $null 
        }
        if (-not $Password) {
            Write-ScriptLog -Message "Invoke-AADUserOnboarding: Password not provided for UPN '$UserPrincipalName', generating random." -Level INFO -LogFilePath $LogFilePath
            $Password = ConvertTo-SecureString (New-Guid).ToString() -AsPlainText -Force
        }
    }
    process {
        if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Create Azure AD User '$DisplayName'")) {
            try {
                $userParams = @{
                    UserPrincipalName = $UserPrincipalName; DisplayName = $DisplayName
                    MailNickname = $UserPrincipalName.Split('@')[0]; AccountEnabled = $true
                }
                if ($PSBoundParameters.ContainsKey('Department')) { $userParams.Department = $Department }
                if ($PSBoundParameters.ContainsKey('JobTitle')) { $userParams.JobTitle = $JobTitle }
                $userParams.PasswordProfile = @{ Password = $Password; ForceChangePasswordNextLogin = $ForceChangePasswordNextLogin }
                
                $manager = $null
                if ($PSBoundParameters.ContainsKey('ManagerUPN') -and -not [string]::IsNullOrEmpty($ManagerUPN)) {
                    try {
                        Write-ScriptLog -Message "Invoke-AADUserOnboarding: Finding manager '$ManagerUPN' for '$UserPrincipalName'." -Level DEBUG -LogFilePath $LogFilePath
                        $manager = Get-AzADUser -UserPrincipalName $ManagerUPN -ErrorAction Stop
                        if ($manager) { Write-ScriptLog -Message "Invoke-AADUserOnboarding: Manager '$($manager.DisplayName)' found for '$UserPrincipalName'." -Level INFO -LogFilePath $LogFilePath }
                    } catch { Write-ScriptLog -Message "Invoke-AADUserOnboarding: Manager UPN '$ManagerUPN' not found for '$UserPrincipalName'. Skipping. Error: $($_.Exception.Message)" -Level WARN -LogFilePath $LogFilePath; $manager = $null }
                }
                Write-ScriptLog -Message "Invoke-AADUserOnboarding: Creating user '$UserPrincipalName'." -Level INFO -LogFilePath $LogFilePath
                $newUser = New-AzADUser @userParams -ErrorAction Stop
                Write-ScriptLog -Message "Invoke-AADUserOnboarding: User '$($newUser.UserPrincipalName)' (ID: $($newUser.Id)) created." -Level INFO -LogFilePath $LogFilePath
                if ($manager) {
                    if ($PSCmdlet.ShouldProcess($newUser.UserPrincipalName, "Set Manager to $($manager.DisplayName)")) {
                        try {
                            Write-ScriptLog -Message "Invoke-AADUserOnboarding: Setting manager for '$($newUser.UserPrincipalName)' to '$($manager.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath
                            Set-AzADUserManager -ObjectId $newUser.Id -ManagerId $manager.Id -ErrorAction Stop
                            Write-ScriptLog -Message "Invoke-AADUserOnboarding: Set manager for '$($newUser.UserPrincipalName)' to '$($manager.DisplayName)'." -Level INFO -LogFilePath $LogFilePath
                        } catch { Write-ScriptLog -Message "Invoke-AADUserOnboarding: Failed to set manager for '$($newUser.UserPrincipalName)'. Error: $($_.Exception.Message)" -Level WARN -LogFilePath $LogFilePath }
                    } else { Write-ScriptLog -Message "Invoke-AADUserOnboarding: Manager assignment for '$($newUser.UserPrincipalName)' skipped (-WhatIf)." -Level WARN -LogFilePath $LogFilePath}
                }
                if ($InitialGroups) {
                    Write-ScriptLog -Message "Invoke-AADUserOnboarding: Assigning groups to '$($newUser.UserPrincipalName)': $($InitialGroups -join ', ')." -Level INFO -LogFilePath $LogFilePath
                    foreach ($groupNameOrId in $InitialGroups) {
                        try {
                            $group = Get-AzADGroup -Filter "Id eq '$groupNameOrId' or DisplayName eq '$groupNameOrId'" -ErrorAction Stop | Select-Object -First 1
                            if ($group) {
                                if ($PSCmdlet.ShouldProcess($group.DisplayName, "Add user '$($newUser.UserPrincipalName)' to group")) {
                                    Add-AzADGroupMember -TargetGroupObjectId $group.Id -MemberObjectId $newUser.Id -ErrorAction Stop
                                    Write-ScriptLog -Message "Invoke-AADUserOnboarding: Added '$($newUser.UserPrincipalName)' to group '$($group.DisplayName)'." -Level INFO -LogFilePath $LogFilePath
                                } else { Write-ScriptLog -Message "Invoke-AADUserOnboarding: Adding '$($newUser.UserPrincipalName)' to group '$($group.DisplayName)' skipped (-WhatIf)." -Level WARN -LogFilePath $LogFilePath }
                            } else { Write-ScriptLog -Message "Invoke-AADUserOnboarding: Group '$groupNameOrId' not found for '$($newUser.UserPrincipalName)'. Skipping." -Level WARN -LogFilePath $LogFilePath }
                        } catch { Write-ScriptLog -Message "Invoke-AADUserOnboarding: Failed processing group '$groupNameOrId' for '$($newUser.UserPrincipalName)'. Error: $($_.Exception.Message)" -Level WARN -LogFilePath $LogFilePath }
                    }
                }
                if ($LicenseSKUs) {
                    Write-ScriptLog -Message "Invoke-AADUserOnboarding: Assigning licenses to '$($newUser.UserPrincipalName)': $($LicenseSKUs -join ', ')." -Level INFO -LogFilePath $LogFilePath
                    $licensesToAssign = $LicenseSKUs | ForEach-Object { @{SkuId = $_; DisabledPlans = @()} }
                    if ($licensesToAssign.Count -gt 0) {
                         if ($PSCmdlet.ShouldProcess($newUser.UserPrincipalName, "Assign Licenses: $($LicenseSKUs -join ',')")) {
                            try {
                                Set-AzADUser -ObjectId $newUser.Id -AssignedLicenses $licensesToAssign -ErrorAction Stop
                                Write-ScriptLog -Message "Invoke-AADUserOnboarding: Assigned licenses to '$($newUser.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath
                            } catch { Write-ScriptLog -Message "Invoke-AADUserOnboarding: Failed to assign licenses to '$($newUser.UserPrincipalName)'. Error: $($_.Exception.Message)" -Level WARN -LogFilePath $LogFilePath }
                        } else { Write-ScriptLog -Message "Invoke-AADUserOnboarding: License assignment for '$($newUser.UserPrincipalName)' skipped (-WhatIf)." -Level WARN -LogFilePath $LogFilePath }
                    }
                }
                Write-ScriptLog -Message "Invoke-AADUserOnboarding: Onboarding complete for '$($newUser.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath
                return $newUser
            } catch {
                Write-ScriptLog -Message "Invoke-AADUserOnboarding: Failed to onboard '$UserPrincipalName'. Error: $($_.Exception.Message)" -Level ERROR -LogFilePath $LogFilePath
                if ($_.Exception.StackTrace) { Write-ScriptLog -Message "Invoke-AADUserOnboarding: StackTrace for '$UserPrincipalName': $($_.Exception.StackTrace)" -Level DEBUG -LogFilePath $LogFilePath}
                return $null
            }
        } else {
            Write-ScriptLog -Message "Invoke-AADUserOnboarding: Creation of '$UserPrincipalName' skipped (-WhatIf)." -Level WARN -LogFilePath $LogFilePath
            return $null
        }
    }
    end { Write-ScriptLog -Message "Invoke-AADUserOnboarding: Finished for UPN '$UserPrincipalName'." -Level VERBOSE -LogFilePath $LogFilePath }
}

#--------------------------------------------------------------------------------
# Single-User Update Function (Update-AADUser)
#--------------------------------------------------------------------------------
function Update-AADUser {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = "UPN")] [string]$UserPrincipalName,
        [Parameter(Mandatory = $true, ParameterSetName = "ObjectId")] [string]$ObjectId,
        [Parameter(Mandatory = $false)] [string]$DisplayName,
        [Parameter(Mandatory = $false)] [string]$Department,
        [Parameter(Mandatory = $false)] [string]$JobTitle,
        [Parameter(Mandatory = $false)] [string]$OfficeLocation,
        [Parameter(Mandatory = $false)] [string]$StreetAddress,
        [Parameter(Mandatory = $false)] [string]$City,
        [Parameter(Mandatory = $false)] [string]$State,
        [Parameter(Mandatory = $false)] [string]$PostalCode,
        [Parameter(Mandatory = $false)] [string]$Country,
        [Parameter(Mandatory = $false)] [string]$MobilePhone,
        [Parameter(Mandatory = $false)] [string]$OfficePhone,
        [Parameter(Mandatory = $false)] [string]$ManagerUPN,
        [Parameter(Mandatory = $false)] [string[]]$GroupsToAdd,
        [Parameter(Mandatory = $false)] [string[]]$GroupsToRemove,
        [Parameter(Mandatory = $false)] [string[]]$LicensesToAssign,
        [Parameter(Mandatory = $false)] [string[]]$LicensesToRemove,
        [Parameter(Mandatory = $true)] [string]$LogFilePath
    )
    function _EnsureWriteScriptLogUpdateUser { if (-not (Get-Command Write-ScriptLog -ErrorAction SilentlyContinue)) { New-Alias -Name Write-ScriptLog -Value Write-Host -Scope Script -Force } }
    _EnsureWriteScriptLogUpdateUser
    begin {
        Write-ScriptLog -Message "Update-AADUser: Starting." -Level VERBOSE -LogFilePath $LogFilePath
        try { Get-AzContext -ErrorAction Stop | Out-Null } catch { Write-ScriptLog -Message "Update-AADUser: Not connected to Azure." -Level ERROR -LogFilePath $LogFilePath; return $null }
        $script:updateUserIdentifier = if ($PSCmdlet.ParameterSetName -eq "ObjectId") { $ObjectId } else { $UserPrincipalName }
        Write-ScriptLog -Message "Update-AADUser: Target user '$($script:updateUserIdentifier)'." -Level DEBUG -LogFilePath $LogFilePath
    }
    process {
        $processMsg = "Update Azure AD User '$($script:updateUserIdentifier)'"
        if ($PSBoundParameters.Count -le 3) { # Identifier, LogFilePath, and one action param
            Write-ScriptLog -Message "Update-AADUser: No attributes specified for update for '$($script:updateUserIdentifier)'. Nothing to do." -Level WARN -LogFilePath $LogFilePath
            return $null 
        }
        if (-not ($PSCmdlet.ShouldProcess($script:updateUserIdentifier, $processMsg))) {
            Write-ScriptLog -Message "Update-AADUser: Update for '$($script:updateUserIdentifier)' skipped (-WhatIf)." -Level WARN -LogFilePath $LogFilePath
            return $null 
        }
        try {
            $user = if ($PSCmdlet.ParameterSetName -eq "ObjectId") { Get-AzADUser -ObjectId $script:updateUserIdentifier -ErrorAction Stop } else { Get-AzADUser -UserPrincipalName $script:updateUserIdentifier -ErrorAction Stop }
            if (-not $user) { Write-ScriptLog -Message "Update-AADUser: User '$($script:updateUserIdentifier)' not found." -Level ERROR -LogFilePath $LogFilePath; return $null }
            Write-ScriptLog -Message "Update-AADUser: User '$($user.UserPrincipalName)' (ID: $($user.Id)) found." -Level INFO -LogFilePath $LogFilePath
            
            $setParams = @{}
            $paramExclusions = @('UserPrincipalName', 'ObjectId', 'LogFilePath', 'WhatIf', 'Confirm', 'Verbose', 'Debug', 'ErrorAction', 'ErrorVariable', 'OutVariable', 'OutBuffer', 'PipelineVariable', 'WarningAction', 'WarningVariable', 'ManagerUPN', 'GroupsToAdd', 'GroupsToRemove', 'LicensesToAssign', 'LicensesToRemove')
            $PSBoundParameters.Keys | Where-Object { $paramExclusions -notcontains $_ -and $PSBoundParameters[$_] -ne $null } | ForEach-Object {
                if ($_ -eq 'OfficePhone') { $setParams.BusinessPhones = @($PSBoundParameters[$_]) } else { $setParams[$_] = $PSBoundParameters[$_] }
            }
            if ($setParams.Count -gt 0) {
                Write-ScriptLog -Message "Update-AADUser: Updating attributes for '$($user.UserPrincipalName)': $($setParams.Keys -join ', ')." -Level INFO -LogFilePath $LogFilePath
                Set-AzADUser -ObjectId $user.Id @setParams -ErrorAction Stop
                Write-ScriptLog -Message "Update-AADUser: Attributes updated for '$($user.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath
            }
            if ($PSBoundParameters.ContainsKey('ManagerUPN')) {
                if (-not [string]::IsNullOrEmpty($ManagerUPN)) {
                    Write-ScriptLog -Message "Update-AADUser: Setting manager for '$($user.UserPrincipalName)' to '$ManagerUPN'." -Level INFO -LogFilePath $LogFilePath
                    try {
                        $newMgr = Get-AzADUser -UserPrincipalName $ManagerUPN -ErrorAction Stop
                        if ($newMgr) { Set-AzADUserManager -ObjectId $user.Id -ManagerId $newMgr.Id -ErrorAction Stop; Write-ScriptLog -Message "Update-AADUser: Set manager for '$($user.UserPrincipalName)' to '$($newMgr.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath }
                    } catch { Write-ScriptLog -Message "Update-AADUser: Failed to set manager '$ManagerUPN' for '$($user.UserPrincipalName)'. Error: $($_.Exception.Message)" -Level WARN -LogFilePath $LogFilePath }
                } else { 
                    Write-ScriptLog -Message "Update-AADUser: Removing manager for '$($user.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath
                    try { Set-AzADUserManager -ObjectId $user.Id -ManagerId $null -ErrorAction Stop; Write-ScriptLog -Message "Update-AADUser: Removed manager for '$($user.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath }
                    catch { Write-ScriptLog -Message "Update-AADUser: Failed to remove manager for '$($user.UserPrincipalName)'. Error: $($_.Exception.Message)" -Level WARN -LogFilePath $LogFilePath }
                }
            }
            if ($PSBoundParameters.ContainsKey('GroupsToAdd')) {
                Write-ScriptLog -Message "Update-AADUser: Adding '$($user.UserPrincipalName)' to groups: $($GroupsToAdd -join ', ')." -Level INFO -LogFilePath $LogFilePath
                $GroupsToAdd | ForEach-Object { try { $grp = Get-AzADGroup -Filter "DisplayName eq '$_' or Id eq '$_'" -ErrorAction Stop | Select -First 1; if ($grp) { Add-AzADGroupMember -TargetGroupObjectId $grp.Id -MemberObjectId $user.Id -ErrorAction Stop; Write-ScriptLog -Message "Update-AADUser: Added '$($user.UserPrincipalName)' to group '$($grp.DisplayName)'." -Level INFO -LogFilePath $LogFilePath } else { Write-ScriptLog -Message "Update-AADUser: Group '$_' not found for adding. Skipping." -Level WARN -LogFilePath $LogFilePath } } catch { Write-ScriptLog -Message "Update-AADUser: Failed to add '$($user.UserPrincipalName)' to group '$_'. Error: $($_.Exception.Message)" -Level WARN -LogFilePath $LogFilePath } }
            }
            if ($PSBoundParameters.ContainsKey('GroupsToRemove')) {
                Write-ScriptLog -Message "Update-AADUser: Removing '$($user.UserPrincipalName)' from groups: $($GroupsToRemove -join ', ')." -Level INFO -LogFilePath $LogFilePath
                $GroupsToRemove | ForEach-Object { try { $grp = Get-AzADGroup -Filter "DisplayName eq '$_' or Id eq '$_'" -ErrorAction Stop | Select -First 1; if ($grp) { Remove-AzADGroupMember -TargetGroupObjectId $grp.Id -MemberObjectId $user.Id -ErrorAction Stop; Write-ScriptLog -Message "Update-AADUser: Removed '$($user.UserPrincipalName)' from group '$($grp.DisplayName)'." -Level INFO -LogFilePath $LogFilePath } else { Write-ScriptLog -Message "Update-AADUser: Group '$_' not found for removal. Skipping." -Level WARN -LogFilePath $LogFilePath } } catch { Write-ScriptLog -Message "Update-AADUser: Failed to remove '$($user.UserPrincipalName)' from group '$_'. Error: $($_.Exception.Message)" -Level WARN -LogFilePath $LogFilePath } }
            }
            if ($PSBoundParameters.ContainsKey('LicensesToAssign') -or $PSBoundParameters.ContainsKey('LicensesToRemove')) {
                Write-ScriptLog -Message "Update-AADUser: Updating licenses for '$($user.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath
                try {
                    $currentLicenses = Get-AzADUserAssignedLicense -ObjectId $user.Id -ErrorAction Stop
                    $finalSkuIds = [System.Collections.Generic.List[string]]::new($currentLicenses | ForEach-Object { $_.SkuId.ToString() })
                    if ($PSBoundParameters.ContainsKey('LicensesToAssign')) { $LicensesToAssign | ForEach-Object { if (-not ($finalSkuIds.Contains($_))) { $finalSkuIds.Add($_) } } }
                    if ($PSBoundParameters.ContainsKey('LicensesToRemove')) { $LicensesToRemove | ForEach-Object { if ($finalSkuIds.Contains($_)) { $finalSkuIds.Remove($_) } } }
                    $licensesForUpdate = $finalSkuIds | ForEach-Object { @{SkuId = $_; DisabledPlans = @()} }
                    Set-AzADUser -ObjectId $user.Id -AssignedLicenses $licensesForUpdate -ErrorAction Stop
                    Write-ScriptLog -Message "Update-AADUser: Licenses updated for '$($user.UserPrincipalName)'. Final SKUs: $($finalSkuIds -join ', ')." -Level INFO -LogFilePath $LogFilePath
                } catch { Write-ScriptLog -Message "Update-AADUser: Failed to update licenses for '$($user.UserPrincipalName)'. Error: $($_.Exception.Message)" -Level WARN -LogFilePath $LogFilePath }
            }
            Write-ScriptLog -Message "Update-AADUser: Update process complete for '$($user.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath
            return Get-AzADUser -ObjectId $user.Id
        } catch [Microsoft.Azure.Commands.MicrosoftGraph.Cmdlets.Users.Models.MicrosoftGraphServiceException] {
             if ($_.Exception.Response.StatusCode -eq [System.Net.HttpStatusCode]::NotFound) { Write-ScriptLog -Message "Update-AADUser: User '$($script:updateUserIdentifier)' not found. $($_.Exception.Message)" -Level ERROR -LogFilePath $LogFilePath }
             else { Write-ScriptLog -Message "Update-AADUser: Graph service error for '$($script:updateUserIdentifier)': $($_.Exception.Message)" -Level ERROR -LogFilePath $LogFilePath }
             return $null
        } catch {
            Write-ScriptLog -Message "Update-AADUser: Failed to update '$($script:updateUserIdentifier)'. Error: $($_.Exception.Message)" -Level ERROR -LogFilePath $LogFilePath
            if ($_.Exception.StackTrace) { Write-ScriptLog -Message "Update-AADUser: StackTrace for '$($script:updateUserIdentifier)': $($_.Exception.StackTrace)" -Level DEBUG -LogFilePath $LogFilePath}
            return $null
        }
    }
    end { Write-ScriptLog -Message "Update-AADUser: Finished for user '$($script:updateUserIdentifier)'." -Level VERBOSE -LogFilePath $LogFilePath }
}

#--------------------------------------------------------------------------------
# Single-User Offboarding Function (Invoke-AADUserOffboarding)
#--------------------------------------------------------------------------------
function Invoke-AADUserOffboarding {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = "UPN")] [string]$UserPrincipalName,
        [Parameter(Mandatory = $true, ParameterSetName = "ObjectId")] [string]$ObjectId,
        [Parameter(Mandatory = $false)] [ValidateSet("Disable", "Delete")] [string]$Action = "Disable",
        [Parameter(Mandatory = $false)] [bool]$RevokeSignInSessions = $true,
        [Parameter(Mandatory = $false)] [bool]$RemoveAllLicenses = $true,
        [Parameter(Mandatory = $false)] [bool]$RemoveFromAllGroups = $false,
        [Parameter(Mandatory = $true)] [string]$LogFilePath
    )
    function _EnsureWriteScriptLogOffboardUser { if (-not (Get-Command Write-ScriptLog -ErrorAction SilentlyContinue)) { New-Alias -Name Write-ScriptLog -Value Write-Host -Scope Script -Force } }
    _EnsureWriteScriptLogOffboardUser
    begin {
        Write-ScriptLog -Message "Invoke-AADUserOffboarding: Starting." -Level VERBOSE -LogFilePath $LogFilePath
        try { Get-AzContext -ErrorAction Stop | Out-Null } catch { Write-ScriptLog -Message "Invoke-AADUserOffboarding: Not connected to Azure." -Level ERROR -LogFilePath $LogFilePath; return "Error: Not connected." }
        $script:offboardUserIdentifier = if ($PSCmdlet.ParameterSetName -eq "ObjectId") { $ObjectId } else { $UserPrincipalName }
        Write-ScriptLog -Message "Invoke-AADUserOffboarding: Target user '$($script:offboardUserIdentifier)', Action: $Action." -Level DEBUG -LogFilePath $LogFilePath
    }
    process {
        $processMsgDetail = "Offboard user '$($script:offboardUserIdentifier)' (Action: $Action"
        $additionalDesc = @()
        if ($RevokeSignInSessions) { $additionalDesc += "Revoke Sessions" } else {$additionalDesc += "Keep Sessions"}
        if ($RemoveAllLicenses) { $additionalDesc += "Remove Licenses" } else {$additionalDesc += "Keep Licenses"}
        if ($RemoveFromAllGroups) { $additionalDesc += "Remove from Groups" } else {$additionalDesc += "Keep in Groups"}
        $processMsgDetail += ", Options: $($additionalDesc -join '/'))"

        if (-not ($PSCmdlet.ShouldProcess($script:offboardUserIdentifier, $processMsgDetail))) {
            Write-ScriptLog -Message "Invoke-AADUserOffboarding: Action '$Action' for '$($script:offboardUserIdentifier)' skipped (-WhatIf)." -Level WARN -LogFilePath $LogFilePath
            return "Skipped: Offboarding for '$($script:offboardUserIdentifier)'."
        }
        try {
            $user = if ($PSCmdlet.ParameterSetName -eq "ObjectId") { Get-AzADUser -ObjectId $script:offboardUserIdentifier -ErrorAction Stop } else { Get-AzADUser -UserPrincipalName $script:offboardUserIdentifier -ErrorAction Stop }
            if (-not $user) { Write-ScriptLog -Message "Invoke-AADUserOffboarding: User '$($script:offboardUserIdentifier)' not found." -Level ERROR -LogFilePath $LogFilePath; return "Error: User not found." }
            Write-ScriptLog -Message "Invoke-AADUserOffboarding: User '$($user.UserPrincipalName)' (ID: $($user.Id)) found. Action: $Action." -Level INFO -LogFilePath $LogFilePath
            
            if ($user.AccountEnabled) {
                if ($PSCmdlet.ShouldProcess($user.UserPrincipalName, "Disable account")) {
                    Write-ScriptLog -Message "Invoke-AADUserOffboarding: Disabling account for '$($user.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath
                    Set-AzADUser -ObjectId $user.Id -AccountEnabled $false -ErrorAction Stop
                    Write-ScriptLog -Message "Invoke-AADUserOffboarding: Account for '$($user.UserPrincipalName)' disabled." -Level INFO -LogFilePath $LogFilePath
                } else { Write-ScriptLog -Message "Invoke-AADUserOffboarding: Account disable for '$($user.UserPrincipalName)' skipped (-WhatIf)." -Level WARN -LogFilePath $LogFilePath }
            } else { Write-ScriptLog -Message "Invoke-AADUserOffboarding: Account for '$($user.UserPrincipalName)' already disabled." -Level INFO -LogFilePath $LogFilePath }

            if ($RevokeSignInSessions) {
                if ($PSCmdlet.ShouldProcess($user.UserPrincipalName, "Revoke sign-in sessions")) {
                    Write-ScriptLog -Message "Invoke-AADUserOffboarding: Revoking sessions for '$($user.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath
                    try { Revoke-AzADUserSignInSession -ObjectId $user.Id -ErrorAction Stop; Write-ScriptLog -Message "Invoke-AADUserOffboarding: Sessions revoked for '$($user.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath }
                    catch { Write-ScriptLog -Message "Invoke-AADUserOffboarding: Failed to revoke sessions for '$($user.UserPrincipalName)'. Error: $($_.Exception.Message). Continuing." -Level WARN -LogFilePath $LogFilePath }
                } else { Write-ScriptLog -Message "Invoke-AADUserOffboarding: Session revocation for '$($user.UserPrincipalName)' skipped (-WhatIf)." -Level WARN -LogFilePath $LogFilePath }
            }
            if ($RemoveAllLicenses) {
                if ($PSCmdlet.ShouldProcess($user.UserPrincipalName, "Remove all licenses")) {
                    Write-ScriptLog -Message "Invoke-AADUserOffboarding: Removing licenses for '$($user.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath
                    try {
                        if ((Get-AzADUserAssignedLicense -ObjectId $user.Id -ErrorAction SilentlyContinue).Count -gt 0) {
                            Set-AzADUser -ObjectId $user.Id -AssignedLicenses @() -ErrorAction Stop; Write-ScriptLog -Message "Invoke-AADUserOffboarding: Licenses removed for '$($user.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath
                        } else { Write-ScriptLog -Message "Invoke-AADUserOffboarding: No licenses found for '$($user.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath }
                    } catch { Write-ScriptLog -Message "Invoke-AADUserOffboarding: Failed to remove licenses for '$($user.UserPrincipalName)'. Error: $($_.Exception.Message). Continuing." -Level WARN -LogFilePath $LogFilePath }
                } else { Write-ScriptLog -Message "Invoke-AADUserOffboarding: License removal for '$($user.UserPrincipalName)' skipped (-WhatIf)." -Level WARN -LogFilePath $LogFilePath }
            }
            if ($RemoveFromAllGroups) {
                if ($PSCmdlet.ShouldProcess($user.UserPrincipalName, "Remove from ALL groups")) {
                    Write-ScriptLog -Message "Invoke-AADUserOffboarding: Removing '$($user.UserPrincipalName)' from all groups." -Level INFO -LogFilePath $LogFilePath
                    try {
                        $memberships = Get-AzADUserMembership -ObjectId $user.Id -ErrorAction Stop
                        if ($memberships.Count -gt 0) {
                            Write-ScriptLog -Message "Invoke-AADUserOffboarding: User '$($user.UserPrincipalName)' in $($memberships.Count) groups. Starting removal." -Level INFO -LogFilePath $LogFilePath
                            $memberships | ForEach-Object {
                                $grpId = $_.Id; $grpName = if ($_.AdditionalProperties.displayName) { $_.AdditionalProperties.displayName } else { "(ID: $grpId)"}
                                if ($PSCmdlet.ShouldProcess("Group: $grpName", "Remove user '$($user.UserPrincipalName)' from this group")) {
                                    try { Remove-AzADGroupMember -TargetGroupObjectId $grpId -MemberObjectId $user.Id -ErrorAction Stop; Write-ScriptLog -Message "Invoke-AADUserOffboarding: Removed '$($user.UserPrincipalName)' from group '$grpName'." -Level INFO -LogFilePath $LogFilePath }
                                    catch { Write-ScriptLog -Message "Invoke-AADUserOffboarding: Failed to remove '$($user.UserPrincipalName)' from group '$grpName'. Error: $($_.Exception.Message). Continuing." -Level WARN -LogFilePath $LogFilePath }
                                } else { Write-ScriptLog -Message "Invoke-AADUserOffboarding: Removal from group '$grpName' for '$($user.UserPrincipalName)' skipped (-WhatIf)." -Level WARN -LogFilePath $LogFilePath }
                            }
                        } else { Write-ScriptLog -Message "Invoke-AADUserOffboarding: User '$($user.UserPrincipalName)' not in any groups." -Level INFO -LogFilePath $LogFilePath }
                    } catch { Write-ScriptLog -Message "Invoke-AADUserOffboarding: Error retrieving/processing groups for '$($user.UserPrincipalName)'. Error: $($_.Exception.Message). Continuing." -Level WARN -LogFilePath $LogFilePath }
                } else { Write-ScriptLog -Message "Invoke-AADUserOffboarding: Removal from all groups for '$($user.UserPrincipalName)' skipped (-WhatIf)." -Level WARN -LogFilePath $LogFilePath }
            }
            if ($Action -eq "Delete") {
                Write-ScriptLog -Message "Invoke-AADUserOffboarding: Proceeding with SOFT DELETION of '$($user.UserPrincipalName)'." -Level WARN -LogFilePath $LogFilePath
                if ($PSCmdlet.ShouldProcess($user.UserPrincipalName, "SOFT-DELETE user account.")) {
                    try {
                        Remove-AzADUser -ObjectId $user.Id -ErrorAction Stop
                        Write-ScriptLog -Message "Invoke-AADUserOffboarding: User '$($user.UserPrincipalName)' soft-deleted." -Level INFO -LogFilePath $LogFilePath
                        return "Success: User '$($user.UserPrincipalName)' soft-deleted."
                    } catch { Write-ScriptLog -Message "Invoke-AADUserOffboarding: Failed to soft-delete '$($user.UserPrincipalName)'. Error: $($_.Exception.Message)" -Level ERROR -LogFilePath $LogFilePath; return "Error: Failed to soft-delete. Message: $($_.Exception.Message)" }
                } else { Write-ScriptLog -Message "Invoke-AADUserOffboarding: Soft deletion of '$($user.UserPrincipalName)' skipped (-WhatIf)." -Level WARN -LogFilePath $LogFilePath; return "Skipped: Soft deletion of '$($user.UserPrincipalName)'." }
            }
            Write-ScriptLog -Message "Invoke-AADUserOffboarding: Action '$Action' completed for '$($user.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath
            return "Success: User '$($user.UserPrincipalName)' action '$Action' completed."
        } catch [Microsoft.Azure.Commands.MicrosoftGraph.Cmdlets.Users.Models.MicrosoftGraphServiceException] {
             if ($_.Exception.Response.StatusCode -eq [System.Net.HttpStatusCode]::NotFound) { Write-ScriptLog -Message "Invoke-AADUserOffboarding: User '$($script:offboardUserIdentifier)' not found/already deleted. $($_.Exception.Message)" -Level ERROR -LogFilePath $LogFilePath; return "Error: User not found/deleted. Graph API: $($_.Exception.Message)" }
             else { Write-ScriptLog -Message "Invoke-AADUserOffboarding: Graph service error for '$($script:offboardUserIdentifier)': $($_.Exception.Message)" -Level ERROR -LogFilePath $LogFilePath; return "Error: Graph API error. Message: $($_.Exception.Message)" }
        } catch {
            Write-ScriptLog -Message "Invoke-AADUserOffboarding: Unexpected error for '$($script:offboardUserIdentifier)'. Error: $($_.Exception.Message)" -Level ERROR -LogFilePath $LogFilePath
            if ($_.Exception.StackTrace) { Write-ScriptLog -Message "Invoke-AADUserOffboarding: StackTrace for '$($script:offboardUserIdentifier)': $($_.Exception.StackTrace)" -Level DEBUG -LogFilePath $LogFilePath }
            return "Error: Unexpected error. Message: $($_.Exception.Message)"
        }
    }
    end { Write-ScriptLog -Message "Invoke-AADUserOffboarding: Finished for user '$($script:offboardUserIdentifier)'." -Level VERBOSE -LogFilePath $LogFilePath }
}

#--------------------------------------------------------------------------------
# Bulk Onboarding Function (Invoke-BulkAADUserOnboarding)
#--------------------------------------------------------------------------------
function Invoke-BulkAADUserOnboarding {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true)] [string]$CSVPath,
        [Parameter(Mandatory = $true)] [string]$LogFilePath
    )
    # Write-ScriptLog is assumed to be defined globally by the main script body
    begin {
        Write-ScriptLog -Message "Invoke-BulkAADUserOnboarding starting." -Level INFO -LogFilePath $LogFilePath
        # Headers logged by main script's Write-ScriptLog call
        if (-not (Test-Path -Path $CSVPath -PathType Leaf)) { Write-ScriptLog -Message "CSV '$CSVPath' not found." -Level ERROR -LogFilePath $LogFilePath; return }
        if (-not (Get-Command Invoke-AADUserOnboarding -ErrorAction SilentlyContinue)) { Write-ScriptLog -Message "Dependency Invoke-AADUserOnboarding missing." -Level ERROR -LogFilePath $LogFilePath; return }
        $script:bulkOnboardSummary = @{ SuccessCount = 0; FailureCount = 0; FailedEntries = [System.Collections.Generic.List[hashtable]]::new(); ProcessedRows = 0; TotalRows = 0 }
    }
    process {
        try {
            $csvData = Import-Csv -Path $CSVPath -ErrorAction Stop
            $script:bulkOnboardSummary.TotalRows = $csvData.Count
            Write-ScriptLog -Message "Loaded $($script:bulkOnboardSummary.TotalRows) rows from CSV '$CSVPath'." -Level INFO -LogFilePath $LogFilePath
            foreach ($row in $csvData) {
                $script:bulkOnboardSummary.ProcessedRows++
                $currentUPN = $null; $rawUPN = if ($row.PSObject.Properties['UserPrincipalName']) {$row.UserPrincipalName} else {"<UPN_COL_MISSING>"}
                try {
                    if ($row.PSObject.Properties['UserPrincipalName']) { $currentUPN = $row.UserPrincipalName.Trim(); if ([string]::IsNullOrWhiteSpace($currentUPN)) { throw "UPN empty."}} else { throw "UPN column missing."}
                    Write-ScriptLog -Message "Processing row $($script:bulkOnboardSummary.ProcessedRows)/$($script:bulkOnboardSummary.TotalRows): '$currentUPN'" -Level INFO -LogFilePath $LogFilePath
                    if ($PSCmdlet.ShouldProcess($currentUPN, "Onboard from CSV row $($script:bulkOnboardSummary.ProcessedRows)")) {
                        $params = @{ LogFilePath = $LogFilePath; UserPrincipalName = $currentUPN }
                        if ($row.PSObject.Properties['DisplayName']) { $params.DisplayName = $row.DisplayName.Trim(); if([string]::IsNullOrWhiteSpace($params.DisplayName)){throw "DisplayName empty."}} else {throw "DisplayName column missing."}
                        if ($row.PSObject.Properties['Password'] -and -not [string]::IsNullOrWhiteSpace($row.Password)) {
                            $params.Password = ConvertTo-SecureString $row.Password.Trim() -AsPlainText -Force
                            if ($row.PSObject.Properties['ForceChangePasswordNextLogin']-and -not [string]::IsNullOrWhiteSpace($row.ForceChangePasswordNextLogin)){try {$params.ForceChangePasswordNextLogin=[System.Convert]::ToBoolean($row.ForceChangePasswordNextLogin.Trim())}catch{throw "Invalid ForceChangePasswordNextLogin boolean."}}
                        } elseif ($row.PSObject.Properties['ForceChangePasswordNextLogin']-and -not [string]::IsNullOrWhiteSpace($row.ForceChangePasswordNextLogin)){try {$params.ForceChangePasswordNextLogin=[System.Convert]::ToBoolean($row.ForceChangePasswordNextLogin.Trim())}catch{throw "Invalid ForceChangePasswordNextLogin boolean."}}
                        if ($row.PSObject.Properties['InitialGroups'] -and -not [string]::IsNullOrWhiteSpace($row.InitialGroups)) {$params.InitialGroups = $row.InitialGroups.Split(',')|%{ $_.Trim()}|?{-not[string]::IsNullOrWhiteSpace($_)}}
                        if ($row.PSObject.Properties['LicenseSKUs'] -and -not [string]::IsNullOrWhiteSpace($row.LicenseSKUs)) {$params.LicenseSKUs = $row.LicenseSKUs.Split(',')|%{ $_.Trim()}|?{-not[string]::IsNullOrWhiteSpace($_)}}
                        if ($row.PSObject.Properties['Department'] -and -not [string]::IsNullOrWhiteSpace($row.Department)) {$params.Department = $row.Department.Trim()}
                        if ($row.PSObject.Properties['JobTitle'] -and -not [string]::IsNullOrWhiteSpace($row.JobTitle)) {$params.JobTitle = $row.JobTitle.Trim()}
                        if ($row.PSObject.Properties['ManagerUPN'] -and -not [string]::IsNullOrWhiteSpace($row.ManagerUPN)) {$params.ManagerUPN = $row.ManagerUPN.Trim()}
                        if ($PSBoundParameters.Verbose) { $params.Verbose = $true }
                        if ($PSBoundParameters.Debug) { $params.Debug = $true }
                        
                        $res = Invoke-AADUserOnboarding @params
                        if ($res -and $res.Id) { Write-ScriptLog -Message "Success: Onboarded '$currentUPN'." -Level INFO -LogFilePath $LogFilePath; $script:bulkOnboardSummary.SuccessCount++ }
                        else { throw "Onboarding failed for '$currentUPN'. Result: $res" }
                    } else { Write-ScriptLog -Message "Skipped '$currentUPN' (row $($script:bulkOnboardSummary.ProcessedRows)) (-WhatIf)." -Level WARN -LogFilePath $LogFilePath }
                } catch {
                    $err = if($currentUPN){"Error for UPN '$currentUPN'" } else {"Error for raw UPN '$rawUPN'"}
                    Write-ScriptLog -Message "$err (row $($script:bulkOnboardSummary.ProcessedRows)): $($_.Exception.Message)" -Level ERROR -LogFilePath $LogFilePath
                    $script:bulkOnboardSummary.FailureCount++; $script:bulkOnboardSummary.FailedEntries.Add(@{Row=$script:bulkOnboardSummary.ProcessedRows; ID=$currentUPN; Error=$_.Exception.Message; CSV=$row|ConvertTo-Json -Compress})
                }
            }
        } catch { Write-ScriptLog -Message "Failed to read/process CSV '$CSVPath'. Error: $($_.Exception.Message)" -Level ERROR -LogFilePath $LogFilePath }
    }
    end {
        Write-ScriptLog -Message "Bulk Onboarding Summary: Total: $($script:bulkOnboardSummary.TotalRows), Success: $($script:bulkOnboardSummary.SuccessCount), Failed: $($script:bulkOnboardSummary.FailureCount)." -Level INFO -LogFilePath $LogFilePath
        if ($script:bulkOnboardSummary.FailureCount -gt 0) { $script:bulkOnboardSummary.FailedEntries | ForEach-Object { Write-ScriptLog -Message "Failed Row $($_.Row) User: $($_.ID) Error: $($_.Error)" -Level WARN -LogFilePath $LogFilePath } }
        Write-ScriptLog -Message "Invoke-BulkAADUserOnboarding finished." -Level INFO -LogFilePath $LogFilePath
        return $script:bulkOnboardSummary
    }
}

#--------------------------------------------------------------------------------
# Bulk Update Function (Invoke-BulkAADUserUpdate)
#--------------------------------------------------------------------------------
function Invoke-BulkAADUserUpdate {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true)] [string]$CSVPath,
        [Parameter(Mandatory = $true)] [string]$LogFilePath
    )
    begin {
        Write-ScriptLog -Message "Invoke-BulkAADUserUpdate starting." -Level INFO -LogFilePath $LogFilePath
        if (-not (Test-Path -Path $CSVPath -PathType Leaf)) { Write-ScriptLog -Message "CSV '$CSVPath' not found." -Level ERROR -LogFilePath $LogFilePath; return }
        if (-not (Get-Command Update-AADUser -ErrorAction SilentlyContinue)) { Write-ScriptLog -Message "Dependency Update-AADUser missing." -Level ERROR -LogFilePath $LogFilePath; return }
        $script:updateUserParamsList = (Get-Command Update-AADUser).Parameters.Keys | ForEach-Object { $_.ToLowerInvariant() }
        $script:bulkUpdateSummary = @{ SuccessCount = 0; FailureCount = 0; FailedEntries = [System.Collections.Generic.List[hashtable]]::new(); ProcessedRows = 0; TotalRows = 0 }
    }
    process {
        try {
            $csvData = Import-Csv -Path $CSVPath -ErrorAction Stop
            $script:bulkUpdateSummary.TotalRows = $csvData.Count
            Write-ScriptLog -Message "Loaded $($script:bulkUpdateSummary.TotalRows) rows from CSV '$CSVPath'." -Level INFO -LogFilePath $LogFilePath
            foreach ($row in $csvData) {
                $script:bulkUpdateSummary.ProcessedRows++; $currentUserIdent = "<MISSING_IDENT>"
                try {
                    $params = @{ LogFilePath = $LogFilePath }; $identProvided = $false
                    if ($row.PSObject.Properties['UserPrincipalName'] -and -not [string]::IsNullOrWhiteSpace($row.UserPrincipalName)) { $params.UserPrincipalName = $row.UserPrincipalName.Trim(); $currentUserIdent = $params.UserPrincipalName; $identProvided = $true }
                    elseif ($row.PSObject.Properties['ObjectId'] -and -not [string]::IsNullOrWhiteSpace($row.ObjectId)) { $params.ObjectId = $row.ObjectId.Trim(); $currentUserIdent = $params.ObjectId; $identProvided = $true }
                    if (-not $identProvided) { throw "Identifier (UPN or ObjectId) missing." }
                    Write-ScriptLog -Message "Processing row $($script:bulkUpdateSummary.ProcessedRows)/$($script:bulkUpdateSummary.TotalRows): '$currentUserIdent'" -Level INFO -LogFilePath $LogFilePath
                    if ($PSCmdlet.ShouldProcess($currentUserIdent, "Update from CSV row $($script:bulkUpdateSummary.ProcessedRows)")) {
                        foreach ($prop in $row.PSObject.Properties) {
                            $pName = $prop.Name; $pValue = $prop.Value
                            if ($pName -eq 'UserPrincipalName' -or $pName -eq 'ObjectId') { continue }
                            if ($script:updateUserParamsList -contains $pName.ToLowerInvariant() -and -not [string]::IsNullOrWhiteSpace($pValue)) {
                                switch ($pName.ToLowerInvariant()) {
                                    'groupstoadd';'groupstoremove';'licensestoassign';'licensestoremove' {$params[$pName]=$pValue.Split(',')|%{$_.Trim()}|?{-not[string]::IsNullOrWhiteSpace($_)}}
                                    default { if ($pName.ToLowerInvariant() -eq 'managerupn') {$params[$pName]=$pValue.Trim()} else {$params[$pName]=$pValue.Trim()}}
                                }
                            } elseif ($script:updateUserParamsList -contains $pName.ToLowerInvariant() -and $pName.ToLowerInvariant() -eq 'managerupn' -and $row.PSObject.Properties['ManagerUPN'] -and [string]::IsNullOrEmpty($pValue)) { $params[$pName] = "" }
                        }
                        if ($params.Count -le 2) { throw "No update attributes found for '$currentUserIdent'." }
                        if ($PSBoundParameters.Verbose) { $params.Verbose = $true }
                        if ($PSBoundParameters.Debug) { $params.Debug = $true }
                        
                        $res = Update-AADUser @params
                        if ($res -and $res.Id) { Write-ScriptLog -Message "Success: Updated '$currentUserIdent'." -Level INFO -LogFilePath $LogFilePath; $script:bulkUpdateSummary.SuccessCount++ }
                        else { throw "Update failed for '$currentUserIdent'. Result: $res" }
                    } else { Write-ScriptLog -Message "Skipped '$currentUserIdent' (row $($script:bulkUpdateSummary.ProcessedRows)) (-WhatIf)." -Level WARN -LogFilePath $LogFilePath }
                } catch {
                    Write-ScriptLog -Message "Error for user '$currentUserIdent' (row $($script:bulkUpdateSummary.ProcessedRows)): $($_.Exception.Message)" -Level ERROR -LogFilePath $LogFilePath
                    $script:bulkUpdateSummary.FailureCount++; $script:bulkUpdateSummary.FailedEntries.Add(@{Row=$script:bulkUpdateSummary.ProcessedRows;ID=$currentUserIdent;Error=$_.Exception.Message;CSV=$row|ConvertTo-Json-Compress})
                }
            }
        } catch { Write-ScriptLog -Message "Failed to read/process CSV '$CSVPath'. Error: $($_.Exception.Message)" -Level ERROR -LogFilePath $LogFilePath }
    }
    end {
        Write-ScriptLog -Message "Bulk Update Summary: Total: $($script:bulkUpdateSummary.TotalRows), Success: $($script:bulkUpdateSummary.SuccessCount), Failed: $($script:bulkUpdateSummary.FailureCount)." -Level INFO -LogFilePath $LogFilePath
        if ($script:bulkUpdateSummary.FailureCount -gt 0) { $script:bulkUpdateSummary.FailedEntries | ForEach-Object { Write-ScriptLog -Message "Failed Row $($_.Row) User: $($_.ID) Error: $($_.Error)" -Level WARN -LogFilePath $LogFilePath } }
        Write-ScriptLog -Message "Invoke-BulkAADUserUpdate finished." -Level INFO -LogFilePath $LogFilePath
        return $script:bulkUpdateSummary
    }
}

#--------------------------------------------------------------------------------
# Bulk Offboarding Function (Invoke-BulkAADUserOffboarding)
#--------------------------------------------------------------------------------
function Invoke-BulkAADUserOffboarding {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true)] [string]$CSVPath,
        [Parameter(Mandatory = $true)] [string]$LogFilePath
    )
    begin {
        Write-ScriptLog -Message "Invoke-BulkAADUserOffboarding starting." -Level INFO -LogFilePath $LogFilePath
        if (-not (Test-Path -Path $CSVPath -PathType Leaf)) { Write-ScriptLog -Message "CSV '$CSVPath' not found." -Level ERROR -LogFilePath $LogFilePath; return }
        if (-not (Get-Command Invoke-AADUserOffboarding -ErrorAction SilentlyContinue)) { Write-ScriptLog -Message "Dependency Invoke-AADUserOffboarding missing." -Level ERROR -LogFilePath $LogFilePath; return }
        $script:bulkOffboardSummary = @{ SuccessCount = 0; FailureCount = 0; FailedEntries = [System.Collections.Generic.List[hashtable]]::new(); ProcessedRows = 0; TotalRows = 0 }
    }
    process {
        try {
            $csvData = Import-Csv -Path $CSVPath -ErrorAction Stop
            $script:bulkOffboardSummary.TotalRows = $csvData.Count
            Write-ScriptLog -Message "Loaded $($script:bulkOffboardSummary.TotalRows) rows from CSV '$CSVPath'." -Level INFO -LogFilePath $LogFilePath
            foreach ($row in $csvData) {
                $script:bulkOffboardSummary.ProcessedRows++; $currentUserIdent = "<MISSING_IDENT>"
                try {
                    $params = @{ LogFilePath = $LogFilePath }; $identProvided = $false
                    if ($row.PSObject.Properties['UserPrincipalName'] -and -not [string]::IsNullOrWhiteSpace($row.UserPrincipalName)) { $params.UserPrincipalName = $row.UserPrincipalName.Trim(); $currentUserIdent = $params.UserPrincipalName; $identProvided = $true }
                    elseif ($row.PSObject.Properties['ObjectId'] -and -not [string]::IsNullOrWhiteSpace($row.ObjectId)) { $params.ObjectId = $row.ObjectId.Trim(); $currentUserIdent = $params.ObjectId; $identProvided = $true }
                    if (-not $identProvided) { throw "Identifier (UPN or ObjectId) missing." }
                    Write-ScriptLog -Message "Processing row $($script:bulkOffboardSummary.ProcessedRows)/$($script:bulkOffboardSummary.TotalRows): '$currentUserIdent'" -Level INFO -LogFilePath $LogFilePath
                    
                    if ($row.PSObject.Properties['Action'] -and -not [string]::IsNullOrWhiteSpace($row.Action)) { $act = $row.Action.Trim(); if($act -in "Disable","Delete"){$params.Action=$act}else{throw "Invalid Action value."}}
                    $boolParamsMap = @("RevokeSignInSessions", "RemoveAllLicenses", "RemoveFromAllGroups")
                    $boolParamsMap | ForEach-Object { if($row.PSObject.Properties[$_] -and -not [string]::IsNullOrWhiteSpace($row.$_)){try{$params[$_]=[System.Convert]::ToBoolean($row.$_.Trim())}catch{throw "Invalid boolean for $_."}}}
                    
                    $effAction = if($params.Action){$params.Action}else{"Disable"}
                    $shouldProcessMsg = "Offboard user '$currentUserIdent' (Action: $effAction"
                    $desc = @()
                    if($params.ContainsKey('RevokeSignInSessions')){$desc += "RevokeSess:$($params.RevokeSignInSessions)"}else{$desc += "RevokeSess:DefaultTrue"}
                    if($params.ContainsKey('RemoveAllLicenses')){$desc += "RemLic:$($params.RemoveAllLicenses)"}else{$desc += "RemLic:DefaultTrue"}
                    if($params.ContainsKey('RemoveFromAllGroups')){$desc += "RemGrps:$($params.RemoveFromAllGroups)"}else{$desc += "RemGrps:DefaultFalse"}
                    if($desc.Count -gt 0){$shouldProcessMsg += ", Options: $($desc -join '/'))"}else{$shouldProcessMsg += ")"}

                    if ($PSCmdlet.ShouldProcess($currentUserIdent, $shouldProcessMsg)) {
                        if ($PSBoundParameters.Verbose) { $params.Verbose = $true }
                        if ($PSBoundParameters.Debug) { $params.Debug = $true }
                        
                        $res = Invoke-AADUserOffboarding @params
                        if ($res -is [string] -and $res.StartsWith("Success:")) { Write-ScriptLog -Message "Success: Offboarded '$currentUserIdent'. Result: $res" -Level INFO -LogFilePath $LogFilePath; $script:bulkOffboardSummary.SuccessCount++ }
                        elseif ($res -is [string] -and ($res.StartsWith("Skipped:") -or $res.StartsWith("Warn:"))) { Write-ScriptLog -Message "Skipped/Warn for '$currentUserIdent': $res" -Level WARN -LogFilePath $LogFilePath; $script:bulkOffboardSummary.FailureCount++; $script:bulkOffboardSummary.FailedEntries.Add(@{Row=$script:bulkOffboardSummary.ProcessedRows;ID=$currentUserIdent;Error="Skipped/Warn: $res";CSV=$row|ConvertTo-Json-Compress}) }
                        else { throw "Offboarding failed for '$currentUserIdent'. Result: $res" }
                    } else { Write-ScriptLog -Message "Skipped '$currentUserIdent' (row $($script:bulkOffboardSummary.ProcessedRows)) (-WhatIf)." -Level WARN -LogFilePath $LogFilePath }
                } catch {
                    Write-ScriptLog -Message "Error for user '$currentUserIdent' (row $($script:bulkOffboardSummary.ProcessedRows)): $($_.Exception.Message)" -Level ERROR -LogFilePath $LogFilePath
                    $script:bulkOffboardSummary.FailureCount++; $script:bulkOffboardSummary.FailedEntries.Add(@{Row=$script:bulkOffboardSummary.ProcessedRows;ID=$currentUserIdent;Error=$_.Exception.Message;CSV=$row|ConvertTo-Json-Compress})
                }
            }
        } catch { Write-ScriptLog -Message "Failed to read/process CSV '$CSVPath'. Error: $($_.Exception.Message)" -Level ERROR -LogFilePath $LogFilePath }
    }
    end {
        Write-ScriptLog -Message "Bulk Offboarding Summary: Total: $($script:bulkOffboardSummary.TotalRows), Success: $($script:bulkOffboardSummary.SuccessCount), Failed: $($script:bulkOffboardSummary.FailureCount)." -Level INFO -LogFilePath $LogFilePath
        if ($script:bulkOffboardSummary.FailureCount -gt 0) { $script:bulkOffboardSummary.FailedEntries | ForEach-Object { Write-ScriptLog -Message "Failed Row $($_.Row) User: $($_.ID) Detail: $($_.Error)" -Level WARN -LogFilePath $LogFilePath } }
        Write-ScriptLog -Message "Invoke-BulkAADUserOffboarding finished." -Level INFO -LogFilePath $LogFilePath
        return $script:bulkOffboardSummary
    }
}

#endregion Function Definitions

# --- Main Script Logic ---
function Main {
    # Default LogFilePath
    if (-not $PSBoundParameters.ContainsKey('LogFilePath') -or [string]::IsNullOrWhiteSpace($LogFilePath)) {
        $defaultLogDir = Join-Path -Path $PSScriptRoot -ChildPath "ScriptLogs"
        $timestampForLog = Get-Date -Format "yyyyMMdd_HHmmss"
        $LogFilePath = Join-Path -Path $defaultLogDir -ChildPath "AzureADLifecycle_$timestampForLog.log"
        Write-Host "LogFilePath not specified. Defaulting to: $LogFilePath"
    }

    # Ensure log directory exists (Write-ScriptLog also does this, but good for transcript path too)
    $logDir = Split-Path -Path $LogFilePath -Parent
    if (-not (Test-Path -Path $logDir -PathType Container)) {
        try {
            New-Item -ItemType Directory -Path $logDir -Force -ErrorAction Stop | Out-Null
            Write-Host "Created log directory: $logDir"
        } catch {
            Write-Error "Failed to create log directory '$logDir'. Aborting. Error: $($_.Exception.Message)"
            return
        }
    }

    # Start Transcript
    $transcriptPath = Join-Path -Path $logDir -ChildPath "Transcript_AzureADLifecycle_$(Get-Date -Format yyyyMMddHHmmss).log"
    try {
        Start-Transcript -Path $transcriptPath -Append -ErrorAction Stop
        Write-ScriptLog -Message "Transcript started at $transcriptPath" -Level INFO -LogFilePath $LogFilePath
    } catch {
        Write-Error "Failed to start transcript at '$transcriptPath'. Error: $($_.Exception.Message)"
        # Continue without transcript if it fails, but log the attempt.
        Write-ScriptLog -Message "Failed to start transcript at '$transcriptPath'. Error: $($_.Exception.Message)" -Level ERROR -LogFilePath $LogFilePath
    }

    try {
        Write-ScriptLog -Message "Manage-AzureADUserLifecycle script started. Action: $LifecycleAction" -Level INFO -LogFilePath $LogFilePath
        
        # Parameter validation for single operations based on LifecycleAction
        switch ($LifecycleAction) {
            "SingleOnboard" {
                if (-not $UserPrincipalName -or -not $DisplayName) {
                    Write-ScriptLog -Message "For SingleOnboard, -UserPrincipalName and -DisplayName are mandatory." -Level ERROR -LogFilePath $LogFilePath
                    return
                }
            }
            "SingleUpdate" {
                if ( ($PSCmdlet.ParameterSetName -eq "SingleUpdateUPN" -and -not $UserPrincipalName) -or `
                     ($PSCmdlet.ParameterSetName -eq "SingleUpdateObjectId" -and -not $ObjectId) ) {
                     # This specific check might be redundant due to PowerShell's own mandatory param enforcement per set
                     # but good for explicit understanding.
                    Write-ScriptLog -Message "For SingleUpdate, either -UserPrincipalName or -ObjectId is mandatory." -Level ERROR -LogFilePath $LogFilePath
                    return
                }
                # Check if at least one updateable attribute is provided beyond identifiers and LogFilePath
                $updateAttributeParams = $PSBoundParameters.Keys | Where-Object { $_ -notin @('LifecycleAction', 'UserPrincipalName', 'ObjectId', 'LogFilePath', 'WhatIf', 'Confirm', 'Verbose', 'Debug', 'ErrorAction', 'ErrorVariable', 'OutVariable', 'OutBuffer', 'PipelineVariable', 'WarningAction', 'WarningVariable')}
                if ($updateAttributeParams.Count -eq 0) {
                    Write-ScriptLog -Message "For SingleUpdate, at least one attribute to update must be specified." -Level ERROR -LogFilePath $LogFilePath
                    return
                }
            }
            "SingleOffboard" {
                 if ( ($PSCmdlet.ParameterSetName -eq "SingleOffboardUPN" -and -not $UserPrincipalName) -or `
                      ($PSCmdlet.ParameterSetName -eq "SingleOffboardObjectId" -and -not $ObjectId) ) {
                    Write-ScriptLog -Message "For SingleOffboard, either -UserPrincipalName or -ObjectId is mandatory." -Level ERROR -LogFilePath $LogFilePath
                    return
                }
            }
            "BulkOnboard" ; "BulkUpdate" ; "BulkOffboard" {
                if (-not $CSVPath) { # Should be caught by Mandatory=true on CSVPath for BulkOperation ParameterSet
                    Write-ScriptLog -Message "For Bulk operations, -CSVPath is mandatory." -Level ERROR -LogFilePath $LogFilePath
                    return
                }
            }
        }


        switch ($LifecycleAction) {
            "SingleOnboard" {
                $onboardParams = @{
                    UserPrincipalName = $UserPrincipalName
                    DisplayName = $DisplayName
                    LogFilePath = $LogFilePath
                }
                if ($PSBoundParameters.ContainsKey('Password')) { $onboardParams.Password = $Password }
                if ($PSBoundParameters.ContainsKey('ForceChangePasswordNextLogin')) { $onboardParams.ForceChangePasswordNextLogin = $ForceChangePasswordNextLogin }
                if ($PSBoundParameters.ContainsKey('InitialGroups')) { $onboardParams.InitialGroups = $InitialGroups }
                if ($PSBoundParameters.ContainsKey('LicenseSKUs')) { $onboardParams.LicenseSKUs = $LicenseSKUs }
                if ($PSBoundParameters.ContainsKey('Department')) { $onboardParams.Department = $Department }
                if ($PSBoundParameters.ContainsKey('JobTitle')) { $onboardParams.JobTitle = $JobTitle }
                if ($PSBoundParameters.ContainsKey('ManagerUPN')) { $onboardParams.ManagerUPN = $ManagerUPN }
                if ($PSBoundParameters.Verbose) { $onboardParams.Verbose = $true }
                if ($PSBoundParameters.Debug) { $onboardParams.Debug = $true }

                Invoke-AADUserOnboarding @onboardParams
            }
            "SingleUpdate" {
                $updateParams = @{ LogFilePath = $LogFilePath }
                if ($PSCmdlet.ParameterSetName -eq "SingleUpdateUPN") { $updateParams.UserPrincipalName = $UserPrincipalName }
                elseif ($PSCmdlet.ParameterSetName -eq "SingleUpdateObjectId") { $updateParams.ObjectId = $ObjectId }

                # Add all other relevant bound parameters for Update-AADUser
                $paramExclusionsForSingleUpdate = @('LifecycleAction', 'UserPrincipalName', 'ObjectId', 'CSVPath', 'LogFilePath', 'WhatIf', 'Confirm', 'Verbose', 'Debug', 'ErrorAction', 'ErrorVariable', 'OutVariable', 'OutBuffer', 'PipelineVariable', 'WarningAction', 'WarningVariable')
                $PSBoundParameters.Keys | Where-Object { $paramExclusionsForSingleUpdate -notcontains $_ } | ForEach-Object {
                    $updateParams[$_] = $PSBoundParameters[$_]
                }
                if ($PSBoundParameters.Verbose) { $updateParams.Verbose = $true }
                if ($PSBoundParameters.Debug) { $updateParams.Debug = $true }
                
                Update-AADUser @updateParams
            }
            "SingleOffboard" {
                $offboardParams = @{ LogFilePath = $LogFilePath }
                if ($PSCmdlet.ParameterSetName -eq "SingleOffboardUPN") { $offboardParams.UserPrincipalName = $UserPrincipalName }
                elseif ($PSCmdlet.ParameterSetName -eq "SingleOffboardObjectId") { $offboardParams.ObjectId = $ObjectId }
                
                if ($PSBoundParameters.ContainsKey('OffboardAction')) { $offboardParams.Action = $OffboardAction }
                if ($PSBoundParameters.ContainsKey('RevokeSignInSessions')) { $offboardParams.RevokeSignInSessions = $RevokeSignInSessions }
                if ($PSBoundParameters.ContainsKey('RemoveAllLicenses')) { $offboardParams.RemoveAllLicenses = $RemoveAllLicenses }
                if ($PSBoundParameters.ContainsKey('RemoveFromAllGroups')) { $offboardParams.RemoveFromAllGroups = $RemoveFromAllGroups }
                if ($PSBoundParameters.Verbose) { $offboardParams.Verbose = $true }
                if ($PSBoundParameters.Debug) { $offboardParams.Debug = $true }

                Invoke-AADUserOffboarding @offboardParams
            }
            "BulkOnboard" {
                Write-ScriptLog -Message "Bulk Onboarding Expected CSV Headers:" -Level INFO -LogFilePath $LogFilePath
                Write-ScriptLog -Message "- UserPrincipalName (string, mandatory), DisplayName (string, mandatory), Password (string, optional), ForceChangePasswordNextLogin (True/False, optional), InitialGroups (comma-separated, optional), LicenseSKUs (comma-separated, optional), Department (string, optional), JobTitle (string, optional), ManagerUPN (string, optional)" -Level INFO -LogFilePath $LogFilePath
                Invoke-BulkAADUserOnboarding -CSVPath $CSVPath -LogFilePath $LogFilePath -Verbose:$PSBoundParameters.Verbose -Debug:$PSBoundParameters.Debug -WhatIf:$PSBoundParameters.WhatIf
            }
            "BulkUpdate" {
                Write-ScriptLog -Message "Bulk Update Expected CSV Headers: UserPrincipalName OR ObjectId (mandatory identifier), plus any valid parameters for Update-AADUser (e.g., DisplayName, Department, ManagerUPN, GroupsToAdd, LicensesToAssign etc.)" -Level INFO -LogFilePath $LogFilePath
                Invoke-BulkAADUserUpdate -CSVPath $CSVPath -LogFilePath $LogFilePath -Verbose:$PSBoundParameters.Verbose -Debug:$PSBoundParameters.Debug -WhatIf:$PSBoundParameters.WhatIf
            }
            "BulkOffboard" {
                Write-ScriptLog -Message "Bulk Offboarding Expected CSV Headers: UserPrincipalName OR ObjectId (mandatory identifier), Action (Disable/Delete, optional), RevokeSignInSessions (True/False, optional), RemoveAllLicenses (True/False, optional), RemoveFromAllGroups (True/False, optional)" -Level INFO -LogFilePath $LogFilePath
                Invoke-BulkAADUserOffboarding -CSVPath $CSVPath -LogFilePath $LogFilePath -Verbose:$PSBoundParameters.Verbose -Debug:$PSBoundParameters.Debug -WhatIf:$PSBoundParameters.WhatIf
            }
            default {
                Write-ScriptLog -Message "Invalid -LifecycleAction specified: $LifecycleAction" -Level ERROR -LogFilePath $LogFilePath
            }
        }

        Write-ScriptLog -Message "Manage-AzureADUserLifecycle script finished action: $LifecycleAction." -Level INFO -LogFilePath $LogFilePath

    } catch {
        Write-ScriptLog -Message "An unhandled error occurred in the main script logic: $($_.Exception.Message)" -Level ERROR -LogFilePath $LogFilePath
        if ($_.Exception.StackTrace) {
            Write-ScriptLog -Message "Stack Trace: $($_.Exception.StackTrace)" -Level DEBUG -LogFilePath $LogFilePath
        }
    } finally {
        if (Get-Transcript) {
            try {
                Stop-Transcript
                Write-Host "Transcript stopped. Path: $transcriptPath" # This won't go to Write-ScriptLog as it might be unavailable
            } catch {
                Write-Error "Failed to stop transcript. Error: $($_.Exception.Message)"
            }
        }
    }
}

# Execute Main function
Main
</script>
