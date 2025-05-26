# Ensure Write-ScriptLog is available in the session if this script is run standalone
# For bulk operations, Write-ScriptLog is assumed to be loaded by the bulk script.

function Invoke-AADUserOnboarding {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName,

        [Parameter(Mandatory = $true)]
        [string]$DisplayName,

        [Parameter(Mandatory = $false)]
        [System.Security.SecureString]$Password,

        [Parameter(Mandatory = $false)]
        [bool]$ForceChangePasswordNextLogin = $true,

        [Parameter(Mandatory = $false)]
        [string[]]$InitialGroups,

        [Parameter(Mandatory = $false)]
        [string[]]$LicenseSKUs,

        [Parameter(Mandatory = $false)]
        [string]$Department,

        [Parameter(Mandatory = $false)]
        [string]$JobTitle,

        [Parameter(Mandatory = $false)]
        [string]$ManagerUPN,

        [Parameter(Mandatory = $true, HelpMessage = "Path to the log file. This is typically passed by the calling bulk script.")]
        [string]$LogFilePath
    )

    # Helper to ensure Write-ScriptLog is available or provide a fallback
    # In a real module, this would be handled by module loading or a global function.
    function _EnsureWriteScriptLog {
        if (-not (Get-Command Write-ScriptLog -ErrorAction SilentlyContinue)) {
            # Minimal fallback if Write-ScriptLog is not loaded (e.g. when testing standalone)
            Write-Warning "Write-ScriptLog function not found. Using basic console output for this session."
            New-Alias -Name Write-ScriptLog -Value Write-Host -Scope Script -Force # Very basic fallback
        }
    }
    _EnsureWriteScriptLog # Call the helper

    begin {
        Write-ScriptLog -Message "Invoke-AADUserOnboarding: Starting for UPN '$UserPrincipalName'." -Level VERBOSE -LogFilePath $LogFilePath
        try {
            $azContext = Get-AzContext -ErrorAction Stop
            Write-ScriptLog -Message "Invoke-AADUserOnboarding: Already connected to Azure Tenant: $($azContext.Tenant.Id) for UPN '$UserPrincipalName'." -Level VERBOSE -LogFilePath $LogFilePath
        }
        catch {
            Write-ScriptLog -Message "Invoke-AADUserOnboarding: Not connected to Azure AD. Please run Connect-AzAccount. Aborting for UPN '$UserPrincipalName'." -Level ERROR -LogFilePath $LogFilePath
            # Return null or throw, depending on desired behavior for non-interactive use.
            # For this subtask, returning null is consistent with how bulk scripts might check.
            return $null 
        }

        if (-not $Password) {
            Write-ScriptLog -Message "Invoke-AADUserOnboarding: Password not provided for UPN '$UserPrincipalName', generating a random one." -Level INFO -LogFilePath $LogFilePath
            $Password = ConvertTo-SecureString (New-Guid).ToString() -AsPlainText -Force
        }
    }

    process {
        if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Create Azure AD User '$DisplayName'")) {
            try {
                $userParams = @{
                    UserPrincipalName = $UserPrincipalName
                    DisplayName       = $DisplayName
                    MailNickname      = $UserPrincipalName.Split('@')[0]
                    AccountEnabled    = $true # Users are enabled by default on creation
                }
                if ($PSBoundParameters.ContainsKey('Department')) { $userParams.Department = $Department }
                if ($PSBoundParameters.ContainsKey('JobTitle')) { $userParams.JobTitle = $JobTitle }
                
                $passwordProfile = @{
                    Password                             = $Password
                    ForceChangePasswordNextLogin         = $ForceChangePasswordNextLogin
                    # ForceChangePasswordNextLoginWithMfa  = $false # Not explicitly requested, keep it simple
                }
                $userParams.PasswordProfile = $passwordProfile

                $manager = $null
                if ($PSBoundParameters.ContainsKey('ManagerUPN') -and -not [string]::IsNullOrEmpty($ManagerUPN)) {
                    try {
                        Write-ScriptLog -Message "Invoke-AADUserOnboarding: Attempting to find manager with UPN '$ManagerUPN' for user '$UserPrincipalName'." -Level VERBOSE -LogFilePath $LogFilePath
                        $manager = Get-AzADUser -UserPrincipalName $ManagerUPN -ErrorAction Stop
                        if ($manager) {
                            Write-ScriptLog -Message "Invoke-AADUserOnboarding: Manager '$($manager.DisplayName)' found with ID: $($manager.Id) for user '$UserPrincipalName'." -Level INFO -LogFilePath $LogFilePath
                        }
                    } catch {
                        Write-ScriptLog -Message "Invoke-AADUserOnboarding: Manager with UPN '$ManagerUPN' not found for user '$UserPrincipalName'. Error: $($_.Exception.Message). Skipping manager assignment." -Level WARN -LogFilePath $LogFilePath
                        $manager = $null 
                    }
                }

                Write-ScriptLog -Message "Invoke-AADUserOnboarding: Creating user '$UserPrincipalName' with DisplayName '$DisplayName'." -Level INFO -LogFilePath $LogFilePath
                $newUser = New-AzADUser @userParams -ErrorAction Stop
                Write-ScriptLog -Message "Invoke-AADUserOnboarding: User '$($newUser.UserPrincipalName)' created successfully with ID: $($newUser.Id)." -Level INFO -LogFilePath $LogFilePath

                if ($manager) {
                    if ($PSCmdlet.ShouldProcess($newUser.UserPrincipalName, "Set Manager to $($manager.DisplayName)")) {
                        try {
                            Write-ScriptLog -Message "Invoke-AADUserOnboarding: Setting manager for '$($newUser.UserPrincipalName)' to '$($manager.UserPrincipalName)' (ID: $($manager.Id))." -Level INFO -LogFilePath $LogFilePath
                            Set-AzADUserManager -ObjectId $newUser.Id -ManagerId $manager.Id -ErrorAction Stop
                            Write-ScriptLog -Message "Invoke-AADUserOnboarding: Successfully set manager for '$($newUser.UserPrincipalName)' to '$($manager.DisplayName)'." -Level INFO -LogFilePath $LogFilePath
                        } catch {
                            Write-ScriptLog -Message "Invoke-AADUserOnboarding: Could not set manager for '$($newUser.UserPrincipalName)'. Error: $($_.Exception.Message)" -Level WARN -LogFilePath $LogFilePath
                        }
                    } else { Write-ScriptLog -Message "Invoke-AADUserOnboarding: Manager assignment for '$($newUser.UserPrincipalName)' skipped due to -WhatIf or user cancellation." -Level WARN -LogFilePath $LogFilePath}
                }

                if ($InitialGroups) {
                    Write-ScriptLog -Message "Invoke-AADUserOnboarding: Assigning user '$($newUser.UserPrincipalName)' to initial groups: $($InitialGroups -join ', ')." -Level INFO -LogFilePath $LogFilePath
                    foreach ($groupNameOrId in $InitialGroups) {
                        try {
                            Write-ScriptLog -Message "Invoke-AADUserOnboarding: Attempting to find group '$groupNameOrId' for user '$($newUser.UserPrincipalName)'." -Level VERBOSE -LogFilePath $LogFilePath
                            $group = Get-AzADGroup -Filter "Id eq '$groupNameOrId' or DisplayName eq '$groupNameOrId'" -ErrorAction Stop | Select-Object -First 1
                            
                            if ($group) {
                                Write-ScriptLog -Message "Invoke-AADUserOnboarding: Group '$($group.DisplayName)' (ID: $($group.Id)) found for user '$($newUser.UserPrincipalName)'." -Level VERBOSE -LogFilePath $LogFilePath
                                if ($PSCmdlet.ShouldProcess($group.DisplayName, "Add user '$($newUser.UserPrincipalName)' to group")) {
                                    Add-AzADGroupMember -TargetGroupObjectId $group.Id -MemberObjectId $newUser.Id -ErrorAction Stop
                                    Write-ScriptLog -Message "Invoke-AADUserOnboarding: Added user '$($newUser.UserPrincipalName)' to group '$($group.DisplayName)'." -Level INFO -LogFilePath $LogFilePath
                                } else { Write-ScriptLog -Message "Invoke-AADUserOnboarding: Adding user '$($newUser.UserPrincipalName)' to group '$($group.DisplayName)' skipped due to -WhatIf or user cancellation." -Level WARN -LogFilePath $LogFilePath }
                            } else {
                                Write-ScriptLog -Message "Invoke-AADUserOnboarding: Group '$groupNameOrId' not found for user '$($newUser.UserPrincipalName)'. Skipping." -Level WARN -LogFilePath $LogFilePath
                            }
                        } catch {
                            Write-ScriptLog -Message "Invoke-AADUserOnboarding: Failed to process group '$groupNameOrId' or add user '$($newUser.UserPrincipalName)' to it. Error: $($_.Exception.Message)" -Level WARN -LogFilePath $LogFilePath
                        }
                    }
                }

                if ($LicenseSKUs) {
                    Write-ScriptLog -Message "Invoke-AADUserOnboarding: Assigning licenses to user '$($newUser.UserPrincipalName)': $($LicenseSKUs -join ', ')." -Level INFO -LogFilePath $LogFilePath
                    $licensesToAssign = @()
                    foreach($skuId in $LicenseSKUs){
                        $licensesToAssign += @{SkuId = $skuId; DisabledPlans = @()} 
                    }
                    
                    if ($licensesToAssign.Count -gt 0) {
                         if ($PSCmdlet.ShouldProcess($newUser.UserPrincipalName, "Assign Licenses: $($LicenseSKUs -join ',')")) {
                            try {
                                Set-AzADUser -ObjectId $newUser.Id -AssignedLicenses $licensesToAssign -ErrorAction Stop
                                Write-ScriptLog -Message "Invoke-AADUserOnboarding: Successfully assigned licenses to user '$($newUser.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath
                            } catch {
                                Write-ScriptLog -Message "Invoke-AADUserOnboarding: Failed to assign licenses to user '$($newUser.UserPrincipalName)'. Error: $($_.Exception.Message)" -Level WARN -LogFilePath $LogFilePath
                                Write-ScriptLog -Message "Invoke-AADUserOnboarding: License SKUs attempted for '$($newUser.UserPrincipalName)': $($LicenseSKUs -join ', ')" -Level DEBUG -LogFilePath $LogFilePath
                            }
                        } else { Write-ScriptLog -Message "Invoke-AADUserOnboarding: License assignment for '$($newUser.UserPrincipalName)' skipped due to -WhatIf or user cancellation." -Level WARN -LogFilePath $LogFilePath }
                    }
                }
                
                Write-ScriptLog -Message "Invoke-AADUserOnboarding: Onboarding process complete for '$($newUser.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath
                return $newUser

            } catch {
                Write-ScriptLog -Message "Invoke-AADUserOnboarding: Failed to onboard user '$UserPrincipalName'. Error: $($_.Exception.Message)" -Level ERROR -LogFilePath $LogFilePath
                if ($_.Exception.StackTrace) { Write-ScriptLog -Message "Invoke-AADUserOnboarding: Stack Trace for UPN '$UserPrincipalName': $($_.Exception.StackTrace)" -Level DEBUG -LogFilePath $LogFilePath}
                return $null # Indicate failure to the caller
            }
        } else {
            Write-ScriptLog -Message "Invoke-AADUserOnboarding: User creation for '$UserPrincipalName' skipped due to -WhatIf parameter or user cancellation." -Level WARN -LogFilePath $LogFilePath
            return $null # Indicate skipped operation
        }
    }

    end {
        Write-ScriptLog -Message "Invoke-AADUserOnboarding: Finished processing for UPN '$UserPrincipalName'." -Level VERBOSE -LogFilePath $LogFilePath
    }
}

# Example Usage (comment out or remove before final script delivery):
# 
# Ensure Write-ScriptLog is loaded if testing standalone, or use the fallback.
# function Write-ScriptLog { param([string]$Message, [string]$Level='INFO', [string]$LogFilePath) Write-Host "[$Level] $Message (Log: $LogFilePath)" }
#
# Ensure you are connected to your Azure AD tenant first:
# Connect-AzAccount -TenantId "your-tenant-id.onmicrosoft.com"
# $logPath = ".\onboarding_single_user.log"
# 
# Example 1: Create a user with a specified password, department, job title, and assign to groups and licenses
# $securePassword = ConvertTo-SecureString "AStrongP@ssword123!" -AsPlainText -Force
# $user1Params = @{
#     UserPrincipalName = "johndoe@yourtenant.onmicrosoft.com"
#     DisplayName       = "John Doe"
#     Password          = $securePassword
#     Department        = "IT"
#     JobTitle          = "Cloud Engineer"
#     InitialGroups     = @("All Company Users", "IT Department Staff") 
#     LicenseSKUs       = @("xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx") 
#     ManagerUPN        = "manager.user@yourtenant.onmicrosoft.com"
#     Verbose           = $true
#     LogFilePath       = $logPath
#     #WhatIf            = $true 
# }
# Invoke-AADUserOnboarding @user1Params
# 
# Example 2: Create a user with an auto-generated password and minimal information
# $user2Params = @{
#     UserPrincipalName = "alicesmith@yourtenant.onmicrosoft.com"
#     DisplayName       = "Alice Smith"
#     Department        = "Human Resources"
#     JobTitle          = "HR Coordinator"
#     Verbose           = $true
#     LogFilePath       = $logPath
# }
# Invoke-AADUserOnboarding @user2Params
#
# To get available license SKU IDs:
# Get-AzSubscribedSku | Select-Object SkuPartNumber, SkuId
#
# To get Group Object IDs (if needed):
# Get-AzADGroup -Filter "DisplayName eq 'Your Group Name'" | Select-Object DisplayName, Id
#
# To get User Object ID (for manager assignment if UPN is not used):
# Get-AzADUser -UserPrincipalName "manager.user@yourtenant.onmicrosoft.com" | Select-Object DisplayName, Id
#
