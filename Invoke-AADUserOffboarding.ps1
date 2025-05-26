# Ensure Write-ScriptLog is available in the session if this script is run standalone
# For bulk operations, Write-ScriptLog is assumed to be loaded by the bulk script.

function Invoke-AADUserOffboarding {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = "UPN", HelpMessage="User Principal Name of the user to offboard.")]
        [string]$UserPrincipalName,

        [Parameter(Mandatory = $true, ParameterSetName = "ObjectId", HelpMessage="Object ID (GUID) of the user to offboard.")]
        [string]$ObjectId,

        [Parameter(Mandatory = $false, HelpMessage="Action to perform: 'Disable' or 'Delete'. Defaults to 'Disable'.")]
        [ValidateSet("Disable", "Delete")]
        [string]$Action = "Disable",

        [Parameter(Mandatory = $false, HelpMessage="Revoke all sign-in sessions for the user. Defaults to $true.")]
        [bool]$RevokeSignInSessions = $true,

        [Parameter(Mandatory = $false, HelpMessage="Remove all assigned licenses from the user. Defaults to $true.")]
        [bool]$RemoveAllLicenses = $true,

        [Parameter(Mandatory = $false, HelpMessage="Remove the user from all group memberships. Defaults to $false. Use with caution.")]
        [bool]$RemoveFromAllGroups = $false,

        [Parameter(Mandatory = $true, HelpMessage = "Path to the log file. This is typically passed by the calling bulk script.")]
        [string]$LogFilePath
    )

    function _EnsureWriteScriptLogOffboardUser { 
        if (-not (Get-Command Write-ScriptLog -ErrorAction SilentlyContinue)) {
            Write-Warning "Write-ScriptLog function not found in Invoke-AADUserOffboarding. Using basic console output."
            New-Alias -Name Write-ScriptLog -Value Write-Host -Scope Script -Force 
        }
    }
    _EnsureWriteScriptLogOffboardUser

    begin {
        Write-ScriptLog -Message "Invoke-AADUserOffboarding: Starting." -Level VERBOSE -LogFilePath $LogFilePath
        try {
            $azContext = Get-AzContext -ErrorAction Stop
            Write-ScriptLog -Message "Invoke-AADUserOffboarding: Azure context retrieved. Tenant: $($azContext.Tenant.Id)." -Level VERBOSE -LogFilePath $LogFilePath
        }
        catch {
            Write-ScriptLog -Message "Invoke-AADUserOffboarding: Not connected to Azure AD. Please run Connect-AzAccount first." -Level ERROR -LogFilePath $LogFilePath
            return "Error: Not connected to Azure AD." # Return error status
        }

        $script:userIdentifierForOffboard = if ($PSCmdlet.ParameterSetName -eq "ObjectId") { $ObjectId } else { $UserPrincipalName } # script-scoped
        Write-ScriptLog -Message "Invoke-AADUserOffboarding: Target user identifier set to '$($script:userIdentifierForOffboard)' (Action: $Action)." -Level DEBUG -LogFilePath $LogFilePath
    }

    process {
        $effectiveAction = $Action # Action defaults to "Disable" via param block
        $processMessageDetail = "Offboard user '$($script:userIdentifierForOffboard)' with action '$effectiveAction'."
        $additionalOpsDesc = @()
        if ($RevokeSignInSessions) { $additionalOpsDesc += "Revoke Sessions" }
        if ($RemoveAllLicenses) { $additionalOpsDesc += "Remove Licenses" }
        if ($RemoveFromAllGroups) { $additionalOpsDesc += "Remove from All Groups" }
        if ($additionalOpsDesc.Count -gt 0) { $processMessageDetail += " Additional steps: $($additionalOpsDesc -join ', ')." }


        if (-not ($PSCmdlet.ShouldProcess($script:userIdentifierForOffboard, $processMessageDetail))) {
            Write-ScriptLog -Message "Invoke-AADUserOffboarding: Offboarding action '$effectiveAction' for user '$($script:userIdentifierForOffboard)' skipped due to -WhatIf or user cancellation." -Level WARN -LogFilePath $LogFilePath
            return "Skipped: Offboarding for '$($script:userIdentifierForOffboard)' cancelled by user or -WhatIf."
        }

        try {
            $user = $null
            Write-ScriptLog -Message "Invoke-AADUserOffboarding: Attempting to retrieve user '$($script:userIdentifierForOffboard)'." -Level INFO -LogFilePath $LogFilePath
            if ($PSCmdlet.ParameterSetName -eq "ObjectId") {
                 $user = Get-AzADUser -ObjectId $script:userIdentifierForOffboard -ErrorAction Stop
            } else { # UPN
                 $user = Get-AzADUser -UserPrincipalName $script:userIdentifierForOffboard -ErrorAction Stop
            }
            if (-not $user) { # Should be caught by ErrorAction Stop
                Write-ScriptLog -Message "Invoke-AADUserOffboarding: User '$($script:userIdentifierForOffboard)' not found." -Level ERROR -LogFilePath $LogFilePath
                return "Error: User '$($script:userIdentifierForOffboard)' not found."
            }
            Write-ScriptLog -Message "Invoke-AADUserOffboarding: User '$($user.UserPrincipalName)' (ID: $($user.Id)) found. Proceeding with offboarding (Action: $effectiveAction)." -Level INFO -LogFilePath $LogFilePath

            # 1. Disable Account (always done if account is enabled, even if Action is Delete)
            if ($user.AccountEnabled) {
                if ($PSCmdlet.ShouldProcess($user.UserPrincipalName, "Disable account as part of offboarding (Action: $effectiveAction)")) {
                    Write-ScriptLog -Message "Invoke-AADUserOffboarding: Disabling account for '$($user.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath
                    Set-AzADUser -ObjectId $user.Id -AccountEnabled $false -ErrorAction Stop
                    Write-ScriptLog -Message "Invoke-AADUserOffboarding: Account for '$($user.UserPrincipalName)' disabled." -Level INFO -LogFilePath $LogFilePath
                } else { Write-ScriptLog -Message "Invoke-AADUserOffboarding: Account disable skipped for '$($user.UserPrincipalName)' due to -WhatIf/cancellation." -Level WARN -LogFilePath $LogFilePath }
            } else {
                Write-ScriptLog -Message "Invoke-AADUserOffboarding: Account for '$($user.UserPrincipalName)' is already disabled." -Level INFO -LogFilePath $LogFilePath
            }

            # 2. Revoke Sign-in Sessions
            if ($RevokeSignInSessions) {
                if ($PSCmdlet.ShouldProcess($user.UserPrincipalName, "Revoke all sign-in sessions")) {
                    Write-ScriptLog -Message "Invoke-AADUserOffboarding: Revoking sign-in sessions for '$($user.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath
                    try {
                        Revoke-AzADUserSignInSession -ObjectId $user.Id -ErrorAction Stop
                        Write-ScriptLog -Message "Invoke-AADUserOffboarding: Sign-in sessions revoked for '$($user.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath
                    } catch { Write-ScriptLog -Message "Invoke-AADUserOffboarding: Failed to revoke sign-in sessions for '$($user.UserPrincipalName)'. Error: $($_.Exception.Message). Continuing." -Level WARN -LogFilePath $LogFilePath }
                } else { Write-ScriptLog -Message "Invoke-AADUserOffboarding: Session revocation skipped for '$($user.UserPrincipalName)' due to -WhatIf/cancellation." -Level WARN -LogFilePath $LogFilePath }
            }

            # 3. Remove All Licenses
            if ($RemoveAllLicenses) {
                if ($PSCmdlet.ShouldProcess($user.UserPrincipalName, "Remove all assigned Azure AD licenses")) {
                    Write-ScriptLog -Message "Invoke-AADUserOffboarding: Removing all licenses for '$($user.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath
                    try {
                        $currentLicenses = Get-AzADUserAssignedLicense -ObjectId $user.Id -ErrorAction SilentlyContinue # Check without stopping if none
                        if ($currentLicenses.Count -gt 0) {
                            Set-AzADUser -ObjectId $user.Id -AssignedLicenses @() -ErrorAction Stop 
                            Write-ScriptLog -Message "Invoke-AADUserOffboarding: All licenses removed for '$($user.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath
                        } else { Write-ScriptLog -Message "Invoke-AADUserOffboarding: User '$($user.UserPrincipalName)' has no licenses assigned." -Level INFO -LogFilePath $LogFilePath }
                    } catch { Write-ScriptLog -Message "Invoke-AADUserOffboarding: Failed to remove licenses for '$($user.UserPrincipalName)'. Error: $($_.Exception.Message). Continuing." -Level WARN -LogFilePath $LogFilePath }
                } else { Write-ScriptLog -Message "Invoke-AADUserOffboarding: License removal skipped for '$($user.UserPrincipalName)' due to -WhatIf/cancellation." -Level WARN -LogFilePath $LogFilePath }
            }

            # 4. Remove from All Groups
            if ($RemoveFromAllGroups) {
                if ($PSCmdlet.ShouldProcess($user.UserPrincipalName, "Remove user from ALL Azure AD groups (use with caution)")) {
                    Write-ScriptLog -Message "Invoke-AADUserOffboarding: Attempting to remove user '$($user.UserPrincipalName)' from all groups." -Level INFO -LogFilePath $LogFilePath
                    try {
                        $groupMemberships = Get-AzADUserMembership -ObjectId $user.Id -ErrorAction Stop
                        if ($groupMemberships.Count -gt 0) {
                            Write-ScriptLog -Message "Invoke-AADUserOffboarding: User '$($user.UserPrincipalName)' is member of $($groupMemberships.Count) groups. Starting removal." -Level INFO -LogFilePath $LogFilePath
                            foreach ($groupMembership in $groupMemberships) {
                                $groupId = $groupMembership.Id
                                $groupDisplayName = if ($groupMembership.AdditionalProperties.ContainsKey("displayName")) { $groupMembership.AdditionalProperties["displayName"] } else { "(displayName N/A)" }
                                $groupNameForLog = "'$groupDisplayName' (ID: $groupId)"
                                if ($PSCmdlet.ShouldProcess("Group: $groupNameForLog", "Remove user '$($user.UserPrincipalName)' from this specific group")) {
                                    try {
                                        Remove-AzADGroupMember -TargetGroupObjectId $groupId -MemberObjectId $user.Id -ErrorAction Stop
                                        Write-ScriptLog -Message "Invoke-AADUserOffboarding: Removed user '$($user.UserPrincipalName)' from group $groupNameForLog." -Level INFO -LogFilePath $LogFilePath
                                    } catch { Write-ScriptLog -Message "Invoke-AADUserOffboarding: Failed to remove user '$($user.UserPrincipalName)' from group $groupNameForLog. Error: $($_.Exception.Message). Continuing." -Level WARN -LogFilePath $LogFilePath }
                                } else { Write-ScriptLog -Message "Invoke-AADUserOffboarding: Removal from group $groupNameForLog skipped for '$($user.UserPrincipalName)' due to -WhatIf/cancellation." -Level WARN -LogFilePath $LogFilePath }
                            }
                        } else { Write-ScriptLog -Message "Invoke-AADUserOffboarding: User '$($user.UserPrincipalName)' is not a member of any groups." -Level INFO -LogFilePath $LogFilePath }
                    } catch { Write-ScriptLog -Message "Invoke-AADUserOffboarding: Error retrieving/processing group memberships for '$($user.UserPrincipalName)'. Error: $($_.Exception.Message). Continuing." -Level WARN -LogFilePath $LogFilePath }
                } else { Write-ScriptLog -Message "Invoke-AADUserOffboarding: Removal from all groups skipped for '$($user.UserPrincipalName)' due to -WhatIf/cancellation." -Level WARN -LogFilePath $LogFilePath }
            }

            if ($effectiveAction -eq "Delete") {
                Write-ScriptLog -Message "Invoke-AADUserOffboarding: Proceeding with SOFT DELETION of user '$($user.UserPrincipalName)' (ID: $($user.Id))." -Level WARN -LogFilePath $LogFilePath # WARN because it's destructive
                if ($PSCmdlet.ShouldProcess($user.UserPrincipalName, "SOFT-DELETE user account. This user can be recovered from 'Deleted users' for up to 30 days.")) {
                    try {
                        Remove-AzADUser -ObjectId $user.Id -ErrorAction Stop
                        Write-ScriptLog -Message "Invoke-AADUserOffboarding: User '$($user.UserPrincipalName)' (ID: $($user.Id)) has been soft-deleted." -Level INFO -LogFilePath $LogFilePath
                        return "Success: User '$($user.UserPrincipalName)' soft-deleted."
                    } catch {
                        Write-ScriptLog -Message "Invoke-AADUserOffboarding: Failed to soft-delete user '$($user.UserPrincipalName)'. Error: $($_.Exception.Message)" -Level ERROR -LogFilePath $LogFilePath
                        return "Error: Failed to soft-delete user '$($user.UserPrincipalName)'. Message: $($_.Exception.Message)"
                    }
                } else {
                     Write-ScriptLog -Message "Invoke-AADUserOffboarding: Soft deletion of user '$($user.UserPrincipalName)' skipped due to -WhatIf or user cancellation during final delete confirmation." -Level WARN -LogFilePath $LogFilePath
                     return "Skipped: Soft deletion of user '$($user.UserPrincipalName)'."
                }
            }

            Write-ScriptLog -Message "Invoke-AADUserOffboarding: Action '$effectiveAction' completed for user '$($user.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath
            return "Success: User '$($user.UserPrincipalName)' offboarding action '$effectiveAction' completed."

        } catch [Microsoft.Azure.Commands.MicrosoftGraph.Cmdlets.Users.Models.MicrosoftGraphServiceException] {
             if ($_.Exception.Response.StatusCode -eq [System.Net.HttpStatusCode]::NotFound) {
                Write-ScriptLog -Message "Invoke-AADUserOffboarding: User '$($script:userIdentifierForOffboard)' not found or already deleted. $($_.Exception.Message)" -Level ERROR -LogFilePath $LogFilePath
                return "Error: User '$($script:userIdentifierForOffboard)' not found/already deleted. Graph API Error: $($_.Exception.Message)"
             } else {
                Write-ScriptLog -Message "Invoke-AADUserOffboarding: A Microsoft Graph service error occurred for user '$($script:userIdentifierForOffboard)': $($_.Exception.Message)" -Level ERROR -LogFilePath $LogFilePath
                if ($_.Exception.ErrorRecords.Count -gt 0) { Write-ScriptLog -Message "Invoke-AADUserOffboarding: Graph Error Record Details for '$($script:userIdentifierForOffboard)': $($_.Exception.ErrorRecords[0].ErrorDetails | ConvertTo-Json -Depth 3)" -Level DEBUG -LogFilePath $LogFilePath }
                return "Error: Failed to offboard user '$($script:userIdentifierForOffboard)' due to Graph API error. Message: $($_.Exception.Message)"
             }
        } catch {
            Write-ScriptLog -Message "Invoke-AADUserOffboarding: An unexpected error occurred for user '$($script:userIdentifierForOffboard)'. Error: $($_.Exception.Message)" -Level ERROR -LogFilePath $LogFilePath
            if ($_.Exception.StackTrace) { Write-ScriptLog -Message "Invoke-AADUserOffboarding: Stack Trace for '$($script:userIdentifierForOffboard)': $($_.Exception.StackTrace)" -Level DEBUG -LogFilePath $LogFilePath }
            return "Error: Failed to offboard user '$($script:userIdentifierForOffboard)' due to an unexpected error. Message: $($_.Exception.Message)"
        }
    } 
    
    end {
        Write-ScriptLog -Message "Invoke-AADUserOffboarding: Finished processing for user '$($script:userIdentifierForOffboard)'." -Level VERBOSE -LogFilePath $LogFilePath
    }
}

# Example Usage (comment out or remove before final script delivery):
# 
# Ensure Write-ScriptLog is loaded if testing standalone.
# function Write-ScriptLog { param([string]$Message, [string]$Level='INFO', [string]$LogFilePath) Write-Host "[$Level] $Message (Log: $LogFilePath)" }
#
# Connect-AzAccount -TenantId "your-tenant-id.onmicrosoft.com" 
# $logPath = ".\offboard_single_user.log"
# $testUserUPN = "someuser@yourtenant.onmicrosoft.com" 
#
# $result = Invoke-AADUserOffboarding -UserPrincipalName $testUserUPN -LogFilePath $logPath -Verbose #-WhatIf
# Write-Host "Offboarding Result: $result"
#
# $resultDel = Invoke-AADUserOffboarding -UserPrincipalName $testUserUPN -Action Delete -RemoveFromAllGroups $true -LogFilePath $logPath -Verbose #-WhatIf
# Write-Host "Offboarding Delete Result: $resultDel"
#
