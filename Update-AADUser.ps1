# Ensure Write-ScriptLog is available in the session if this script is run standalone
# For bulk operations, Write-ScriptLog is assumed to be loaded by the bulk script.

function Update-AADUser {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = "UPN", HelpMessage="User Principal Name of the user to update.")]
        [string]$UserPrincipalName,

        [Parameter(Mandatory = $true, ParameterSetName = "ObjectId", HelpMessage="Object ID (GUID) of the user to update.")]
        [string]$ObjectId,

        [Parameter(Mandatory = $false, HelpMessage="New display name for the user.")]
        [string]$DisplayName,

        [Parameter(Mandatory = $false, HelpMessage="New department for the user.")]
        [string]$Department,

        [Parameter(Mandatory = $false, HelpMessage="New job title for the user.")]
        [string]$JobTitle,

        [Parameter(Mandatory = $false, HelpMessage="New office location for the user.")]
        [string]$OfficeLocation,

        [Parameter(Mandatory = $false, HelpMessage="New street address for the user.")]
        [string]$StreetAddress,

        [Parameter(Mandatory = $false, HelpMessage="New city for the user.")]
        [string]$City,

        [Parameter(Mandatory = $false, HelpMessage="New state for the user.")]
        [string]$State,

        [Parameter(Mandatory = $false, HelpMessage="New postal code for the user.")]
        [string]$PostalCode,

        [Parameter(Mandatory = $false, HelpMessage="New country for the user.")]
        [string]$Country,

        [Parameter(Mandatory = $false, HelpMessage="New mobile phone number for the user.")]
        [string]$MobilePhone,

        [Parameter(Mandatory = $false, HelpMessage="New office phone number for the user. This will be the primary business phone.")]
        [string]$OfficePhone,

        [Parameter(Mandatory = $false, HelpMessage="UPN of the new manager. Provide an empty string to remove the current manager.")]
        [string]$ManagerUPN,

        [Parameter(Mandatory = $false, HelpMessage="Array of Group ObjectIDs or DisplayNames to add the user to.")]
        [string[]]$GroupsToAdd,

        [Parameter(Mandatory = $false, HelpMessage="Array of Group ObjectIDs or DisplayNames to remove the user from.")]
        [string[]]$GroupsToRemove,

        [Parameter(Mandatory = $false, HelpMessage="Array of License SKU IDs (GUIDs) to assign to the user.")]
        [string[]]$LicensesToAssign,

        [Parameter(Mandatory = $false, HelpMessage="Array of License SKU IDs (GUIDs) to remove from the user.")]
        [string[]]$LicensesToRemove,

        [Parameter(Mandatory = $true, HelpMessage = "Path to the log file. This is typically passed by the calling bulk script.")]
        [string]$LogFilePath
    )

    function _EnsureWriteScriptLogUpdateUser { 
        if (-not (Get-Command Write-ScriptLog -ErrorAction SilentlyContinue)) {
            Write-Warning "Write-ScriptLog function not found in Update-AADUser. Using basic console output."
            New-Alias -Name Write-ScriptLog -Value Write-Host -Scope Script -Force 
        }
    }
    _EnsureWriteScriptLogUpdateUser

    begin {
        Write-ScriptLog -Message "Update-AADUser: Starting." -Level VERBOSE -LogFilePath $LogFilePath
        try {
            $azContext = Get-AzContext -ErrorAction Stop
            Write-ScriptLog -Message "Update-AADUser: Azure context retrieved. Tenant: $($azContext.Tenant.Id)." -Level VERBOSE -LogFilePath $LogFilePath
        }
        catch {
            Write-ScriptLog -Message "Update-AADUser: Not connected to Azure AD. Please run Connect-AzAccount first." -Level ERROR -LogFilePath $LogFilePath
            return $null # Indicate failure/inability to proceed
        }

        $script:userIdentifierForUpdate = if ($PSCmdlet.ParameterSetName -eq "ObjectId") { $ObjectId } else { $UserPrincipalName } # script-scoped
        Write-ScriptLog -Message "Update-AADUser: Target user identifier set to '$($script:userIdentifierForUpdate)' (ParameterSet: $($PSCmdlet.ParameterSetName))." -Level DEBUG -LogFilePath $LogFilePath
    }

    process {
        $processMessage = "Update Azure AD User '$($script:userIdentifierForUpdate)'"
        if ($PSBoundParameters.Count -gt 3) { # 3 = Identifier, LogFilePath, and one actual change param
             $processMessage += " with specified attributes."
        } else {
            $processMessage += " - NO ACTION (no attributes specified for update beyond identifier and LogFilePath)."
            Write-ScriptLog -Message "Update-AADUser: No attributes specified for update for user '$($script:userIdentifierForUpdate)'. Nothing to do." -Level WARN -LogFilePath $LogFilePath
            # Consider returning the user object if found, or null/error if not. For now, just returning null as no update was attempted.
            return $null 
        }


        if (-not ($PSCmdlet.ShouldProcess($script:userIdentifierForUpdate, $processMessage))) {
            Write-ScriptLog -Message "Update-AADUser: Update operation for user '$($script:userIdentifierForUpdate)' skipped due to -WhatIf or user cancellation." -Level WARN -LogFilePath $LogFilePath
            return $null 
        }

        try {
            $user = $null
            Write-ScriptLog -Message "Update-AADUser: Attempting to retrieve user '$($script:userIdentifierForUpdate)'." -Level INFO -LogFilePath $LogFilePath
            if ($PSCmdlet.ParameterSetName -eq "ObjectId") {
                 $user = Get-AzADUser -ObjectId $script:userIdentifierForUpdate -ErrorAction Stop
            } else { # ParameterSetName is "UPN"
                 $user = Get-AzADUser -UserPrincipalName $script:userIdentifierForUpdate -ErrorAction Stop
            }
            if (-not $user) { # Should be caught by ErrorAction Stop
                Write-ScriptLog -Message "Update-AADUser: User '$($script:userIdentifierForUpdate)' not found." -Level ERROR -LogFilePath $LogFilePath
                return $null
            }
            Write-ScriptLog -Message "Update-AADUser: User '$($user.UserPrincipalName)' (ID: $($user.Id)) found." -Level INFO -LogFilePath $LogFilePath

            $setAzUserParams = @{}
            # Populate $setAzUserParams based on $PSBoundParameters, excluding specific ones
            $excludedParams = @('UserPrincipalName', 'ObjectId', 'LogFilePath', 'WhatIf', 'Confirm', 'Verbose', 'Debug', 'ErrorAction', 'ErrorVariable', 'OutVariable', 'OutBuffer', 'PipelineVariable', 'WarningAction', 'WarningVariable', # Common params
                                'ManagerUPN', 'GroupsToAdd', 'GroupsToRemove', 'LicensesToAssign', 'LicensesToRemove') # Handled separately

            foreach($paramName in $PSBoundParameters.Keys){
                if($excludedParams -notcontains $paramName -and $PSBoundParameters[$paramName] -ne $null){
                     # Specific handling for BusinessPhones if it's different from a simple string.
                     # Update-AADUser current spec implies OfficePhone becomes BusinessPhones = @($OfficePhone)
                    if($paramName -eq 'OfficePhone') {
                        $setAzUserParams.BusinessPhones = @($PSBoundParameters[$paramName])
                    } else {
                        $setAzUserParams[$paramName] = $PSBoundParameters[$paramName]
                    }
                }
            }
            
            if ($setAzUserParams.Count -gt 0) {
                Write-ScriptLog -Message "Update-AADUser: Updating user attributes for '$($user.UserPrincipalName)': $($setAzUserParams.Keys -join ', ')." -Level INFO -LogFilePath $LogFilePath
                Set-AzADUser -ObjectId $user.Id @setAzUserParams -ErrorAction Stop
                Write-ScriptLog -Message "Update-AADUser: Successfully updated standard attributes for user '$($user.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath
            }

            if ($PSBoundParameters.ContainsKey('ManagerUPN')) {
                if (-not [string]::IsNullOrEmpty($ManagerUPN)) {
                    Write-ScriptLog -Message "Update-AADUser: Attempting to set manager for '$($user.UserPrincipalName)' to '$ManagerUPN'." -Level INFO -LogFilePath $LogFilePath
                    try {
                        $newManager = Get-AzADUser -UserPrincipalName $ManagerUPN -ErrorAction Stop
                        if ($newManager) {
                            Set-AzADUserManager -ObjectId $user.Id -ManagerId $newManager.Id -ErrorAction Stop
                            Write-ScriptLog -Message "Update-AADUser: Successfully set manager for '$($user.UserPrincipalName)' to '$($newManager.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath
                        }
                    } catch { Write-ScriptLog -Message "Update-AADUser: Failed to find or set manager '$ManagerUPN' for '$($user.UserPrincipalName)'. Error: $($_.Exception.Message)" -Level WARN -LogFilePath $LogFilePath }
                } else { 
                    Write-ScriptLog -Message "Update-AADUser: ManagerUPN parameter provided as empty string. Attempting to remove current manager for '$($user.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath
                    try {
                       Set-AzADUserManager -ObjectId $user.Id -ManagerId $null -ErrorAction Stop 
                       Write-ScriptLog -Message "Update-AADUser: Successfully removed manager for user '$($user.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath
                    } catch { Write-ScriptLog -Message "Update-AADUser: Failed to remove manager for '$($user.UserPrincipalName)'. Error: $($_.Exception.Message)" -Level WARN -LogFilePath $LogFilePath }
                }
            }

            if ($PSBoundParameters.ContainsKey('GroupsToAdd')) {
                Write-ScriptLog -Message "Update-AADUser: Adding user '$($user.UserPrincipalName)' to groups: $($GroupsToAdd -join ', ')." -Level INFO -LogFilePath $LogFilePath
                foreach ($groupNameOrId in $GroupsToAdd) {
                    try {
                        $group = Get-AzADGroup -Filter "DisplayName eq '$groupNameOrId' or Id eq '$groupNameOrId'" -ErrorAction Stop | Select-Object -First 1
                        if ($group) {
                            Add-AzADGroupMember -TargetGroupObjectId $group.Id -MemberObjectId $user.Id -ErrorAction Stop
                            Write-ScriptLog -Message "Update-AADUser: Added user '$($user.UserPrincipalName)' to group '$($group.DisplayName)' (ID: $($group.Id))." -Level INFO -LogFilePath $LogFilePath
                        } else { Write-ScriptLog -Message "Update-AADUser: Group '$groupNameOrId' not found for adding to user '$($user.UserPrincipalName)'. Skipping." -Level WARN -LogFilePath $LogFilePath }
                    } catch { Write-ScriptLog -Message "Update-AADUser: Failed to add user '$($user.UserPrincipalName)' to group '$groupNameOrId'. Error: $($_.Exception.Message)" -Level WARN -LogFilePath $LogFilePath }
                }
            }
            if ($PSBoundParameters.ContainsKey('GroupsToRemove')) {
                Write-ScriptLog -Message "Update-AADUser: Removing user '$($user.UserPrincipalName)' from groups: $($GroupsToRemove -join ', ')." -Level INFO -LogFilePath $LogFilePath
                foreach ($groupNameOrId in $GroupsToRemove) {
                    try {
                        $group = Get-AzADGroup -Filter "DisplayName eq '$groupNameOrId' or Id eq '$groupNameOrId'" -ErrorAction Stop | Select-Object -First 1
                        if ($group) {
                            Remove-AzADGroupMember -TargetGroupObjectId $group.Id -MemberObjectId $user.Id -ErrorAction Stop
                            Write-ScriptLog -Message "Update-AADUser: Removed user '$($user.UserPrincipalName)' from group '$($group.DisplayName)' (ID: $($group.Id))." -Level INFO -LogFilePath $LogFilePath
                        } else { Write-ScriptLog -Message "Update-AADUser: Group '$groupNameOrId' not found for removing from user '$($user.UserPrincipalName)'. Skipping." -Level WARN -LogFilePath $LogFilePath }
                    } catch { Write-ScriptLog -Message "Update-AADUser: Failed to remove user '$($user.UserPrincipalName)' from group '$groupNameOrId'. Error: $($_.Exception.Message)" -Level WARN -LogFilePath $LogFilePath }
                }
            }

            if ($PSBoundParameters.ContainsKey('LicensesToAssign') -or $PSBoundParameters.ContainsKey('LicensesToRemove')) {
                Write-ScriptLog -Message "Update-AADUser: Updating licenses for user '$($user.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath
                try {
                    $currentUserLicenses = Get-AzADUserAssignedLicense -ObjectId $user.Id -ErrorAction Stop
                    $currentSkuIds = $currentUserLicenses | ForEach-Object { $_.SkuId.ToString() }
                    Write-ScriptLog -Message "Update-AADUser: User '$($user.UserPrincipalName)' currently has SKU IDs: $($currentSkuIds -join ', ')." -Level DEBUG -LogFilePath $LogFilePath

                    $finalSkuIds = [System.Collections.Generic.List[string]]::new($currentSkuIds)
                    if ($PSBoundParameters.ContainsKey('LicensesToAssign')) { 
                        foreach ($sku in $LicensesToAssign) { 
                            if (-not ($finalSkuIds.Contains($sku))) { $finalSkuIds.Add($sku) }
                        }
                    }
                    if ($PSBoundParameters.ContainsKey('LicensesToRemove')) { 
                        foreach ($sku in $LicensesToRemove) { 
                            if ($finalSkuIds.Contains($sku)) { $finalSkuIds.Remove($sku) }
                        }
                    }
                    
                    $licensesForUpdate = @()
                    foreach($skuId in $finalSkuIds){
                        $licensesForUpdate += @{SkuId = $skuId; DisabledPlans = @()} 
                    }

                    Set-AzADUser -ObjectId $user.Id -AssignedLicenses $licensesForUpdate -ErrorAction Stop
                    Write-ScriptLog -Message "Update-AADUser: Successfully updated licenses for user '$($user.UserPrincipalName)'. Final SKUs: $($finalSkuIds -join ', ')." -Level INFO -LogFilePath $LogFilePath
                } catch {
                    Write-ScriptLog -Message "Update-AADUser: Failed to update licenses for user '$($user.UserPrincipalName)'. Error: $($_.Exception.Message)" -Level WARN -LogFilePath $LogFilePath
                }
            }

            Write-ScriptLog -Message "Update-AADUser: Update process complete for '$($user.UserPrincipalName)'." -Level INFO -LogFilePath $LogFilePath
            return Get-AzADUser -ObjectId $user.Id # Return the updated user object
            
        } catch [Microsoft.Azure.Commands.MicrosoftGraph.Cmdlets.Users.Models.MicrosoftGraphServiceException] {
             if ($_.Exception.Response.StatusCode -eq [System.Net.HttpStatusCode]::NotFound) { 
                Write-ScriptLog -Message "Update-AADUser: User '$($script:userIdentifierForUpdate)' not found. $($_.Exception.Message)" -Level ERROR -LogFilePath $LogFilePath
             } else {
                Write-ScriptLog -Message "Update-AADUser: A Microsoft Graph service error occurred while processing user '$($script:userIdentifierForUpdate)': $($_.Exception.Message)" -Level ERROR -LogFilePath $LogFilePath
             }
             if ($_.Exception.ErrorRecords.Count -gt 0) { 
                Write-ScriptLog -Message "Update-AADUser: Graph Error Record Details for '$($script:userIdentifierForUpdate)': $($_.Exception.ErrorRecords[0].ErrorDetails | ConvertTo-Json -Depth 3)" -Level DEBUG -LogFilePath $LogFilePath
             }
             return $null
        } catch {
            Write-ScriptLog -Message "Update-AADUser: Failed to update user '$($script:userIdentifierForUpdate)'. Error: $($_.Exception.Message)" -Level ERROR -LogFilePath $LogFilePath
            if ($_.Exception.StackTrace) { Write-ScriptLog -Message "Update-AADUser: Stack Trace for '$($script:userIdentifierForUpdate)': $($_.Exception.StackTrace)" -Level DEBUG -LogFilePath $LogFilePath}
            return $null
        }
    }
    end {
        Write-ScriptLog -Message "Update-AADUser: Finished processing for user '$($script:userIdentifierForUpdate)'." -Level VERBOSE -LogFilePath $LogFilePath
    }
}

# Example Usage (comment out or remove before final script delivery):
# 
# Ensure Write-ScriptLog is loaded if testing standalone.
# function Write-ScriptLog { param([string]$Message, [string]$Level='INFO', [string]$LogFilePath) Write-Host "[$Level] $Message (Log: $LogFilePath)" }
#
# Connect-AzAccount -TenantId "your-tenant-id.onmicrosoft.com" 
# $logPath = ".\update_single_user.log"
# $targetUserUPN = "adelev@yourtenant.onmicrosoft.com" 
#
# Update-AADUser -UserPrincipalName $targetUserUPN -Department "Marketing" -JobTitle "Marketing Specialist" -LogFilePath $logPath -Verbose #-WhatIf
#
# Update-AADUser -UserPrincipalName $targetUserUPN -ManagerUPN "" -LogFilePath $logPath -Verbose # Remove manager
#
# Update-AADUser -UserPrincipalName $targetUserUPN -GroupsToAdd "Sales Team" -LogFilePath $logPath -Verbose
#
# Update-AADUser -UserPrincipalName $targetUserUPN -LicensesToAssign "sku-guid-here" -LogFilePath $logPath -Verbose
#
