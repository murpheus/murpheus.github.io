#Requires -Modules Az.Accounts, Az.Resources 
# Make sure Update-AADUser.ps1 is loaded in the session.
# Example: . .\Update-AADUser.ps1
# Ensure Write-ScriptLog is available in the session if this script is run standalone
# For bulk operations, Write-ScriptLog is assumed to be loaded by the calling (e.g. bulk onboarding) script.

function Invoke-BulkAADUserUpdate {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true, HelpMessage = "Path to the CSV file for bulk user updates.")]
        [string]$CSVPath,

        [Parameter(Mandatory = $true, HelpMessage = "Path to the log file.")]
        [string]$LogFilePath
    )

    # Helper to ensure Write-ScriptLog is available or provide a fallback
    function _EnsureWriteScriptLogBulkUpdate { # Renamed to avoid conflict if scripts are dot-sourced together
        if (-not (Get-Command Write-ScriptLog -ErrorAction SilentlyContinue)) {
            Write-Warning "Write-ScriptLog function not found. Using basic console output for this session (Invoke-BulkAADUserUpdate)."
            New-Alias -Name Write-ScriptLog -Value Write-Host -Scope Script -Force 
        }
    }
    _EnsureWriteScriptLogBulkUpdate 

    begin {
        Write-ScriptLog -Message "Invoke-BulkAADUserUpdate starting." -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "Expected CSV Headers (at least one identifier - UserPrincipalName or ObjectId - is mandatory):" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- UserPrincipalName (string, identifier)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- ObjectId (string, identifier)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- DisplayName (string)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- Department (string)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- JobTitle (string)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- OfficeLocation (string)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- StreetAddress (string)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- City (string)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- State (string)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- PostalCode (string)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- Country (string)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- MobilePhone (string)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- OfficePhone (string)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- ManagerUPN (string, use empty string in CSV cell to remove manager)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- GroupsToAdd (comma-separated string of Group DisplayNames or ObjectIDs)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- GroupsToRemove (comma-separated string of Group DisplayNames or ObjectIDs)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- LicensesToAssign (comma-separated string of License SKU GUIDs)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- LicensesToRemove (comma-separated string of License SKU GUIDs)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message " (Other valid parameters for Update-AADUser can also be used as headers)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "------------------------------------" -Level INFO -LogFilePath $LogFilePath

        if (-not (Test-Path -Path $CSVPath -PathType Leaf)) {
            Write-ScriptLog -Message "CSV file not found at path: $CSVPath" -Level ERROR -LogFilePath $LogFilePath
            return 
        }

        if (-not (Get-Command Update-AADUser -ErrorAction SilentlyContinue)) {
            Write-ScriptLog -Message "Critical dependency missing: Function 'Update-AADUser' is not loaded. Please load it first (e.g., '. .\Update-AADUser.ps1')." -Level ERROR -LogFilePath $LogFilePath
            return
        }
        
        $script:updateFunctionParams = (Get-Command Update-AADUser).Parameters.Keys | ForEach-Object { $_.ToLowerInvariant() } # script-scoped for process block

        $script:summary = @{ # script-scoped
            SuccessCount  = 0
            FailureCount  = 0
            FailedEntries = [System.Collections.Generic.List[hashtable]]::new()
            ProcessedRows = 0
            TotalRows     = 0
        }
    }

    process {
        try {
            $csvData = Import-Csv -Path $CSVPath -ErrorAction Stop
            $script:summary.TotalRows = $csvData.Count
            $currentRowNumber = 0
            Write-ScriptLog -Message "Loaded $($script:summary.TotalRows) rows from CSV '$CSVPath'." -Level INFO -LogFilePath $LogFilePath


            foreach ($row in $csvData) {
                $currentRowNumber++
                $script:summary.ProcessedRows = $currentRowNumber
                $userIdentifierForLog = "<MISSING_IDENTIFIER>" 

                try {
                    $updateParams = @{ LogFilePath = $LogFilePath } # Pass LogFilePath
                    $identifierProvided = $false

                    if ($row.PSObject.Properties['UserPrincipalName'] -and -not [string]::IsNullOrWhiteSpace($row.UserPrincipalName)) {
                        $updateParams.UserPrincipalName = $row.UserPrincipalName.Trim()
                        $userIdentifierForLog = $updateParams.UserPrincipalName
                        $identifierProvided = $true
                    } elseif ($row.PSObject.Properties['ObjectId'] -and -not [string]::IsNullOrWhiteSpace($row.ObjectId)) {
                        $updateParams.ObjectId = $row.ObjectId.Trim()
                        $userIdentifierForLog = $updateParams.ObjectId
                        $identifierProvided = $true
                    }

                    if (-not $identifierProvided) {
                        throw "Mandatory identifier (UserPrincipalName or ObjectId) is missing or empty."
                    }

                    Write-ScriptLog -Message "Processing user $currentRowNumber of $($script:summary.TotalRows): '$userIdentifierForLog'" -Level INFO -LogFilePath $LogFilePath

                    if ($PSCmdlet.ShouldProcess($userIdentifierForLog, "Update Azure AD User from CSV row $currentRowNumber")) {
                        foreach ($property in $row.PSObject.Properties) {
                            $paramName = $property.Name
                            $paramValue = $property.Value

                            if ($paramName -eq 'UserPrincipalName' -or $paramName -eq 'ObjectId') { continue }

                            if ($script:updateFunctionParams -contains $paramName.ToLowerInvariant() -and -not [string]::IsNullOrWhiteSpace($paramValue)) {
                                switch ($paramName.ToLowerInvariant()) {
                                    'groupstoadd'; 'groupstoremove'; 'licensestoassign'; 'licensestoremove' {
                                        $updateParams[$paramName] = $paramValue.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                                    }
                                    default {
                                        if ($paramName.ToLowerInvariant() -eq 'managerupn') {
                                            $updateParams[$paramName] = $paramValue.Trim() 
                                        } else {
                                            $updateParams[$paramName] = $paramValue.Trim()
                                        }
                                    }
                                }
                            } elseif ($script:updateFunctionParams -contains $paramName.ToLowerInvariant() -and $paramName.ToLowerInvariant() -eq 'managerupn' -and $row.PSObject.Properties['ManagerUPN'] -and [string]::IsNullOrEmpty($paramValue)) {
                                $updateParams[$paramName] = "" # Explicitly empty ManagerUPN
                            }
                        }
                        
                        if ($updateParams.Count -le 2) { # Only identifier and LogFilePath are present
                           throw "No update attributes found for user '$userIdentifierForLog'. At least one attribute to update must be provided in the CSV."
                        }

                        if ($PSBoundParameters.ContainsKey('Verbose') && $PSBoundParameters['Verbose']) { $updateParams.Verbose = $true }
                        if ($PSBoundParameters.ContainsKey('Debug') && $PSBoundParameters['Debug']) { $updateParams.Debug = $true }
                        
                        $result = Update-AADUser @updateParams
                        
                        if ($result -and $result.Id) { # Assuming Update-AADUser returns user object on success
                            Write-ScriptLog -Message "Successfully updated user '$userIdentifierForLog'." -Level INFO -LogFilePath $LogFilePath
                            $script:summary.SuccessCount++
                        } else {
                            throw "Update failed for user '$userIdentifierForLog'. Result from single-user function was not a success object or was null: $result"
                        }
                    } else {
                        Write-ScriptLog -Message "Skipped update for user '$userIdentifierForLog' (row $currentRowNumber) due to -WhatIf or user cancellation." -Level WARN -LogFilePath $LogFilePath
                    }

                } catch {
                    $errorMessage = "Error processing row $currentRowNumber (User: '$userIdentifierForLog'): $($_.Exception.Message)"
                    Write-ScriptLog -Message $errorMessage -Level ERROR -LogFilePath $LogFilePath
                    $script:summary.FailureCount++
                    $script:summary.FailedEntries.Add(@{
                        RowNumber         = $currentRowNumber
                        UserIdentifier    = $userIdentifierForLog
                        Error             = $_.Exception.Message
                        FullCSVRowContent = $row | ConvertTo-Json -Compress
                    })
                }
            }
        } catch {
            Write-ScriptLog -Message "Failed to read or process CSV file '$CSVPath'. Error: $($_.Exception.Message)" -Level ERROR -LogFilePath $LogFilePath
        }
    }

    end {
        Write-ScriptLog -Message "------------------------------------" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "Bulk Update Summary:" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "Total rows processed: $($script:summary.ProcessedRows) of $($script:summary.TotalRows)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "Successfully updated: $($script:summary.SuccessCount)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "Failed to update: $($script:summary.FailureCount)" -Level INFO -LogFilePath $LogFilePath
        if ($script:summary.FailureCount -gt 0) {
            Write-ScriptLog -Message "Details for failed entries (also check main log file):" -Level WARN -LogFilePath $LogFilePath
            $script:summary.FailedEntries | ForEach-Object { Write-ScriptLog -Message "- Row $($_.RowNumber), User: $($_.UserIdentifier), Error: $($_.Error)" -Level WARN -LogFilePath $LogFilePath }
        }
        Write-ScriptLog -Message "Invoke-BulkAADUserUpdate finished." -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "------------------------------------" -Level INFO -LogFilePath $LogFilePath
        return $script:summary
    }
}

<#
.EXAMPLE
# Assuming Update-AADUser.ps1 is in the same directory and loaded.
# . .\Update-AADUser.ps1
#
# Create a CSV file named 'update_users.csv' with the following content:
#
# UserPrincipalName,DisplayName,Department,JobTitle,ManagerUPN,GroupsToAdd,LicensesToRemove
# user1@contoso.com,User One Updated,Marketing,Specialist,newmgr@contoso.com,"Sales Team,Project Alpha",old-license-sku-guid
# user2@contoso.com,,Finance,Analyst,,,,
# user3@contoso.com,User Three New Name,,,,,,
# user4@contoso.com,,,,,,, # This row might cause an error as no update attributes are specified
# user5@contoso.com,User Five Manager Removed,,,,"",,
#
# ObjectId,OfficeLocation,MobilePhone
# xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx,New York Office,555-1234 # Example using ObjectId
#
# Invoke-BulkAADUserUpdate -CSVPath .\update_users.csv -LogFilePath .\bulk_update.log -Verbose
# Invoke-BulkAADUserUpdate -CSVPath .\update_users.csv -LogFilePath .\bulk_update.log -WhatIf


CSV Column Details:
- UserPrincipalName (string, identifier): UPN of the user to update. (Mandatory if ObjectId is not provided)
- ObjectId (string, identifier): Object ID of the user to update. (Mandatory if UserPrincipalName is not provided)
- Any other column header should match a parameter name of the `Update-AADUser` function.
  Examples: DisplayName, Department, JobTitle, OfficeLocation, StreetAddress, City, State, PostalCode, Country,
            MobilePhone, OfficePhone, ManagerUPN, GroupsToAdd, GroupsToRemove, LicensesToAssign, LicensesToRemove.
- For ManagerUPN: Provide the new manager's UPN. To remove a manager, leave the cell for ManagerUPN *explicitly empty* in the CSV.
- For multi-value fields (GroupsToAdd, GroupsToRemove, LicensesToAssign, LicensesToRemove): Use comma-separated strings.
  (e.g., "GroupA,GroupB,GroupC" or "sku1-guid,sku2-guid").
- Only non-empty cells for update attributes will be processed.
#>
