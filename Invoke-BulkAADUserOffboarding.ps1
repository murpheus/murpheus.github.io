#Requires -Modules Az.Accounts, Az.Resources 
# Make sure Invoke-AADUserOffboarding.ps1 is loaded in the session.
# Example: . .\Invoke-AADUserOffboarding.ps1
# Ensure Write-ScriptLog is available in the session if this script is run standalone
# For bulk operations, Write-ScriptLog is assumed to be loaded by the calling (e.g. bulk onboarding) script.

function Invoke-BulkAADUserOffboarding {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true, HelpMessage = "Path to the CSV file for bulk user offboarding.")]
        [string]$CSVPath,

        [Parameter(Mandatory = $true, HelpMessage = "Path to the log file.")]
        [string]$LogFilePath
    )

    function _EnsureWriteScriptLogBulkOffboard { # Renamed to avoid conflict
        if (-not (Get-Command Write-ScriptLog -ErrorAction SilentlyContinue)) {
            Write-Warning "Write-ScriptLog function not found. Using basic console output for this session (Invoke-BulkAADUserOffboarding)."
            New-Alias -Name Write-ScriptLog -Value Write-Host -Scope Script -Force 
        }
    }
    _EnsureWriteScriptLogBulkOffboard

    begin {
        Write-ScriptLog -Message "Invoke-BulkAADUserOffboarding starting." -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "Expected CSV Headers (at least one identifier - UserPrincipalName or ObjectId - is mandatory):" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- UserPrincipalName (string, identifier)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- ObjectId (string, identifier)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- Action (string, optional, 'Disable' or 'Delete', defaults to 'Disable')" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- RevokeSignInSessions (True/False, optional, defaults to True)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- RemoveAllLicenses (True/False, optional, defaults to True)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- RemoveFromAllGroups (True/False, optional, defaults to False)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "------------------------------------" -Level INFO -LogFilePath $LogFilePath

        if (-not (Test-Path -Path $CSVPath -PathType Leaf)) {
            Write-ScriptLog -Message "CSV file not found at path: $CSVPath" -Level ERROR -LogFilePath $LogFilePath
            return 
        }

        if (-not (Get-Command Invoke-AADUserOffboarding -ErrorAction SilentlyContinue)) {
            Write-ScriptLog -Message "Critical dependency missing: Function 'Invoke-AADUserOffboarding' is not loaded. Please load it first (e.g., '. .\Invoke-AADUserOffboarding.ps1')." -Level ERROR -LogFilePath $LogFilePath
            return
        }
        
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
                    $offboardingParams = @{ LogFilePath = $LogFilePath } # Pass LogFilePath
                    $identifierProvided = $false

                    if ($row.PSObject.Properties['UserPrincipalName'] -and -not [string]::IsNullOrWhiteSpace($row.UserPrincipalName)) {
                        $offboardingParams.UserPrincipalName = $row.UserPrincipalName.Trim()
                        $userIdentifierForLog = $offboardingParams.UserPrincipalName
                        $identifierProvided = $true
                    } elseif ($row.PSObject.Properties['ObjectId'] -and -not [string]::IsNullOrWhiteSpace($row.ObjectId)) {
                        $offboardingParams.ObjectId = $row.ObjectId.Trim()
                        $userIdentifierForLog = $offboardingParams.ObjectId
                        $identifierProvided = $true
                    }

                    if (-not $identifierProvided) {
                        throw "Mandatory identifier (UserPrincipalName or ObjectId) is missing or empty."
                    }

                    Write-ScriptLog -Message "Processing user $currentRowNumber of $($script:summary.TotalRows): '$userIdentifierForLog'" -Level INFO -LogFilePath $LogFilePath
                    
                    if ($row.PSObject.Properties['Action'] -and -not [string]::IsNullOrWhiteSpace($row.Action)) {
                        $actionValue = $row.Action.Trim()
                        if ($actionValue -in "Disable", "Delete") {
                            $offboardingParams.Action = $actionValue
                        } else {
                            throw "Invalid value for 'Action' for user '$userIdentifierForLog'. Must be 'Disable' or 'Delete'."
                        }
                    } 

                    $booleanParams = @("RevokeSignInSessions", "RemoveAllLicenses", "RemoveFromAllGroups")
                    foreach ($paramName in $booleanParams) {
                        if ($row.PSObject.Properties[$paramName] -and -not [string]::IsNullOrWhiteSpace($row.$paramName)) {
                            try {
                                $offboardingParams[$paramName] = [System.Convert]::ToBoolean($row.$paramName.Trim())
                            } catch {
                                throw "Invalid boolean value for '$paramName' for user '$userIdentifierForLog'. Error: $($_.Exception.Message)"
                            }
                        } 
                    }

                    $effectiveActionForShouldProcess = $offboardingParams.Action # Will be $null if not in CSV, then single-user func default applies
                    if (-not $effectiveActionForShouldProcess) { $effectiveActionForShouldProcess = "Disable" } 
                    
                    $shouldProcessMessage = "Offboard user '$userIdentifierForLog' with action '$effectiveActionForShouldProcess'"
                    # Build more descriptive message for ShouldProcess based on what's actually being passed
                    $descParams = @()
                    if($offboardingParams.ContainsKey('RevokeSignInSessions')) {$descParams += "RevokeSignInSessions:$($offboardingParams.RevokeSignInSessions)"}
                    if($offboardingParams.ContainsKey('RemoveAllLicenses')) {$descParams += "RemoveAllLicenses:$($offboardingParams.RemoveAllLicenses)"}
                    if($offboardingParams.ContainsKey('RemoveFromAllGroups')) {$descParams += "RemoveFromAllGroups:$($offboardingParams.RemoveFromAllGroups)"}
                    if($descParams.Count -gt 0) {$shouldProcessMessage += " ($($descParams -join ', '))"}


                    if ($PSCmdlet.ShouldProcess($userIdentifierForLog, $shouldProcessMessage)) {
                        if ($PSBoundParameters.ContainsKey('Verbose') && $PSBoundParameters['Verbose']) { $offboardingParams.Verbose = $true }
                        if ($PSBoundParameters.ContainsKey('Debug') && $PSBoundParameters['Debug']) { $offboardingParams.Debug = $true }

                        $result = Invoke-AADUserOffboarding @offboardingParams
                        
                        if ($result -is [string] -and $result.StartsWith("Success:")) {
                            Write-ScriptLog -Message "Successfully processed offboarding for user '$userIdentifierForLog'. Result: $result" -Level INFO -LogFilePath $LogFilePath
                            $script:summary.SuccessCount++
                        } elseif ($result -is [string] -and ($result.StartsWith("Skipped:") -or $result.StartsWith("Warn:"))) {
                             Write-ScriptLog -Message "Offboarding action for user '$userIdentifierForLog' resulted in a non-success status: $result" -Level WARN -LogFilePath $LogFilePath
                             $script:summary.FailureCount++ 
                             $script:summary.FailedEntries.Add(@{
                                RowNumber         = $currentRowNumber
                                UserIdentifier    = $userIdentifierForLog
                                Error             = "Non-Success status from single-user function: $result"
                                FullCSVRowContent = $row | ConvertTo-Json -Compress
                            })
                        } else { # Covers "Error:" or any other unexpected string/object from single-user function
                            throw "Offboarding failed for user '$userIdentifierForLog'. Result from single-user function: $result"
                        }
                    } else {
                        Write-ScriptLog -Message "Skipped offboarding for user '$userIdentifierForLog' (row $currentRowNumber) due to -WhatIf or user cancellation." -Level WARN -LogFilePath $LogFilePath
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
        Write-ScriptLog -Message "Bulk Offboarding Summary:" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "Total rows processed: $($script:summary.ProcessedRows) of $($script:summary.TotalRows)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "Successfully processed offboarding: $($script:summary.SuccessCount)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "Failed or non-success offboarding steps: $($script:summary.FailureCount)" -Level INFO -LogFilePath $LogFilePath
        if ($script:summary.FailureCount -gt 0) {
            Write-ScriptLog -Message "Details for failed or non-success entries (also check main log file):" -Level WARN -LogFilePath $LogFilePath
            $script:summary.FailedEntries | ForEach-Object { Write-ScriptLog -Message "- Row $($_.RowNumber), User: $($_.UserIdentifier), Detail: $($_.Error)" -Level WARN -LogFilePath $LogFilePath }
        }
        Write-ScriptLog -Message "Invoke-BulkAADUserOffboarding finished." -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "------------------------------------" -Level INFO -LogFilePath $LogFilePath
        return $script:summary
    }
}

<#
.EXAMPLE
# Assuming Invoke-AADUserOffboarding.ps1 is in the same directory and loaded.
# . .\Invoke-AADUserOffboarding.ps1
#
# Create a CSV file named 'offboard_users.csv' with the following content:
#
# UserPrincipalName,Action,RevokeSignInSessions,RemoveAllLicenses,RemoveFromAllGroups
# user1@contoso.com,Disable,True,True,False
# user2@contoso.com,Delete,True,True,True
# user3@contoso.com,,False,False, # Action defaults to Disable, flags default to True but overridden here
# user4@contoso.com,Delete,,,False # Flags RevokeSessions & RemoveLicenses default to True
#
# ObjectId,Action
# xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx,Disable # Example using ObjectId
#
# Invoke-BulkAADUserOffboarding -CSVPath .\offboard_users.csv -LogFilePath .\bulk_offboard.log -Verbose
# Invoke-BulkAADUserOffboarding -CSVPath .\offboard_users.csv -LogFilePath .\bulk_offboard.log -WhatIf

CSV Column Details:
- UserPrincipalName (string, identifier): UPN of the user. (Mandatory if ObjectId is not provided)
- ObjectId (string, identifier): Object ID of the user. (Mandatory if UserPrincipalName is not provided)
- Action (string, optional): 'Disable' or 'Delete'. Defaults to 'Disable' if column is missing or cell is empty.
- RevokeSignInSessions (True/False, optional): Defaults to True if column is missing or cell is empty (as per Invoke-AADUserOffboarding's default).
- RemoveAllLicenses (True/False, optional): Defaults to True if column is missing or cell is empty (as per Invoke-AADUserOffboarding's default).
- RemoveFromAllGroups (True/False, optional): Defaults to False if column is missing or cell is empty (as per Invoke-AADUserOffboarding's default).
#>
