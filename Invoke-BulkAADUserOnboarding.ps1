# Function to write logs to both console and file
function Write-ScriptLog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet('INFO', 'WARN', 'ERROR', 'VERBOSE', 'DEBUG')]
        [string]$Level = 'INFO',

        [Parameter(Mandatory = $true)]
        [string]$LogFilePath
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp [$Level] - $Message"

    try {
        if (-not (Test-Path -Path (Split-Path -Path $LogFilePath -Parent) -PathType Container)) {
            New-Item -ItemType Directory -Path (Split-Path -Path $LogFilePath -Parent) -Force -ErrorAction Stop | Out-Null
        }
        Add-Content -Path $LogFilePath -Value $logEntry -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to write to log file '$LogFilePath'. Error: $($_.Exception.Message)"
    }

    switch ($Level) {
        'INFO'    { Write-Host $logEntry }
        'WARN'    { Write-Warning $logEntry }
        'ERROR'   { Write-Error $logEntry }
        'VERBOSE' { Write-Verbose $logEntry }
        'DEBUG'   { Write-Debug $logEntry }
        default   { Write-Host $logEntry } # Default to INFO behavior
    }
}

#Requires -Modules Az.Accounts, Az.Resources 
# Make sure Invoke-AADUserOnboarding is loaded in the session.
# Example: . .\Invoke-AADUserOnboarding.ps1

function Invoke-BulkAADUserOnboarding {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true, HelpMessage = "Path to the CSV file for bulk user onboarding.")]
        [string]$CSVPath,

        [Parameter(Mandatory = $true, HelpMessage = "Path to the log file.")]
        [string]$LogFilePath
    )

    begin {
        Write-ScriptLog -Message "Invoke-BulkAADUserOnboarding starting." -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "Expected CSV Headers (mandatory in bold):" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- UserPrincipalName (string, mandatory)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- DisplayName (string, mandatory)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- Password (string, optional, will be converted to SecureString; auto-generated if blank)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- ForceChangePasswordNextLogin (True/False, optional, defaults to True if Password is provided)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- InitialGroups (comma-separated string of Group DisplayNames or ObjectIDs, optional)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- LicenseSKUs (comma-separated string of License SKU GUIDs, optional)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- Department (string, optional)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- JobTitle (string, optional)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "- ManagerUPN (string, optional)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "------------------------------------" -Level INFO -LogFilePath $LogFilePath

        if (-not (Test-Path -Path $CSVPath -PathType Leaf)) {
            Write-ScriptLog -Message "CSV file not found at path: $CSVPath" -Level ERROR -LogFilePath $LogFilePath
            return # Exit if CSV not found
        }
        
        if (-not (Get-Command Invoke-AADUserOnboarding -ErrorAction SilentlyContinue)) {
            Write-ScriptLog -Message "Critical dependency missing: Function 'Invoke-AADUserOnboarding' is not loaded. Please load it first (e.g., '. .\Invoke-AADUserOnboarding.ps1')." -Level ERROR -LogFilePath $LogFilePath
            return 
        }

        $script:summary = @{ # Made script-scoped for easier access in End block if needed
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
                $userPrincipalName = $null # For error logging if UPN parsing fails
                $userPrincipalName_raw_for_error = if ($row.PSObject.Properties['UserPrincipalName']) {$row.UserPrincipalName} else {"<UPN_COLUMN_MISSING_OR_INVALID>"}


                try {
                    if ($row.PSObject.Properties['UserPrincipalName']) {
                        $userPrincipalName = $row.UserPrincipalName.Trim()
                        if ([string]::IsNullOrWhiteSpace($userPrincipalName)) {
                            throw "UserPrincipalName is missing or empty."
                        }
                    } else {
                        throw "CSV column 'UserPrincipalName' not found."
                    }

                    Write-ScriptLog -Message "Processing user $currentRowNumber of $($script:summary.TotalRows): '$userPrincipalName'" -Level INFO -LogFilePath $LogFilePath

                    if ($PSCmdlet.ShouldProcess($userPrincipalName, "Onboard Azure AD User from CSV row $currentRowNumber")) {
                        $onboardingParams = @{ LogFilePath = $LogFilePath } # Pass LogFilePath

                        if ($row.PSObject.Properties['DisplayName']) {
                            $onboardingParams.DisplayName = $row.DisplayName.Trim()
                            if ([string]::IsNullOrWhiteSpace($onboardingParams.DisplayName)) {
                                throw "DisplayName is missing or empty for UPN '$userPrincipalName'."
                            }
                        } else {
                            throw "CSV column 'DisplayName' not found for UPN '$userPrincipalName'."
                        }
                        
                        $onboardingParams.UserPrincipalName = $userPrincipalName

                        if ($row.PSObject.Properties['Password'] -and -not [string]::IsNullOrWhiteSpace($row.Password)) {
                            $onboardingParams.Password = ConvertTo-SecureString $row.Password.Trim() -AsPlainText -Force
                            if ($row.PSObject.Properties['ForceChangePasswordNextLogin'] -and -not [string]::IsNullOrWhiteSpace($row.ForceChangePasswordNextLogin)) {
                                try { $onboardingParams.ForceChangePasswordNextLogin = [System.Convert]::ToBoolean($row.ForceChangePasswordNextLogin.Trim()) }
                                catch { throw "Invalid boolean value for 'ForceChangePasswordNextLogin' for UPN '$userPrincipalName'. Error: $($_.Exception.Message)" }
                            }
                        } elseif ($row.PSObject.Properties['ForceChangePasswordNextLogin'] -and -not [string]::IsNullOrWhiteSpace($row.ForceChangePasswordNextLogin)) {
                             try { $onboardingParams.ForceChangePasswordNextLogin = [System.Convert]::ToBoolean($row.ForceChangePasswordNextLogin.Trim()) }
                             catch { throw "Invalid boolean value for 'ForceChangePasswordNextLogin' for UPN '$userPrincipalName'. Error: $($_.Exception.Message)" }
                        }

                        if ($row.PSObject.Properties['InitialGroups'] -and -not [string]::IsNullOrWhiteSpace($row.InitialGroups)) {
                            $onboardingParams.InitialGroups = $row.InitialGroups.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                        }
                        if ($row.PSObject.Properties['LicenseSKUs'] -and -not [string]::IsNullOrWhiteSpace($row.LicenseSKUs)) {
                            $onboardingParams.LicenseSKUs = $row.LicenseSKUs.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                        }
                        if ($row.PSObject.Properties['Department'] -and -not [string]::IsNullOrWhiteSpace($row.Department)) {
                            $onboardingParams.Department = $row.Department.Trim()
                        }
                        if ($row.PSObject.Properties['JobTitle'] -and -not [string]::IsNullOrWhiteSpace($row.JobTitle)) {
                            $onboardingParams.JobTitle = $row.JobTitle.Trim()
                        }
                        if ($row.PSObject.Properties['ManagerUPN'] -and -not [string]::IsNullOrWhiteSpace($row.ManagerUPN)) {
                            $onboardingParams.ManagerUPN = $row.ManagerUPN.Trim()
                        }
                        
                        if ($PSBoundParameters.ContainsKey('Verbose') && $PSBoundParameters['Verbose']) {
                            $onboardingParams.Verbose = $true # Pass Verbose to single-user function if specified for bulk
                        }
                         if ($PSBoundParameters.ContainsKey('Debug') && $PSBoundParameters['Debug']) {
                            $onboardingParams.Debug = $true 
                        }


                        $result = Invoke-AADUserOnboarding @onboardingParams
                        
                        # Check result from Invoke-AADUserOnboarding
                        # Assuming it returns the user object on success, or $null/throws an error which would be caught.
                        # If it returns a status string, adjust this logic.
                        if ($result -and $result.Id) { # Check for a valid user object with an ID
                            Write-ScriptLog -Message "Successfully onboarded user '$userPrincipalName'." -Level INFO -LogFilePath $LogFilePath
                            $script:summary.SuccessCount++
                        } else {
                            # This 'else' might not be reached if Invoke-AADUserOnboarding throws an error on failure (which is good)
                            # Or if it returns a custom status object/string that needs specific parsing.
                            # For now, assume any non-object or null result is a failure that wasn't an exception.
                            throw "Onboarding failed for user '$userPrincipalName'. Result from single-user function was not a success object: $result"
                        }
                    } else {
                        Write-ScriptLog -Message "Skipped onboarding for user '$userPrincipalName' (row $currentRowNumber) due to -WhatIf or user cancellation." -Level WARN -LogFilePath $LogFilePath
                    }

                } catch {
                    $errorMessage = if ($userPrincipalName) {
                        "Error processing UserPrincipalName '$userPrincipalName' (row $currentRowNumber): $($_.Exception.Message)"
                    } else {
                        "Error processing row $currentRowNumber (Problematic UserPrincipalName: '$($userPrincipalName_raw_for_error -replace "'", "''")'): $($_.Exception.Message)"
                    }
                    Write-ScriptLog -Message $errorMessage -Level ERROR -LogFilePath $LogFilePath
                    $script:summary.FailureCount++
                    $script:summary.FailedEntries.Add(@{
                        RowNumber         = $currentRowNumber
                        UserIdentifier    = if ($userPrincipalName) {$userPrincipalName} else {$userPrincipalName_raw_for_error}
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
        Write-ScriptLog -Message "Bulk Onboarding Summary:" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "Total rows processed: $($script:summary.ProcessedRows) of $($script:summary.TotalRows)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "Successfully onboarded: $($script:summary.SuccessCount)" -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "Failed to onboard: $($script:summary.FailureCount)" -Level INFO -LogFilePath $LogFilePath # Changed to INFO as it's a summary fact
        if ($script:summary.FailureCount -gt 0) {
            Write-ScriptLog -Message "Details for failed entries (also check main log file):" -Level WARN -LogFilePath $LogFilePath
            $script:summary.FailedEntries | ForEach-Object { 
                Write-ScriptLog -Message "- Row $($_.RowNumber), User: $($_.UserIdentifier), Error: $($_.Error)" -Level WARN -LogFilePath $LogFilePath
            }
        }
        Write-ScriptLog -Message "Invoke-BulkAADUserOnboarding finished." -Level INFO -LogFilePath $LogFilePath
        Write-ScriptLog -Message "------------------------------------" -Level INFO -LogFilePath $LogFilePath
        return $script:summary
    }
}

<#
.EXAMPLE
# Assuming Invoke-AADUserOnboarding.ps1 is in the same directory and loaded.
# . .\Invoke-AADUserOnboarding.ps1 
#
# Create a CSV file named 'onboard_users.csv' with the following content:
#
# UserPrincipalName,DisplayName,Password,ForceChangePasswordNextLogin,InitialGroups,LicenseSKUs,Department,JobTitle,ManagerUPN
# user1@contoso.com,User One,P@sswOrd1,True,"Group A, Group B",sku1-guid,IT,Developer,manager1@contoso.com
# user2@contoso.com,User Two,P@sswOrd2,False,,sku2-guid,Sales,Account Rep,
# user3@contoso.com,User Three,,,,,,
# user4@contoso.com,User Four Bad FCPNL,P@sswOrd3,MaybeBoolean,"Group C",,,
# missingupn@contoso.com,,PasswordForMissingUPN,,,,
#
# Invoke-BulkAADUserOnboarding -CSVPath .\onboard_users.csv -LogFilePath .\bulk_onboard.log -Verbose
# Invoke-BulkAADUserOnboarding -CSVPath .\onboard_users.csv -LogFilePath .\bulk_onboard.log -WhatIf

.EXAMPLE
# Minimal CSV:
# UserPrincipalName,DisplayName
# newuser@example.com,New User Example
# another@example.com,Another User
#
# Invoke-BulkAADUserOnboarding -CSVPath .\minimal_onboard.csv -LogFilePath .\minimal_onboard.log

CSV Column Details:
- UserPrincipalName (string, mandatory): User's UPN (e.g., user@domain.com).
- DisplayName (string, mandatory): User's full display name.
- Password (string, optional): User's initial password. If blank, Invoke-AADUserOnboarding should generate one.
- ForceChangePasswordNextLogin (True/False, optional): If true, user must change password on next login. Defaults to true if a password is provided, otherwise depends on Invoke-AADUserOnboarding.
- InitialGroups (comma-separated string, optional): Display names or ObjectIDs of groups to add user to (e.g., "Sales Team,Marketing Group").
- LicenseSKUs (comma-separated string, optional): GUIDs of license SKUs to assign (e.g., "guid1,guid2").
- Department (string, optional): User's department.
- JobTitle (string, optional): User's job title.
- ManagerUPN (string, optional): UPN of the user's manager.

#>
