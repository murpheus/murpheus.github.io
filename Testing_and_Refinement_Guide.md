# Testing and Refinement Guide: Azure AD User Lifecycle PowerShell Scripts

This guide outlines the steps and considerations for testing and refining the suite of PowerShell scripts designed for Azure AD user lifecycle management. These scripts include: `Invoke-AADUserOnboarding.ps1`, `Update-AADUser.ps1`, `Invoke-AADUserOffboarding.ps1`, and their bulk counterparts `Invoke-BulkAADUserOnboarding.ps1` (which also defines `Write-ScriptLog`), `Invoke-BulkAADUserUpdate.ps1`, and `Invoke-BulkAADUserOffboarding.ps1`.

## 1. Test Environment Setup

Proper environment setup is crucial for effective testing.

*   **Azure AD Tenant:**
    *   **Type:** Use a **non-production/test Azure AD tenant**. *Never test these scripts directly in a production environment.*
    *   **Licensing:** Ensure the test tenant has licenses available (e.g., Azure AD Premium P1/P2, M365 E3/E5) to test license assignment features. Note down available SKU IDs.
*   **Permissions:**
    *   The account used for testing must have sufficient Azure AD permissions. Typically, roles like **User Administrator** or **Global Administrator** (use with caution, prefer least privilege) are required.
    *   Specific permissions needed:
        *   Read all users, groups, licenses.
        *   Create, update, and delete users.
        *   Assign/remove users from groups.
        *   Assign/remove licenses for users.
        *   Set/remove user managers.
        *   Revoke user sign-in sessions.
*   **PowerShell Environment:**
    *   **Az PowerShell Module:** Ensure the `Az.Accounts` and `Az.Resources` (or `Az.MicrosoftGraph` if cmdlets were updated to directly use it) modules are installed and up-to-date.
        ```powershell
        # Install if needed
        # Install-Module Az.Accounts -Scope CurrentUser -Force
        # Install-Module Az.Resources -Scope CurrentUser -Force 
        Get-Module Az.Accounts, Az.Resources -ListAvailable
        ```
    *   **Script Location:** Place all six `.ps1` script files in the same directory.
    *   **Execution Policy:** Ensure PowerShell execution policy allows running local scripts (e.g., `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser`).
*   **Connection:**
    *   Always start testing sessions by connecting to your test Azure AD tenant:
        ```powershell
        Connect-AzAccount -TenantId "your-test-tenant-id.onmicrosoft.com"
        # Or simply Connect-AzAccount and select the appropriate tenant interactively.
        ```
*   **Test Data:**
    *   Prepare lists of test user UPNs, display names.
    *   Identify existing group names/ObjectIDs in the test tenant.
    *   Identify valid license SKU IDs (`Get-AzSubscribedSku | Select SkuPartNumber, SkuId`).
    *   Identify existing user UPNs to act as managers.

## 2. Testing `Write-ScriptLog`

The `Write-ScriptLog` function (defined in `Invoke-BulkAADUserOnboarding.ps1`) is central to logging.

*   **Dot-source `Invoke-BulkAADUserOnboarding.ps1`** to make `Write-ScriptLog` available in the current session.
    ```powershell
    . .\Invoke-BulkAADUserOnboarding.ps1
    $testLogPath = ".\test_write_script_log.log"
    ```
*   **Test Cases:**
    *   **Basic INFO Log:**
        ```powershell
        Write-ScriptLog -Message "Test INFO message" -Level INFO -LogFilePath $testLogPath
        ```
        *   **Verify:** Console shows the INFO message. Log file `$testLogPath` is created and contains the timestamped INFO message.
    *   **WARN Log:**
        ```powershell
        Write-ScriptLog -Message "Test WARN message" -Level WARN -LogFilePath $testLogPath
        ```
        *   **Verify:** Console shows warning. Log file contains timestamped WARN message.
    *   **ERROR Log:**
        ```powershell
        Write-ScriptLog -Message "Test ERROR message" -Level ERROR -LogFilePath $testLogPath
        ```
        *   **Verify:** Console shows error. Log file contains timestamped ERROR message.
    *   **VERBOSE Log:**
        ```powershell
        Write-ScriptLog -Message "Test VERBOSE message" -Level VERBOSE -LogFilePath $testLogPath -Verbose
        # Also test without -Verbose to ensure it doesn't output to console then
        ```
        *   **Verify:** Console shows verbose message (if `$VerbosePreference = "Continue"` or `-Verbose` is used). Log file contains timestamped VERBOSE message.
    *   **DEBUG Log:**
        ```powershell
        Write-ScriptLog -Message "Test DEBUG message" -Level DEBUG -LogFilePath $testLogPath -Debug
        # Also test without -Debug
        ```
        *   **Verify:** Console shows debug message (if `$DebugPreference = "Continue"` or `-Debug` is used). Log file contains timestamped DEBUG message.
    *   **Log File Path:**
        *   Test with a path where the directory doesn't exist. Verify the directory is created.
        *   Test with an invalid path (e.g., restricted area) to check error handling within `Write-ScriptLog` itself (should write an error to console).
    *   **Message Content:** Ensure special characters in messages are handled correctly in the log file.

## 3. Testing Single-User Lifecycle Functions

For each single-user function, dot-source it and `Write-ScriptLog` (if not already available via the bulk onboarding script). Set a unique `$LogFilePath` for each test run.

```powershell
. .\Invoke-BulkAADUserOnboarding.ps1 # To get Write-ScriptLog
. .\Invoke-AADUserOnboarding.ps1
. .\Update-AADUser.ps1
. .\Invoke-AADUserOffboarding.ps1
$singleUserLogPath = ".\single_user_test.log"
```

### 3.1. `Invoke-AADUserOnboarding`

*   **Test Cases:**
    1.  **Minimal Parameters:**
        *   Call with only `-UserPrincipalName`, `-DisplayName`, `-LogFilePath`.
        *   **Verify:** User created in Azure AD, password auto-generated (check if policy allows this), account enabled. Log entries for auto-password generation.
    2.  **All Parameters (Valid):**
        *   Provide valid `-Password` (SecureString), `-ForceChangePasswordNextLogin $true/$false`, `-InitialGroups` (valid group names/IDs), `-LicenseSKUs` (valid SKU IDs), `-Department`, `-JobTitle`, `-ManagerUPN` (valid manager).
        *   **Verify:** User created with all attributes set correctly in Azure AD. User added to groups, licenses assigned, manager set. Check logs for each step.
    3.  **Password Management:**
        *   Provide password, test `ForceChangePasswordNextLogin $true` vs `$false`.
        *   **Verify:** User object's password policy settings in Azure AD.
    4.  **Invalid Group/License/Manager:**
        *   Provide non-existent group name, invalid SKU ID, non-existent manager UPN.
        *   **Verify:** User created. Warnings logged for failed assignments. Script should not halt.
    5.  **User Already Exists:**
        *   Attempt to create a user with an existing UPN.
        *   **Verify:** `New-AzADUser` should fail. Error correctly logged. Function returns `$null` or appropriate error status.
    6.  **`-WhatIf`:**
        *   Run a valid creation scenario with `-WhatIf`.
        *   **Verify:** No user created in Azure AD. Logs indicate "WhatIf" skip for creation and subsequent operations.
*   **Verification (General):**
    *   `Get-AzADUser -UserPrincipalName "testuser@tenant.com" | Format-List *`
    *   `Get-AzADUserAssignedLicense -UserPrincipalName "testuser@tenant.com"`
    *   `Get-AzADUserMembership -UserPrincipalName "testuser@tenant.com"`
    *   `Get-AzADUserManager -UserPrincipalName "testuser@tenant.com"`
    *   Review `$singleUserLogPath` for detailed logging of each step.

### 3.2. `Update-AADUser`

*   **Test Cases (for an existing test user):**
    1.  **Update Single Attribute:**
        *   E.g., update `-Department` only.
        *   **Verify:** Attribute updated in Azure AD. Other attributes unchanged. Logs show specific update.
    2.  **Update Multiple Attributes:**
        *   Update `-DisplayName`, `-JobTitle`, `-MobilePhone`, etc.
        *   **Verify:** All specified attributes updated.
    3.  **Manager Assignment:**
        *   Assign a new manager using `-ManagerUPN`.
        *   **Verify:** Manager updated.
    4.  **Manager Removal:**
        *   Use `-ManagerUPN ""` (empty string).
        *   **Verify:** Manager removed.
    5.  **Group Additions/Removals:**
        *   `-GroupsToAdd` (new groups), `-GroupsToRemove` (existing groups). Test with group names and ObjectIDs.
        *   **Verify:** User's group memberships correctly reflect changes.
    6.  **License Assignments/Revocations:**
        *   `-LicensesToAssign` (new SKUs), `-LicensesToRemove` (existing SKUs).
        *   **Verify:** User's licenses correctly reflect changes.
    7.  **Invalid Inputs:**
        *   Non-existent manager UPN, non-existent groups for adding/removing, invalid license SKUs.
        *   **Verify:** Valid changes applied. Warnings logged for invalid operations. Script completes.
    8.  **User Not Found:**
        *   Attempt to update a non-existent user (by UPN or ObjectId).
        *   **Verify:** Error logged. Function returns `$null` or error status.
    9.  **`-WhatIf`:**
        *   Run a valid update scenario with `-WhatIf`.
        *   **Verify:** No changes in Azure AD. Logs show "WhatIf" skips.
*   **Verification (General):** As with onboarding, use `Get-AzADUser`, `Get-AzADUserAssignedLicense`, etc., and check logs.

### 3.3. `Invoke-AADUserOffboarding`

*   **Test Cases (for an existing test user):**
    1.  **Default (Disable):**
        *   Call with `-UserPrincipalName` and `-LogFilePath`.
        *   **Verify:** Account disabled (`AccountEnabled: False`). Sessions revoked, licenses removed (by default). Logs detail these actions.
    2.  **Disable with Options:**
        *   `-RevokeSignInSessions $false`, `-RemoveAllLicenses $false`, `-RemoveFromAllGroups $true` (ensure user is in groups).
        *   **Verify:** Account disabled. Sessions NOT revoked, licenses NOT removed. User removed from all groups.
    3.  **Action Delete (Soft Delete):**
        *   Call with `-Action Delete`.
        *   **Verify:** Account disabled, sessions revoked, licenses removed (by default). User then soft-deleted (not found by `Get-AzADUser`, visible in "Deleted users" in Azure Portal).
    4.  **Delete with Options:**
        *   `-Action Delete`, `-RemoveFromAllGroups $true`.
        *   **Verify:** All steps performed before soft deletion.
    5.  **User Not Found:**
        *   Attempt to offboard a non-existent user.
        *   **Verify:** Error logged. Function returns error status.
    6.  **User Already Disabled:**
        *   Attempt to disable an already disabled user.
        *   **Verify:** Script notes user is already disabled and proceeds with other actions (e.g., license removal if specified).
    7.  **`-WhatIf`:**
        *   Run disable/delete scenarios with `-WhatIf`.
        *   **Verify:** No changes in Azure AD. Logs show "WhatIf" skips.
*   **Verification (General):**
    *   `Get-AzADUser -UserPrincipalName "..." | Select AccountEnabled` (will fail if soft-deleted).
    *   Azure Portal: "Deleted users" section.
    *   `Get-AzADUserAssignedLicense`, `Get-AzADUserMembership` before potential deletion.
    *   Check logs for detailed action reports.

## 4. Testing Bulk Operation Functions

For each bulk function, prepare sample CSV files. Dot-source the relevant single-user function and `Invoke-BulkAADUserOnboarding.ps1` (for `Write-ScriptLog`).

```powershell
. .\Invoke-BulkAADUserOnboarding.ps1 # For Write-ScriptLog and its own test
. .\Invoke-AADUserOnboarding.ps1
. .\Invoke-BulkAADUserUpdate.ps1
. .\Update-AADUser.ps1
. .\Invoke-BulkAADUserOffboarding.ps1
. .\Invoke-AADUserOffboarding.ps1

$bulkLogPath = ".\bulk_operations_test.log"
```

*   **Sample CSVs:**
    *   Create CSVs with a mix of data:
        *   Valid rows.
        *   Rows with missing mandatory data (e.g., missing `DisplayName` for onboarding).
        *   Rows with invalid data (e.g., incorrect boolean format for `ForceChangePasswordNextLogin`, bad UPN format).
        *   Rows targeting non-existent users (for update/offboarding).
        *   Empty CSV file.
        *   CSV with only headers.
        *   (For Update) Rows with only an identifier and no actual update columns.
        *   (For Update) Rows with `ManagerUPN` explicitly empty.
*   **Test Cases (General for all bulk functions):**
    1.  **Valid CSV:**
        *   Process a CSV with several valid entries.
        *   **Verify:** All users processed successfully. Azure AD reflects changes. Summary report shows correct counts. Logs are detailed.
    2.  **Mixed CSV (Valid & Invalid Rows):**
        *   Process a CSV with a mix of good and bad rows.
        *   **Verify:** Valid rows processed successfully. Errors logged for invalid rows (with row number, identifier, and error message). Script continues to next row. Summary report accurately reflects success/failure counts and lists failed entries.
    3.  **CSV Parsing Issues:**
        *   CSV with missing mandatory headers (e.g., `UserPrincipalName` missing for onboarding).
        *   **Verify:** Function should handle this gracefully, log an error about the CSV structure or missing column, and potentially halt or report widespread failures.
    4.  **Dependency Check:**
        *   Ensure the bulk functions correctly error out if their corresponding single-user function is not loaded.
    5.  **`-WhatIf` on Bulk Function:**
        *   Run with a valid CSV and `-WhatIf`.
        *   **Verify:** No changes in Azure AD. Logs show "WhatIf" skips for each row. Summary report might indicate all as "skipped" or "processed with WhatIf."
    6.  **Empty CSV / Headers Only:**
        *   **Verify:** Script handles this without error, summary shows zero processed.
*   **Verification (General):**
    *   Inspect Azure AD for a sample of users from the CSV.
    *   Carefully review the `$bulkLogPath` for:
        *   Start/end messages for the bulk process.
        *   Per-row processing messages.
        *   Correct logging of errors from single-user functions, attributed to the correct CSV row.
        *   Logging of data conversions (e.g., SecureString password, boolean parsing).
    *   Validate the accuracy of the returned summary hashtable (`SuccessCount`, `FailureCount`, `FailedEntries`, `ProcessedRows`, `TotalRows`).
    *   Ensure `FailedEntries` in the summary contains useful information for diagnosing row-level failures.

## 5. Logging & Error Handling Verification (Overall)

This is an overarching check across all scripts.

*   **Log File Integrity:**
    *   Log files are created at the specified `$LogFilePath`.
    *   Directory for log file is created if it doesn't exist.
    *   Timestamps and log levels are correctly formatted.
    *   All intended messages (INFO, WARN, ERROR, VERBOSE, DEBUG) appear in the log file as expected.
*   **Error Capture:**
    *   Errors from Azure AD cmdlets (e.g., user not found, invalid group ID) are caught and logged appropriately (typically as WARN or ERROR depending on severity for the operation).
    *   Input validation errors (e.g., bad CSV data, missing mandatory params) are caught and logged.
*   **Script Continuity (Bulk Operations):**
    *   Confirm that an error in processing one row of a CSV does not stop the entire bulk operation. The script should log the error for that row and continue to the next.
*   **`-WhatIf` Behavior:**
    *   Ensure `-WhatIf` is respected by all functions and prevents changes to Azure AD. Log messages should clearly indicate that actions were skipped due to `-WhatIf`.
*   **Verbose/Debug Output:**
    *   Test with `-Verbose` and `-Debug` switches to ensure detailed diagnostic information is logged, helping with troubleshooting.

## 6. Refinement Considerations

After functional testing, consider these aspects for refinement:

*   **Clarity of Output/Logs:**
    *   Are log messages clear, concise, and provide enough context?
    *   Is it easy to trace the lifecycle of a single user in the logs, especially during bulk operations?
    *   Is the summary report from bulk operations easy to understand?
*   **Error Messages:**
    *   Are error messages user-friendly? Do they suggest potential causes or fixes?
    *   For CSV errors, is the row number and problematic data clearly indicated?
*   **Unhandled Exceptions:**
    *   Try to force unexpected errors (e.g., disconnect network during a call). Does the script handle them gracefully or crash? (This is advanced).
*   **Performance (Conceptual for Bulk Operations):**
    *   While full performance testing is complex, consider:
        *   Are there any obviously inefficient operations within loops (e.g., repeatedly fetching the same information that could be cached)?
        *   For very large CSVs (e.g., thousands of users), would there be any throttling concerns with Azure AD APIs? (The scripts call cmdlets one by one per user, which is generally safer than trying to batch too aggressively without proper Graph API batching knowledge).
        *   The current design processes users serially. This is simpler and more robust for logging but slower than parallel processing. For massive datasets, parallel invocation might be a future consideration, but it significantly increases complexity in error handling and logging.
*   **Usability:**
    *   Are parameter names intuitive?
    *   Is the CSV header expectation clear and well-documented in the script's initial output/comments?
    *   Is the process for loading dependencies (single-user functions, `Write-ScriptLog`) clear for users? (Consider packaging as a module in the future).
*   **Idempotency (for single-user functions where applicable):**
    *   Can an operation be run multiple times with the same input and produce the same result without error (e.g., assigning a user to a group they are already in)? `Add-AzADGroupMember` usually handles this, but it's a good thought process.
    *   `Invoke-AADUserOnboarding` is not idempotent (cannot create the same user twice).
    *   `Update-AADUser` should ideally be idempotent for attribute setting.
    *   `Invoke-AADUserOffboarding` (Disable) should be idempotent. (Delete is not).

By systematically going through these testing phases and considering the refinement points, the reliability and usability of the Azure AD PowerShell script suite can be significantly improved. Remember to document any bugs found and their resolutions.Okay, I've created the "Testing and Refinement Guide" for the PowerShell scripts.
