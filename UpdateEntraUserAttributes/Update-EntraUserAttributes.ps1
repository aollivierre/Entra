

<#
.SYNOPSIS
    Updates Microsoft Entra ID user attributes from a CSV file.
.DESCRIPTION
    This script reads user data from a CSV file and updates department, manager, title, 
    and fax attributes in Microsoft Entra ID using Microsoft Graph API.
.PARAMETER CsvPath
    Path to the CSV file containing user information.
.PARAMETER TestRun
    If specified, only updates the first user in the CSV file.
.PARAMETER LogPath
    Path to write log file. If not specified, logs will only be written to console.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$CsvPath,
    
    [Parameter(Mandatory = $false)]
    [switch]$TestRun,
    
    [Parameter(Mandatory = $false)]
    [string]$LogPath
)


# Check and install required modules
$requiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users')
$modulesToInstall = @()

foreach ($module in $requiredModules) {
    if (-not (Get-Module -Name $module -ListAvailable)) {
        $modulesToInstall += $module
    }
}

if ($modulesToInstall.Count -gt 0) {
    Write-Host "Installing required modules: $($modulesToInstall -join ', ')"
    foreach ($module in $modulesToInstall) {
        try {
            Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber
            Write-Host "Successfully installed module $module" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to install module $module. Error: $_" -ForegroundColor Red
            exit 1
        }
    }
}

# Import required modules
foreach ($module in $requiredModules) {
    try {
        Import-Module -Name $module -ErrorAction Stop
        Write-Host "Successfully imported module $module" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to import module $module. Error: $_" -ForegroundColor Red
        exit 1
    }
}



# Function for logging
function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('INFO', 'WARNING', 'ERROR')]
        [string]$Level = 'INFO'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Write to console
    switch ($Level) {
        'INFO' { Write-Host $logMessage }
        'WARNING' { Write-Host $logMessage -ForegroundColor Yellow }
        'ERROR' { Write-Host $logMessage -ForegroundColor Red }
    }
    
    # Write to log file if specified
    if ($LogPath) {
        $logMessage | Out-File -FilePath $LogPath -Append
    }
}

# Check if CSV file exists
if (-not (Test-Path -Path $CsvPath)) {
    Write-Log "CSV file not found at path $CsvPath" -Level 'ERROR'
    exit 1
}

# Import CSV file
try {
    $users = Import-Csv -Path $CsvPath
    Write-Log "Successfully imported CSV with $($users.Count) users"
} 
catch {
    Write-Log "Failed to import CSV file: $_" -Level 'ERROR'
    exit 1
}

# Connect to Microsoft Graph with required permissions
$requiredScopes = @(
    'User.ReadWrite.All',
    'Directory.ReadWrite.All'
)

try {
    Write-Log "Connecting to Microsoft Graph..."
    Connect-MgGraph -Scopes $requiredScopes
    Write-Log "Successfully connected to Microsoft Graph"
}
catch {
    Write-Log "Failed to connect to Microsoft Graph: $_" -Level 'ERROR'
    exit 1
}

# If TestRun is specified, only process the first user
if ($TestRun) {
    Write-Log "Test run mode enabled - only processing the first user"
    $users = $users | Select-Object -First 1
}

# Process each user in the CSV
$processedCount = 0
$errorCount = 0
$notFoundCount = 0
$summaryReport = @()

foreach ($user in $users) {
    $upn = $user.'User principal name'
    
    if (-not $upn) {
        Write-Log "Skipping user with empty UPN" -Level 'WARNING'
        continue
    }
    
    try {
        # Try to get the user from Entra ID
        Write-Log "Looking up user $upn in Entra ID"
        
        try {
            $entraUser = Get-MgUser -UserId $upn -ErrorAction Stop
        }
        catch {
            if ($_.Exception.Message -like "*Resource*not*found*" -or $_.Exception.Message -like "*404*") {
                Write-Log "User $upn not found in Entra ID" -Level 'WARNING'
                $notFoundCount++
                continue
            }
            else {
                throw
            }
        }
        
        # Record current values before updating
        $currentUser = Get-MgUser -UserId $entraUser.Id -Property Id, DisplayName, UserPrincipalName, Department, JobTitle, FaxNumber
        $currentManagerId = $null
        
        try {
            $currentManager = Get-MgUserManager -UserId $entraUser.Id -ErrorAction SilentlyContinue
            if ($currentManager) {
                $currentManagerId = $currentManager.Id
                $currentManagerUser = Get-MgUser -UserId $currentManagerId -ErrorAction SilentlyContinue
                $currentManagerName = if ($currentManagerUser) { $currentManagerUser.DisplayName } else { "Unknown Manager ($currentManagerId)" }
            }
            else {
                $currentManagerName = "No manager assigned"
            }
        }
        catch {
            $currentManagerName = "Error retrieving manager"
            Write-Log "Could not retrieve current manager for user $upn : $_" -Level 'WARNING'
        }
        
        # Log current values
        Write-Log "Current values for user $upn"
        Write-Log "  - Department : $($currentUser.Department)"
        Write-Log "  - Title      : $($currentUser.JobTitle)"
        Write-Log "  - Fax        : $($currentUser.FaxNumber)"
        Write-Log "  - Manager    : $currentManagerName"
        
        # Prepare update parameters
        $updateParams = @{
            Department = $user.Department
            JobTitle = $user.Title
            FaxNumber = $user.Fax
        }
        
        # Update user attributes
        Update-MgUser -UserId $entraUser.Id -BodyParameter $updateParams
        Write-Log "Updated department, title, and fax for user $upn"
        
        # Handle manager update if present
        $newManagerName = "No change"
        if ($user.Manager) {
            try {
                # Find manager in Entra ID
                $manager = Get-MgUser -Filter "displayName eq '$($user.Manager)'" -ErrorAction Stop
                
                if ($manager) {
                    # Only update if the manager is different
                    if ($manager.Id -ne $currentManagerId) {
                        # Set manager reference
                        Set-MgUserManagerByRef -UserId $entraUser.Id -BodyParameter @{
                            "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($manager.Id)"
                        }
                        Write-Log "Updated manager for user $upn to $($user.Manager)"
                        $newManagerName = $user.Manager
                    }
                    else {
                        Write-Log "Manager for user $upn is already set to $($user.Manager) - no update needed"
                        $newManagerName = $user.Manager + " (no change needed)"
                    }
                }
                else {
                    Write-Log "Manager '$($user.Manager)' not found for user $upn" -Level 'WARNING'
                    $newManagerName = "Not found: $($user.Manager)"
                }
            }
            catch {
                Write-Log "Failed to update manager for user $upn : $_" -Level 'ERROR'
                $errorCount++
                $newManagerName = "Error updating: $($user.Manager)"
            }
        }
        
        # Get updated user values for the report
        $updatedUser = Get-MgUser -UserId $entraUser.Id -Property Id, DisplayName, UserPrincipalName, Department, JobTitle, FaxNumber
        
        # Add to summary report
        $summaryReport += [PSCustomObject]@{
            UserPrincipalName = $upn
            DisplayName = $entraUser.DisplayName
            Department_Before = $currentUser.Department
            Department_After = $updatedUser.Department
            Title_Before = $currentUser.JobTitle
            Title_After = $updatedUser.JobTitle
            Fax_Before = $currentUser.FaxNumber
            Fax_After = $updatedUser.FaxNumber
            Manager_Before = $currentManagerName
            Manager_After = $newManagerName
        }
        
        $processedCount++
    }
    catch {
        Write-Log "Failed to process user $upn : $_" -Level 'ERROR'
        $errorCount++
    }
}

# Summary
Write-Log "Processing complete." 
Write-Log "Total users in CSV: $($users.Count)"
Write-Log "Successfully processed: $processedCount"
Write-Log "Users not found in Entra ID: $notFoundCount"
Write-Log "Errors encountered: $errorCount"

# Display detailed report
Write-Log "Detailed Summary Report:" -Level 'INFO'
$reportWidth = 100
Write-Log ('-' * $reportWidth)
Write-Log ("{0,-40} {1,-30} {2,-30}" -f "USER", "ATTRIBUTE", "CHANGES")
Write-Log ('-' * $reportWidth)

foreach ($item in $summaryReport) {
    Write-Log ("{0,-40} {1,-30} {2,-30}" -f $item.DisplayName, "Department", "[$($item.Department_Before)] → [$($item.Department_After)]")
    Write-Log ("{0,-40} {1,-30} {2,-30}" -f "", "Title", "[$($item.Title_Before)] → [$($item.Title_After)]")
    Write-Log ("{0,-40} {1,-30} {2,-30}" -f "", "Fax", "[$($item.Fax_Before)] → [$($item.Fax_After)]")
    Write-Log ("{0,-40} {1,-30} {2,-30}" -f "", "Manager", "[$($item.Manager_Before)] → [$($item.Manager_After)]")
    Write-Log ('-' * $reportWidth)
}

# Disconnect from Microsoft Graph
Disconnect-MgGraph | Out-Null
Write-Log "Disconnected from Microsoft Graph" 