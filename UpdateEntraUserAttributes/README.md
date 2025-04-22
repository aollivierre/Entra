# Microsoft Entra ID User Attribute Update Script

This PowerShell script updates user attributes in Microsoft Entra ID (formerly Azure AD) based on data from a CSV file.

## Prerequisites

- PowerShell 5.1 or higher
- Microsoft Graph PowerShell SDK modules:
  - Microsoft.Graph.Authentication
  - Microsoft.Graph.Users

To install the required modules:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module Microsoft.Graph.Users -Scope CurrentUser
```

## CSV File Format

The script expects a CSV file with the following headers:
- `User principal name` (required)
- `Department`
- `Title`
- `Fax`
- `Manager`

## Usage

```powershell
# Basic usage
.\Update-EntraUserAttributes.ps1 -CsvPath "path\to\your\file.csv"

# Test run (updates only the first user)
.\Update-EntraUserAttributes.ps1 -CsvPath "path\to\your\file.csv" -TestRun

# Save logs to file
.\Update-EntraUserAttributes.ps1 -CsvPath "path\to\your\file.csv" -LogPath "path\to\logfile.log"
```

## Notes

- The script will prompt for authentication to Microsoft Graph with the appropriate permissions
- The user running the script must have sufficient permissions in Entra ID
- Managers are located by their display name, so ensure names in the CSV match exactly with names in Entra ID 