# Export-AADAppsSecretsAndCertificates

A PowerShell script that collects all Azure AD registered applications with secrets expiration dates and certificates expiration dates and exports them to an Excel file.
# Prerequisites
- Install the AzureAD and ImportExcel modules before running this script.

# Usage
```powershell
.\Export-AADAppsSecretsAndCertificates.ps1
```

This will connect to Azure AD and export the collected data to an Excel file named "AADAppsSecrets_YYYY-MM-DD.xlsx" in the current directory.
# Optional Parameters
- OutputFile

Specify the full path and name for the output Excel file.
# Notes

The Excel file will contain two sheets: one for the list of applications and their secrets and certificates count, and one for the details of each secret and certificate.  
The script sorts the data by the application display name before exporting.