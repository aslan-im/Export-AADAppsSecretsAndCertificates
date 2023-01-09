<#
.SYNOPSIS
    Collects all Azure AD registered applications with secrets expiration dates and certificates expiration dates and exports them to an Excel file.
.DESCRIPTION
    This script uses the AzureAD and ImportExcel PowerShell modules to connect to Azure AD, retrieve information about all registered applications and their secrets and certificates, and export the data to an Excel file. The Excel file will contain two sheets: one for the list of applications and their secrets and certificates count, and one for the details of each secret and certificate. The script also sorts the data by the application display name before exporting.
.NOTES
    Make sure to install the AzureAD and ImportExcel modules before running this script.
.EXAMPLE
    .\Export-AADAppsSecretsAndCertificates.ps1
    Connects to Azure AD and exports the collected data to an Excel file named "AADAppsSecrets_YYYY-MM-DD.xlsx" in the current directory.
#>

[CmdletBinding()]
param (
    # Export file full path and name
    [Parameter(Mandatory = $False)]
    [string]
    $OutputFile = "AADAppsSecrets_$(Get-Date -Format "yyyy-MM-dd").xlsx"
)

#Requires -module AzureAD, ImportExcel

# Connect Azure Ad
Connect-AzureAD

# Variables
$AllApps = @()
$AllAppsWithSecrets = @()
$AppsWithSecretsAndCertificatesCount = @()

# Get all apps
$AllApps = Get-AzureADApplication -All $True

# Get all secrets and certificates and add to array
foreach($App in $AllApps){
    $AppSecrets = Get-AzureADApplicationPasswordCredential -ObjectId $App.ObjectId
    $AppCertificates = Get-AzureADApplicationKeyCredential -ObjectId $App.ObjectId
    $AppSecretsCount = $AppSecrets.count
    $AppCertificatesCount = $AppCertificates.count
    $AppsWithSecretsAndCertificatesCount += New-Object PSObject -Property @{
        AppId = $App.AppId
        DisplayName = $App.DisplayName
        SecretsCount = $AppSecretsCount
        CertificatesCount = $AppCertificatesCount
    }
    Write-Output "Working with $($App.DisplayName) $($App.AppId) - $AppSecretsCount secrets and $AppCertificatesCount certificates."

    if($AppSecrets){
        foreach($Secret in $AppSecrets){
            $AllAppsWithSecrets += New-Object PSObject -Property @{
                AppId = $App.AppId
                DisplayName = $App.DisplayName
                SecretKeyID = $Secret.KeyId
                SecretStartDate = $Secret.StartDate
                SecretEndDate = $Secret.EndDate
                CertificateId = $null
                CertificateStartDate = $null
                CertificateEndDate = $null
                Description = $null
            }
        }
    }

    if($AppCertificates){
        foreach($Certificate in $AppCertificates){
            $AllAppsWithSecrets += New-Object PSObject -Property @{
                AppId = $App.AppId
                DisplayName = $App.DisplayName
                SecretKeyID = $null
                SecretStartDate = $null
                SecretEndDate = $null
                CertificateId = $Certificate.KeyId
                CertificateStartDate = $Certificate.StartDate
                CertificateEndDate = $Certificate.EndDate
                Description = $null
            }
        }
    }
}

# Export to Excel
$SecretsSelector = @(
    "AppId",
    "DisplayName",
    "SecretKeyID",
    "SecretStartDate",
    "SecretEndDate",
    "CertificateId",
    "CertificateStartDate",
    "CertificateEndDate"
)

$AppsListSelector = @(
    "AppId",
    "DisplayName",
    "SecretsCount",
    "CertificatesCount",
    "Description"
)

$AppsListExportSplat = @{
    Path = $OutputFile
    AutoSize = $True
    AutoFilter = $True
    TableStyle = "Medium2"
    WorksheetName = "AppsList"
}

$SecretsExportSplat = @{
    Path = $OutputFile
    AutoSize = $True
    AutoFilter = $True
    TableStyle = "Medium2"
    WorksheetName = "Secrets&Certificates"
    Show = $True
}

if($AppsWithSecretsAndCertificatesCount.count -gt 1){
    $AppsWithSecretsAndCertificatesCount = $AppsWithSecretsAndCertificatesCount | Sort-Object "DisplayName"
    $AppsWithSecretsAndCertificatesCount | Select-Object $AppsListSelector | Export-Excel @AppsListExportSplat
}

if($AllAppsWithSecrets.count -gt 1){
    $AllAppsWithSecrets = $AllAppsWithSecrets | Sort-Object "DisplayName"
    $AllAppsWithSecrets | Select-Object $SecretsSelector | Export-Excel @SecretsExportSplat
}

