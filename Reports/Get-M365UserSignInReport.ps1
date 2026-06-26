<#
.SYNOPSIS
Exports Microsoft 365 user sign-in activity.

.DESCRIPTION
Retrieves Microsoft 365 user sign-in activity from Microsoft Graph and exports the results to a CSV report.

.AUTHOR
Rahul Namdev

.VERSION
1.0.0

.REQUIREMENTS
Microsoft Graph PowerShell SDK

.NOTES
Repository:
https://github.com/Rahulnamdev365/Microsoft365-AdminToolkit
#>

[CmdletBinding()]
param (
    [int]$InactiveDays,
    [switch]$EnabledUsersOnly,
    [switch]$DisabledUsersOnly,
    [switch]$LicensedUsersOnly,
    [switch]$ExternalUsersOnly,
    [switch]$CreateSession,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)

Set-StrictMode -Version Latest

$ErrorActionPreference = 'Stop'

$ScriptVersion = '1.0.0'

function Connect-M365Graph 
{
    # Check for module installation
    $Module = Get-Module -Name Microsoft.Graph -ListAvailable
    if ($Module.Count -eq 0) {
        Write-Host "Microsoft Graph PowerShell SDK is not available" -ForegroundColor Yellow
        $Confirm = Read-Host "Are you sure you want to install module? [Y] Yes [N] No"
        if ($Confirm -match "[yY]") {
            Write-Host "Installing Microsoft Graph PowerShell module..."
            Install-Module Microsoft.Graph -Repository PSGallery -Scope CurrentUser -AllowClobber -Force
        } else {
            Write-Host "Microsoft Graph PowerShell module is required to run this script. Please install it using Install-Module Microsoft.Graph"
            return
        }
    }

    # Disconnect existing session if requested
    if ($CreateSession.IsPresent) {
        Disconnect-MgGraph
    }

    Write-Host "Connecting to Microsoft Graph..."
    if (-not [string]::IsNullOrEmpty($TenantId) -and -not [string]::IsNullOrEmpty($ClientId) -and -not [string]::IsNullOrEmpty($CertificateThumbprint)) {
        Connect-MgGraph -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
    } else {
        Connect-MgGraph -Scopes "User.Read.All", "AuditLog.Read.All" -NoWelcome
    }
}

Connect_MgGraph

$ExportCSV = ".\M365_UserSignInReport_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
$PrintedUser = 0

# Load Microsoft 365 license friendly names
$ResourcePath = Join-Path $PSScriptRoot "..\Resources\LicenseFriendlyName.txt"

try
{
    $FriendlyNameHash = Get-Content -Raw -Path $ResourcePath -ErrorAction Stop | ConvertFrom-StringData
}
catch
{
    Write-Error "Required resource file 'LicenseFriendlyName.txt' was not found in the Resources folder."
    return
}

$Count = 0
$RequiredProperties = @('UserPrincipalName', 'EmployeeId', 'CreatedDateTime', 'AccountEnabled', 'Department', 'JobTitle', 'SignInActivity')

Get-MgUser -All -Property $RequiredProperties | Select-Object -Property $RequiredProperties | ForEach-Object {
    $Count++
    $UPN = $_.UserPrincipalName
    Write-Progress -Activity "`n     Processing user: $Count - $UPN"

    $EmployeeId = $_.EmployeeId
    $LastSuccessfulSigninDate = $_.SignInActivity.LastSuccessfulSignInDateTime
    $LastInteractiveSignIn = $_.SignInActivity.LastSignInDateTime
    $LastNonInteractiveSignIn = $_.SignInActivity.LastNonInteractiveSignInDateTime
    $CreatedDate = $_.CreatedDateTime
    $AccountEnabled = $_.AccountEnabled
    $Department = $_.Department
    $JobTitle = $_.JobTitle

    # Calculate inactive days
    if ($LastSuccessfulSigninDate -eq $null) {
        $LastSuccessfulSigninDate = "Data not available"
        $InactiveUserDays = "-"
    } else {
        $InactiveUserDays = (New-TimeSpan -Start $LastSuccessfulSigninDate).Days
    }

    $AccountStatus = if ($AccountEnabled) { 'Enabled' } else { 'Disabled' }

    # Get licenses
    $Licenses = (Get-MgUserLicenseDetail -UserId $UPN).SkuPartNumber
    $AssignedLicense = @()

    if ($Licenses.Count -eq 0) {
        $LicenseDetails = "No License Assigned"
    } else {
        foreach ($License in $Licenses) {
            $EasyName = $FriendlyNameHash[$License]
            $AssignedLicense += if ($EasyName) { $EasyName } else { $License }
        }
        $LicenseDetails = $AssignedLicense -join ", "
    }

    $Print = 1

    # Apply filters
    if ($InactiveUserDays -ne "-" -and $InactiveDays -ne $null -and $InactiveDays -gt $InactiveUserDays) {
        $Print = 0
    }

    if ($ExternalUsersOnly.IsPresent -and $UPN -notmatch '#EXT#') {
        $Print = 0
    }

    if ($EnabledUsersOnly.IsPresent -and $AccountStatus -eq 'Disabled') {
        $Print = 0
    }

    if ($DisabledUsersOnly.IsPresent -and $AccountStatus -eq 'Enabled') {
        $Print = 0
    }

    if ($LicensedUsersOnly.IsPresent -and $Licenses.Count -eq 0) {
        $Print = 0
    }

    # Export to CSV
    if ($Print -eq 1) {
        $PrintedUser++
        $ExportResult = [PSCustomObject]@{
            'UPN'                          = $UPN
            'Creation Date'               = $CreatedDate
            'Last Successful Signin Date' = $LastSuccessfulSigninDate
            'Inactive Days'               = $InactiveUserDays
            'Last Interactive SignIn Date' = $LastInteractiveSignIn
            'Last Non Interactive SignIn Date' = $LastNonInteractiveSignIn
            'Emp id'                      = $EmployeeId
            'License Details'             = $LicenseDetails
            'Account Status'              = $AccountStatus
            'Department'                  = $Department
            'Job Title'                   = $JobTitle
        }
        $ExportResult | Export-Csv -Path $ExportCSV -NoTypeInformation -Append
    }
}

# Final output
Write-Host "`nReport generated successfully." -ForegroundColor Green
Write-Host "Exported users : $PrintedUser"
Write-Host "Report location: $ExportCSV"

if (Test-Path -Path $ExportCSV)
{
    $Prompt = New-Object -ComObject WScript.Shell
    $UserInput = $Prompt.Popup(
        "Do you want to open the generated report?",
        0,
        "Microsoft365-AdminToolkit",
        4
    )

    if ($UserInput -eq 6)
    {
        Invoke-Item $ExportCSV
    }
}
else
{
    Write-Host "No users matched the selected filters." -ForegroundColor Yellow
}
