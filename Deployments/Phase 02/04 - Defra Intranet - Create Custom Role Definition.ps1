<#
    SCRIPT OVERVIEW:
    This script creates a custom role definition which we'll apply to the Site Page library. This role will prevent editors from editing the library views, so they cannot explose our internal fields.

    SOFTWARE REQUIREMENTS:
    This script was developed on the following platform:
        PowerShell v5.1.22621.1778
        SharePointPnPPowerShellOnline v3.29.2101.0

    SHAREPOINT PERMISSIONS REQUIREMENTS:
    - Site Collection Admins rights to the Defra and ALB Intranet SharePoint sites
    OR
    - Access to the SharePoint Tenant Administration site
#>

$ErrorActionPreference="SilentlyContinue"
Stop-Transcript | out-null
$ErrorActionPreference = "Stop"

Import-Module SharePointPnPPowerShellOnline

if ($null -ne $psISE)
{
    $logfileName = $($psISE.CurrentFile.FullPath.Split('\'))[$psISE.CurrentFile.FullPath.Split('\').Count-1]
    $logfileName = $logfileName.Replace(".ps1",".txt")

    $global:scriptPath = Split-Path -Path $psISE.CurrentFile.FullPath

    Import-Module "$global:scriptPath\PSModules\Configuration.psm1" -Force
    Import-Module "$global:scriptPath\PSModules\Helper.psm1" -Force
}
else
{
    Clear-Host

    $logFileName = $MyInvocation.MyCommand.Name
    $global:scriptPath = "." 

    Import-Module "./PSModules/Configuration.psm1" -Force
    Import-Module "./PSModules/Helper.psm1" -Force
}

$logfileName = $logfileName.Replace(".ps1",".txt")
Start-Transcript -path "$global:scriptPath/Logs/$logfileName" -append | Out-Null

Invoke-Configuration

$site = $global:sites | Where-Object { $_.SiteType -eq "Parent" -and $_.RelativeURL.Length -gt 0 } | Sort-Object -Property @{Expression="SiteType";Descending=$true},@{Expression="DisplayName";Descending=$false}

if($null -eq $sites)
{
    throw "An entry in the configuration could not be found for the 'Defra Intranet' or is not configured correctly"
}

Connect-PnPOnline -Url "$global:rootURL/$($site.RelativeURL)" -UseWebLogin
Write-Host "SCRIPT EXECUTED BY '$(Get-CurrentUser)' AT $(get-date -f "HH:mm:ss") ON $(get-date -f "dd/MM/yyyy")" -ForegroundColor Cyan
Write-Host "ACCESSING SHAREPOINT SITE: $($global:rootURL)/$($global:site.RelativeURL)" -ForegroundColor Cyan

# Create a custom permission level
$roleName = "Custom Permission - Contribute - For Site Page Library Only"
$role = Get-PnPRoleDefinition -Identity $roleName -ErrorAction SilentlyContinue

if($null -eq $role)
{
    $BasePermissionLevel = Get-PnPRoleDefinition -Identity "Contribute"
 
    # Set Parameters for new permission level
    $NewPermissionLevel= @{
        Exclude     = 'ManagePersonalViews','AddDelPrivateWebParts','UpdatePersonalWebParts'
        Description = "Can view, add, update, and delete list items and documents, but cannot create or edit list views."
        RoleName    = $roleName
        Clone       = $BasePermissionLevel
    }
 
    # Create new permission level
    Add-PnPRoleDefinition @NewPermissionLevel
    Write-Host "`nNEW ROLE DEFINTION '$roleName' CREATED" -ForegroundColor Green
}
else
{
    Write-Host "`nNEW ROLE DEFINTION '$roleName' ALREADY EXISTS" -ForegroundColor Yellow
}

Write-Host ""

Write-Host "SCRIPT FINISHED" -ForegroundColor Yellow
Stop-Transcript