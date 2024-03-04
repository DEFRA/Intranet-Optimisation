<#
    SCRIPT OVERVIEW:
    This script uninstalls our custom lists

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

$sites = $global:sites | Where-Object { $_.SiteType -eq "ALB" -or $_.SiteType -eq "Parent" -and $_.RelativeURL.Length -gt 0 } | Sort-Object -Property @{Expression="SiteType";Descending=$true},@{Expression="DisplayName";Descending=$false}

if($null -eq $sites)
{
    throw "Entries could not be found in the configuration module that matches the requirements for this script to run. The Defra Intranet and all associated ALB intranets are required."
}

foreach($site in $sites)
{
    Connect-PnPOnline -Url "$global:rootURL/$($site.RelativeURL)" -UseWebLogin
    Write-Host "SCRIPT EXECUTED BY '$(Get-CurrentUser)' AT $(get-date -f "HH:mm:ss") ON $(get-date -f "dd/MM/yyyy")" -ForegroundColor Cyan
    Write-Host "ACCESSING SHAREPOINT SITE: $($global:rootURL)/$($global:site.RelativeURL)" -ForegroundColor Cyan
    Write-Host ""

    $displayName = "Internal Comms Intranet Content Submissions"
    $listURL = "Lists/ICICS"

    $web = Get-PnPWeb
    $list = Get-PnPList -Identity $listURL


    if($null -ne $list)
    {
        Remove-PnPList -Identity $listURL -Force
        Write-Host "THE '$displayName' LIST HAS BEEN REMOVED" -ForegroundColor Yellow
    }
    else
    {
        Write-Host "THE '$displayName' LIST DOES NOT EXIST IN THE SITE '$($web.Title)'" -ForegroundColor Green
    }

    Write-Host ""
}

Write-Host "SCRIPT FINISHED" -ForegroundColor Yellow
Stop-Transcript