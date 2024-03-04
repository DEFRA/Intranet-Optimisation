<#
    SCRIPT OVERVIEW:
    This script registers each of the configured sites as hub sites and sets the ALB's parent association to point at the Defra Intranet

    SOFTWARE REQUIREMENTS:
    This script was developed on the following platform:
        PowerShell v5.1.22621.1778
        SharePointPnPPowerShellOnline v3.29.2101.0

    SHAREPOINT PERMISSIONS REQUIREMENTS:
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

$sites = $global:sites | Where-Object { $_.SiteType -eq "ALB" -or $_.SiteType -eq "Parent" -and $_.RelativeURL.Length -gt 0 -and $_.ApplyHubSiteNavigationChanges -eq $true } | Sort-Object -Property @{Expression="SiteType";Descending=$false},@{Expression="DisplayName";Descending=$false}

if($sites.Count -gt 0)
{
    SharePointPnPPowerShellOnline\Connect-PnPOnline -Url $global:adminURL -UseWebLogin
    Write-Host "SCRIPT EXECUTED BY '$(Get-CurrentUser)' AT $(get-date -f "HH:mm:ss") ON $(get-date -f "dd/MM/yyyy")" -ForegroundColor Cyan
    Write-Host "ACCESSING SHAREPOINT SITE: $global:adminURL" -ForegroundColor Cyan
}

foreach($site in $sites)
{
    $fullURL = "$global:rootURL/$($site.RelativeURL)"
    $IsHubSite = $(Get-PnPHubSite | Where-Object { $_.SiteUrl -eq $fullURL }).Count

    if($IsHubSite -eq 0)
    {
        Register-PnPHubSite -Site $fullURL
        Write-Host "The site '$($site.DisplayName)' has been registered as a hub" -ForegroundColor Green
    }
    else
    {
        Write-Host "The site '$($site.DisplayName)' is already a hub site" -ForegroundColor Yellow
    }
}

Write-Host "SCRIPT FINISHED" -ForegroundColor Yellow
Stop-Transcript
