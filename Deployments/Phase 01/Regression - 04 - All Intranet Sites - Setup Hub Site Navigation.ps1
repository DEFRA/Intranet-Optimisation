<#
    SCRIPT OVERVIEW:
    This script will remove all of the top-level navigation links matching the URLs held within the script configuration.

    SOFTWARE REQUIREMENTS:
    This script was developed on the following platform:
        PowerShell v5.1.22621.1778
        SharePointPnPPowerShellOnline v3.29.2101.0

    SHAREPOINT PERMISSIONS REQUIREMENTS:
    - Access to the SharePoint Tenant Administration site
#>

Import-Module SharePointPnPPowerShellOnline

$ErrorActionPreference = "Stop"

$script:results = @()
$scriptPath = $global:PSScriptRoot

Import-Module SharePointPnPPowerShellOnline

if($null -ne $psISE)
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

Invoke-Configuration

$sites = $global:sites | Where-Object { $_.SiteType -eq "ALB" -or $_.SiteType -eq "Parent" -and $_.RelativeURL.Length -gt 0 -and $_.ApplyHubSiteNavigationChanges -eq $true }

if($sites.Count -gt 0)
{
    SharePointPnPPowerShellOnline\Connect-PnPOnline -Url $global:adminURL -UseWebLogin
    Write-Host "SCRIPT EXECUTED BY '$(Get-CurrentUser)' AT $(get-date -f "HH:mm:ss") ON $(get-date -f "dd/MM/yyyy")" -ForegroundColor Cyan
    Write-Host "ACCESSING SHAREPOINT SITE: $global:adminURL" -ForegroundColor Cyan

    $hubSites = Get-PnPHubSite

    Write-Host ""
}

foreach($site in $sites)
{
    $fullURL = "$global:rootURL/$($site.RelativeURL)"

    # We only run this if it's a hub site
    $hub = $hubSites | Where-Object { $_.SiteUrl -eq $fullURL }

    if($null -ne $hub)
    {
        Set-PnPHubSite -Identity $fullURL -HideNameInNavigation:$false
        Write-Host "HUB SITE SETTING UPDATE: '$($hub.Title)' NAME SWITCHED BACK ON IN HUB SITE NAVIGATION IN: $fullURL" -ForegroundColor Yellow
    }
}

Write-Host ""

foreach($site in $sites)
{
    $fullURL = "$global:rootURL/$($site.RelativeURL)"

    SharePointPnPPowerShellOnline\Connect-PnPOnline -Url $fullURL -UseWebLogin
    $web = Get-PnPWeb

    # We only run this if it's a hub site
    $hub = $hubSites | Where-Object { $_.SiteUrl -eq $fullURL }

    if($null -ne $hub)
    {
        Write-Host "REVERTING THE NAVIGATION UPDATES WITHIN THE '$($web.Title)' SITE" -ForegroundColor Cyan

        # This site is a hub site?
        foreach($siteNav in $sites)
        {
            $fullURL = "$global:rootURL/$($siteNav.RelativeURL)"
            $navNodes = Get-PnPNavigationNode -Location TopNavigationBar | Where-Object { $_.Url -eq "/$($siteNav.RelativeURL)" -or $_.Url -contains $fullURL }
    
            foreach($navNode in $navNodes)
            {
                Remove-PnPNavigationNode -Identity $navNode.Id -Force
                Write-Host "NAVIGATION NODE '$($navNode.Title)' WITH URL '/$($siteNav.RelativeURL)' REMOVED FROM THE SITE '$($web.Title)'" -ForegroundColor Yellow
            }
        }

        Write-Host ""
    }
    else
    {
        Write-Host "THE SITE '$($web.Title)' IS NOT A HUB SITE" -ForegroundColor Red
    }
}

Write-Host "SCRIPT FINISHED" -ForegroundColor Yellow
