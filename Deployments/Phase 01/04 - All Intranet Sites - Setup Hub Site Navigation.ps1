<#
    SCRIPT OVERVIEW:
    This script populates the top-level navigation on every hub site with links to the other hubs, as per the configuration file

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

$sites = $global:sites | Where-Object { $_.SiteType -eq "ALB" -or $_.SiteType -eq "Parent" -and $_.RelativeURL.Length -gt 0 -and $_.ApplyHubSiteNavigationChanges -eq $true } | Sort-Object -Property @{Expression="SiteType";Descending=$true},@{Expression="DisplayName";Descending=$false}

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
        Set-PnPHubSite -Identity $fullURL -HideNameInNavigation:$true
        Write-Host "HUB SITE SETTING UPDATE: '$($hub.Title)' NAME HIUDDEN FROM HUB SITE NAVIGATION IN: $fullURL" -ForegroundColor Yellow
    }
}

Write-Host ""

foreach($site in $sites)
{
    $fullURL = "$global:rootURL/$($site.RelativeURL)"

    SharePointPnPPowerShellOnline\Connect-PnPOnline -Url $fullURL -UseWebLogin
    $navNodes = Get-PnPNavigationNode -Location TopNavigationBar
    $web = Get-PnPWeb

    # We only run this if it's a hub site
    $hub = $hubSites | Where-Object { $_.SiteUrl -eq $fullURL }

    if($null -ne $hub)
    {
        Write-Host "UPDATING THE NAVIGATION WITHIN THE '$($web.Title)' SITE" -ForegroundColor Cyan

        foreach($siteNav in $sites)
        {
            $fullURL = "$global:rootURL/$($siteNav.RelativeURL)"
            $URLMatch = $navNodes.Url -contains "/$($siteNav.RelativeURL)" -or $navNodes.Url -contains $fullURL

            # Does the URL of our site already exist in the hub site navigation top-level?
            if($URLMatch -eq $false)
            {
                $newNode = Add-PnPNavigationNode -Location TopNavigationBar -Title $siteNav.DisplayName -Url $fullURL
                Write-Host "NEW TOP NAVIGATION ADDED. TITLE: $($newNode.Title) URL: $($newNode.Url)" -ForegroundColor Yellow
            }
            else
            {
                foreach($nav in $($navNodes | Where-Object { $_.Url -eq "/$($siteNav.RelativeURL)" -or $_.Url -eq $fullURL}))
                {
                    Write-Host "TOP NAVIGATION NODE EXISTS. TITLE: '$($nav.Title)' URL: $global:rootURL/$($siteNav.RelativeURL)" -ForegroundColor Green
                }
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
Stop-Transcript
