<#
    SCRIPT OVERVIEW:
    REGRESSION SCRIPT FOR: 03 - DEFRA Intranet - Site Columns.ps1
    This script uninstalls the site columns required by our custom list(s) and libraries  

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

Write-Host "SCRIPT EXECUTED BY '$(Get-CurrentUser)' AT $(get-date -f "HH:mm:ss") ON $(get-date -f "dd/MM/yyyy")" -ForegroundColor Cyan
Write-Host ""

$site = $sites | Where-Object { $_.SiteType -eq "Parent" -and $_.RelativeURL.Length -gt 0 } | Sort-Object -Property @{Expression="SiteType";Descending=$true},@{Expression="DisplayName";Descending=$false}

foreach($site in $sites)
{
    Connect-PnPOnline -Url "$global:rootURL/$($site.RelativeURL)" -UseWebLogin
    Write-Host "ACCESSING SHAREPOINT SITE: $($global:rootURL)/$($global:site.RelativeURL)" -ForegroundColor Green

    $fieldNames = @("AltContact","ContentTypes","LineManager","EventLink","PublishBy","ContentSubmissionApprovalOptions","ContentSubmissionStatus","ContentSubmissionDescription","EventDateTime","EventEndDateTime","EventVenueAndJoiningDetails","EventDetails","ContentSubmissionApproveRejectBy")
    $web = Get-PnPWeb

    foreach($fieldName in $fieldNames)
    {
        $field = Get-PnPField -Identity $fieldName -ErrorAction SilentlyContinue

        if($null -ne $field)
        {
            Remove-PnPField -Identity $fieldName -Force
            Write-Host "SITE COLUMN REMOVED: $fieldName" -ForegroundColor Yellow
        }
        else
        {
            Write-Host "THE FIELD '$fieldName' DOES NOT EXIST IN THE SITE '$($web.Title)'" -ForegroundColor Cyan
        }
    }

    Write-Host ""
}

Write-Host "SCRIPT FINISHED" -ForegroundColor Yellow
Stop-Transcript