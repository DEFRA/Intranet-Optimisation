<#
    SCRIPT OVERVIEW:
    REGRESSION SCRIPT FOR: 02 - Content Type Hub - Site Columns and Content Types.ps1
    This script removes the "Base Site Page Content Type Hub" content type and the global site columns from SharePoint's Content Type Hub

    SOFTWARE REQUIREMENTS:
    This script was developed on the following platform:
        PowerShell v5.1.22621.1778
        SharePointPnPPowerShellOnline v3.29.2101.0

    SHAREPOINT PERMISSIONS REQUIREMENTS:
    - Site Collection Administrator rights to the "Content Type Hub" site
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

$site = $global:sites | Where-Object { $_.Abbreviation -eq "ContentTypeHub" -and $_.RelativeURL.Length -gt 0 }

if($null -eq $site)
{
    throw "An entry in the configuration could not be found for the 'Content Type Hub' or is not configured correctly"
}

SharePointPnPPowerShellOnline\Connect-PnPOnline -Url "$global:rootURL/$($site.RelativeURL)" -UseWebLogin
Write-Host "SCRIPT EXECUTED BY '$(Get-CurrentUser)' AT $(get-date -f "HH:mm:ss") ON $(get-date -f "dd/MM/yyyy")" -ForegroundColor Cyan
Write-Host "ACCESSING SHAREPOINT SITE: $($global:rootURL)/$($global:site.RelativeURL)" -ForegroundColor Cyan

$ctx = Get-PnPContext

# SCRIPT CONFIG
# Field(s)
$fieldNames = @("OrganisationIntranet")

# Site Page(s)
$ctSPDisplayName = "Site Page"
$library = Get-PnPList -Identity "Site Pages"
$page = Get-PnPListItem -List $library -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($global:placeholderSitePageName)</Value></Eq></Where></Query></View>"

if($null -ne $page)
{
    Remove-PnPListItem -List $library -Identity $page.Id -Force
    Write-Host "SITE PAGE '$($global:placeholderSitePageName)' HAS BEEN DELETED FROM THE '$($library.Title)' LIBRARY" -ForegroundColor Green
}

# Remove the field(s) from "Site Page" content type
$ctSP = Get-PnPContentType -Identity $ctSPDisplayName -List $library -ErrorAction SilentlyContinue

if($null -ne $ctSP)
{
    $ctx.Load($ctSP.FieldLinks)
    $ctx.ExecuteQuery()

    foreach($fieldName in $fieldNames)
    {
        $fieldExistsOnCT = $ctSP.FieldLinks | Where-Object { $_.Name -eq $fieldName }

        if($null -ne $fieldExistsOnCT)
        {
            Remove-PnPField -List $library -Identity $fieldName -Force
            Write-Host "SITE COLUMN '$fieldName' HAS BEEN REMOVED FROM THE '$($library.Title)' LIBRARY CONTENT TYPE '$ctSPDisplayName'" -ForegroundColor Green
        }
        else
        {
            Write-Host "SITE COLUMN '$fieldName' DOES NOT EXIST IN THE '$($library.Title)' LIBRARY" -ForegroundColor Yellow
        }
    }
}

$shell = New-Object -ComObject "WScript.Shell"
$Button = $shell.Popup("Please now follow the deployment guide section 'Regression - Content Type Hub' and once completed, press OK to continue.", 0, "UNPUBLISH THE CONTENT TYPE", 0)

$ct = Get-PnPContentType -Identity $global:ctDisplayName -ErrorAction SilentlyContinue

if($null -ne $ct)
{
    $ctx.Load($ct.FieldLinks)
    $ctx.ExecuteQuery()

    foreach($fieldName in $fieldNames)
    {
        $fieldExistsOnCT = $ct.FieldLinks | Where-Object { $_.Name -eq $fieldName }

        if($null -ne $fieldExistsOnCT)
        {
            Remove-PnPFieldFromContentType -Field $fieldName -ContentType $ct
            Write-Host "SITE COLUMN '$fieldName' REMOVED FROM THE CONTENT TYPE '$($global:ctDisplayName)'" -ForegroundColor Green
        }
    }

    Remove-PnPContentType -Identity $global:ctDisplayName -Force
    Write-Host "CONTENT TYPE REMOVED FROM THE CONTENT TYPE HUB: $($global:ctDisplayName)" -ForegroundColor Green
}

# Remove the field(s) from the "Site Pages" library, "Site Pages" content type and then the site
foreach($fieldName in $fieldNames)
{
    $field = Get-PnPField -Identity $fieldName -ErrorAction SilentlyContinue
    $libField = Get-PnPField -Identity $fieldName -List $library -ErrorAction SilentlyContinue
    $ct = Get-PnPContentType -Identity $ctSPDisplayName -ErrorAction SilentlyContinue

    $ctx.Load($ct.FieldLinks)
    $ctx.ExecuteQuery()

    if($null -ne $libField)
    {
        Remove-PnPField -Identity $fieldName -List $library -Force
        Write-Host "SITE COLUMN REMOVED '$fieldName' REMOVED FROM THE '$($library.Title)' LIBRARY" -ForegroundColor Green
    }

    foreach($ctFieldName in $fieldNames)
    {
        $fieldExistsOnCT = $ct.FieldLinks | Where-Object { $_.Name -eq $ctFieldName }

        if($null -ne $fieldExistsOnCT)
        {
            Remove-PnPFieldFromContentType -Field $ctFieldName -ContentType $ct
            Write-Host "SITE COLUMN '$ctFieldName' REMOVED FROM THE CONTENT TYPE '$ctSPDisplayName'" -ForegroundColor Green
        }
        else
        {
            Write-Host "SITE COLUMN '$ctFieldName' DOES NOT EXIST ON THE CONTENT TYPE '$ctSPDisplayName'" -ForegroundColor Yellow
        }
    }

    if($null -ne $field)
    {
        Remove-PnPField -Identity $fieldName -Force
        Write-Host "SITE COLUMN REMOVED: $fieldName" -ForegroundColor Cyan
    }
}

# Request a reindex of the site
$web = Get-PnPWeb
[Int]$SearchVersion = 0
    
# Get the Search Version Property - If exists
if ($Web.AllProperties.FieldValues.ContainsKey("vti_searchversion") -eq $True)
{
    $SearchVersion = $Web.AllProperties["vti_searchversion"]
}

# Increment Search Version
$SearchVersion++
$Web.AllProperties["vti_searchversion"] = $SearchVersion
$web.Update()
$ctx.ExecuteQuery()

Write-Host "REINDEXING THE 'CONTENT TYPE HUB' SITE ON THE NEXT SCHEDULED SEARCH CRAWL" -ForegroundColor Yellow
Write-Host "SCRIPT FINISHED" -ForegroundColor Yellow
Stop-Transcript
