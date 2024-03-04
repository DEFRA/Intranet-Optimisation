<#
    SCRIPT OVERVIEW:
    This script removes the custom column(s) from the Site Page library

    SOFTWARE REQUIREMENTS:
    This script was developed on the following platform:
        PowerShell v5.1.22621.1778
        SharePointPnPPowerShellOnline v3.29.2101.0

    SHAREPOINT PERMISSIONS REQUIREMENTS:
    - Site Collection Admins rights to the DEFRA Intranet SharePoint site
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

$site = $global:sites | Where-Object { $_.Abbreviation -eq "Defra" -and $_.RelativeURL.Length -gt 0 }

if($null -eq $site)
{
    throw "An entry in the configuration could not be found for the 'Defra Intranet' or is not configured correctly"
}

SharePointPnPPowerShellOnline\Connect-PnPOnline -Url "$global:rootURL/$($site.RelativeURL)" -UseWebLogin
Write-Host "SCRIPT EXECUTED BY '$(Get-CurrentUser)' AT $(get-date -f "HH:mm:ss") ON $(get-date -f "dd/MM/yyyy")" -ForegroundColor Cyan
Write-Host "ACCESSING SHAREPOINT SITE: $($global:rootURL)/$($global:site.RelativeURL)" -ForegroundColor Cyan

$ctx = Get-PnPContext

$ctDisplayName = "Site Page"
$displayName = "Site Pages"
$fieldNames = @("OrganisationIntranets")

# Remove the field(s) from "Site Page" content type
$ct = Get-PnPContentType -Identity $ctDisplayName -ErrorAction SilentlyContinue

if($null -ne $ct -and $ct.Count -eq 1)
{
    $ctx.Load($ct.FieldLinks)
    $ctx.ExecuteQuery()

    foreach($fieldName in $fieldNames)
    {
        $fieldExistsOnCT = $ct.FieldLinks | Where-Object { $_.Name -eq $fieldName }

        if($null -ne $fieldExistsOnCT)
        {
            Remove-PnPFieldFromContentType -Field $fieldName -ContentType $ct -ErrorAction SilentlyContinue
            Write-Host "SITE COLUMN '$fieldName' REMOVED FROM THE CONTENT TYPE '$ctDisplayName'" -ForegroundColor Green
        }
    }
}
else
{
    Write-Host "THE CONTENT TYPE '$ctDisplayname' COULD NOT BE FOUND OR TOO MANY RESULTS WERE RETURNED" -ForegroundColor Red
}

$list = Get-PnPList -Identity $displayName

if($null -ne $list)
{
    foreach($fieldName in $fieldNames)
    {
        $removedField = Remove-PnPField -List $list -Identity $fieldName -Force -ErrorAction SilentlyContinue

        if($null -ne $removedField)
        {
            Write-Host "THE FIELD '$fieldName' HAS BEEN REMOVED FROM THE '$displayName' LIBRARY" -ForegroundColor Yellow
        }
        else
        {
            Write-Host "THE FIELD '$fieldName' DOES NOT EXIST IN THE '$displayName' LIBRARY" -ForegroundColor Yellow
        }
    }
}

Write-Host "SCRIPT FINISHED" -ForegroundColor Yellow
Stop-Transcript
