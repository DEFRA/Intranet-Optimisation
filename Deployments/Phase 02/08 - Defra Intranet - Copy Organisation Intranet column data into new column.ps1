<#
    SCRIPT OVERVIEW:
    This script updates the Defra Intranet Site Pages library with our custom columns for the approval system

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

$ctx = Get-PnPContext
$list = Get-PnPList -Identity "Site Pages"
$listItems = Get-PnPListItem -List $list -PageSize 1000 -Fields "File","OrganisationIntranets","OrganisationIntranetsContentEditorInput","Title","_UIVersionString"| Where-Object {$_["OrganisationIntranets"].Count -gt 0}

foreach($item in $listItems)
{  
    $termValues = @{}
    $currentVersionNumber = $item["_UIVersionString"]

    $terms = $(Get-PnPTerm -TermSet "Organisational Unit" -TermGroup "DEFRA EDRM UAT" -Identity "Defra Orgs" -IncludeChildTerms).Terms
    
    foreach($term in $item["OrganisationIntranets"])
    {
        $termStoreTerm = $terms | Where-Object {$_.Name -eq $term.Label.Trim()}
        $termValues.Add($termStoreTerm.Id.ToString(),$term.Label)
    }

    Set-PnpTaxonomyFieldValue -ListItem $item -InternalFieldName "OrganisationIntranetsContentEditorInput" -Terms $termValues

    Write-Host "SITE PAGE '$($item["Title"])' COLUMN 'OrganisationIntranetsContentEditorInput' VALUE UPDATED TO BE: $(($termValues.GetEnumerator() | % { $($_.Value) }) -join ',')" -ForegroundColor Green

    # Only publish the page again if it was a major version, otherwise a publish would release a user's draft
    if($currentVersionNumber.Split('.')[1] -eq 0)
    {
        $item.File.Publish("Published")
        Invoke-PnPQuery
        Write-Host "PAGE REPUBLISHED" -ForegroundColor Yellow
    }
    else
    {
        Write-Host "THE PAGE WAS IN DRAFT BEFORE THIS UPDATE, SO SKIPPED REPUBLISHING" -ForegroundColor Cyan
    }
}

Write-Host "SCRIPT FINISHED" -ForegroundColor Yellow
Stop-Transcript