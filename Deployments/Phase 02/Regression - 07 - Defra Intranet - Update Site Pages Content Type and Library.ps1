<#
    SCRIPT OVERVIEW:
    This script uninstalls the Defra Intranet Site Pages customisations we setup for site page approval system

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
$fieldNames = @("OrganisationIntranetsContentEditorInput","PageApprovalInfo","WorkflowArticleSentForALBApproval","WorkflowPublishtoALBIntranets","WorkflowApprovalProgress")

$list = Get-PnPList -Identity $displayName

# Reverse the list settings
Write-Host "REVERSING '$displayName' LIBRARY SETTINGS" -ForegroundColor Green
$list.DisableGridEditing = $false
$list.Update()
Invoke-PnPQuery

Set-PnPList -Identity $displayName -ResetRoleInheritance

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

# SITE-SPECIFIC FIELDS
$fieldInternalName = "OrganisationIntranets"
$field = Get-PnPField -Identity $fieldInternalName -ErrorAction SilentlyContinue

if($null -ne $field)
{
    Set-PnPField -Identity $field.Id -Values @{
        Title = "Organisation (Intranets)"
    }

    Write-Host "SITE COLUMN '$fieldInternalName' RESET" -ForegroundColor Green
}

$field = Get-PnPField -Identity $fieldInternalName -List $list -ErrorAction SilentlyContinue

if($null -ne $field)
{
    Set-PnPField -List $list -Identity $field.Id -Values @{
        Title = "Organisation (Intranets)"
        Hidden = $false
    }

    Write-Host "FIELD '$fieldInternalName' RESET" -ForegroundColor Green
}
else
{
    Write-Host "THE FIELD '$fieldInternalName' DOES NOT EXIST IN THE LIBRARY '$displayName'" -ForegroundColor Red
}

# RESET VIEWS
if($null -ne $list -and $null -ne $field)
{
    $views = Get-PnPView -List $list | Where-Object { $_.Title -ne "" }

    foreach($view in $views)
    {
        $ctx.Load($view.ViewFields)
        $ctx.ExecuteQuery()

        $fieldExists = $view.ViewFields | Where-Object { $_ -eq $fieldInternalName }

        if($null -eq $fieldExists)
        {
            $fieldNames = New-Object Collections.Generic.List[String]

            foreach($viewField in $view.ViewFields)
            {
                $fieldNames.Add($viewField)
            }

            $fieldNames.Add($fieldInternalName)

            $view = Set-PnPView -List $list -Identity $view.Title -Fields $fieldNames
            Write-Host "THE FIELD '$($fieldInternalName)' HAS BEEN ADDED BACK INTO THE '$displayName' LIBRARY VIEW '$($view.Title)'" -ForegroundColor Green 
        }
        else
        {
            Write-Host "THE FIELD '$($fieldInternalName)' ALREADY EXISTS ON THE VIEW '$($view.Title)' OF THE '$displayName' LIBRARY " -ForegroundColor Yellow
        }
    }
}

Write-Host "SCRIPT FINISHED" -ForegroundColor Yellow
Stop-Transcript