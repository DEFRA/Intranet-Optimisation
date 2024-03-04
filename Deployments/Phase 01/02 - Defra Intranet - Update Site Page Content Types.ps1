 <#
    SCRIPT OVERVIEW:
    This script adds the new "Organisation" column, coming from the Content Type Hub, into the Site Pages library and it's views

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

$SPCtx = Get-PnPContext

$SPCtDisplayName = "Site Page"
$libraryTitle = "Site Pages"

# Field(s)
$fields = @(
    [PSCustomObject]@{
        'DisplayName' = "Organisation (Intranets)"
        'InternalName' = "OrganisationIntranets"
    }
)

foreach($objField in $fields)
{
    # Add the field to "Site Page" content type
    $SPCt = Get-PnPContentType -Identity $SPCtDisplayName -ErrorAction SilentlyContinue
    $field = Get-PnPField -Identity $objField.InternalName -ErrorAction SilentlyContinue
    $list = Get-PnPList -Identity $libraryTitle -ErrorAction SilentlyContinue

    if($null -ne $field -and ($null -ne $SPCt -and $SPCt.Count -eq 1))
    {
        $SPCtx.Load($SPCt.FieldLinks)
        $SPCtx.ExecuteQuery()

        $fieldExistsOnCT = $SPCt.FieldLinks | Where-Object { $_.Name -eq $objField.InternalName }

        if($null -eq $fieldExistsOnCT)
        {
            Add-PnPFieldToContentType -Field $field -ContentType $SPCt -ErrorAction SilentlyContinue
            Write-Host "SITE COLUMN '$($objField.InternalName)' ADDED TO THE CONTENT TYPE '$SPCtDisplayName'" -ForegroundColor Green
        }
        else 
        {
            Write-Host "SITE COLUMN '$($objField.InternalName)' IS ALREADY INSTALLED IN THE CONTENT TYPE '$SPCtDisplayName'" -ForegroundColor Yellow
        }

        $field = Get-PnPField -List $list -Identity $objField.InternalName -ErrorAction SilentlyContinue
    }
    else
    {
        Write-Host "THE SITE COLUMN '$($objField.InternalName)' WAS NOT ADDED TO THE CONTENT TYPE '$SPCtDisplayName'. THE FIELD COULD NOT BE FOUND" -ForegroundColor Yellow
        
        $ct = Get-PnPContentType -Identity $global:ctDisplayName

        if($null -ne $list -and $null -ne $ct)
        {
            $ct = Get-PnPContentType -Identity $global:ctDisplayName
            Add-PnPContentTypeToList -List $list -ContentType $ct
        }
    }

    if($null -eq $field)
    {
        throw "SCRIPT ERROR: DEPLOYMENT OF THE SITE COLUMN '$($objField.InternalName) HAS FAILED. PLEASE ENSURE THE CONTENT TYPE '$($global:ctDisplayName)' HAS BEEN PUBLISHED TO THE SITES FROM THE CONTENT TYPE HUB AND THE COLUMN '$($objField.InternalName)' IS INCLUDED WITHIN THE '$($global:ctDisplayName)'. PLEASE REFER TO THE DEPLOYMENT GUIDE FOR ADDITIONAL SUPPORT."
    }

    $field = Get-PnPField -List $list -Identity $objField.InternalName -ErrorAction SilentlyContinue

    if($null -ne $list -and $null -eq $field)
    {
        Add-PnPField -List $list -Field $objField.InternalName
        Write-Host "SITE COLUMN '$($objField.InternalName)' ADDED TO THE LIBRARY '$libraryTitle'" -ForegroundColor Green

        $field = Get-PnPField -List $list -Identity $objField.InternalName
    }

    if($null -ne $list -and $null -ne $field)
    {
        # Update the list's default view with our new fields
        $views = Get-PnPView -List $list | Where-Object { $_.Title -ne "" }

        foreach($view in $views)
        {
            $SPCtx.Load($view.ViewFields)
            $SPCtx.ExecuteQuery()

            $fieldExists = $view.ViewFields | Where-Object { $_ -eq $objField.InternalName }

            if($null -eq $fieldExists)
            {
                $fieldNames = New-Object Collections.Generic.List[String]

                foreach($viewField in $view.ViewFields)
                {
                    $fieldNames.Add($viewField)
                }

                $fieldNames.Add($objField.InternalName)

                $view = Set-PnPView -List $list -Identity $view.Title -Fields $fieldNames
                Write-Host "THE FIELD '$($objField.InternalName)' HAS BEEN ADDED TO THE '$libraryTitle' LIBRARY VIEW '$($view.Title)'" -ForegroundColor Green 
            }
            else
            {
                Write-Host "THE FIELD '$($objField.InternalName)' HAS ALREADY BEEN ADDED TO THE VIEW '$($view.Title)'" -ForegroundColor Yellow
            }
        }
    }
}

Write-Host "SCRIPT FINISHED" -ForegroundColor Yellow
Stop-Transcript
