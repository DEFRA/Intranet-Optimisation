<#
    SCRIPT OVERVIEW:
    This script creates the site columns used within the lists and libraries of the Defra Intranet site.

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

Connect-PnPOnline -Url "$global:rootURL/$($site.RelativeURL)" -UseWebLogin
Write-Host "SCRIPT EXECUTED BY '$(Get-CurrentUser)' AT $(get-date -f "HH:mm:ss") ON $(get-date -f "dd/MM/yyyy")" -ForegroundColor Cyan
Write-Host "ACCESSING SHAREPOINT SITE: $($global:rootURL)/$($global:site.RelativeURL)" -ForegroundColor Cyan
Write-Host ""

# SITE PAGE COLUMNS 
# "Organisation (Intranets)" column
$displayName = "Organisation (Intranets)"
$field = Get-PnPField | Where-Object { $_.InternalName -eq "OrganisationIntranetsContentEditorInput" }
$termSetPath = $global:termSetPath

$web = Get-PnPWeb

if($null -eq $field)
{
    $field = Add-PnPTaxonomyField -DisplayName $displayName -InternalName "OrganisationIntranetsContentEditorInput" -TermSetPath $termSetPath -MultiValue
    Write-Host "SITE COLUMN INSTALLED: $($displayName)" -ForegroundColor Green
}
else
{
    Write-Host "SITE COLUMN ALREADY INSTALLED: $($displayName)" -ForegroundColor Yellow
}

# "Workflow - Publish to ALB Intranets" column
$displayName = "Workflow - Publish to ALB Intranets"
$field = Get-PnPField -Identity "WorkflowPublishtoALBIntranets" -ErrorAction SilentlyContinue

if($null -eq $field)
{
    $field = Add-PnPField -Type Text -InternalName "WorkflowPublishtoALBIntranets" -DisplayName $displayName
    Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
}
else
{
    Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow        
}

# "Article Sent for ALB Approval" column
$displayName = "Article Sent for ALB Approval"
$field = Get-PnPField -Identity "WorkflowArticleSentForALBApproval" -ErrorAction SilentlyContinue

if($null -eq $field)
{
    $field = Add-PnPField -Type Boolean -InternalName "WorkflowArticleSentForALBApproval" -DisplayName $displayName
    Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
}
else
{
    Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow      
}

# "Approval Workflow Progress" column
$displayName = "Approval Workflow Progress"
$field = Get-PnPField -Identity "WorkflowApprovalProgress" -ErrorAction SilentlyContinue

if($null -eq $field)
{
    $field = Add-PnPField -Type Text -InternalName "WorkflowApprovalProgress" -DisplayName $displayName
    Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
}
else
{
    Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow
}

# "Approval Information" column
$displayName = "Approval Information"
$field = Get-PnPField | Where-Object { $_.InternalName -eq "PageApprovalInfo" }

if($null -eq $field)
{
    $field = Add-PnPField -DisplayName $displayName -InternalName "PageApprovalInfo" -Type URL
    Write-Host "SITE COLUMN INSTALLED: $($displayName)" -ForegroundColor Green
}
else
{
    Write-Host "SITE COLUMN ALREADY INSTALLED: $($displayName)" -ForegroundColor Yellow
}

# "Site Page ID" column
$displayName = "Site Page ID"
$field = Get-PnPField -Identity "AssociatedSitePage" -ErrorAction SilentlyContinue

if($null -eq $field)
{
    $parentList = Get-PnPList -Identity "Site Pages"
    $fieldXML = "<Field Type='Lookup' DisplayName='$($displayName)' Required='FALSE' EnforceUniqueValues='FALSE' List='{$($parentList.Id)}' WebId='$($web.Id)' ShowField='ID' UnlimitedLengthInDocumentLibrary='FALSE' Group='Custom Columns' ID='{6cabf092-f7a8-4ecf-8acd-5b87238a38a6}' StaticName='AssociatedSitePage' Name='AssociatedSitePage'/>"
    $field = Add-PnPFieldFromXml -FieldXml $fieldXML
    Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
}
else 
{
    Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow        
}

# "Article Title" column
$displayName = "Article Title"
$field = Get-PnPField -Identity "NewsArticleTitle" -ErrorAction SilentlyContinue

if($null -eq $field)
{
    $parentList = Get-PnPList -Identity "Site Pages"
    $parentField = Get-PnPField -Identity "AssociatedSitePage" 

    $fieldXML = '<Field Type="Lookup" DisplayName="News Article Title" Name="NewsArticleTitle" ShowField="Title" EnforceUniqueValues="FALSE" Required="FALSE" Hidden="FALSE" ReadOnly="TRUE" CanToggleHidden="FALSE"  ID="' + [guid]::NewGuid().Guid + '" UnlimitedLengthInDocumentLibrary="FALSE" FieldRef="' + $parentField.Id + '" List="' + $parentList.Id + '" />'
    $field = Add-PnPFieldFromXml -FieldXml $fieldXML
    Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
}
else 
{
    Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow        
}

# "Date of Approval Request" column
$displayName = "Date of Approval Request"
$field = Get-PnPField -Identity "DateOfApprovalRequest" -ErrorAction SilentlyContinue

if($null -eq $field)
{
    $field = Add-PnPField -Type "DateTime" -InternalName "DateOfApprovalRequest" -DisplayName $displayName

    Set-PnPField -Identity $field.Id -Values @{
        FriendlyDisplayFormat = [Microsoft.SharePoint.Client.DateTimeFieldFriendlyFormatType]::Disabled;
    }

    Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
}
else
{
    Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow        
}

# "Version Number" column
$displayName = "Version Number"
$field = Get-PnPField -Identity "SPVersionNumber" -ErrorAction SilentlyContinue

if($null -eq $field)
{
    $field = Add-PnPField -Type Number -InternalName "SPVersionNumber" -DisplayName $displayName
    Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
}
else
{
    Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow        
}

# "Approved/Rejected On" column
$displayName = "Approved/Rejected On"
$field = Get-PnPField -Identity "DateTimeALBApprovalDecision" -ErrorAction SilentlyContinue

if($null -eq $field)
{
    $field = Add-PnPField -Type "DateTime" -InternalName "DateTimeALBApprovalDecision" -DisplayName $displayName

    Set-PnPField -Identity $field.Id -Values @{
        FriendlyDisplayFormat = [Microsoft.SharePoint.Client.DateTimeFieldFriendlyFormatType]::Disabled;
    }

    Write-Host "SITE COLUMN INSTALLED: $displayName" -ForegroundColor Green
}
else
{
    Write-Host "SITE COLUMN ALREADY INSTALLED: $displayName" -ForegroundColor Yellow        
}

Write-Host ""

Write-Host "SCRIPT FINISHED" -ForegroundColor Yellow
Stop-Transcript