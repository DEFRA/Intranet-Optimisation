<#
    SCRIPT OVERVIEW:
    This script creates our custom lists required within the Defra Intranet

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

# LIST - Create the "News Article Approval Information" list in which of the Intranet sites
$displayName = "News Article Approval Information"
$listURL = "Lists/SPAI"

$fieldNames = @("AssociatedSitePage","NewsArticleTitle","ContentSubmissionStatus","DateOfApprovalRequest","OrganisationIntranets","SPVersionNumber","DateTimeALBApprovalDecision")

Write-Host "`nCREATING THE '$displayName' LIST" -ForegroundColor Green

$site = $sites | Where-Object { $_.SiteType -eq "Parent" -and $_.RelativeURL.Length -gt 0 } | Sort-Object -Property @{Expression="SiteType";Descending=$true},@{Expression="DisplayName";Descending=$false}

Connect-PnPOnline -Url "$global:rootURL/$($site.RelativeURL)" -UseWebLogin
Write-Host "SCRIPT EXECUTED BY '$(Get-CurrentUser)' AT $(get-date -f "HH:mm:ss") ON $(get-date -f "dd/MM/yyyy")" -ForegroundColor Cyan
Write-Host "ACCESSING SHAREPOINT SITE: $($global:rootURL)/$($global:site.RelativeURL)" -ForegroundColor Cyan

$ctx = Get-PnPContext

$list = Get-PnPList -Identity $listURL -ErrorAction SilentlyContinue

if($null -eq $list)
{
    $list = New-PnPList -Template GenericList -Title $displayName -Url $listURL -Hidden
    Write-Host "LIST CREATED: $displayName (URL: $listURL)" -ForegroundColor Green
}
else
{
    Write-Host "THE LIST '$displayName' ALREADY EXISTS" -ForegroundColor Yellow
}

# FIELDS - ADD OUR CUSTOM FIELDS TO THE LIST 
Write-Host "`nADDING OUR FIELDS TO THE LIST" -ForegroundColor Green

foreach($fieldName in $fieldNames)
{
    $field = Get-PnPField -List $list -Identity $fieldName -ErrorAction SilentlyContinue

    if($null -eq $field)
    {
        Add-PnPField -List $list -Field $fieldName
        Write-Host "FIELD ADDED TO THE '$displayName' LIST: $fieldName" -ForegroundColor Green 
    }
    else
    {
        Write-Host "THE FIELD '$fieldName' ALREADY EXISTS IN THE LIST '$displayName'" -ForegroundColor Yellow 
    }
}

# FIELD CUSTOMISATIONS
Set-PnPField -List $list -Identity "Title" -Values @{
    Hidden = $true
    Required = $false
    Title = "Title"
}

Write-Host "`nDEFAULT FIELD 'Title' HIDDEN" -ForegroundColor Green

Set-PnPField -List $list -Identity "OrganisationIntranets" -Values @{
    Title = "Approving ALB"
    AllowMultipleValues = $false
}

$field = Get-PnPField -List $list -Identity "ContentSubmissionStatus"

Set-PnPField -List $list -Identity $field.Id -Values @{
    CustomFormatter = '{"elmType":"div","style":{"flex-wrap":"wrap","display":"flex"},"children":[{"elmType":"div","style":{"box-sizing":"border-box","padding":"4px 8px 5px 8px","overflow":"hidden","text-overflow":"ellipsis","display":"flex","border-radius":"16px","height":"24px","align-items":"center","white-space":"nowrap","margin":"4px 4px 4px 4px"},"attributes":{"class":{"operator":":","operands":[{"operator":"==","operands":["[$ContentSubmissionStatus]","Pending Approval"]},"sp-css-backgroundColor-BgGold sp-css-borderColor-GoldFont sp-field-fontSizeSmall sp-css-color-GoldFont",{"operator":":","operands":[{"operator":"==","operands":["[$ContentSubmissionStatus]","Approved"]},"sp-css-backgroundColor-BgMintGreen sp-field-fontSizeSmall sp-css-color-MintGreenFont",{"operator":":","operands":[{"operator":"==","operands":["[$ContentSubmissionStatus]","Approved (Automated)"]},"sp-css-backgroundColor-BgMintGreen sp-css-borderColor-MintGreenFont sp-field-fontSizeSmall sp-css-color-MintGreenFont",{"operator":":","operands":[{"operator":"==","operands":["[$ContentSubmissionStatus]","Cancelled Pending Approval"]},"sp-css-backgroundColor-BgLilac sp-css-borderColor-LilacFont sp-field-fontSizeSmall sp-css-color-LilacFont",{"operator":":","operands":[{"operator":"==","operands":["[$ContentSubmissionStatus]","Rejected"]},"sp-css-backgroundColor-BgCoral sp-css-borderColor-CoralFont sp-field-fontSizeSmall sp-css-color-CoralFont","sp-field-borderAllRegular sp-field-borderAllSolid sp-css-borderColor-neutralSecondary"]}]}]}]}]}},"txtContent":"[$ContentSubmissionStatus]"}]}'
}

# Cast the field to Choice Field
$choiceField = New-Object Microsoft.SharePoint.Client.FieldChoice($ctx, $field.Path)
$ctx.Load($choiceField)
Invoke-PnPQuery

$choiceField.Choices = "Pending Approval","Approved","Approved (Automated)","Cancelled Pending Approval","Rejected"
$choiceField.UpdateAndPushChanges($false)
Invoke-PnPQuery

# UPDATE VIEW INFORMATION
$view = Get-PnPView -List $list -Identity "All Items"

if($null -ne $view)
{
    $view = Set-PnPView -List $list -Identity $view.Title -Fields @("AssociatedSitePage","NewsArticleTitle","OrganisationIntranets","ContentSubmissionStatus","DateOfApprovalRequest","DateTimeALBApprovalDecision")
    $view.ViewQuery = '<GroupBy Collapse="FALSE" GroupLimit="30"><FieldRef Name="DateOfApprovalRequest" Ascending="FALSE" /></GroupBy><OrderBy><FieldRef Name="DateOfApprovalRequest" Ascending="FALSE" /><FieldRef Name="DateTimeALBApprovalDecision" /></OrderBy>'
    $view.Update()
    $ctx.ExecuteQuery()

    Write-Host "`nLIST DEFAULT VIEW '$($view.Title)' UPDATED WITH NEW FIELDS" -ForegroundColor Green 
}
else
{
    Write-Host "`nLIST DEFAULT VIEW '$($view.Title)' DOES NOT EXIST" -ForegroundColor Yellow
}

# LIST SETTING AND PERMISSION UPDATES
Set-PnPList -Identity $list -EnableAttachments 0
Write-Host "LIST ATTACHMENTS DISABLED" -ForegroundColor Green

$list.NoCrawl = $true
$list.Update()
$ctx.ExecuteQuery()
Write-Host "LIST EXCLUDED FROM SEARCH INDEX. CHANGES TAKE EFFECT AFTER THE NEXT CRAWL" -ForegroundColor Green

# Break Permission Inheritance of the List and set the new permissions for the members
Set-PnPList -Identity $list -BreakRoleInheritance

$membersGroup = Get-PnPGroup | Where-Object { $_.Title -like "* Members"}
$ownersGroup = Get-PnPGroup | Where-Object { $_.Title -like "* Owners"}

Set-PnPListPermission -Identity $list -AddRole "Read" -Group $membersGroup
Set-PnPListPermission -Identity $list -AddRole "Read" -Group $ownersGroup

Write-Host "LIST PERMISSIONS UPDATED" -ForegroundColor Green

Write-Host "SCRIPT FINISHED" -ForegroundColor Yellow
Stop-Transcript