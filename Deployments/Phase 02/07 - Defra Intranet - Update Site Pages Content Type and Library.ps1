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

$CTName = "Site Page"
$ct = Get-PnPContentType -Identity $CTName

$listName = "Site Pages"
$list = Get-PnPList -Identity $listName

if($null -eq $ct)
{
    throw "The content type $CTName could not be found on the site"
}

$ctx.Load($ct.FieldLinks)
$ctx.ExecuteQuery()

# Variable configuration.
$fieldNames = @("OrganisationIntranetsContentEditorInput","PageApprovalInfo","WorkflowPublishtoALBIntranets","WorkflowArticleSentForALBApproval","WorkflowApprovalProgress")

# FIELDS - ADD FIELDS TO SITE PAGE LIBRARY
Write-Host "`nADDING FIELDS TO THE CONTENT TYPE $CTName" -ForegroundColor Green

foreach($fieldName in $fieldNames)
{
    $field = $ct.FieldLinks | Where-Object { $_.Name -eq $fieldName }

    if($null -eq $field)
    {
        Add-PnPFieldToContentType -Field $fieldName -ContentType $ct -ErrorAction SilentlyContinue
        Write-Host "THE FIELD '$fieldName' HAS BEEN ADDED TO THE CONTENT TYPE: $CTName" -ForegroundColor Green 
    }
    else
    {
        Write-Host "THE FIELD '$fieldName' ALREADY EXISTS ON THE CONTENT TYPE: $CTName" -ForegroundColor Yellow
    }
}

# Customise the existing "OrganisationIntranets" column for this library. The new "Organisation (Intranets)" column will be taking over user interaction.
$fieldInternalName = "OrganisationIntranets"
$field = Get-PnPField -List $list -Identity $fieldInternalName -ErrorAction SilentlyContinue

if($null -ne $field)
{
    Set-PnPField -List $listName -Identity $field.Id -Values @{
        Title = "Organisation (Intranets) - Approving ALBs"
    }

    # Apply conditional formula
    $formula = "=if([{0}] == 'Site Page', 'true', 'false')" -f '$ContentType'
    $field.ClientValidationFormula = $formula
    $field.Update()
    Invoke-PnPQuery

    Write-Host "FIELD '$fieldInternalName' UPDATED" -ForegroundColor Green
}
else
{
    Write-Host "THE FIELD '$fieldInternalName' DOES NOT EXIST IN THE LIBRARY '$listName'" -ForegroundColor Red
}

# Customise the existing "WorkflowPublishtoALBIntranets" column for this library.
$fieldInternalName = "WorkflowPublishtoALBIntranets"
$field = Get-PnPField -List $list -Identity $fieldInternalName -ErrorAction SilentlyContinue

if($null -ne $field)
{
    $condition1 = "=if([{0}] != 'Workflow is starting...', 'inherit','none')" -f '$WorkflowApprovalProgress'
    $condition2 = "=if([{0}] == 'Workflow is starting...', 'inherit','none')" -f '$WorkflowApprovalProgress'

    Set-PnPField -List $listName -Identity $field.Id -Values @{
        CustomFormatter = '{"elmType":"div","style":{"flex-directon":"row","justify-content":"left","align-items":"center","flex-wrap":"nowrap"},"children":[{"elmType":"div","style":{"display":"' + $condition1 + '","flex-directon":"row","justify-content":"left","align-items":"center","flex-wrap":"wrap"},"children":[{"elmType":"button","customRowAction":{"action":"setValue","actionInput":{"WorkflowApprovalProgress":"Workflow is starting..."}},"attributes":{"class":"ms-fontColor-themePrimary ms-fontColor-themeDarker--hover"},"style":{"border":"1px solid","background-color":"transparent","cursor":"pointer","display":"flex","flex-directon":"row","justify-content":"left","align-items":"center","flex-wrap":"wrap","padding":"10px"},"children":[{"elmType":"span","attributes":{"iconName":"PublishContent"},"style":{"padding-right":"6px"}},{"elmType":"span","txtContent":"Publish to ALB Intranets","style":{"word-break":"keep-all"}}]}]},{"elmType":"div","style":{"border":"1px solid","background-color":"transparent","display":"' + $condition2 + '","flex-directon":"row","justify-content":"left","align-items":"center","flex-wrap":"wrap","padding":"10px","min-width":"159px","color":"#A0A0A0"},"attributes":{"class":"ms-fontColor-themePrimary ms-fontColor-themeDarker--hover"},"children":[{"elmType":"span","attributes":{"iconName":"PublishContent"},"style":{"padding-right":"6px"}},{"elmType":"span","txtContent":"Publish to ALB Intranets","style":{"word-break":"keep-all","font-size":"13.33px"}}]}]}'
    }
}

# Hide the internal fields, which are managed by our workflows and not the users, from the site page property details panel
$hiddenFieldNames = @("PageApprovalInfo","WorkflowPublishtoALBIntranets","WorkflowArticleSentForALBApproval","WorkflowApprovalProgress")

foreach($hiddenField in $hiddenFieldNames)
{
    $field = Get-PnPField -List $list -Identity $hiddenField -ErrorAction SilentlyContinue

    if($null -ne $field)
    {
        Set-PnPField -List $listName -Identity $field.Id -Values @{
            Hidden = $true
        }

        Write-Host "FIELD '$hiddenField' UPDATED" -ForegroundColor Green
    }
    else
    {
        Write-Host "THE FIELD '$hiddenField' DOES NOT EXIST IN THE LIBRARY '$listName'" -ForegroundColor Red
    }
}

# VIEW UPDATES
$views = Get-PnPView -List $list | Where-Object { $_.Title -ne "" }

# Remove the old Organisation (Intranet) field. The reason for this is SharePoint is going to manage this away from the user now
$fieldInternalName = "OrganisationIntranets"

foreach($view in $views)
{
    $ctx.Load($view.ViewFields)
    $ctx.ExecuteQuery()

    $viewFields = $view.ViewFields

    if($null -ne $($viewFields | Where-Object { $_ -eq $fieldInternalName }))
    {
        $viewFieldNames = New-Object Collections.Generic.List[String]

        foreach($viewField in $viewFields)
        {
            if($viewField -ne $fieldInternalName)
            {
                $viewFieldNames.Add($viewField)
            }
        }

        $view = Set-PnPView -List $listName -Identity $view.Title -Fields $viewFieldNames
        Write-Host "THE FIELD '$($fieldInternalName)' HAS REMOVED FROM THE '$listName' LIBRARY VIEW '$($view.Title)'" -ForegroundColor Green 
    }
    else
    {
        Write-Host "THE FIELD '$($fieldInternalName)' HAS ALREADY BEEN REMOVED FROM THE VIEW '$($view.Title)'" -ForegroundColor Yellow
    }
}

Write-Host ""

$fieldNames = @("OrganisationIntranetsContentEditorInput","PageApprovalInfo","WorkflowPublishtoALBIntranets","WorkflowApprovalProgress")

foreach($view in $views)
{
    $ctx.Load($view.ViewFields)
    $ctx.ExecuteQuery()

    $viewFields = $view.ViewFields

    $viewFieldNames = New-Object Collections.Generic.List[String]

    foreach($viewField in $viewFields)
    {
        $viewFieldNames.Add($viewField)
    }

    foreach($fieldName in $fieldNames)
    {

        if($null -eq $($viewFields | Where-Object { $_ -eq $fieldName }))
        {
            $viewFieldNames.Add($fieldName)
            Write-Host "THE FIELD '$($fieldName)' HAS BEEN ADDED TO THE '$listName' LIBRARY VIEW '$($view.Title)'" -ForegroundColor Green
        }
        else
        {
            Write-Host "THE FIELD '$($fieldName)' HAS ALREADY BEEN ADDED TO THE VIEW '$($view.Title)'" -ForegroundColor Yellow
        }
    }

    $view = Set-PnPView -List $listName -Identity $view.Title -Fields $viewFieldNames
}

# LIST SETTINGS AND PERMISSIONS
Write-Host "Updating Permissions" -ForegroundColor Green

# Break Permission Inheritance of the List and set the new permissions for the members
Set-PnPList -Identity $list -BreakRoleInheritance -CopyRoleAssignments

$group = Get-PnPGroup | Where-Object { $_.Title -like "* Members"}

Set-PnPListPermission -Identity $list -AddRole "Custom Permission - Contribute - For Site Page Library Only" -Group $group -RemoveRole "Edit"

Write-Host "Disabled Edit in Grid View" -ForegroundColor Green
$list.DisableGridEditing = $true
$list.Update()
Invoke-PnPQuery

Write-Host ""
Write-Host "SCRIPT FINISHED" -ForegroundColor Yellow
Stop-Transcript