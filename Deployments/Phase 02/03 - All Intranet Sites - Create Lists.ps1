<#
    SCRIPT OVERVIEW:
    This script creates our custom lists in the Defra Intranet and the ALB sites

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

# LIST - Create the "Internal Comms Intranet Content Submissions" list in which of the Intranet sites
foreach($site in $sites)
{
    Connect-PnPOnline -Url "$global:rootURL/$($site.RelativeURL)" -UseWebLogin
    Write-Host "SCRIPT EXECUTED BY '$(Get-CurrentUser)' AT $(get-date -f "HH:mm:ss") ON $(get-date -f "dd/MM/yyyy")" -ForegroundColor Cyan
    Write-Host "ACCESSING SHAREPOINT SITE: $($global:rootURL)/$($global:site.RelativeURL)" -ForegroundColor Cyan

    $ctx = Get-PnPContext

    # Create new "Submission" list
    $displayName = "Internal Comms Intranet Content Submissions"
    $listURL = "Lists/ICICS"

    # Site-specific variable configuration.
    switch ($site.Abbreviation)
    {
        "Defra" { 
            $fieldNames = @("ContentTypes","OrganisationIntranets","PublishBy","LineManager","AltContact","ContentSubmissionApprovalOptions","ContentSubmissionStatus","ContentSubmissionDescription","ContentSubmissionApproveRejectBy")
        }

        default { 
            $fieldNames = @("ContentTypes","LineManager","AltContact","PublishBy","ContentSubmissionApprovalOptions","ContentSubmissionStatus","ContentSubmissionDescription","ContentSubmissionApproveRejectBy")
        }
    }

    Write-Host "`nCREATING THE LIST" -ForegroundColor Green

    # LIST - LIST CREATION
    $list = Get-PnPList -Identity $listURL

    if($null -eq $list)
    {
        $list = New-PnPList -Template GenericList -Title $displayName -Url $listURL -EnableVersioning -EnableContentTypes
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

    # SITE-SPECIFIC FIELDS
    switch ($site.Abbreviation)
    {
        "Defra" { 
            # Customise the OrganisationIntranets column for this list
            $field = Get-PnPField -List $list -Identity "OrganisationIntranets" -ErrorAction SilentlyContinue

            if($null -ne $field)
            {
                Set-PnPField -List $list -Identity $field.Id -Values @{
                    Title = "Which Defra organisation is this relevant to?"
                    Description = "Select whether the content is relevant to the whole of the Defra group or specific organisations."
                    Required = $true
                }
            }
            else
            {
                Write-Host "THE FIELD 'OrganisationIntranets' DOES NOT EXIST IN THE LIST '$displayName'" -ForegroundColor Red
            }
        }
    }

    # LIST-LEVEL FIELD CUSTOMISATION
    Write-Host "`nCUSTOMISING FIELDS" -ForegroundColor Green

    # Customise the "Approval Options" field for this list
    $field = Get-PnPField -List $list -Identity "ContentSubmissionApprovalOptions" -ErrorAction SilentlyContinue

    if($null -ne $field)
    {
        Set-PnPField -List $list -Identity $field.Id -Values @{
            Hidden = $true
            CustomFormatter = '{"elmType":"div","style":{"flex-directon":"row","justify-content":"left","align-items":"center","flex-wrap":"nowrap"},"children":[{"elmType":"div","txtContent":"N/A","style":{"visibility":"=if([$PermMask] < 0x1b03c431aef, ''visible'',''hidden'')","width":"=if([$PermMask] < 0x1b03c431aef, '' '',''0'')","height":"=if([$PermMask] < 0x1b03c431aef, '' '',''0'')","word-break":"keep-all"}},{"elmType":"div","style":{"visibility":{"operator":"?","operands":[{"operator":"<","operands":["[$PermMask]","0x1b03c431aef"]},"hidden","visible"]},"width":{"operator":"?","operands":[{"operator":"<","operands":["[$PermMask]","0x1b03c431aef"]},"0",""]},"height":{"operator":"?","operands":[{"operator":"<","operands":["[$PermMask]","0x1b03c431aef"]},"0",""]},"display":"=if(([$ContentSubmissionStatus] == ''Pending Approval''),''inherit'',''none'')","flex-directon":"row","justify-content":"left","align-items":"center","flex-wrap":"wrap"},"children":[{"elmType":"button","customRowAction":{"action":"setValue","actionInput":{"ContentSubmissionStatus":"Approved"}},"attributes":{"class":"ms-fontColor-themePrimary ms-fontColor-themeDarker--hover"},"style":{"border":"none","background-color":"transparent","cursor":"pointer","display":"flex","flex-directon":"row","justify-content":"left","align-items":"center","flex-wrap":"wrap"},"children":[{"elmType":"span","attributes":{"iconName":"SkypeCircleCheck"},"style":{"padding":"4px"}},{"elmType":"span","txtContent":"Approve","style":{"word-break":"keep-all"}}]},{"elmType":"button","customRowAction":{"action":"setValue","actionInput":{"ContentSubmissionStatus":"Rejected"}},"attributes":{"class":"ms-fontColor-themePrimary ms-fontColor-themeDarker--hover"},"style":{"border":"none","background-color":"transparent","cursor":"pointer","display":"flex","flex-directon":"row","justify-content":"left","align-items":"center","flex-wrap":"wrap"},"children":[{"elmType":"span","attributes":{"iconName":"Blocked"},"style":{"padding":"4px"}},{"elmType":"span","txtContent":"Reject","style":{"word-break":"keep-all"}}]}]}]}'
        }

        Write-Host "THE FIELD '$($field.Title)' HAS BEEN CUSTOMISED FOR THE LIST '$displayName'" -ForegroundColor Green
    }
    else
    {
        Write-Host "THE FIELD 'ContentSubmissionStatus' DOES NOT EXIST IN THE LIST '$displayName'" -ForegroundColor Red
    }

    # Customise the "Status" field for this list
    $field = Get-PnPField -List $list -Identity "ContentSubmissionStatus" -ErrorAction SilentlyContinue

    if($null -ne $field)
    {
        Set-PnPField -List $list -Identity $field.Id -Values @{
            Hidden = $true
            CustomFormatter = '{"elmType":"div","style":{"flex-wrap":"wrap","display":"flex"},"children":[{"elmType":"div","style":{"box-sizing":"border-box","padding":"4px 8px 5px 8px","overflow":"hidden","text-overflow":"ellipsis","display":"flex","border-radius":"16px","height":"24px","align-items":"center","white-space":"nowrap","margin":"4px 4px 4px 4px"},"attributes":{"class":{"operator":":","operands":[{"operator":"==","operands":["[$ContentSubmissionStatus]","Pending Approval"]},"sp-css-backgroundColor-BgGold sp-css-borderColor-GoldFont sp-field-fontSizeSmall sp-css-color-GoldFont",{"operator":":","operands":[{"operator":"==","operands":["[$ContentSubmissionStatus]","Approved"]},"sp-css-backgroundColor-BgMintGreen sp-field-fontSizeSmall sp-css-color-MintGreenFont",{"operator":":","operands":[{"operator":"==","operands":["[$ContentSubmissionStatus]","Rejected"]},"sp-css-backgroundColor-BgDustRose sp-css-borderColor-DustRoseFont sp-field-fontSizeSmall sp-css-color-DustRoseFont",""]}]}]}},"txtContent":"[$ContentSubmissionStatus]"}]}'
        }

        Write-Host "THE FIELD '$($field.Title)' HAS BEEN CUSTOMISED FOR THE LIST '$displayName'" -ForegroundColor Green
    }
    else
    {
        Write-Host "THE FIELD 'ContentSubmissionStatus' DOES NOT EXIST IN THE LIST '$displayName'" -ForegroundColor Red
    }

    # Customise the "Approved or Rejected By" field for this list
    $field = Get-PnPField -List $list -Identity "ContentSubmissionApproveRejectBy" -ErrorAction SilentlyContinue

    if($null -ne $field)
    {
        Set-PnPField -List $list -Identity $field.Id -Values @{
            Hidden = $true
        }

        Write-Host "THE FIELD '$($field.Title)' HAS BEEN CUSTOMISED FOR THE LIST '$displayName'" -ForegroundColor Green
    }
    else
    {
        Write-Host "THE FIELD 'ContentSubmissionStatus' DOES NOT EXIST IN THE LIST '$displayName'" -ForegroundColor Red
    }

    # Customise the "Attachments" field
    $field = Get-PnPField -List $list -Identity "Attachments" -ErrorAction SilentlyContinue

    if($null -ne $field)
    {
        Set-PnPField -List $list -Identity $field.Id -Values @{
            Description = "Please attach any content submission here and any associated imagery that you would like to appear on your page."
        }
         
        Write-Host "THE FIELD '$($field.Title)' HAS BEEN CUSTOMISED FOR THE LIST '$displayName'" -ForegroundColor Green
    }
    else
    {
        Write-Host "THE FIELD 'Attachments' DOES NOT EXIST IN THE LIST '$displayName'" -ForegroundColor Red
    }

    # CONTENT TYPES
    Write-Host "`nCUSTOMISING CONTENT TYPES" -ForegroundColor Green

    $CTsToHide = New-Object System.Collections.ArrayList

    # SITE-SPECIFIC FIELDS
    switch ($site.Abbreviation)
    {
        # The Content Submission content types are to be hidden on the APHA and RPA sites
        { @("APHA","RPA") -contains $_ }  
        {
            Write-Host "Skipping the content submission content types in the $_ site" -ForegroundColor Cyan
        }

        default 
        {
            # Content Submission Request - Stage 2
            $ctName = "Content Submission Request - Stage 2"
            $listCT = Get-PnPContentType -Identity $ctName -List $displayName -ErrorAction SilentlyContinue

            if($null -eq $listCT)
            {
                $ct = Get-PnPContentType -Identity $ctName

                if($null -ne $ct)
                {
                    Add-PnPContentTypeToList -List $displayName -ContentType $ct
                    $listCT = Get-PnPContentType -Identity $ctName -List $displayName
                    Write-Host "SITE CONTENT TYPE INSTALLED '$ctName' HAS BEEN INSTALLED ON THE LIST '$displayName'" -ForegroundColor Green
                }
                else
                {
                    throw "ERROR: The content Type '$ctName' is missing from the site. Please run the 'Create Content Types.ps1' script then try again." 
                }
            }
            else
            {
                Write-Host "THE CONTENT TYPE '$displayName' ALREADY EXISTS ON THE LIST '$displayName'" -ForegroundColor Yellow   
            }

            $ctx.Load($list.ContentTypes)
            $ctx.Load($list.RootFolder)
            $ctx.ExecuteQuery()

            # We'll hide this CT from the New menu, as it's only needed by Power Automate.
            $CTsToHide.Add($listCT.Id.StringValue) | Out-Null
        }
    }

    # Event Submission Request
    $ctName = "Event Submission Request"
    $listCT = Get-PnPContentType -Identity $ctName -List $displayName -ErrorAction SilentlyContinue

    if($null -eq $listCT)
    {
        $ct = Get-PnPContentType -Identity $ctName

        if($null -ne $ct)
        {
            Add-PnPContentTypeToList -List $displayName -ContentType $ct
            $listCT = Get-PnPContentType -Identity $ctName -List $displayName
            Write-Host "SITE CONTENT TYPE INSTALLED '$ctName' HAS BEEN INSTALLED ON THE LIST '$displayName'" -ForegroundColor Green
        }
        else
        {
            throw "ERROR: The content Type '$ctName' is missing from the site. Please run the 'Create Content Types.ps1' script then try again." 
        }
    }
    else
    {
        Write-Host "THE CONTENT TYPE '$displayName' ALREADY EXISTS ON THE LIST '$displayName'" -ForegroundColor Yellow   
    }

    $ctx.Load($list.ContentTypes)
    $ctx.Load($list.RootFolder)
    $ctx.ExecuteQuery()

    # Event Submission Request - Stage 2
    $ctName = "Event Submission Request - Stage 2"
    $listCT = Get-PnPContentType -Identity $ctName -List $displayName -ErrorAction SilentlyContinue

    if($null -eq $listCT)
    {
        $ct = Get-PnPContentType -Identity $ctName

        if($null -ne $ct)
        {
            Add-PnPContentTypeToList -List $displayName -ContentType $ct
            $listCT = Get-PnPContentType -Identity $ctName -List $displayName
            Write-Host "SITE CONTENT TYPE INSTALLED '$ctName' HAS BEEN INSTALLED ON THE LIST '$displayName'" -ForegroundColor Green
        }
        else
        {
            throw "ERROR: The content Type '$ctName' is missing from the site. Please run the 'All Intranet Sites - Create Content Types.ps1' script then try again." 
        }
    }
    else
    {
        Write-Host "THE CONTENT TYPE '$displayName' ALREADY EXISTS ON THE LIST '$displayName'" -ForegroundColor Yellow   
    }

    $ctx.Load($list.ContentTypes)
    $ctx.Load($list.RootFolder)
    $ctx.ExecuteQuery()

    # We'll hide this CT from the New menu, as it's only needed by Power Automate.
    $CTsToHide.Add($listCT.Id.StringValue) | Out-Null
   
    if($null -eq $list.RootFolder.UniqueContentTypeOrder)
    {
        $contentTypesInPlace = New-Object -TypeName 'System.Collections.Generic.List[Microsoft.SharePoint.Client.ContentTypeId]'
        
        foreach($ct in $list.ContentTypes | where {$CTsToHide -notcontains $_.Id.StringValue -and $_.Name -ne "Folder"})
        {
            Write-Host "$($ct.Name) added the 'New' menu" -ForegroundColor Cyan
            $contentTypesInPlace.Add($ct.Id)
        }
    }
    else 
    {
        $contentTypesInPlace = [System.Collections.ArrayList] $list.RootFolder.UniqueContentTypeOrder
        $contentTypesInPlace = $contentTypesInPlace | where {$_.StringValue -ne $ct.Id.StringValue}
    }

    $list.RootFolder.UniqueContentTypeOrder = [System.Collections.Generic.List[Microsoft.SharePoint.Client.ContentTypeId]] $contentTypesInPlace
    $list.RootFolder.Update()             
    Invoke-PnPQuery

    # Rename default "Item" content type to "Content Submission Request"
    $ct = Get-PnPContentType -List $list -Identity "Item" -ErrorAction SilentlyContinue

    if($null -ne $ct)
    {
        $ctx = Get-PnPContext
        $ctx.Load($ct)
        $ctx.ExecuteQuery()

        try
        {
            $ct.ReadOnly = $false
            $ct.Update($false)
            $ctx.ExecuteQuery()

            $ct.Name = "Content Submission Request"
            $ct.Update($false)
            $ctx.ExecuteQuery()

            Write-Host "`nList default content type 'Item' renamed to 'Content Submission Request'" -ForegroundColor Green
        }
        finally
        {
            $ct.ReadOnly = $true
            $ct.Update($false)
            $ctx.ExecuteQuery()
        }
    }

    # SITE-SPECIFIC CUSTOM CONTENT TYPE CHANGES
    switch ($site.Abbreviation)
    {
        # APHA and RPA are only using the list for event submissions, so we change the default content type to events and hide the "Content" content type
        { @("APHA","RPA") -contains $_ } 
        {
            # Set Event Submission Request as the default for this list.....
            $ctName = "Event Submission Request"           
            Set-PnPDefaultContentTypeToList -List $list -ContentType $ctName
            Write-Host "Default content type set to $ctName" -ForegroundColor Yellow

            # .... and hide the previous default
            $ctName = "Content Submission Request" 
            $ct = Get-PnPContentType -List $list -Identity $ctName

            $ct.ReadOnly = $false
            $ct.Hidden = $true
            $ct.Update($false)
            $ctx.ExecuteQuery()
            Write-Host "THE ORIGINAL DEFAULT '$ctName' HAS BEEN HIDDEN FROM THE 'NEW' MENU" -ForegroundColor Yellow
        }
    }

    # CONTENT TYPE - CONFIGURE LAYOUT
    Write-Host "`nCUSTOMISING CONTENT TYPE FORM LAYOUTS" -ForegroundColor Green
    $cts = Get-PnPContentType -List $list | Where-Object {$_.Name -notlike "*Stage 2" -and $_.Name -ne "Folder"}
    $ctIsReadOnly = $false

    foreach($ct in $cts)
    {
        $ctx.Load($ct)
        $ctx.ExecuteQuery()
    
        try
        {
            if($ct.ReadOnly -eq $true)
            {
                $ctIsReadOnly = $true

                $ct.ReadOnly = $false
                $ct.Update($false)
                $ctx.ExecuteQuery()
            }

            $ct.ClientFormCustomFormatter = '{"headerJSONFormatter":{},"footerJSONFormatter":"","bodyJSONFormatter":""}'
            $ct.Update($false)
            $ctx.ExecuteQuery()

            Write-Host "Applied custom form layout to the content type '$($ct.Name)'" -ForegroundColor Green
        }
        finally
        {
            if($ctIsReadOnly -eq $true)
            {
                $ct.ReadOnly = $false
                $ct.Update($false)
                $ctx.ExecuteQuery()

                $ctIsReadOnly = $false
            }
        }
    }

    # VIEWS - Setup custom list views
    Write-Host "`nCUSTOMISING LIST VIEWS" -ForegroundColor Green

    switch ($site.Abbreviation)
    {
        "Defra" 
        {
            $viewFields = @{
                'AllItemsAssigned' = "Attachments","LinkTitle","ContentType","PublishBy","Author","OrganisationIntranets","AltContact","ContentSubmissionApprovalOptions","ContentSubmissionStatus","ContentSubmissionApproveRejectBy"
                'Content' = "Attachments","LinkTitle","AssignedTo","OrganisationIntranets","ContentSubmissionDescription","Author","PublishBy","ContentTypes","AltContact","LineManager","ContentSubmissionApprovalOptions","ContentSubmissionStatus","ContentSubmissionApproveRejectBy"
                'Default' = "Attachments","LinkTitle","ContentType","PublishBy","AssignedTo","Author","OrganisationIntranets","AltContact","ContentSubmissionApprovalOptions","ContentSubmissionStatus","ContentSubmissionApproveRejectBy"
                'Events' = "Attachments","LinkTitle","AssignedTo","OrganisationIntranets","Author","PublishBy","EventDateTime","EventVenueAndJoiningDetails","EventDetails","ContentSubmissionApprovalOptions","ContentSubmissionStatus","ContentSubmissionApproveRejectBy"
            }
        }

        { @("APHA","RPA") -contains $_ }  
        {
            $viewFields = @{
                'AllItemsAssigned' = "Attachments","LinkTitle","AssignedTo","Author","PublishBy","EventDateTime","EventVenueAndJoiningDetails","EventDetails","ContentSubmissionApprovalOptions","ContentSubmissionStatus","ContentSubmissionApproveRejectBy"
                'Default' = "Attachments","LinkTitle","AssignedTo","Author","PublishBy","EventDateTime","EventVenueAndJoiningDetails","EventDetails","ContentSubmissionApprovalOptions","ContentSubmissionStatus","ContentSubmissionApproveRejectBy"
                'Events' = "Attachments","LinkTitle","AssignedTo","Author","PublishBy","EventDateTime","EventVenueAndJoiningDetails","EventDetails","ContentSubmissionApprovalOptions","ContentSubmissionStatus","ContentSubmissionApproveRejectBy"
            }
        }

        default 
        {
            $viewFields = @{
                'AllItemsAssigned' = "Attachments","LinkTitle","ContentType","Author","AltContact","ContentSubmissionApprovalOptions","ContentSubmissionStatus","ContentSubmissionApproveRejectBy"
                'Content' = "Attachments","LinkTitle","AssignedTo","Author","ContentSubmissionDescription","PublishBy","ContentTypes","AltContact","LineManager","ContentSubmissionApprovalOptions","ContentSubmissionStatus","ContentSubmissionApproveRejectBy"
                'Default' = "Attachments","LinkTitle","ContentType","AssignedTo","Author","AltContact","ContentSubmissionApprovalOptions","ContentSubmissionStatus","ContentSubmissionApproveRejectBy"
                'Events' = "Attachments","LinkTitle","AssignedTo","Author","PublishBy","EventDateTime","EventVenueAndJoiningDetails","EventDetails","ContentSubmissionApprovalOptions","ContentSubmissionStatus","ContentSubmissionApproveRejectBy"
            }
        }
    }

    switch ($site.Abbreviation)
    {
        { @("APHA","RPA") -contains $_ } 
        {
            $viewConfiguration = @(
                [PSCustomObject]@{
                    'Query' = '<GroupBy Collapse="FALSE" GroupLimit="30"><FieldRef Name="AssignedTo" /></GroupBy><OrderBy><FieldRef Name="ID" Ascending="FALSE" /></OrderBy><Where><Eq><FieldRef Name="ContentSubmissionStatus" /><Value Type="Text">Pending Approval</Value></Eq></Where>'
                    'TargetSite' = ''
                    'Title' = 'Events - By Assigned To'
                    'ViewFields' = $viewFields.AllItemsAssigned
                },
                [PSCustomObject]@{
                    'Query' = '<OrderBy><FieldRef Name="ID" Ascending="FALSE" /></OrderBy><Where><And><Eq><FieldRef Name="ContentSubmissionStatus" /><Value Type="Text">Pending Approval</Value></Eq><Or><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Event Submission Request</Value></Eq><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Event Submission Request - Stage 2</Value></Eq></Or></And></Where>'
                    'TargetSite' = ''
                    'Title' = 'Events - All Pending Submissions'
                    'ViewFields' = $viewFields.Events
                },
                [PSCustomObject]@{
                    'Query' = '<OrderBy><FieldRef Name="ID" Ascending="FALSE" /></OrderBy><Where><And><Or><Eq><FieldRef Name="ContentSubmissionStatus" /><Value Type="Text">Approved</Value></Eq><Eq><FieldRef Name="ContentSubmissionStatus" /><Value Type="Text">Rejected</Value></Eq></Or><Or><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Event Submission Request</Value></Eq><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Event Submission Request - Stage 2</Value></Eq></Or></And></Where>'
                    'TargetSite' = ''
                    'Title' = 'Events - All Processed Submissions'
                    'ViewFields' = $viewFields.Events
                },
                [PSCustomObject]@{
                    'Query' = '<GroupBy Collapse="FALSE" GroupLimit="30"><FieldRef Name="AssignedTo" /></GroupBy><Where><Or><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Event Submission Request</Value></Eq><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Event Submission Request - Stage 2</Value></Eq></Or></Where>'
                    'TargetSite' = ''
                    'Title' = 'Events - By Assigned To'
                    'ViewFields' = $viewFields.Events
                },
                [PSCustomObject]@{
                    'Query' = '<GroupBy Collapse="FALSE" GroupLimit="30"><FieldRef Name="AssignedTo" /></GroupBy><OrderBy><FieldRef Name="EventDateTime" /></OrderBy><Where><And><And><And><Or><And><And><And><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Event Submission Request</Value></Eq><Geq><FieldRef Name="EventDateTime" /><Value Type="DateTime"><Today /></Value></Geq></And><Leq><FieldRef Name="EventDateTime" /><Value Type="DateTime"><Today OffsetDays="7" /></Value></Leq></And><Neq><FieldRef Name="ContentSubmissionStatus" /><Value Type="Text">Rejected</Value></Neq></And><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Event Submission Request - Stage 2</Value></Eq></Or><Geq><FieldRef Name="EventDateTime" /><Value Type="DateTime"><Today /></Value></Geq></And><Leq><FieldRef Name="EventDateTime" /><Value Type="DateTime"><Today OffsetDays="7" /></Value></Leq></And><Neq><FieldRef Name="ContentSubmissionStatus" /><Value Type="Text">Rejected</Value></Neq></And></Where>'
                    'TargetSite' = ''
                    'Title' = 'Events - Due in the Next 07 Days'
                    'ViewFields' = $viewFields.Events
                },
                [PSCustomObject]@{
                    'Query' = '<GroupBy Collapse="FALSE" GroupLimit="30"><FieldRef Name="AssignedTo" /></GroupBy><OrderBy><FieldRef Name="EventDateTime" /></OrderBy><Where><And><And><And><Or><And><And><And><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Event Submission Request</Value></Eq><Geq><FieldRef Name="EventDateTime" /><Value Type="DateTime"><Today /></Value></Geq></And><Leq><FieldRef Name="EventDateTime" /><Value Type="DateTime"><Today OffsetDays="14" /></Value></Leq></And><Neq><FieldRef Name="ContentSubmissionStatus" /><Value Type="Text">Rejected</Value></Neq></And><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Event Submission Request - Stage 2</Value></Eq></Or><Geq><FieldRef Name="EventDateTime" /><Value Type="DateTime"><Today /></Value></Geq></And><Leq><FieldRef Name="EventDateTime" /><Value Type="DateTime"><Today OffsetDays="14" /></Value></Leq></And><Neq><FieldRef Name="ContentSubmissionStatus" /><Value Type="Text">Rejected</Value></Neq></And></Where>'
                    'TargetSite' = ''
                    'Title' = 'Events - Due in the Next 14 Days'
                    'ViewFields' = $viewFields.Events
                }
            )
        }

        default 
        {
            $viewConfiguration = @(
                [PSCustomObject]@{
                    'Query' = '<GroupBy Collapse="FALSE" GroupLimit="30"><FieldRef Name="AssignedTo" /></GroupBy><OrderBy><FieldRef Name="ID" Ascending="FALSE" /></OrderBy><Where><Eq><FieldRef Name="ContentSubmissionStatus" /><Value Type="Text">Pending Approval</Value></Eq></Where>'
                    'TargetSite' = ''
                    'Title' = 'All Items - By Assigned To'
                    'ViewFields' = $viewFields.AllItemsAssigned
                },
                [PSCustomObject]@{
                    'Query' = '<OrderBy><FieldRef Name="ID" Ascending="FALSE" /></OrderBy><Where><And><Eq><FieldRef Name="ContentSubmissionStatus" /><Value Type="Text">Pending Approval</Value></Eq><Or><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Content Submission Request</Value></Eq><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Content Submission Request - Stage 2</Value></Eq></Or></And></Where>'
                    'TargetSite' = ''
                    'Title' = 'Content - All Pending Submissions'
                    'ViewFields' = $viewFields.Content
                },
                [PSCustomObject]@{
                    'Query' = '<OrderBy><FieldRef Name="ID" Ascending="FALSE" /></OrderBy><Where><And><Or><Eq><FieldRef Name="ContentSubmissionStatus" /><Value Type="Text">Approved</Value></Eq><Eq><FieldRef Name="ContentSubmissionStatus" /><Value Type="Text">Rejected</Value></Eq></Or><Or><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Content Submission Request</Value></Eq><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Content Submission Request - Stage 2</Value></Eq></Or></And></Where>'
                    'TargetSite' = ''
                    'Title' = 'Content - All Processed Submissions'
                    'ViewFields' = $viewFields.Content
                },
                [PSCustomObject]@{
                    'Query' = '<GroupBy Collapse="FALSE" GroupLimit="30"><FieldRef Name="AssignedTo" /></GroupBy><Where><Or><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Content Submission Request</Value></Eq><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Content Submission Request - Stage 2</Value></Eq></Or></Where>'
                    'TargetSite' = ''
                    'Title' = 'Content - By Assigned To'
                    'ViewFields' = $viewFields.Content
                },
                [PSCustomObject]@{
                    'Query' = '<GroupBy Collapse="FALSE" GroupLimit="30"><FieldRef Name="ContentTypes" /></GroupBy><Where><Or><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Content Submission Request</Value></Eq><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Content Submission Request - Stage 2</Value></Eq></Or></Where>'
                    'TargetSite' = ''
                    'Title' = 'Content - By Content Types'
                    'ViewFields' = $viewFields.Content
                },
                [PSCustomObject]@{
                    'Query' = '<GroupBy Collapse="FALSE" GroupLimit="30"><FieldRef Name="AssignedTo" /></GroupBy><OrderBy><FieldRef Name="PublishBy" /></OrderBy><Where><And><And><Or><And><And><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Content Submission Request</Value></Eq><Geq><FieldRef Name="PublishBy" /><Value Type="DateTime"><Today /></Value></Geq></And><Leq><FieldRef Name="PublishBy" /><Value Type="DateTime"><Today OffsetDays="7" /></Value></Leq></And><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Content Submission Request - Stage 2</Value></Eq></Or><Geq><FieldRef Name="PublishBy" /><Value Type="DateTime"><Today /></Value></Geq></And><Leq><FieldRef Name="PublishBy" /><Value Type="DateTime"><Today OffsetDays="7" /></Value></Leq></And></Where>'
                    'TargetSite' = ''
                    'Title' = 'Content - Due in the Next 07 Days'
                    'ViewFields' = $viewFields.Content
                },
                [PSCustomObject]@{
                    'Query' = '<GroupBy Collapse="FALSE" GroupLimit="30"><FieldRef Name="AssignedTo" /></GroupBy><OrderBy><FieldRef Name="PublishBy" /></OrderBy><Where><And><And><Or><And><And><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Content Submission Request</Value></Eq><Geq><FieldRef Name="PublishBy" /><Value Type="DateTime"><Today /></Value></Geq></And><Leq><FieldRef Name="PublishBy" /><Value Type="DateTime"><Today OffsetDays="14" /></Value></Leq></And><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Content Submission Request - Stage 2</Value></Eq></Or><Geq><FieldRef Name="PublishBy" /><Value Type="DateTime"><Today /></Value></Geq></And><Leq><FieldRef Name="PublishBy" /><Value Type="DateTime"><Today OffsetDays="14" /></Value></Leq></And></Where>'
                    'TargetSite' = ''
                    'Title' = 'Content - Due in the Next 14 Days'
                    'ViewFields' = $viewFields.Content
                },
                [PSCustomObject]@{
                    'Query' = '<OrderBy><FieldRef Name="ID" Ascending="FALSE" /></OrderBy><Where><And><Eq><FieldRef Name="ContentSubmissionStatus" /><Value Type="Text">Pending Approval</Value></Eq><Or><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Event Submission Request</Value></Eq><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Event Submission Request - Stage 2</Value></Eq></Or></And></Where>'
                    'TargetSite' = ''
                    'Title' = 'Events - All Pending Submissions'
                    'ViewFields' = $viewFields.Events
                },
                [PSCustomObject]@{
                    'Query' = '<OrderBy><FieldRef Name="ID" Ascending="FALSE" /></OrderBy><Where><And><Or><Eq><FieldRef Name="ContentSubmissionStatus" /><Value Type="Text">Approved</Value></Eq><Eq><FieldRef Name="ContentSubmissionStatus" /><Value Type="Text">Rejected</Value></Eq></Or><Or><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Event Submission Request</Value></Eq><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Event Submission Request - Stage 2</Value></Eq></Or></And></Where>'
                    'TargetSite' = ''
                    'Title' = 'Events - All Processed Submissions'
                    'ViewFields' = $viewFields.Events
                },
                [PSCustomObject]@{
                    'Query' = '<GroupBy Collapse="FALSE" GroupLimit="30"><FieldRef Name="AssignedTo" /></GroupBy><Where><Or><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Event Submission Request</Value></Eq><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Event Submission Request - Stage 2</Value></Eq></Or></Where>'
                    'TargetSite' = ''
                    'Title' = 'Events - By Assigned To'
                    'ViewFields' = $viewFields.Events
                },
                [PSCustomObject]@{
                    'Query' = '<GroupBy Collapse="FALSE" GroupLimit="30"><FieldRef Name="AssignedTo" /></GroupBy><OrderBy><FieldRef Name="EventDateTime" /></OrderBy><Where><And><And><And><Or><And><And><And><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Event Submission Request</Value></Eq><Geq><FieldRef Name="EventDateTime" /><Value Type="DateTime"><Today /></Value></Geq></And><Leq><FieldRef Name="EventDateTime" /><Value Type="DateTime"><Today OffsetDays="7" /></Value></Leq></And><Neq><FieldRef Name="ContentSubmissionStatus" /><Value Type="Text">Rejected</Value></Neq></And><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Event Submission Request - Stage 2</Value></Eq></Or><Geq><FieldRef Name="EventDateTime" /><Value Type="DateTime"><Today /></Value></Geq></And><Leq><FieldRef Name="EventDateTime" /><Value Type="DateTime"><Today OffsetDays="7" /></Value></Leq></And><Neq><FieldRef Name="ContentSubmissionStatus" /><Value Type="Text">Rejected</Value></Neq></And></Where>'
                    'TargetSite' = ''
                    'Title' = 'Events - Due in the Next 07 Days'
                    'ViewFields' = $viewFields.Events
                },
                [PSCustomObject]@{
                    'Query' = '<GroupBy Collapse="FALSE" GroupLimit="30"><FieldRef Name="AssignedTo" /></GroupBy><OrderBy><FieldRef Name="EventDateTime" /></OrderBy><Where><And><And><And><Or><And><And><And><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Event Submission Request</Value></Eq><Geq><FieldRef Name="EventDateTime" /><Value Type="DateTime"><Today /></Value></Geq></And><Leq><FieldRef Name="EventDateTime" /><Value Type="DateTime"><Today OffsetDays="14" /></Value></Leq></And><Neq><FieldRef Name="ContentSubmissionStatus" /><Value Type="Text">Rejected</Value></Neq></And><Eq><FieldRef Name="ContentType" /><Value Type="Computed">Event Submission Request - Stage 2</Value></Eq></Or><Geq><FieldRef Name="EventDateTime" /><Value Type="DateTime"><Today /></Value></Geq></And><Leq><FieldRef Name="EventDateTime" /><Value Type="DateTime"><Today OffsetDays="14" /></Value></Leq></And><Neq><FieldRef Name="ContentSubmissionStatus" /><Value Type="Text">Rejected</Value></Neq></And></Where>'
                    'TargetSite' = ''
                    'Title' = 'Events - Due in the Next 14 Days'
                    'ViewFields' = $viewFields.Events
                }
            )
        }
    }

    $view = Get-PnPView -List $list -Identity "All Items"

    if($null -ne $view)
    {
        $view = Set-PnPView -List $list -Identity $view.Title -Fields $viewFields.Default

        $ctx.Load($view)
        $ctx.ExecuteQuery()
        $view.ViewQuery = '<OrderBy><FieldRef Name="ID" Ascending="FALSE" /></OrderBy><Where><And><Neq><FieldRef Name="ContentType" /><Value Type="Computed">Content Submission Request</Value></Neq><Neq><FieldRef Name="ContentType" /><Value Type="Computed">Event Submission Request</Value></Neq></And></Where>'
        $view.Update()

        $ctx.ExecuteQuery()

        Write-Host "`nLIST DEFAULT VIEW '$($view.Title)' UPDATED WITH NEW FIELDS" -ForegroundColor Green 
    }

    foreach($viewConfig in $viewConfiguration)
    {
        # If this is a view for a specific-site and the site we're on is not that site then we skip
        if($viewConfig.TargetSite.Length -gt 0 -and $viewConfig.TargetSite -ne $site.Abbreviation)
        {
            continue;
        }

        $title = $viewConfig.viewTitle
        $view = Get-PnPView -List $list -Identity $viewConfig.Title -ErrorAction SilentlyContinue

        if($null -eq $view -and $null -ne $viewConfig.ViewFields)
        {
            $view = Add-PnPView -List $list -Title $viewConfig.Title -Fields $viewConfig.ViewFields -Query $viewConfig.Query
            Write-Host "VIEW '$($viewConfig.Title)' ADD TO THE LIST" -ForegroundColor Green
        }
        elseif($null -eq $viewConfig.ViewFields)
        {
            Write-Host "THE VIEW FIELDS FOR '$($viewConfig.Title)' IS MISSING. SKIPPING THIS VIEW." -ForegroundColor Red
        }
        elseif($null -ne $view)
        {
            $view = Set-PnPView -List $list -Identity $view.Title -Fields $viewConfig.ViewFields

            $ctx.Load($view)
            $ctx.ExecuteQuery()

            $view.ViewQuery = $viewConfig.Query
            $view.Update()

            $ctx.ExecuteQuery()

            Write-Host "LIST DEFAULT VIEW '$($view.Title)' FIELDS AND FILTER REFRESHED" -ForegroundColor Yellow 
        }
        else
        {
            Write-Host "THE VIEW '$($viewConfig.Title)' ALREADY EXISTS" -ForegroundColor Yellow
        }
    }
  
    # Set unique permissions for the list so anyone on the site can add an item
    if($null -ne $site.GroupPrefix -and $site.GroupPrefix.Length -gt 0)
    {
        Write-Host "`nCUSTOMISING LIST PERMISSIONS" -ForegroundColor Green
        Set-PnpList -Identity $list -BreakRoleInheritance
    
        Set-PnPListPermission -Identity $list -Group "$($site.GroupPrefix) Owners" -AddRole "Full Control"
        Write-Host "'$($site.GroupPrefix) Owners' given Full Control" -ForegroundColor Yellow
    
        Set-PnPListPermission -Identity $list -Group "$($site.GroupPrefix) Members" -AddRole "Edit"
        Write-Host "'$($site.GroupPrefix) Members' given Edit permissions to the list" -ForegroundColor Yellow

        Set-PnPListPermission -Identity $list -Group "$($site.GroupPrefix) Visitors" -AddRole "Contribute"
        Write-Host "'$($site.GroupPrefix) Visitors' given Contribute permissions to the list" -ForegroundColor Yellow
    }

    # Set the list's valiation formula and message
    Write-Host "Setting the list's validation formula and message" -ForegroundColor Green
    $list = Get-PnPList -Identity $list
    $list.ValidationFormula = "=IF(NOT(ISBLANK([Event Start Date/Time])),[Event End Date/Time]>[Event Start Date/Time],TRUE)"
    $list.ValidationMessage = "The Event End Date/Time must be after the Event Start Date/Time."
    $list.Update()
    $ctx.ExecuteQuery()

    Write-Host ""
}

Write-Host "SCRIPT FINISHED" -ForegroundColor Yellow
Stop-Transcript