<#
    SCRIPT OVERVIEW:
    This script creates the global site column(s) and a "Base Site Page Content Type Hub" content type within the Content Type Hub. 
    The column's crawled property will need to be added to a mapped properrty within the SharePoint Search Admin after creation (see the deployment guide for the mapping).
    The content type is used purely as a transporter for the column so it's available to our Intranet sites. 

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
$fields = @(
    [PSCustomObject]@{
        'DisplayName' = "Organisation (Intranets)"
        'InternalName' = "OrganisationIntranets"
    }
)

$termSetPath = $global:termSetPath

# Content Type(s)
$ctSPDisplayName = "Site Page"
$ct = Get-PnPContentType -Identity $global:ctDisplayName -ErrorAction SilentlyContinue
$library = Get-PnPList -Identity "Site Pages"
$parentCt = Get-PnPContentType -Identity $ctSPDisplayName -ErrorAction SilentlyContinue

if($null -eq $ct -and $null -ne $parentCt)
{
    $ct = Add-PnPContentType -Name $global:ctDisplayName -Description "Used for the deployment of Content Type Hub site columns to sites" -ContentTypeId $global:ctID
    Write-Host "CONTENT TYPE INSTALLED: $($global:ctDisplayName)" -ForegroundColor Green
}
elseif($null -eq $parentCt)
{
    throw "THE PARENT CONTENT TYPE '$ctSPDisplayName' CANNOT BE FOUND"
}
else
{
    Write-Host "CONTENT TYPE ALREADY INSTALLED: $($global:ctDisplayName)" -ForegroundColor Yellow
}

# "Page Category" column
foreach($objField in $fields)
{
    $field = Get-PnPField | Where-Object { $_.InternalName -eq $objField.InternalName }

    if($null -eq $field)
    {
        $field = Add-PnPTaxonomyField -DisplayName $objField.DisplayName -InternalName $objField.InternalName -TermSetPath $termSetPath -MultiValue
        Write-Host "SITE COLUMN INSTALLED: $($objField.DisplayName)" -ForegroundColor Green
    }
    else
    {
        Write-Host "SITE COLUMN ALREADY INSTALLED: $($objField.DisplayName)" -ForegroundColor Yellow
    }

    # Add the field to "Base Site Page Content Type Hub" content type
    if($null -ne $field -and $null -ne $ct)
    {
        $ctx.Load($ct.FieldLinks)
        $ctx.ExecuteQuery()

        $fieldExistsOnCT = $ct.FieldLinks | Where-Object { $_.Name -eq $objField.InternalName }

        if($null -eq $fieldExistsOnCT)
        {
            Add-PnPFieldToContentType -Field $field -ContentType $ct -Required:$false
            Write-Host "SITE COLUMN '$($objField.DisplayName)' ADDED TO THE CONTENT TYPE '$($global:ctDisplayName)'" -ForegroundColor Green
        }
        else 
        {
            Write-Host "SITE COLUMN '$($objField.DisplayName)' IS ALREADY INSTALLED IN THE CONTENT TYPE '$($global:ctDisplayName)'" -ForegroundColor Yellow
        }
    }
    else
    {
        Write-Host "THE SITE COLUMN '$($objField.DisplayName)' WAS NOT ADDED TO THE CONTENT TYPE '$($global:ctDisplayName)'. THE FIELD COULD NOT BE FOUND." -ForegroundColor Red
    }

    # Add the field to "Site Page" content type at the library-level (DefraDev looks to be sealed at the Site Column level so we can't add it that way)
    $ctSP = Get-PnPContentType -Identity $ctSPDisplayName -List $library

    if($null -ne $field -and $null -ne $ctSP)
    {
        $ctx.Load($ctSP.FieldLinks)
        $ctx.ExecuteQuery()

        $fieldExistsOnCT = $ctSP.FieldLinks | Where-Object { $_.Name -eq $objField.InternalName }

        if($null -eq $fieldExistsOnCT)
        {
            $fieldLink = New-Object Microsoft.SharePoint.Client.FieldLinkCreationInformation
            $fieldLink.Field = $field
            $ctSP.FieldLinks.Add($fieldLink) | Out-Null
            $ctSP.Update($false)
            $ctx.ExecuteQuery()

            Write-Host "SITE COLUMN '$($objField.DisplayName)' ADDED TO THE '$($library.Title)' LIBRARY CONTENT TYPE '$ctSPDisplayName'" -ForegroundColor Green
        }
        else 
        {
            Write-Host "SITE COLUMN '$($objField.DisplayName)' IS ALREADY INSTALLED IN THE LIBRARY CONTENT TYPE '$ctSPDisplayName'" -ForegroundColor Yellow
        }
    }
    else
    {
        Write-Host "THE SITE COLUMN '$($objField.DisplayName)' WAS NOT ADDED TO THE CONTENT TYPE '$ctSPDisplayName'. THE FIELD COULD NOT BE FOUND." -ForegroundColor Red
    }
}

# Create a site page so we have content associated with our new site column, this is so the column becomes available as a crawled property and is then exposes to SharePoint's Central Search Admin   
if($null -ne $library -and $null -ne $ctSP)
{
    Write-Host "INITIALISING THE COLUMN'S CRAWLED PROPERTY" -ForegroundColor Green
    $page = Get-PnPListItem -List $library -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($global:placeholderSitePageName)</Value></Eq></Where></Query></View>"

    if($null -eq $page)
    {
        # NOTE: This page can be deleted once our new crawled property appears within the site's search schema
        $page = $library.RootFolder.Files.AddTemplateFile($library.RootFolder.ServerRelativeUrl + "/$($global:placeholderSitePageName).aspx",[Microsoft.SharePoint.Client.TemplateFileType]::ClientSidePage).ListItemAllFields
        $page["ContentTypeId"] = "0x0101009D1CB255DA76424F860D91F20E6C4118";
        $page["Title"] = $global:placeholderSitePageName
        $page["ClientSideApplicationId"] = "b6917cb1-93a0-4b97-a84d-7cf49975d4ec"
        $page["PageLayoutType"] = "Article"
        $page["PromotedState"] = "0"
        $page["CanvasContent1"] = "<div></div>"
        $page["BannerImageUrl"] = "/_layouts/15/images/sitepagethumbnail.png"
        $page.Update()
        $ctx.ExecuteQuery()
    }

    if($null -ne $page)
    {
        $ctx.Load($page)
        $ctx.ExecuteQuery()

        $taxonomySession = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($ctx);
        $taxonomySession.UpdateCache()
        $termStore =$taxonomySession.GetDefaultSiteCollectionTermStore()
        $ctx.Load($taxonomySession)
        $ctx.Load($termStore)
        $ctx.ExecuteQuery()
 
        $termGroup = $termStore.Groups.GetByName($global:termSetPath.Split('|')[0])
        $ctx.Load($termGroup)
 
        $termSet = $termGroup.TermSets.GetByName($global:termSetPath.Split('|')[1])
        $ctx.Load($termSet)
        $ctx.Load($termSet.Terms)
        $ctx.ExecuteQuery()

        $term = $termSet.Terms | Where-Object {$_.Name -eq $($global:termSetPath.Split('|')[2]) }

        $termValue = "-1;#" + $term.Name + "|" + $term.Id

        foreach($objField in $fields) 
        {
            $field = $library.Fields.GetByInternalNameOrTitle($objField.InternalName)
            $ctx.Load($field)
            $ctx.ExecuteQuery()
            
            $taxField = [Microsoft.SharePoint.Client.ClientContext].GetMethod("CastTo").MakeGenericMethod([Microsoft.SharePoint.Client.Taxonomy.TaxonomyField]).Invoke($ctx,$field)
            $taxFieldValues = New-Object Microsoft.SharePoint.Client.Taxonomy.TaxonomyFieldValueCollection($ctx,$termValue,$taxField)
            $taxField.SetFieldValueByValueCollection($page,$taxFieldValues)
            $page.Update()
            $ctx.ExecuteQuery()
        }
    }

    # Request a reindex of the site
    $web = Get-PnPWeb
    [Int]$SearchVersion = 0
    
    # Get the Search Version Property - If it exists
    if ($Web.AllProperties.FieldValues.ContainsKey("vti_searchversion") -eq $True)
    {
        $SearchVersion = $Web.AllProperties["vti_searchversion"]
    }

    # Increment the search version and trigger a search reindex 
    $SearchVersion++
    $Web.AllProperties["vti_searchversion"] = $SearchVersion
    $web.Update()
    $ctx.ExecuteQuery()

    Write-Host "REINDEXING THE 'CONTENT TYPE HUB' SITE ON THE NEXT SCHEDULED SEARCH CRAWL" -ForegroundColor Yellow
}

Write-Host "SCRIPT FINISHED" -ForegroundColor Yellow
Stop-Transcript
