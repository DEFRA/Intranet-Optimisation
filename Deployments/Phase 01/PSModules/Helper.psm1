<#
    SCRIPT OVERVIEW:
    This PowerShell module containers Helper methods which help them main script with repeated tasks

    SOFTWARE REQUIREMENTS:
    This script was developed on the following platform:
        PowerShell v7.3.6 - https://learn.microsoft.com/en-gb/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.3#supported-versions-of-windows
        SharePointPnPPowerShellOnline 3.29.2101.0 -  https://www.powershellgallery.com/packages/SharePointPnPPowerShellOnline/3.29.2101.0

    REQUIREMENTS:
    SHAREPOINT PERMISSIONS REQUIREMENTS:
    - Full Control of the SharePoint site the context is being run against
    OR
    - Access to the SharePoint Tenant Administration site
#>

function Get-CurrentUser
{
    param (
        [Microsoft.SharePoint.Client.ClientContext]$Ctx = $(Get-PnPContext)
    )
    
    $Ctx.Load($Ctx.Web.CurrentUser)
    $Ctx.ExecuteQuery()   
    return $Ctx.Web.CurrentUser.Email
}