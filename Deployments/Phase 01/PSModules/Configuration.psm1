<#
    SCRIPT OVERVIEW:
    This PowerShell module is the global configuration file for this deployment.

    SOFTWARE REQUIREMENTS:
    This script was developed on the following platform:
        PowerShell v5.1.22621.1778
        SharePointPnPPowerShellOnline v3.29.2101.0

    REQUIREMENTS:
    The global settings for all scripts can be found at the top of this script. For any environment-specific settings, please set these within the appropriate case statement of the Invoke-Configuration method.
#>

# GLOBAL SETTINGS
$global:environment = "preprod" # Which environment are we targetting?

# GLOBAL VARIABLES
$global:placeholderSitePageName = "Intranet Optimisation Deployment Site Page"

$global:ctDisplayName = "Base Site Page Content Type Hub"
$global:ctID = "0x0101009D1CB255DA76424F860D91F20E6C411800C78751CF7906E04285CB8A413FE2791F"

# GLOBAL SCRIPT SETTINGS
if($global:PSScriptRoot.Length -gt 0)
{
    New-Item -ItemType Directory -Force -Path "$global:PSScriptRoot\Logs" | Out-Null
}
else
{
    New-Item -ItemType Directory -Force -Path "./Logs"
}

function Invoke-Configuration
{
    param (
        [string]$env = $global:environment
    )

    # TENANT-SPECIFC SETTINGS
    switch($env) {
        dev {
            $global:adminURL = "https://defradev-admin.sharepoint.com"
            $global:rootURL = "https://defradev.sharepoint.com"
            $global:termSetPath = "DEFRA EDRM UAT|Organisational Unit|Defra Orgs"

            $global:sites = @(
                [PSCustomObject]@{
                    'Abbreviation' = "APHA"
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = "Animal & Plant Health Agency"
                    'RelativeURL' = "sites/APHAIntranet"
                    'SiteType' = "ALB"
                },
                [PSCustomObject]@{
                    'Abbreviation' = "ContentTypeHub"
                    'ApplyHubSiteNavigationChanges' = $false
                    'DisplayName' = ""
                    'RelativeURL' = "sites/ContentTypeHub"
                    'SiteType' = "System"
                },
                [PSCustomObject]@{
                    'Abbreviation' = "Defra"
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = "Defra Intranet"
                    'RelativeURL' = "sites/defraintranet"
                    'SiteType' = "Parent"
                },
                [PSCustomObject]@{
                    'Abbreviation' = "EA"
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = "Environment Agency"
                    'RelativeURL' = "sites/EAIntranet"
                    'SiteType' = "ALB"
                },
                [PSCustomObject]@{
                    'Abbreviation' = "MMO"
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = "Marine Management Organisation"
                    'RelativeURL' = "sites/MMOIntranet"
                    'SiteType' = "ALB"
                },
                [PSCustomObject]@{
                    'Abbreviation' = "NE"
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = "Natural England Intranet"
                    'RelativeURL' = "sites/NEIntranet"
                    'SiteType' = "ALB"
                },
                [PSCustomObject]@{
                    'Abbreviation' = "RPA"
                    'ApplyHubSiteNavigationChanges' = $false
                    'DisplayName' = "Rural Payments Agency"
                    'RelativeURL' = "sites/RPAIntranet"
                    'SiteType' = "ALB"
                }
            )
        }
      
        preprod {
            $global:adminURL = "https://defra-admin.sharepoint.com"
            $global:rootURL = "https://defra.sharepoint.com"
            $global:termSetPath = "DEFRA EDRM UAT|Organisational Unit|Defra Orgs"

            $global:sites = @(
                [PSCustomObject]@{
                    'Abbreviation' = "APHA"
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = "Animal & Plant Health Agency"
                    'RelativeURL' = "/sites/PPAPHAIntranet"
                    'SiteType' = "ALB"
                },
                [PSCustomObject]@{
                    'Abbreviation' = "ContentTypeHub"
                    'ApplyHubSiteNavigationChanges' = $false
                    'DisplayName' = ""
                    'RelativeURL' = "sites/ContentTypeHub"
                    'SiteType' = "System"
                },
                [PSCustomObject]@{
                    'Abbreviation' = "Defra"
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = "Defra Intranet"
                    'RelativeURL' = "/sites/PPDefraIntranet"
                    'SiteType' = "Parent"
                },
                [PSCustomObject]@{
                    'Abbreviation' = "EA"
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = "Environment Agency"
                    'RelativeURL' = "/sites/PPEAIntranet"
                    'SiteType' = "ALB"
                },
                [PSCustomObject]@{
                    'Abbreviation' = "MMO"
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = "Marine Management Organisation"
                    'RelativeURL' = "/sites/PPMMOIntranet"
                    'SiteType' = "ALB"
                },
                [PSCustomObject]@{
                    'Abbreviation' = "NE"
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = "Natural England Intranet"
                    'RelativeURL' = "/sites/PPNEIntranet"
                    'SiteType' = "ALB"
                },
                [PSCustomObject]@{
                    'Abbreviation' = "RPA"
                    'ApplyHubSiteNavigationChanges' = $false
                    'DisplayName' = "Rural Payments Agency"
                    'RelativeURL' = "/sites/PPRPAIntranet"
                    'SiteType' = "ALB"
                }
            )
        }

        production {
            $global:adminURL = "https://defra-admin.sharepoint.com"
            $global:rootURL = "https://defra.sharepoint.com"
            $global:termSetPath = "DEFRA EDRM UAT|Organisational Unit|Defra Orgs"

            $global:sites = @(
                [PSCustomObject]@{
                    'Abbreviation' = "APHA"
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = "Animal & Plant Health Agency"
                    'RelativeURL' = "/sites/APHAIntranet"
                    'SiteType' = "ALB"
                },
                [PSCustomObject]@{
                    'Abbreviation' = "ContentTypeHub"
                    'ApplyHubSiteNavigationChanges' = $false
                    'DisplayName' = ""
                    'RelativeURL' = "sites/ContentTypeHub"
                    'SiteType' = "System"
                },
                [PSCustomObject]@{
                    'Abbreviation' = "Defra"
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = "Defra Intranet"
                    'RelativeURL' = "/sites/DefraIntranet"
                    'SiteType' = "Parent"
                },
                [PSCustomObject]@{
                    'Abbreviation' = "EA"
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = "Environment Agency"
                    'RelativeURL' = "/sites/EAIntranet"
                    'SiteType' = "ALB"
                },
                [PSCustomObject]@{
                    'Abbreviation' = "MMO"
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = "Marine Management Organisation"
                    'RelativeURL' = "/sites/MMOIntranet"
                    'SiteType' = "ALB"
                },
                [PSCustomObject]@{
                    'Abbreviation' = "NE"
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = "Natural England Intranet"
                    'RelativeURL' = "/sites/NEIntranet"
                    'SiteType' = "ALB"
                },
                [PSCustomObject]@{
                    'Abbreviation' = "RPA"
                    'ApplyHubSiteNavigationChanges' = $false
                    'DisplayName' = "Rural Payments Agency"
                    'RelativeURL' = "/sites/RPAIntranet"
                    'SiteType' = "ALB"
                }
            )
        }

        # JB Development Environment
        local001 {
            $global:adminURL = "https://buckinghamdevelopment-admin.sharepoint.com"
            $global:rootURL = "https://buckinghamdevelopment.sharepoint.com"
            $global:termSetPath = "DEFRA EDRM UAT|Organisational Unit|Defra Orgs"

            $global:sites = @(
                [PSCustomObject]@{
                    'Abbreviation' = "APHA"
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = "Animal  & Plant Health Agency"
                    'RelativeURL' = "sites/DEFRA002"
                    'SiteType' = "ALB"
                },
                [PSCustomObject]@{
                    'Abbreviation' = "ContentTypeHub"
                    'ApplyHubSiteNavigationChanges' = $false
                    'DisplayName' = ""
                    'RelativeURL' = "sites/ContentTypeHub"
                    'SiteType' = "System"
                },
                [PSCustomObject]@{
                    'Abbreviation' = "Defra"
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = "Defra Intranet"
                    'RelativeURL' = "sites/DefraIntranet"
                    'SiteType' = "Parent"
                },
                [PSCustomObject]@{
                    'Abbreviation' = "EA"
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = "Environment Agency"
                    'RelativeURL' = "sites/EAIntranet"
                    'SiteType' = "ALB"
                },
                [PSCustomObject]@{
                    'Abbreviation' = "MMO"
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = "Marine Management Organisation"
                    'RelativeURL' = "sites/DEFRA002"
                    'SiteType' = "ALB"
                },
                [PSCustomObject]@{
                    'Abbreviation' = "NE"
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = "Natural England Intranet"
                    'RelativeURL' = "sites/DEFRA002"
                    'SiteType' = "ALB"
                },
                [PSCustomObject]@{
                    'Abbreviation' = "RPA"
                    'ApplyHubSiteNavigationChanges' = $true
                    'DisplayName' = "Rural Payment Agency"
                    'RelativeURL' = "sites/RPAIntranet"
                    'SiteType' = "ALB"
                }
            )
        }
    }
}
