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
$global:environment = 'prod' # Which environment are we targetting?

# GLOBAL VARIABLES

# GLOBAL SCRIPT SETTINGS
if($global:PSScriptRoot.Length -gt 0)
{
    New-Item -ItemType Directory -Force -Path "$global:PSScriptRoot\Logs" | Out-Null
}
else
{ 
    New-Item -ItemType Directory -Force -Path "./Logs" | Out-Null
}

function Invoke-Configuration
{
    param (
        [string]$env = $global:environment
    )

    # TENANT-SPECIFC SETTINGS
    switch($env) {
        dev {
            $global:adminURL = 'https://defradev-admin.sharepoint.com'
            $global:rootURL = 'https://defradev.sharepoint.com'
            $global:termSetPath = 'DEFRA EDRM UAT|Organisational Unit|Defra Orgs'

            $global:sites = @(
                [PSCustomObject]@{
                    'Abbreviation' = 'APHA'
                    'DisplayName' = 'Animal & Plant Health Agency'
                    'GroupPrefix' = 'APHAIntranet'
                    'RelativeURL' = 'sites/APHAIntranet'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'Defra'
                    'DisplayName' = 'Defra Intranet'
                    'GroupPrefix' = 'DefraIntranet'
                    'RelativeURL' = 'sites/defraintranet'
                    'SiteType' = 'Parent'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'EA'
                    'DisplayName' = 'Environment Agency'
                    'GroupPrefix' = 'EAIntranet'
                    'RelativeURL' = 'sites/EAIntranet'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'MMO'
                    'DisplayName' = 'Marine Management Organisation'
                    'GroupPrefix' = 'MMOIntranet'
                    'RelativeURL' = 'sites/MMOIntranet'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'NE'
                    'DisplayName' = 'Natural England Intranet'
                    'GroupPrefix' = 'NEIntranet'
                    'RelativeURL' = 'sites/NEIntranet'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'RPA'
                    'DisplayName' = 'Rural Payments Agency'
                    'GroupPrefix' = 'RPAIntranet'
                    'RelativeURL' = 'sites/RPAIntranet'
                    'SiteType' = 'ALB'
                }
            )
        }
      
        preprod {
            $global:adminURL = 'https://defra-admin.sharepoint.com'
            $global:rootURL = 'https://defra.sharepoint.com'
            $global:termSetPath = 'DEFRA EDRM UAT|Organisational Unit|Defra Orgs'

            $global:sites = @(
                [PSCustomObject]@{
                    'Abbreviation' = 'APHA'
                    'DisplayName' = 'PreProd-APHA'
                    'GroupPrefix' = 'PreProd-APHA'
                    'RelativeURL' = '/sites/PPAPHAIntranet'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'Defra'
                    'DisplayName' = 'PreProd-Defra Intranet'
                    'GroupPrefix' = 'PreProd-Defra Intranet'
                    'RelativeURL' = '/sites/PPDefraIntranet'
                    'SiteType' = 'Parent'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'EA'
                    'DisplayName' = 'PreProd-Environment Agency Intranet'
                    'GroupPrefix' = 'PreProd-Environment Agency Intranet'
                    'RelativeURL' = '/sites/PPEAIntranet'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'MMO'
                    'DisplayName' = 'Preprod-MMO Connect'
                    'GroupPrefix' = 'Preprod-MMO Connect'
                    'RelativeURL' = '/sites/PPMMOIntranet'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'NE'
                    'DisplayName' = 'PreProd-Natural England Intranet'
                    'GroupPrefix' = 'NEIntranet'
                    'RelativeURL' = '/sites/PPNEIntranet'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'RPA'
                    'DisplayName' = 'Preprod-Rural Payments Agency'
                    'GroupPrefix' = 'RPAIntranet'
                    'RelativeURL' = '/sites/PPRPAIntranet'
                    'SiteType' = 'ALB'
                }
            )
        }

        production {
            $global:adminURL = 'https://defra-admin.sharepoint.com'
            $global:rootURL = 'https://defra.sharepoint.com'
            $global:termSetPath = 'DEFRA EDRM UAT|Organisational Unit|Defra Orgs'

            $global:sites = @(
                [PSCustomObject]@{
                    'Abbreviation' = 'APHA'
                    'DisplayName' = 'Animal & Plant Health Agency'
                    'GroupPrefix' = 'APHAIntranet'
                    'RelativeURL' = '/sites/APHAIntranet'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'Defra'
                    'DisplayName' = 'Defra Intranet'
                    'GroupPrefix' = 'Defraintranet'
                    'RelativeURL' = '/sites/DefraIntranet'
                    'SiteType' = 'Parent'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'EA'
                    'DisplayName' = 'Environment Agency'
                    'GroupPrefix' = 'EAIntranet'
                    'RelativeURL' = '/sites/EAIntranet'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'MMO'
                    'DisplayName' = 'Marine Management Organisation'
                    'GroupPrefix' = 'mmointranet'
                    'RelativeURL' = '/sites/MMOIntranet'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'NE'
                    'DisplayName' = 'Natural England Intranet'
                    'GroupPrefix' = 'NEIntranet'
                    'RelativeURL' = '/sites/NEIntranet'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'RPA'
                    'DisplayName' = 'Rural Payments Agency'
                    'GroupPrefix' = 'RPAIntranet'
                    'RelativeURL' = 'sites/RPAIntranet'
                    'SiteType' = 'ALB'
                }
            )
        }

        # JB Development Environment
        local001 {
            $global:adminURL = 'https://buckinghamdevelopment-admin.sharepoint.com'
            $global:rootURL = 'https://buckinghamdevelopment.sharepoint.com'
            $global:termSetPath = 'DEFRA EDRM UAT|Organisational Unit|Defra Orgs'

            $global:sites = @(
                [PSCustomObject]@{
                    'Abbreviation' = 'APHA'
                    'DisplayName' = 'Animal & Plant Health Agency'
                    'GroupPrefix' = 'APHAIntranet'
                    'RelativeURL' = 'sites/APHAIntranet'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'Defra'
                    'DisplayName' = 'Defra Intranet'
                    'GroupPrefix' = 'Defra Intranet'
                    'RelativeURL' = 'sites/DefraIntranet'
                    'SiteType' = 'Parent'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'EA'
                    'DisplayName' = 'Environment Agency'
                    'GroupPrefix' = 'DEFRA001'
                    'RelativeURL' = 'sites/EAIntranet'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'MMO'
                    'DisplayName' = 'Marine Management Organisation'
                    'GroupPrefix' = 'MMOIntranet'
                    'RelativeURL' = 'sites/MMOIntranet'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'NE'
                    'DisplayName' = 'Natural England Intranet'
                    'GroupPrefix' = 'NEIntranet'
                    'RelativeURL' = 'sites/NEIntranet'
                    'SiteType' = 'ALB'
                },
                [PSCustomObject]@{
                    'Abbreviation' = 'RPA'
                    'DisplayName' = 'Rural Payments Agency'
                    'GroupPrefix' = 'DEFRA002'
                    'RelativeURL' = 'sites/RPAIntranet'
                    'SiteType' = 'ALB'
                }
            )
        }
    }
}
