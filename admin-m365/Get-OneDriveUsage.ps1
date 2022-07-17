

##### New approach. It must be possible without SamAccountName... 



# https://docs.microsoft.com/en-us/onedrive/change-user-storage
# https://docs.microsoft.com/en-us/onedrive/list-onedrive-urls
# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/sort-object?view=powershell-7.2


$onedrives = Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like '-my.sharepoint.com/personal/'"

$report = foreach ($OneDrive in $OneDrives) {
    try {
        [PSCustomObject]@{
            # https://ss64.com/ps/syntax-f-operator.html
            Owner          = $OneDrive.Owner
            CurrentUsageGB = "{0:n3}" -f (($OneDrive.StorageUsageCurrent / 1024) -as [decimal])
            TotalStorageGB = "{0:n0}" -f (($OneDrive.StorageQuota / 1024) -as [int])
            Status         = $OneDrive.Status
        }

    } catch {
        Write-Error $_.Exception.Message

    }
}


$report | Sort-Object {[decimal]$_.CurrentUsageGB}


break

Function Get-OneDriveUsage {
    <#
    .SYNOPSIS
        This will check OneDrive current usage and total limit.
     
    .NOTES
        Name: Get-OneDriveUsage
        Author: theSysadminChannel
        Version: 1.0
        DateCreated: 2020-Sep-20

        Modified by: CyberblueX
        Version: 1.1
        Date: 2022-07
        Comments: Added Parameter Support for Tenant and Domainname
     
    .LINK
        https://thesysadminchannel.com/check-onedrive-usage-for-users-in-office-365 -
    #>
     
        [CmdletBinding()]
        param(
            [Parameter(
                Mandatory=$false,
                ValueFromPipeline=$true,
                ValueFromPipelineByPropertyName=$true,
                Position=0
            )]
     
            [string[]]
            $SamAccountName = $env:USERNAME,

            # Parameter help description
            [Parameter(
                Mandatory=$true,
                ValueFromPipelineByPropertyName=$true
            )]
            [string]
            $TenantName,

            # Bsp. microsoft.com
            [Parameter(
                Mandatory=$true,
                ValueFromPipelineByPropertyName=$true
            )]
            [string]
            $DomainName
     
        )
     
        BEGIN {
            #
            $DomainName = $DomainName.Replace(".","_")
        }
     
        PROCESS {
            foreach ($User in $SamAccountName) {
                try {
                    $URL = "https://$($TenantName)-my.sharepoint.com/personal/$($User)_$($DomainName)"
                    $Stats = Get-SPOSite -Identity $URL | select Owner, StorageUsageCurrent, StorageQuota, Status
                    [PSCustomObject]@{
                        # https://ss64.com/ps/syntax-f-operator.html
                        Owner          = $Stats.Owner
                        CurrentUsageGB = "{0:n3}" -f (($Stats.StorageUsageCurrent / 1024) -as [decimal])
                        TotalStorageGB = "{0:n0}" -f (($Stats.StorageQuota / 1024) -as [int])
                        Status         = $Stats.Status
                    }
     
                } catch {
                    Write-Error $_.Exception.Message
     
                }
            }
        }
     
        END {}
    }

