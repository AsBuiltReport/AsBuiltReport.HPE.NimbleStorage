#Requires -Modules PScribo, HPENimblePowerShellToolkit

function Invoke-AsBuiltReport.HPE.NimbleStorage {
    <#
    .SYNOPSIS
        PowerShell script which documents the configuration of HPE Nimble Storage Arrays in Word/HTML/XML/Text formats
    .DESCRIPTION
        Documents the configuration of HPE Nimble Storage Arrays in Word/HTML/XML/Text formats using PScribo.
    .NOTES
        Version:        0.1
        Author:         Matt Allford
        Twitter:        @mattallford
        Github:         https://github.com/mattallford
        Credits:        Iain Brighton (@iainbrighton) - PScribo module

    .LINK
        https://github.com/AsBuiltReport/
    #>

    #region Script Parameters
    [CmdletBinding()]
    param (
        $Target,
        [pscredential] $Credential
    )

    # If custom style not set, use default style
    if (!$StyleName) {
        & "$PSScriptRoot\..\Assets\Styles\NimbleStorage.ps1"
    }

    $Script:Array = $Null
    foreach ($NimbleGroup in $Target) {
        Try {
            $ConnectedNimbleGroup = Connect-NSGroup -Group $NimbleGroup -Credential $Credential -IgnoreServerCertificate
        } Catch {
            Write-Verbose "Unable to connect to the Nimble Storage Group $NimbleGroup"
        }
    
        if ($ConnectedNimbleGroup) {
            $script:Arrays = Get-NSArray | Sort-Object Name

            Foreach ($Array in $Arrays) {
                $script:ArrayControllers = Get-NSController -array_id $Array.id  | Sort-Object Name
                $script:Volumes = Get-NSVolume
                $Script:VolumeCollections = Get-NSVolumeCollection
                $script:Users = Get-NSUser
                $script:UserGroups = Get-NSUserGroup
                $script:Disks = Get-NSDisk -array_id $Array.id  | Sort-Object Type

                Section -Style Heading1 $Array.name {
                    Section -Style Heading2 'System Summary' {
                        Paragraph "The following section provides a summary of the array configuration for $($Array.name)."
                        BlankLine
                        $ArraySummary = [PSCustomObject] @{
                            'Name' = $Array.name
                            'Model' = $Array.model
                            'Serial' = $Array.serial
                            'ID' = $Array.id
                            'Role' = $Array.Role
                            'Pool Name' = $Array.pool_name
                            'Pool ID' = $array.pool_id
                            'Version' = $Array.version
                            'All Flash' = $Array.all_flash
                            #'Number of Volumes' = 
                        }
                        $ArraySummary | Table -Name 'Array Summary'
                    }#End Section Heading2 'System Summary'

                    Section -Style Heading2 'Storage Summary' {
                        Paragraph "The following section provides a summary of the array storage for $($Array.name)."
                        BlankLine
                        $ArrayStorageSummary = [PSCustomObject] @{
                            'Raw Capacity' = "$([math]::Round(($Array.raw_capacity_bytes) / 1TB, 2)) TB"
                            'Usable Capacity' = "$([math]::Round(($Array.usable_capacity_bytes) / 1TB, 2)) TB"
                            'Used' = "$([math]::Round(($Array.usage) / 1TB, 2)) TB"
                            'Free' = "$([math]::Round(($Array.available_bytes) / 1TB, 2)) TB"
                            'Volume Usage' = "$([math]::Round(($Array.vol_usage_bytes) / 1TB, 2)) TB"
                            'Volume Compression' = $array.vol_compression
                            "Volume Saved Space" = "$([math]::Round(($Array.vol_saved_bytes) / 1TB, 2)) TB"
                            'Snapshot Usage' = "$([math]::Round(($Array.snap_usage_bytes) / 1TB, 2)) TB"
                            'Snapshot Compression' = $array.snap_compression
                            'Snapshot Saved Space' = "$([math]::Round(($Array.snap_saved_bytes) / 1TB, 2)) TB"
                        }
                        $ArrayStorageSummary | Table -Name 'Array Storage Summary'
                    }#End Section Heading2 'Storage Summary'

                    Section -Style Heading2 'Controller Summary' {
                        Paragraph "The following section provides a summary of the controllers in $($Array.name)."
                        BlankLine
                        $ArrayControllerSummary = foreach ($ArrayController in $ArrayControllers) {
                            [PSCustomObject] @{
                                'Name' = $ArrayController.name
                                'ID' = $ArrayController.id
                                'Serial' = $ArrayController.serial
                                'Support Address' = $ArrayController.support_address
                                'State' = $ArrayController.state
                            }
                        }
                        $ArrayControllerSummary | Table -Name 'Controller Summary'
                    }#End Section Heading2 'Controller Summary'

                    Section -Style Heading2 'Disk Summary' {
                        Paragraph "The following section provides a summary of the disks in $($Array.name)."
                        BlankLine
                        $ArrayDiskSummary = foreach ($Disk in $disks) {
                            [PSCustomObject] @{
                                'Type' = $Disk.Type
                                'Slot' = $Disk.slot
                                'Serial' = $Disk.serial
                                'Model' = $Disk.model
                                'Size' = "$([math]::Round(($Disk.size) / 1GB, 0)) GB"
                            }
                        }
                        $ArrayDiskSummary | Table -Name 'Disk Summary'
                    }#End Section Heading2 'Disk Summary'
                }#End Section Heading1 $Array.name
            }#End Foreach ($Array in $Arrays)
        }#End if ($ConnectedNimbleGroup)
    }#End Foreach ($NimbleGroup in $Target)
}#End Function Invoke-AsBuiltReport.HPE.NimbleStorage
