# This function creates a consolidated Teams report
# using adaptive cards 1.4.
function New-TeamsCardJson {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $InputObject,

        [Parameter(Mandatory)]
        [string]$Title
    )

    $moduleInfo = Get-Module $($MyInvocation.MyCommand.ModuleName)

    function New-FactItem {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            $InputObject
        )

        $factHeader = [pscustomobject][ordered]@{
            type  = "Container"
            style = "emphasis"
            bleed = $true
            items = @(
                $([pscustomobject][ordered]@{
                        type      = 'TextBlock'
                        wrap      = $true
                        separator = $true
                        weight    = 'Bolder'
                        text      = "$($InputObject.id) | $($InputObject.Service) | $($InputObject.Title)"
                    } )
            )
        }

        $latestMessage = Get-ServiceHealthLatestMessageHtml -Issue $InputObject
        $factSet = [pscustomobject][ordered]@{
            type      = 'FactSet'
            separator = $true
            facts     = @(
                $([pscustomobject][ordered]@{title ='Issue type'; value = ($InputObject.Classification) } ),
                $([pscustomobject][ordered]@{title ='Status'; value = $(
                            if ($InputObject.IsResolved) {
                                "Resolved | $($InputObject.Status)"
                            }
                            else {
                                "Active | $($InputObject.Status)"
                            }
                        )
                    } ),
                $([pscustomobject][ordered]@{title ='User impact'; value = $($InputObject.impactDescription) } ),
                $([pscustomobject][ordered]@{title ='Start time'; value = ("{0:MMMM dd, yyyy hh:mm tt}" -f [datetime]$InputObject.startDateTime) }),
                $([pscustomobject][ordered]@{title ='End time'; value = $(
                            if ($InputObject.endDateTime) {
                                ("{0:MMMM dd, yyyy hh:mm tt}" -f [datetime]$InputObject.startDateTime)
                            }
                            else {
                                [System.String]::Empty
                            }
                        )
                    }
                ),
                $([pscustomobject][ordered]@{title ='Update time'; value = ("{0:MMMM dd, yyyy hh:mm tt}" -f [datetime]$InputObject.lastModifiedDateTime) }),
                $([pscustomobject][ordered]@{title ='Update'; value = $latestMessage }),
                $([pscustomobject][ordered]@{title ='Link'; value = "https://admin.cloud.microsoft/?#/servicehealth/:/alerts/$($InputObject.Id)" })
            )
        }
        return @($factHeader, $factSet)
    }

    $teamsAdaptiveCard = ((Get-Content (($moduleInfo.ModuleBase.ToString()) + '\source\private\TeamsConsolidated.json') -Raw) | ConvertFrom-Json)

    $teamsAdaptiveCard.attachments[0].content.body += ([pscustomobject][ordered]@{
            type  = "Container"
            style = "emphasis"
            bleed = $true
            items = @(
                $([pscustomobject][ordered]@{
                        type                = 'TextBlock'
                        wrap                = $true
                        weight              = 'Bolder'
                        text                = "$($Title)"
                        size                = 'Large'
                        horizontalAlignment = 'Center'
                    } ),
                $([pscustomobject][ordered]@{
                        type                = 'TextBlock'
                        wrap                = $true
                        text                = "$(Get-Date ($InputObject[0].ReportGeneratedDate.ToLocalTime()) -Format F)"
                        horizontalAlignment = 'Center'
                    } )
            )
        })

    foreach ($item in $InputObject) {
        $teamsAdaptiveCard.attachments[0].content.body += (New-FactItem -InputObject $item)
    }

    # $teamsAdaptiveCard.attachments[0].content = $teamsAdaptiveCard.attachments[0].content | ConvertTo-Json -Depth 10
    return ($teamsAdaptiveCard | ConvertTo-Json -Depth 10)
}