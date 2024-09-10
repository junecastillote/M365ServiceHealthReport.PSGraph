# This function creates a consolidated Teams report
# using adaptive cards 1.4.
Function NewTeamsCardJson {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $InputObject,

        [Parameter(Mandatory)]
        [string]$Title
    )

    $moduleInfo = Get-Module $($MyInvocation.MyCommand.ModuleName)

    Function New-FactItem {
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

        $factSet = [pscustomobject][ordered]@{
            type      = 'FactSet'
            separator = $true
            facts     = @(
                $([pscustomobject][ordered]@{Title = 'Impact'; Value = $($InputObject.impactDescription) } ),
                $([pscustomobject][ordered]@{Title = 'Classification'; Value = ($InputObject.Classification) } ),
                $([pscustomobject][ordered]@{Title = 'Status'; Value = ($InputObject.Status) } ),
                $([pscustomobject][ordered]@{Title = 'Update'; Value = ("{0:MMMM dd, yyyy hh:mm tt}" -f [datetime]$InputObject.lastModifiedDateTime) }),
                $([pscustomobject][ordered]@{Title = 'Start'; Value = ("{0:MMMM dd, yyyy hh:mm tt}" -f [datetime]$InputObject.startDateTime) }),
                $([pscustomobject][ordered]@{Title = 'End'; Value = $(
                            if ($InputObject.endDateTime) {
                                 ("{0:MMMM dd, yyyy hh:mm tt}" -f [datetime]$InputObject.startDateTime)
                            }
                            else {
                                $null
                            }
                        )
                    }
                )
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