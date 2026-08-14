# This function creates a consolidated Teams report
# using Adaptive Cards.
function New-TeamsCardJson {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        $InputObject,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Title
    )

    begin {
        $moduleInfo = Get-Module $($MyInvocation.MyCommand.ModuleName)

        function ConvertTo-AdaptiveCardElementId {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                [ValidateNotNullOrEmpty()]
                [string]$Value
            )

            return ($Value -replace '[^a-zA-Z0-9_-]', '-')
        }

        function Format-ServiceHealthCardDate {
            [CmdletBinding()]
            param(
                [AllowNull()]
                $DateTime
            )

            if (!$DateTime) {
                return ''
            }

            return '{0:MMMM dd, yyyy hh:mm tt} UTC' -f [datetime]$DateTime
        }

        function Format-ServiceHealthText {
            [CmdletBinding()]
            param(
                [AllowNull()]
                [string]$Text
            )

            if ([System.String]::IsNullOrWhiteSpace($Text)) {
                return ''
            }

            if ($Text.Length -eq 1) {
                return $Text.ToUpperInvariant()
            }

            return $Text.Substring(0, 1).ToUpperInvariant() + $Text.Substring(1)
        }

        function Get-ServiceHealthPortalStyleUpdateText {
            [CmdletBinding()]
            param(
                [AllowNull()]
                [string]$Text
            )

            if ([System.String]::IsNullOrWhiteSpace($Text)) {
                return ''
            }

            $normalizedText = $Text.Trim()

            $currentStatusPattern = '(?is)Current status:\s*(?<CurrentStatus>.*?)(?=\r?\n\s*\r?\n(?:Scope of impact:|Start time:|Root cause:|Next update by:|Title:|User impact:|More info:|Final status:|End time:|Next steps:)|\z)'
            $nextUpdatePattern = '(?is)Next update by:\s*(?<NextUpdate>.*?)(?=\r?\n\s*\r?\n(?:Title:|User impact:|More info:|Current status:|Final status:|Scope of impact:|Start time:|Root cause:|End time:|Next steps:)|\z)'

            $finalStatusPattern = '(?is)Final status:\s*(?<FinalStatus>.*?)(?=\r?\n\s*\r?\n(?:Scope of impact:|Start time:|End time:|Root cause:|Next steps:|Next update by:|Title:|User impact:|More info:|Current status:)|\z)'
            $finalUpdatePattern = '(?im)^\s*This is the final update for the event\.\s*$'

            $currentStatusMatch = [System.Text.RegularExpressions.Regex]::Match(
                $normalizedText,
                $currentStatusPattern
            )

            if ($currentStatusMatch.Success) {
                $currentStatus = $currentStatusMatch.Groups['CurrentStatus'].Value.Trim()

                $nextUpdateMatch = [System.Text.RegularExpressions.Regex]::Match(
                    $normalizedText,
                    $nextUpdatePattern
                )

                if ($nextUpdateMatch.Success) {
                    $nextUpdate = $nextUpdateMatch.Groups['NextUpdate'].Value.Trim()

                    if (![System.String]::IsNullOrWhiteSpace($nextUpdate)) {
                        return $currentStatus + "`r`n`r`nNext update by: " + $nextUpdate
                    }
                }

                return $currentStatus
            }

            $finalStatusMatch = [System.Text.RegularExpressions.Regex]::Match(
                $normalizedText,
                $finalStatusPattern
            )

            if ($finalStatusMatch.Success) {
                $finalStatus = $finalStatusMatch.Groups['FinalStatus'].Value.Trim()

                $finalUpdateMatch = [System.Text.RegularExpressions.Regex]::Match(
                    $normalizedText,
                    $finalUpdatePattern
                )

                if ($finalUpdateMatch.Success) {
                    return $finalStatus + "`r`n`r`n" + $finalUpdateMatch.Value.Trim()
                }

                return $finalStatus
            }

            return $normalizedText
        }

        function Get-ServiceHealthLatestUpdateText {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                [ValidateNotNull()]
                $Issue
            )

            $latestPost = if ($Issue.Posts -and $Issue.Posts.Count -gt 0) {
                $Issue.Posts[-1]
            }

            if (
                !$latestPost -or
                !$latestPost.Description -or
                [System.String]::IsNullOrWhiteSpace($latestPost.Description.Content)
            ) {
                return 'No latest update available.'
            }

            return Get-ServiceHealthPortalStyleUpdateText -Text $latestPost.Description.Content
        }

        function Get-ServiceHealthClassificationColor {
            [CmdletBinding()]
            param(
                [AllowNull()]
                [string]$Classification
            )

            if ([System.String]::IsNullOrWhiteSpace($Classification)) {
                return 'Warning'
            }

            switch ($Classification.Trim().ToLowerInvariant()) {
                'incident' { return 'Attention' }
                default    { return 'Warning' }
            }
        }

        function Get-ServiceHealthHeaderStyle {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                $Issue
            )

            if ($Issue.IsResolved) {
                return 'good'
            }

            return 'attention'
        }

        function New-ServiceHealthCardHeader {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                [string]$Title,

                [Parameter(Mandatory)]
                $ReportGeneratedDate
            )

            [pscustomobject][ordered]@{
                type  = 'Container'
                style = 'emphasis'
                bleed = $true
                items = @(
                    [pscustomobject][ordered]@{
                        type                = 'TextBlock'
                        wrap                = $true
                        weight              = 'Bolder'
                        text                = $Title
                        size                = 'Large'
                        horizontalAlignment = 'Center'
                    },
                    [pscustomobject][ordered]@{
                        type                = 'TextBlock'
                        wrap                = $true
                        text                = (Get-Date ($ReportGeneratedDate.ToLocalTime()) -Format F)
                        horizontalAlignment = 'Center'
                    }
                )
            }
        }

        function New-ServiceHealthIssueHeader {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                $Issue
            )

            $elementId = ConvertTo-AdaptiveCardElementId -Value $Issue.Id
            $detailsId = 'v' + $elementId
            $toggleDownId = 'toggle_' + $elementId + 'Down'
            $toggleUpId = 'toggle_' + $elementId + 'Up'

            $classificationColor = Get-ServiceHealthClassificationColor -Classification $Issue.Classification
            $containerStyle = Get-ServiceHealthHeaderStyle -Issue $Issue

            [pscustomobject][ordered]@{
                type         = 'Container'
                style        = $containerStyle
                bleed        = $true
                showBorder   = $true
                selectAction = [pscustomobject][ordered]@{
                    type           = 'Action.ToggleVisibility'
                    targetElements = @(
                        $detailsId,
                        $toggleDownId,
                        $toggleUpId
                    )
                }
                items        = @(
                    [pscustomobject][ordered]@{
                        type    = 'ColumnSet'
                        id      = 'colSet_' + $elementId
                        bleed   = $true
                        columns = @(
                            [pscustomobject][ordered]@{
                                type                     = 'Column'
                                width                    = 'stretch'
                                verticalContentAlignment = 'Bottom'
                                spacing                  = 'ExtraSmall'
                                items                    = @(
                                    [pscustomobject][ordered]@{
                                        type   = 'TextBlock'
                                        text   = $Issue.Service
                                        wrap   = $true
                                        size   = 'Default'
                                        weight = 'Bolder'
                                        color  = 'Accent'
                                    },
                                    [pscustomobject][ordered]@{
                                        type   = 'TextBlock'
                                        text   = $Issue.Id
                                        wrap   = $true
                                        weight = 'Bolder'
                                        size   = 'Large'
                                        color  = $classificationColor
                                    },
                                    [pscustomobject][ordered]@{
                                        type   = 'TextBlock'
                                        text   = $Issue.Title
                                        wrap   = $true
                                        weight = 'Bolder'
                                        size   = 'Medium'
                                        color  = 'Accent'
                                    }
                                )
                            },
                            [pscustomobject][ordered]@{
                                type                     = 'Column'
                                width                    = 'auto'
                                verticalContentAlignment = 'Center'
                                spacing                  = 'ExtraSmall'
                                items                    = @(
                                    [pscustomobject][ordered]@{
                                        type         = 'Icon'
                                        name         = 'ChevronDown'
                                        id           = $toggleDownId
                                        size         = 'xxSmall'
                                        selectAction = [pscustomobject][ordered]@{
                                            type           = 'Action.ToggleVisibility'
                                            targetElements = @(
                                                $detailsId,
                                                $toggleDownId,
                                                $toggleUpId
                                            )
                                        }
                                        fallback     = [pscustomobject][ordered]@{
                                            type = 'TextBlock'
                                            text = '▼'
                                            wrap = $true
                                        }
                                    },
                                    [pscustomobject][ordered]@{
                                        type         = 'Icon'
                                        name         = 'ChevronUp'
                                        id           = $toggleUpId
                                        size         = 'xxSmall'
                                        isVisible    = $false
                                        selectAction = [pscustomobject][ordered]@{
                                            type           = 'Action.ToggleVisibility'
                                            targetElements = @(
                                                $detailsId,
                                                $toggleDownId,
                                                $toggleUpId
                                            )
                                        }
                                        fallback     = [pscustomobject][ordered]@{
                                            type      = 'TextBlock'
                                            text      = '▲'
                                            wrap      = $true
                                            isVisible = $false
                                        }
                                    }
                                )
                            }
                        )
                    }
                )
            }
        }

        function New-ServiceHealthIssueFactSet {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                $Issue
            )

            $elementId = ConvertTo-AdaptiveCardElementId -Value $Issue.Id
            $detailsId = 'v' + $elementId

            $classification = Format-ServiceHealthText -Text $Issue.Classification

            $resolutionState = if ($Issue.IsResolved) {
                'Resolved'
            }
            else {
                'Active'
            }

            $status = $resolutionState + ' | ' + $Issue.Status
            $latestUpdate = Get-ServiceHealthLatestUpdateText -Issue $Issue
            $adminCenterUrl = 'https://admin.cloud.microsoft/?#/servicehealth/:/alerts/' + $Issue.Id
            $adminCenterLink = '' + $adminCenterUrl + ''

            [pscustomobject][ordered]@{
                type      = 'FactSet'
                id        = $detailsId
                isVisible = $false
                separator = $true
                facts     = @(
                    [pscustomobject][ordered]@{
                        title = 'Issue type'
                        value = $classification
                    },
                    [pscustomobject][ordered]@{
                        title = 'Status'
                        value = $status
                    },
                    [pscustomobject][ordered]@{
                        title = 'User impact'
                        value = $Issue.ImpactDescription
                    },
                    [pscustomobject][ordered]@{
                        title = 'Start time'
                        value = Format-ServiceHealthCardDate -DateTime $Issue.StartDateTime
                    },
                    [pscustomobject][ordered]@{
                        title = 'End time'
                        value = Format-ServiceHealthCardDate -DateTime $Issue.EndDateTime
                    },
                    [pscustomobject][ordered]@{
                        title = 'Update time'
                        value = Format-ServiceHealthCardDate -DateTime $Issue.LastModifiedDateTime
                    },
                    [pscustomobject][ordered]@{
                        title = 'Update'
                        value = $latestUpdate
                    },
                    [pscustomobject][ordered]@{
                        title = 'Link'
                        value = $adminCenterLink
                    }
                )
            }
        }
    }

    process {
        $teamsAdaptiveCardPath = Join-Path -Path $moduleInfo.ModuleBase -ChildPath 'source\private\TeamsConsolidated.json'

        $teamsAdaptiveCard = Get-Content -Path $teamsAdaptiveCardPath -Raw | ConvertFrom-Json

        # Ensure body starts empty even if the template file is later modified.
        $teamsAdaptiveCard.attachments[0].content.body = @()

        $teamsAdaptiveCard.attachments[0].content.version = '1.5'

        if (!$teamsAdaptiveCard.attachments[0].content.msTeams) {
            $teamsAdaptiveCard.attachments[0].content | Add-Member -MemberType NoteProperty -Name 'msTeams' -Value ([pscustomobject]@{
                    width = 'full'
                })
        }

        $teamsAdaptiveCard.attachments[0].content.body += New-ServiceHealthCardHeader `
            -Title $Title `
            -ReportGeneratedDate $InputObject[0].ReportGeneratedDate

        foreach ($item in ($InputObject | Sort-Object LastModifiedDateTime -Descending)) {
            $teamsAdaptiveCard.attachments[0].content.body += New-ServiceHealthIssueHeader -Issue $item
            $teamsAdaptiveCard.attachments[0].content.body += New-ServiceHealthIssueFactSet -Issue $item
        }

        return ($teamsAdaptiveCard | ConvertTo-Json -Depth 20)
    }
}