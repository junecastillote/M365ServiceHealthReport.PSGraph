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

        function New-ServiceHealthCardHeader {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                [ValidateNotNullOrEmpty()]
                [string]$Title,

                [Parameter(Mandatory)]
                $ReportGeneratedDate,

                [Parameter(Mandatory)]
                [ValidateNotNullOrEmpty()]
                $Issues
            )

            $totalIssues = $Issues.Count
            $resolvedIssues = ($Issues | Where-Object { $_.IsResolved }).Count
            $activeIssues = ($Issues | Where-Object { !$_.IsResolved }).Count

            $incidentCount = ($Issues | Where-Object { $_.Classification -eq 'Incident' }).Count
            $advisoryCount = ($Issues | Where-Object { $_.Classification -eq 'Advisory' }).Count

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
                        spacing             = 'Small'
                    },
                    [pscustomobject][ordered]@{
                        type                = 'TextBlock'
                        wrap                = $true
                        text                = "Total Events: $totalIssues | Active: $activeIssues | Resolved: $resolvedIssues"
                        horizontalAlignment = 'Center'
                        spacing             = 'Small'
                        isSubtle            = $true
                    },
                    [pscustomobject][ordered]@{
                        type                = 'TextBlock'
                        wrap                = $true
                        text                = "Incidents: $incidentCount | Advisories: $advisoryCount"
                        horizontalAlignment = 'Center'
                        spacing             = 'None'
                        isSubtle            = $true
                    }
                )
            }
        }

        function New-ServiceHealthIssueHeader {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                [ValidateNotNull()]
                $Issue
            )

            $elementId = ConvertTo-AdaptiveCardElementId -Value $Issue.Id
            $detailsId = 'v' + $elementId
            $actionsId = 'a' + $elementId
            $toggleDownId = 'toggle_' + $elementId + 'Down'
            $toggleUpId = 'toggle_' + $elementId + 'Up'

            $classificationColor = Get-ServiceHealthClassificationColor -Classification $Issue.Classification
            $containerStyle = Get-ServiceHealthHeaderStyle -Issue $Issue
            $classificationText = Format-ServiceHealthText -Text $Issue.Classification

            $resolutionState = if ($Issue.IsResolved) {
                'Resolved'
            }
            else {
                'Active'
            }

            $statusText = $resolutionState + ' | ' + $Issue.Status
            $lastUpdatedText = 'Updated: ' + (Format-ServiceHealthCardDate -DateTime $Issue.LastModifiedDateTime)

            [pscustomobject][ordered]@{
                type         = 'Container'
                style        = $containerStyle
                bleed        = $true
                showBorder   = $true
                selectAction = [pscustomobject][ordered]@{
                    type           = 'Action.ToggleVisibility'
                    targetElements = @(
                        $detailsId,
                        $actionsId,
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
                                        type    = 'TextBlock'
                                        text    = $Issue.Service
                                        wrap    = $true
                                        size    = 'Default'
                                        weight  = 'Bolder'
                                        color   = 'Accent'
                                        spacing = 'None'
                                    },
                                    [pscustomobject][ordered]@{
                                        type    = 'TextBlock'
                                        text    = $classificationText
                                        wrap    = $true
                                        weight  = 'Bolder'
                                        size    = 'Small'
                                        color   = $classificationColor
                                        spacing = 'None'
                                    },
                                    [pscustomobject][ordered]@{
                                        type    = 'TextBlock'
                                        text    = $Issue.Id
                                        wrap    = $true
                                        weight  = 'Bolder'
                                        size    = 'Large'
                                        color   = $classificationColor
                                        spacing = 'Small'
                                    },
                                    [pscustomobject][ordered]@{
                                        type    = 'TextBlock'
                                        text    = $statusText
                                        wrap    = $true
                                        weight  = 'Bolder'
                                        size    = 'Small'
                                        color   = if ($Issue.IsResolved) { 'Good' } else { 'Attention' }
                                        spacing = 'Small'
                                    },
                                    [pscustomobject][ordered]@{
                                        type     = 'TextBlock'
                                        text     = $lastUpdatedText
                                        wrap     = $true
                                        size     = 'Small'
                                        isSubtle = $true
                                        spacing  = 'None'
                                    },
                                    [pscustomobject][ordered]@{
                                        type    = 'TextBlock'
                                        text    = $Issue.Title
                                        wrap    = $true
                                        weight  = 'Bolder'
                                        size    = 'Medium'
                                        color   = 'Accent'
                                        spacing = 'Small'
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
                                                $actionsId,
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
                                                $actionsId,
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
                default { return 'Warning' }
            }
        }

        function Get-ServiceHealthHeaderStyle {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                [ValidateNotNull()]
                $Issue
            )

            if ($Issue.IsResolved) {
                return 'good'
            }

            return 'attention'
        }


        function New-ServiceHealthIssueFactSet {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                [ValidateNotNull()]
                $Issue
            )

            $elementId = ConvertTo-AdaptiveCardElementId -Value $Issue.Id
            $detailsId = 'v' + $elementId

            # $latestUpdate = Get-ServiceHealthLatestUpdateText -Issue $Issue
            $latestUpdateInfo = Get-ServiceHealthLatestUpdateObject -Issue $Issue

            # Incrimentally build the facts
            $facts = @(
                [pscustomobject][ordered]@{
                    title = 'User impact'
                    value = $Issue.ImpactDescription
                },
                [pscustomobject][ordered]@{
                    title = 'Start time'
                    value = Format-ServiceHealthCardDate -DateTime $Issue.StartDateTime
                }
            )

            # Add 'End time' only if it exists
            if ($Issue.EndDateTime) {
                $facts += [pscustomobject][ordered]@{
                    title = 'End time'
                    value = Format-ServiceHealthCardDate -DateTime $Issue.EndDateTime
                }
            }

            # Add 'Latest update'
            $facts += [pscustomobject][ordered]@{
                title = 'Latest update'
                value = $latestUpdateInfo.Update
            }

            # Add 'Next update by' is it exists
            if (![System.String]::IsNullOrWhiteSpace($latestUpdateInfo.NextUpdateBy)) {
                $facts += [pscustomobject][ordered]@{
                    title = 'Next update by'
                    value = $latestUpdateInfo.NextUpdateBy
                }
            }

            [pscustomobject][ordered]@{
                type      = 'FactSet'
                id        = $detailsId
                isVisible = $false
                separator = $true
                facts     = $facts
            }
        }

        function New-ServiceHealthIssueActionSet {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                [ValidateNotNull()]
                $Issue
            )

            $elementId = ConvertTo-AdaptiveCardElementId -Value $Issue.Id
            $actionsId = 'a' + $elementId
            $adminCenterUrl = 'https://admin.cloud.microsoft/?#/servicehealth/:/alerts/' + $Issue.Id

            [pscustomobject][ordered]@{
                type      = 'ActionSet'
                id        = $actionsId
                isVisible = $false
                separator = $false
                spacing   = 'Medium'
                actions   = @(
                    [pscustomobject][ordered]@{
                        type  = 'Action.OpenUrl'
                        title = 'Open ' + $Issue.Id + ' in Admin Center'
                        url   = $adminCenterUrl
                        style = 'positive'
                    }
                )
            }
        }

        function New-ServiceHealthServiceHeader {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                [string]$ServiceName,

                [Parameter(Mandatory)]
                [int]$Count
            )

            [pscustomobject][ordered]@{
                type  = 'Container'
                style = 'emphasis'
                bleed = $true
                items = @(
                    [pscustomobject][ordered]@{
                        type   = 'TextBlock'
                        text   = "$ServiceName ($Count)"
                        weight = 'Bolder'
                        size   = 'Medium'
                        wrap   = $true
                    }
                )
            }
        }

        function Get-ServiceHealthLatestUpdateObject {
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
                return [pscustomobject]@{
                    Update       = 'No latest update available.'
                    NextUpdateBy = ''
                }
            }

            $messageText = $latestPost.Description.Content.Trim()

            $nextUpdateMatch = [System.Text.RegularExpressions.Regex]::Match(
                $messageText,
                '(?is)Next update by:\s*(?<NextUpdate>.*?)(?=\r?\n\s*\r?\n|\z)'
            )

            $nextUpdateBy = ''

            if ($nextUpdateMatch.Success) {
                $nextUpdateBy = $nextUpdateMatch.Groups['NextUpdate'].Value.Trim()
            }

            $updateText = Get-ServiceHealthPortalStyleUpdateText -Text $messageText

            if ($nextUpdateBy) {
                $updateText = $updateText -replace '(?is)\r?\n\r?\nNext update by:.*$', ''
                $updateText = $updateText.Trim()
            }

            [pscustomobject]@{
                Update       = $updateText
                NextUpdateBy = $nextUpdateBy
            }
        }
    }

    process {
        $teamsAdaptiveCardPath = Join-Path -Path $moduleInfo.ModuleBase -ChildPath 'source\private\TeamsConsolidated.json'

        $teamsAdaptiveCard = Get-Content -Path $teamsAdaptiveCardPath -Raw | ConvertFrom-Json

        # Ensure the template body starts empty even if the JSON template file is later modified.
        $teamsAdaptiveCard.attachments[0].content.body = @()

        # Icon requires Adaptive Card 1.5.
        $teamsAdaptiveCard.attachments[0].content.version = '1.5'

        if (!$teamsAdaptiveCard.attachments[0].content.msTeams) {
            $teamsAdaptiveCard.attachments[0].content | Add-Member -MemberType NoteProperty -Name 'msTeams' -Value ([pscustomobject]@{
                    width = 'full'
                })
        }

        $teamsAdaptiveCard.attachments[0].content.body += New-ServiceHealthCardHeader `
            -Title $Title `
            -ReportGeneratedDate $InputObject[0].ReportGeneratedDate `
            -Issues $InputObject

        $itemGroups = $InputObject |
        Group-Object Service |
        Sort-Object Count, Name -Descending

        foreach ($group in $itemGroups) {

            $teamsAdaptiveCard.attachments[0].content.body += (
                New-ServiceHealthServiceHeader `
                    -ServiceName $group.Name `
                    -Count $group.Count
            )

            foreach ($item in ($group.Group | Sort-Object LastModifiedDateTime -Descending)) {

                $teamsAdaptiveCard.attachments[0].content.body += (
                    New-ServiceHealthIssueHeader -Issue $item
                )

                $teamsAdaptiveCard.attachments[0].content.body += (
                    New-ServiceHealthIssueFactSet -Issue $item
                )

                $teamsAdaptiveCard.attachments[0].content.body += (
                    New-ServiceHealthIssueActionSet -Issue $item
                )
            }
        }

        return ($teamsAdaptiveCard | ConvertTo-Json -Depth 20)
    }
}