# This function creates a single Microsoft 365 Service Health alert card
# using Adaptive Cards.
function New-ServiceHealthAlertCardJson {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        $Issue,

        [Parameter(Mandatory)]
        [string]$OrganizationName,

        [Parameter(Mandatory)]
        [string]$RunId
    )

    begin {
        $moduleInfo = Get-Module $($MyInvocation.MyCommand.ModuleName)

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
            $finalStatusPattern = '(?is)Final status:\s*(?<FinalStatus>.*?)(?=\r?\n\s*\r?\n(?:Scope of impact:|Start time:|End time:|Root cause:|Next steps:|Next update by:|Title:|User impact:|More info:|Current status:)|\z)'
            $finalUpdatePattern = '(?im)^\s*This is the final update for the event\.\s*$'

            $currentStatusMatch = [System.Text.RegularExpressions.Regex]::Match(
                $normalizedText,
                $currentStatusPattern
            )

            if ($currentStatusMatch.Success) {
                return $currentStatusMatch.Groups['CurrentStatus'].Value.Trim()
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
                return [pscustomobject][ordered]@{
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

            [pscustomobject][ordered]@{
                Update       = $updateText
                NextUpdateBy = $nextUpdateBy
            }
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

        function Get-ServiceHealthContainerStyle {
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

        function New-ServiceHealthAlertHeader {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                [ValidateNotNull()]
                $Issue,

                [Parameter(Mandatory)]
                [string]$OrganizationName
            )

            $elementId = ConvertTo-AdaptiveCardElementId -Value $Issue.Id
            $detailsId = 'v' + $elementId
            $actionsId = 'a' + $elementId
            $toggleDownId = 'toggle_' + $elementId + 'Down'
            $toggleUpId = 'toggle_' + $elementId + 'Up'

            $classificationText = Format-ServiceHealthText -Text $Issue.Classification
            $classificationColor = Get-ServiceHealthClassificationColor -Classification $Issue.Classification
            $containerStyle = Get-ServiceHealthContainerStyle -Issue $Issue

            $resolutionState = if ($Issue.IsResolved) {
                'Resolved'
            }
            else {
                'Active'
            }

            $statusText = $resolutionState + ' | ' + $Issue.Status
            $lastUpdatedText = 'Updated: ' + (Format-ServiceHealthCardDate -DateTime $Issue.LastModifiedDateTime)

            [pscustomobject][ordered]@{
                type       = 'Container'
                style      = $containerStyle
                bleed      = $true
                showBorder = $true
                # selectAction = [pscustomobject][ordered]@{
                #     type           = 'Action.ToggleVisibility'
                #     targetElements = @(
                #         $detailsId,
                #         $actionsId,
                #         $toggleDownId,
                #         $toggleUpId
                #     )
                # }
                items      = @(
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
                                        type     = 'TextBlock'
                                        text     = $OrganizationName + ' | ' + $Issue.Service
                                        wrap     = $true
                                        size     = 'Small'
                                        weight   = 'Bolder'
                                        isSubtle = $true
                                        spacing  = 'None'
                                    },
                                    [pscustomobject][ordered]@{
                                        type     = 'TextBlock'
                                        text     = 'Run ID: ' + $RunId.Substring(0, 8).ToUpperInvariant()
                                        wrap     = $true
                                        size     = 'Small'
                                        isSubtle = $true
                                        spacing  = 'None'
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
                                        # fallback     = [pscustomobject][ordered]@{
                                        #     type = 'TextBlock'
                                        #     text = '▼'
                                        #     wrap = $true
                                        # }
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
                                        # fallback     = [pscustomobject][ordered]@{
                                        #     type      = 'TextBlock'
                                        #     text      = '▲'
                                        #     wrap      = $true
                                        #     isVisible = $false
                                        # }
                                    }
                                )
                            }
                        )
                    }
                )
            }
        }

        function New-ServiceHealthAlertFactSet {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                [ValidateNotNull()]
                $Issue
            )

            $latestUpdateInfo = Get-ServiceHealthLatestUpdateObject -Issue $Issue
            $elementId = ConvertTo-AdaptiveCardElementId -Value $Issue.Id
            $detailsId = 'v' + $elementId

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

            # Calculate issue age
            if ($Issue.IsResolved -and $Issue.EndDateTime) {

                $duration = New-TimeSpan `
                    -Start $Issue.StartDateTime `
                    -End $Issue.EndDateTime

                $facts += [pscustomobject][ordered]@{
                    title = 'Duration'
                    value = (Format-ServiceHealthDuration -TimeSpan $duration)
                }
            }
            else {

                $age = New-TimeSpan `
                    -Start $Issue.StartDateTime `
                    -End (Get-Date).ToUniversalTime()

                $facts += [pscustomobject][ordered]@{
                    title = 'Age'
                    value = (Format-ServiceHealthDuration -TimeSpan $age)
                }
            }

            if ($Issue.EndDateTime) {
                $facts += [pscustomobject][ordered]@{
                    title = 'End time'
                    value = Format-ServiceHealthCardDate -DateTime $Issue.EndDateTime
                }
            }

            $facts += [pscustomobject][ordered]@{
                title = 'Latest update'
                value = $latestUpdateInfo.Update
            }

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

        function New-ServiceHealthAlertActionSet {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                [ValidateNotNull()]
                $Issue
            )

            $adminCenterUrl = 'https://admin.cloud.microsoft/?#/servicehealth/:/alerts/' + $Issue.Id
            $elementId = ConvertTo-AdaptiveCardElementId -Value $Issue.Id
            $actionsId = 'a' + $elementId

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

        function ConvertTo-AdaptiveCardElementId {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                [ValidateNotNullOrEmpty()]
                [string]$Value
            )

            return ($Value -replace '[^a-zA-Z0-9_-]', '-')
        }
    }

    process {
        $teamsAdaptiveCardPath = Join-Path -Path $moduleInfo.ModuleBase -ChildPath 'source\private\TeamsCardSchema.json'

        $teamsAdaptiveCard = Get-Content -Path $teamsAdaptiveCardPath -Raw | ConvertFrom-Json

        # Ensure the template body starts empty even if the JSON template file is later modified.
        $teamsAdaptiveCard.attachments[0].content.body = @()

        $teamsAdaptiveCard.attachments[0].content.version = '1.5'

        if (!$teamsAdaptiveCard.attachments[0].content.msTeams) {
            $teamsAdaptiveCard.attachments[0].content | Add-Member -MemberType NoteProperty -Name 'msTeams' -Value ([pscustomobject]@{
                    width = 'full'
                })
        }

        $teamsAdaptiveCard.attachments[0].content.body += New-ServiceHealthAlertHeader -Issue $Issue -OrganizationName $OrganizationName
        $teamsAdaptiveCard.attachments[0].content.body += New-ServiceHealthAlertFactSet -Issue $Issue
        $teamsAdaptiveCard.attachments[0].content.body += New-ServiceHealthAlertActionSet -Issue $Issue

        return ($teamsAdaptiveCard | ConvertTo-Json -Depth 20)
    }
}