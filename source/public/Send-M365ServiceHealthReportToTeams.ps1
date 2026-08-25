function Send-M365ServiceHealthReportToTeams {
    [CmdletBinding()]
    [Alias('Send-M365ServiceHealthReportTeams')]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [PSTypeNameAttribute('M365ServiceHealthReport')]
        $InputObject,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string[]]$TeamsWebhookUrl,

        [Parameter()]
        [switch]
        $PostBatchHeader
    )

    begin {
        $star_divider = ('*' * 70)

        function New-TeamsBatchHeaderCard {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                [ValidateNotNull()]
                $Report
            )

            # $shortRunId = $Report.RunId.Substring(0, 8)
            $shortRunId = $Report.RunId.Substring(0, 8).ToUpperInvariant()


            $summary = $Report.GetSummary()

            $incidentCount = ($Report.Issues | Where-Object {
                    $_.Classification -eq 'Incident'
                }).Count

            $advisoryCount = ($Report.Issues | Where-Object {
                    $_.Classification -eq 'Advisory'
                }).Count

            $payload = @{
                type        = 'message'
                attachments = @(
                    @{
                        contentType = 'application/vnd.microsoft.card.adaptive'
                        contentUrl  = $null
                        content     = @{
                            '$schema' = 'https://adaptivecards.io/schemas/adaptive-card.json'
                            type      = 'AdaptiveCard'
                            version   = '1.5'
                            msTeams   = @{
                                width = 'full'
                            }
                            body      = @(
                                @{
                                    type  = 'Container'
                                    style = 'emphasis'
                                    bleed = $true
                                    items = @(
                                        @{
                                            type                = 'TextBlock'
                                            text                = 'Microsoft 365 Service Health Alert Summary'
                                            size                = 'Large'
                                            weight              = 'Bolder'
                                            wrap                = $true
                                            horizontalAlignment = 'Center'
                                        },
                                        @{
                                            type                = 'TextBlock'
                                            text                = $Report.OrganizationName
                                            wrap                = $true
                                            horizontalAlignment = 'Center'
                                            isSubtle            = $true
                                        },
                                        @{
                                            type                = 'TextBlock'
                                            text                = ('Generated: {0:F}' -f $Report.ReportGeneratedDate.ToLocalTime())
                                            wrap                = $true
                                            horizontalAlignment = 'Center'
                                            isSubtle            = $true
                                        },
                                        @{
                                            type                = 'TextBlock'
                                            text                = "Run ID: $shortRunId"
                                            wrap                = $true
                                            horizontalAlignment = 'Center'
                                            isSubtle            = $true
                                            spacing             = 'None'
                                        }
                                    )
                                },
                                @{
                                    type  = 'Container'
                                    items = @(
                                        @{
                                            type                = 'TextBlock'
                                            text                = "Total: $($summary.Count) | Active: $($summary.Unresolved) | Resolved: $($summary.Resolved)"
                                            wrap                = $true
                                            horizontalAlignment = 'Center'
                                            weight              = 'Bolder'
                                        },
                                        @{
                                            type                = 'TextBlock'
                                            text                = "Incidents: $incidentCount | Advisories: $advisoryCount"
                                            wrap                = $true
                                            horizontalAlignment = 'Center'
                                            isSubtle            = $true
                                            spacing             = 'None'
                                        }
                                    )
                                }
                            )
                        }
                    }
                )
            }

            return ($payload | ConvertTo-Json -Depth 20)
        }

        function Get-TeamsCardPayload {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                [ValidateNotNull()]
                $ReportObject
            )

            if (
                !$ReportObject.TeamsCardContent -or
                $ReportObject.TeamsCardContent -eq 'None'
            ) {
                return @()
            }

            if ($ReportObject.TeamsCardContent -is [string]) {
                return @($ReportObject.TeamsCardContent)
            }

            return @($ReportObject.TeamsCardContent)
        }

        function Test-JsonPayload {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                [ValidateNotNullOrEmpty()]
                [string]$Json
            )

            try {
                $null = $Json | ConvertFrom-Json -ErrorAction Stop
                return $true
            }
            catch {
                return $false
            }
        }

        function Get-Utf8ByteCount {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                [AllowEmptyString()]
                [string]$Text
            )

            return [System.Text.Encoding]::UTF8.GetByteCount($Text)
        }
    }

    process {
        foreach ($report in $InputObject) {
            $batchHeaderCard = New-TeamsBatchHeaderCard -Report $report
            $teamsCardPayloads = Get-TeamsCardPayload -ReportObject $report

            if ($teamsCardPayloads.Count -eq 0) {
                SayError "The input object does not contain Teams card content. Make sure the report was generated with -Format TeamsCard or without specifying -Format."
                continue
            }

            $payloadIndex = 0
            $payloadCount = $teamsCardPayloads.Count

            foreach ($url in $TeamsWebhookUrl) {

                #
                # Post batch header
                #
                if ($PostBatchHeader) {

                    SayInfo "RunId [$($report.RunId)] - Posting Teams alert batch header."

                    try {
                        Invoke-RestMethod `
                            -Uri $url `
                            -Method POST `
                            -Body $batchHeaderCard `
                            -ContentType 'application/json' `
                            -ErrorAction Stop
                    }
                    catch {
                        SayError "RunId [$($report.RunId)] - Failed to post Teams alert batch header.`n$star_divider`n$_`n$star_divider"
                    }

                    Start-Sleep -Seconds 2
                }

                $payloadIndex = 0

                foreach ($payload in $teamsCardPayloads) {

                    $payloadIndex++

                    if ([System.String]::IsNullOrWhiteSpace($payload)) {
                        SayError "Skipping empty Teams card payload [$payloadIndex/$payloadCount]."
                        continue
                    }

                    if (!(Test-JsonPayload -Json $payload)) {
                        SayError "Skipping invalid Teams card JSON payload [$payloadIndex/$payloadCount]."
                        continue
                    }

                    $payloadSizeBytes = Get-Utf8ByteCount -Text $payload

                    SayInfo "RunId [$($report.RunId)] - Posting Teams alert card [$payloadIndex/$payloadCount]. Payload size: $payloadSizeBytes bytes."

                    $params = @{
                        Uri         = $url
                        Method      = 'POST'
                        Body        = $payload
                        ContentType = 'application/json'
                    }

                    try {
                        Invoke-RestMethod @params -ErrorAction Stop
                    }
                    catch {
                        SayError "Failed to post Teams alert card [$payloadIndex/$payloadCount].`n$star_divider`n$_`n$star_divider"
                    }

                    Start-Sleep -Milliseconds 500
                }
            }
        }
    }
}