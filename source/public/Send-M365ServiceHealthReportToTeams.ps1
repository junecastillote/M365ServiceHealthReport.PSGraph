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
        [string[]]$TeamsWebhookUrl
    )

    begin {
        $star_divider = ('*' * 70)

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
            $teamsCardPayloads = Get-TeamsCardPayload -ReportObject $report

            if ($teamsCardPayloads.Count -eq 0) {
                SayError "The input object does not contain Teams card content. Make sure the report was generated with -Format TeamsCard or without specifying -Format."
                continue
            }

            $payloadIndex = 0
            $payloadCount = $teamsCardPayloads.Count

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

                foreach ($url in $TeamsWebhookUrl) {
                    SayInfo "Posting Teams alert card [$payloadIndex/$payloadCount] to Teams webhook. Payload size: $payloadSizeBytes bytes."

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
                }
            }
        }
    }
}