Function Send-M365ServiceHealthReportToTeams {
    [CmdletBinding()]
    [Alias('Send-M365ServiceHealthReportTeams')]
    param (
        [Parameter(Mandatory,
            ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [PSTypeNameAttribute('M365ServiceHealthReport')]
        $InputObject,

        [Parameter(Mandatory)]
        [string[]]
        $TeamsWebhookUrl
    )
    begin {
        $star_divider = ('*' * 70)
    }
    process {

        foreach ($item in $InputObject) {

            # $item.TeamsCardContent

            foreach ($url in $TeamsWebhookUrl) {
                SayInfo "Posting alert to Teams with URL [$($url)]"
                $Params = @{
                    "URI"         = $url
                    "Method"      = 'POST'
                    "Body"        = $item.TeamsCardContent
                    "ContentType" = 'application/json'
                }
                try {
                    Invoke-RestMethod @Params -ErrorAction Stop
                }
                catch {
                    SayError "Failed to post to channel. `n$star_divider`n$_$star_divider"
                }
            }
        }
    }
    end {

    }
}