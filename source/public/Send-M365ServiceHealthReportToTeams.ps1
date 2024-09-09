Function Send-M365ServiceHealthReportToTeams {
    [CmdletBinding()]
    [Alias('Send-M365ServiceHealthReportTeams')]
    param (
        [Parameter(Mandatory,
            ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [PSTypeNameAttribute('M365ServiceHealthReport')]
        $InputObject,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $ChannelId,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $ChatId
    )
    begin {
        $star_divider = ('*' * 70)
    }
    process {

        foreach ($item in $InputObject) {

            if ($ChannelId) {
                foreach ($id in $ChannelId) {
                    $team_id = $id.Split('/')[0]
                    $channel_id = $id.Split('/')[1]

                    try {
                        New-MgTeamChannelMessage -TeamId $team_id -ChannelId $channel_id -BodyParameter $item.TeamsCardContent -ErrorAction Stop | Out-Null
                    }
                    catch {
                        SayError "Failed channel message to [$($id)]. `n$star_divider`n$_$star_divider"
                    }
                }
            }

            if ($ChatId) {
                foreach ($id in $ChatId) {
                    try {
                        New-MgChatMessage -ChatId $id -BodyParameter $item.TeamsCardContent -ErrorAction Stop | Out-Null
                    }
                    catch {
                        SayError "Failed chat message to [$($id)]. `n$star_divider`n$_$star_divider"
                    }
                }
            }
        }
    }
    end {

    }
}