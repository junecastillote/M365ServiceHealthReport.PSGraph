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

    $finalStatusPattern = '(?is)Final status:\s*(?<FinalStatus>.*?)(?=\r?\n\s*\r?\n(?:Scope of impact:|Start*time:|End time:|Root cause:|Next s*eps:|Next update by:|Title:|User impact:|More info:|Current status:)|\z)'
    $finalUpdatePattern = '(?im)^\s*This is the final update for the event\.\s*$'

    $currentStatusMatch = [System.Text.RegularExpressions.Regex]::Match($normalizedText, $currentStatusPattern)

    if ($currentStatusMatch.Success) {
        $currentStatus = $currentStatusMatch.Groups['CurrentStatus'].Value.Trim()
        $nextUpdateMatch = [System.Text.RegularExpressions.Regex]::Match($normalizedText, $nextUpdatePattern)

        if ($nextUpdateMatch.Success) {
            $nextUpdate = $nextUpdateMatch.Groups['NextUpdate'].Value.Trim()

            if (![System.String]::IsNullOrWhiteSpace($nextUpdate)) {
                return $currentStatus + "`r`n`r`nNext update by: " + $nextUpdate
            }
        }

        return $currentStatus
    }

    $finalStatusMatch = [System.Text.RegularExpressions.Regex]::Match($normalizedText, $finalStatusPattern)

    if ($finalStatusMatch.Success) {
        $finalStatus = $finalStatusMatch.Groups['FinalStatus'].Value.Trim()
        $finalUpdateMatch = [System.Text.RegularExpressions.Regex]::Match($normalizedText, $finalUpdatePattern)

        if ($finalUpdateMatch.Success) {
            return $finalStatus + "`r`n`r`n" + $finalUpdateMatch.Value.Trim()
        }

        return $finalStatus
    }

    return $normalizedText
}

function Get-ServiceHealthLatestMessageHtml {
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
        return 'No latest message available.'
    }

    $messageText = $latestPost.Description.Content.Trim()

    $parsedText = Get-ServiceHealthPortalStyleUpdateText -Text $messageText

    $encodedMessage = ConvertTo-HtmlEncodedText -Text $parsedText
    # $encodedMessage = $encodedMessage -replace "(\r\n|\n|\r)", '<br />'

    return $encodedMessage
}

function Format-ServiceHealthDate {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $DateTime
    )

    if (!$DateTime) {
        return ''
    }

    # return '{0:MMM dd, yyyy, hh:mm tt} UTC' -f [datetime]$DateTime
    return '{0:MMM dd, yyyy, hh:mm tt (UTCzzzz)}' -f ([datetime]$DateTime).ToLocalTime()
}

function ConvertTo-HtmlEncodedText {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [string]$Text
    )

    if ([system.string]::IsNullOrEmpty($Text)) {
        return ''
    }

    return [System.Net.WebUtility]::HtmlEncode($Text)
}

function ConvertTo-HtmlAnchorId {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Value
    )

    return ($Value -replace '[^a-zA-Z0-9_-]', '-')
}

# function Get-ServiceHealthLatestMessageHtml {
#     [CmdletBinding()]
#     param(
#         [Parameter(Mandatory)]
#         [ValidateNotNull()]
#         $Issue
#     )

#     $latestPost = if ($Issue.Posts -and $Issue.Posts.Count -gt 0) {
#         $Issue.Posts[-1]
#     }

#     if (
#         !$latestPost -or
#         !$latestPost.Description -or
#         [System.String]::IsNullOrWhiteSpace($latestPost.Description.Content)
#     ) {
#         return 'No latest message available.'
#     }

#     $message = ConvertTo-HtmlEncodedText -Text $latestPost.Description.Content
#     $message = $message -replace "(\r\n|\n|\r)", '<br />'

#     return $message
# }

function Get-ServiceHealthLatestMessage {
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
        return 'No latest message available.'
    }

    $message = ConvertTo-HtmlEncodedText -Text $latestPost.Description.Content
    # $message = $message -replace "(\r\n|\n|\r)", '<br />'

    return $message
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

function Get-ImageBase64String {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    if (!(Test-Path -Path $Path -PathType Leaf)) {
        throw "Image file not found: $Path"
    }

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    return [System.Convert]::ToBase64String($bytes)
}

function Get-ServiceHealthClassificationIconSource {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [string]$Classification,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$YellowDotSource,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$RedDotSource
    )

    if ([System.String]::IsNullOrWhiteSpace($Classification)) {
        return $YellowDotSource
    }

    switch ($Classification.Trim().ToLowerInvariant()) {
        'incident' { return $RedDotSource }
        default { return $YellowDotSource }
    }
}

function Get-ServiceHealthClassificationHtml {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [string]$Classification,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$YellowDotSource,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$RedDotSource
    )

    $classificationText = Format-ServiceHealthText -Text $Classification
    $encodedClassification = ConvertTo-HtmlEncodedText -Text $classificationText

    $iconSource = Get-ServiceHealthClassificationIconSource `
        -Classification $classificationText `
        -YellowDotSource $YellowDotSource `
        -RedDotSource $RedDotSource

    $altText = if ($classificationText) {
        ConvertTo-HtmlEncodedText -Text $classificationText
    }
    else {
        'Classification'
    }

    return '<img src="' + $iconSource + '" ' + 'alt="' + $altText + '"' + ' width="12" height="12">&nbsp;' + $encodedClassification
}