function Send-M365ServiceHealthReportToEmail {
    [CmdletBinding()]
    [Alias('Send-M365ServiceHealthReportEmail')]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [PSTypeNameAttribute('M365ServiceHealthReport')]
        $InputObject,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$MailFrom,

        [Parameter()]
        [string[]]$MailTo,

        [Parameter()]
        [string[]]$MailCc,

        [Parameter()]
        [string[]]$MailBcc
    )

    begin {
        $star_divider = ('*' * 70)

        function ConvertTo-EmailAddressHashTable {
            [CmdletBinding()]
            param(
                [Parameter()]
                [string[]]$Address
            )

            foreach ($item in $Address) {
                @{
                    EmailAddress = @{
                        Address = $item
                    }
                }
            }
        }

        function Get-FileBase64String {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                [ValidateNotNullOrEmpty()]
                [string]$Path
            )

            if (!(Test-Path -Path $Path -PathType Leaf)) {
                throw "File not found: $Path"
            }

            $bytes = [System.IO.File]::ReadAllBytes($Path)
            return [System.Convert]::ToBase64String($bytes)
        }

        function New-GraphInlineFileAttachment {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                [ValidateNotNullOrEmpty()]
                [string]$Path,

                [Parameter(Mandatory)]
                [ValidateNotNullOrEmpty()]
                [string]$Name,

                [Parameter(Mandatory)]
                [ValidateNotNullOrEmpty()]
                [string]$ContentId,

                [Parameter()]
                [ValidateNotNullOrEmpty()]
                [string]$ContentType = 'image/png'
            )

            @{
                '@odata.type' = '#microsoft.graph.fileAttachment'
                Name          = $Name
                ContentType   = $ContentType
                ContentId     = $ContentId
                IsInline      = $true
                ContentBytes  = Get-FileBase64String -Path $Path
            }
        }

        function Convert-ServiceHealthReportImagesToCid {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                [AllowEmptyString()]
                [string]$HtmlContent,

                [Parameter(Mandatory)]
                [ValidateNotNullOrEmpty()]
                [string]$YellowDotDataUri,

                [Parameter(Mandatory)]
                [ValidateNotNullOrEmpty()]
                [string]$RedDotDataUri
            )

            $htmlBody = $HtmlContent

            $htmlBody = $htmlBody.Replace($YellowDotDataUri, 'cid:yellow-dot.png')
            $htmlBody = $htmlBody.Replace($RedDotDataUri, 'cid:red-dot.png')

            return $htmlBody
        }

        if (!$MailTo -and !$MailCc -and !$MailBcc) {
            throw "At least one recipient parameter is required: MailTo, MailCc, or MailBcc."
        }

        $moduleInfo = Get-Module $($MyInvocation.MyCommand.ModuleName)

        if (!$moduleInfo) {
            throw "Unable to determine module information for $($MyInvocation.MyCommand.ModuleName)."
        }

        $privatePath = Join-Path -Path $moduleInfo.ModuleBase -ChildPath 'source\private'

        $yellowDotPath = Join-Path -Path $privatePath -ChildPath 'yellow-dot.png'
        $redDotPath = Join-Path -Path $privatePath -ChildPath 'red-dot.png'

        $yellowDotBase64 = Get-FileBase64String -Path $yellowDotPath
        $redDotBase64 = Get-FileBase64String -Path $redDotPath

        $yellowDotDataUri = 'data:image/png;base64,' + $yellowDotBase64
        $redDotDataUri = 'data:image/png;base64,' + $redDotBase64

        $inlineImageAttachments = @(
            New-GraphInlineFileAttachment `
                -Path $yellowDotPath `
                -Name 'yellow-dot.png' `
                -ContentId 'yellow-dot.png'

            New-GraphInlineFileAttachment `
                -Path $redDotPath `
                -Name 'red-dot.png' `
                -ContentId 'red-dot.png'
        )
    }

    process {
        if (
            [System.String]::IsNullOrWhiteSpace($InputObject.HtmlContent) -or
            $InputObject.HtmlContent -eq 'None'
        ) {
            SayError "The input object does not contain HTML content. Make sure the report was generated with -Format Html or without specifying -Format."
            return
        }

        $htmlBody = Convert-ServiceHealthReportImagesToCid `
            -HtmlContent $InputObject.HtmlContent `
            -YellowDotDataUri $yellowDotDataUri `
            -RedDotDataUri $redDotDataUri

        $mailMessage = @{
            Subject                = $InputObject.Title
            Body                   = @{
                ContentType = 'HTML'
                Content     = $htmlBody
            }
            InternetMessageHeaders = @(
                @{
                    Name  = 'X-Mailer'
                    Value = $moduleInfo.Name
                }
            )
            Attachments            = $inlineImageAttachments
        }

        if ($MailTo) {
            $mailMessage.Add('toRecipients', @(ConvertTo-EmailAddressHashTable -Address $MailTo))
        }

        if ($MailCc) {
            $mailMessage.Add('ccRecipients', @(ConvertTo-EmailAddressHashTable -Address $MailCc))
        }

        if ($MailBcc) {
            $mailMessage.Add('bccRecipients', @(ConvertTo-EmailAddressHashTable -Address $MailBcc))
        }

        $mailParams = @{
            Message = $mailMessage
            UserId  = $MailFrom
        }

        try {
            Send-MgUserMail @mailParams -ErrorAction Stop
        }
        catch {
            SayError "Failed to send email report.`n$star_divider`n$_`n$star_divider"
        }
    }
}