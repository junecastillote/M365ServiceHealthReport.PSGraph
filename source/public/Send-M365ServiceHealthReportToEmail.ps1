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
        $starDivider = '*' * 70

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
                [object[]]$ImageMap
            )

            $htmlBody = $HtmlContent

            foreach ($image in $ImageMap) {
                $htmlBody = $htmlBody.Replace(
                    $image.DataUri,
                    'cid:' + $image.ContentId
                )
            }

            return $htmlBody
        }

        if (!$MailTo -and !$MailCc -and !$MailBcc) {
            throw 'At least one recipient parameter is required: MailTo, MailCc, or MailBcc.'
        }

        $moduleInfo = Get-Module $MyInvocation.MyCommand.ModuleName

        if (!$moduleInfo) {
            throw "Unable to determine module information for [$($MyInvocation.MyCommand.ModuleName)]."
        }

        $privatePath = Join-Path `
            -Path $moduleInfo.ModuleBase `
            -ChildPath 'source\private'

        $imageDefinitions = @(
            [PSCustomObject][ordered]@{
                Name = 'active.png'
                Path = Join-Path -Path $privatePath -ChildPath 'active.png'
            },
            [PSCustomObject][ordered]@{
                Name = 'resolved.png'
                Path = Join-Path -Path $privatePath -ChildPath 'resolved.png'
            },
            [PSCustomObject][ordered]@{
                Name = 'advisory.png'
                Path = Join-Path -Path $privatePath -ChildPath 'advisory.png'
            },
            [PSCustomObject][ordered]@{
                Name = 'incident.png'
                Path = Join-Path -Path $privatePath -ChildPath 'incident.png'
            }
        )

        $imageMap = foreach ($image in $imageDefinitions) {
            $base64Content = Get-FileBase64String -Path $image.Path

            [PSCustomObject][ordered]@{
                Name      = $image.Name
                Path      = $image.Path
                ContentId = $image.Name
                DataUri   = 'data:image/png;base64,' + $base64Content
            }
        }

        $inlineImageAttachments = @(
            foreach ($image in $imageMap) {
                New-GraphInlineFileAttachment `
                    -Path $image.Path `
                    -Name $image.Name `
                    -ContentId $image.ContentId
            }
        )
    }

    process {
        foreach ($report in $InputObject) {
            if (
                [System.String]::IsNullOrWhiteSpace($report.HtmlContent) -or
                $report.HtmlContent -eq 'None'
            ) {
                SayError 'The input object does not contain HTML content. Make sure the report was generated with -Format Html or without specifying -Format.'
                continue
            }

            $htmlBody = Convert-ServiceHealthReportImagesToCid `
                -HtmlContent $report.HtmlContent `
                -ImageMap $imageMap

            $mailMessage = @{
                Subject                 = $report.Title
                Body                    = @{
                    ContentType = 'HTML'
                    Content     = $htmlBody
                }
                InternetMessageHeaders  = @(
                    @{
                        Name  = 'X-Mailer'
                        Value = $moduleInfo.Name
                    },
                    @{
                        Name  = 'X-M365-Service-Health-RunId'
                        Value = $report.RunId.ToString()
                    }
                )
                Attachments             = $inlineImageAttachments
            }

            if ($MailTo) {
                $mailMessage.Add(
                    'toRecipients',
                    @(ConvertTo-EmailAddressHashTable -Address $MailTo)
                )
            }

            if ($MailCc) {
                $mailMessage.Add(
                    'ccRecipients',
                    @(ConvertTo-EmailAddressHashTable -Address $MailCc)
                )
            }

            if ($MailBcc) {
                $mailMessage.Add(
                    'bccRecipients',
                    @(ConvertTo-EmailAddressHashTable -Address $MailBcc)
                )
            }

            $mailParams = @{
                Message = $mailMessage
                UserId  = $MailFrom
            }

            try {
                Send-MgUserMail @mailParams -ErrorAction Stop

                SayInfo "Email report sent successfully. Run ID: $($report.RunId)"
            }
            catch {
                SayError "Failed to send email report for Run ID [$($report.RunId)].`n$starDivider`n$($_.Exception.Message)`n$starDivider"
            }
        }
    }
}
