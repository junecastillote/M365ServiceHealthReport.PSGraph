Function Send-M365ServiceHealthReportToEmail {
    [CmdletBinding()]
    [Alias('Send-M365ServiceHealthReportEmail')]
    param (
        [Parameter(Mandatory,
            ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [PSTypeNameAttribute('M365ServiceHealthReport')]
        $InputObject,

        [Parameter(Mandatory)]
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
        Function ToEmailAddressHashTable {
            param(
                [parameter()]
                [string[]]
                $Address
            )

            $Address | ForEach-Object {
                @{
                    EmailAddress = @{
                        Address = $_
                    }
                }
            }
        }

        $moduleInfo = Get-Module $($MyInvocation.MyCommand.ModuleName)

        $mail_prop = @{}
        $mail_prop.Add('MailFrom', $MailFrom)
        if ($MailTo) {
            $mail_prop.Add('MailTo', $MailTo)
        }
        if ($MailCc) {
            $mail_prop.Add('MailCc', $MailCc)
        }
        if ($MailBcc) {
            $mail_prop.Add('MailBcc', $MailBcc)
        }

        if (!$MailTo -and !$MailCc -and !$MailBcc) {
            SayError "At least one recipients [MailTo, MailCc, MailBcc] is required. Script stopped."
            Continue
        }
    }
    process {

        $mail_params = @{
            Message = @{
                Subject                = $InputObject.Title
                Body                   = @{
                    ContentType = "HTML"
                    Content     = $InputObject.HtmlContent
                }
                InternetMessageHeaders = @(
                    @{
                        Name  = "X-Mailer"
                        Value = $moduleInfo.Name
                    }
                )
            }
        }

        if ($MailTo) {
            $mail_params.Message.Add('toRecipients', @(ToEmailAddressHashTable $MailTo))
        }

        if ($MailCc) {
            $mail_params.Message.Add('ccRecipients', @(ToEmailAddressHashTable $MailCc))
        }

        if ($MailBcc) {
            $mail_params.Message.Add('bccRecipients', @(ToEmailAddressHashTable $MailBcc))
        }

        try {
            Send-MgUserMail @mail_params -UserId $MailFrom -ErrorAction Stop
        }
        catch {
            SayError "Failed to send email report. `n$star_divider`n$_$star_divider"
        }
    }
    end {

    }
}