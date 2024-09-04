Function Send-M365ServiceHealthReportEmail {
    [CmdletBinding()]
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
                    Content     = $InputObject.Content
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
            SayInfo "Sending email report..."
            Send-MgUserMail @mail_params -UserId $MailFrom -ErrorAction Stop
            SayInfo "Sent!"
        }
        catch {
            SayError $_
        }
    }
    end {

    }
}