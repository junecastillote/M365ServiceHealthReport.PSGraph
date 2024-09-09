# Send-M365ServiceHealthReportToEmail

## Syntax

```PowerShell
Send-M365ServiceHealthReportToEmail [-InputObject] <M365ServiceHealthReport> [-MailFrom] <string> [[-MailTo] <string[]>] [[-MailCc] <string[]>] [[-MailBcc] <string[]>] [<CommonParameters>]
```

## Parameters

### -InputObject

This parameter accepts the `M365ServiceHealthReport` object output of the `ConvertTo-M365ServiceHealthReportObject` command. It can also accept the input through the pipeline.

|                        |                         |
| ---------------------- | ----------------------- |
| Type:                  | M365ServiceHealthReport |
| Position:              | Named                   |
| Default value:         | None                    |
| Required:              | True                    |
| Accept pipeline input: | True                    |

### -MailFrom

Specify the report's sender email address. It has to be a valid Exchange Online mailbox.

|                        |        |
| ---------------------- | ------ |
| Type:                  | String |
| Position:              | Named  |
| Default value:         | None   |
| Required:              | True   |
| Accept pipeline input: | False  |

### -MailTo

The collection of email addresses for the TO recipients.

|                        |          |
| ---------------------- | -------- |
| Type:                  | String[] |
| Position:              | Named    |
| Default value:         | None     |
| Required:              | False    |
| Accept pipeline input: | False    |

### -MailCc

The collection of email addresses for the CC recipients.

|                        |          |
| ---------------------- | -------- |
| Type:                  | String[] |
| Position:              | Named    |
| Default value:         | None     |
| Required:              | False    |
| Accept pipeline input: | False    |

### -MailBcc

The collection of email addresses for the BCC recipients.

|                        |          |
| ---------------------- | -------- |
| Type:                  | String[] |
| Position:              | Named    |
| Default value:         | None     |
| Required:              | False    |
| Accept pipeline input: | False    |

## Examples

### Example 1 - Send the report

The following example retrieves the service health events for the past 10 days affecting 'Exchange Online', and sends the converted report.

```PowerShell
$report = Get-M365ServiceHealthEvent -PastDays 10 -Service 'Exchange Online' | ConvertTo-M365ServiceHealthReportObject -OrganizationName PoshLab
$report | Send-M365ServiceHealthReportToEmail -MailFrom "sender@domain.com" -MailTo 'recipient1@domain.com','recipient2@domain.com'
```

## Output Type

None
