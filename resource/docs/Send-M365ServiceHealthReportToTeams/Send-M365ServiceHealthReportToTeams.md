# Send-M365ServiceHealthReportToTeams

## Syntax

```PowerShell
Send-M365ServiceHealthReportToTeams [-InputObject] <M365ServiceHealthReport> [-TeamsWebhookUrl] <string[]> [<CommonParameters>]
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

### -TeamsWebhookUrl

This parameter accepts one or more Teams WebHook URL to where the report will be sent. You can use the URL from a Power Automate flow that uses the "**When a Teams webhook request is received**" trigger.

|                        |          |
| ---------------------- | -------- |
| Type:                  | String[] |
| Position:              | Named    |
| Default value:         | None     |
| Required:              | True     |
| Accept pipeline input: | False    |
