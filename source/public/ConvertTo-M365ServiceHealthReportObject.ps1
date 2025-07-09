Function ConvertTo-M365ServiceHealthReportObject {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [PSTypeNameAttribute('Microsoft.Graph.PowerShell.Models.MicrosoftGraphServiceHealthIssue')]
        $InputObject,

        [Parameter(Mandatory)]
        [string]$OrganizationName,

        [Parameter()]
        [string]$Title,

        [Parameter()]
        [ValidateSet('Html', 'TeamsCard')]
        [string]$Format,

        [Parameter()]
        [string]
        $HtmlReportFileName,

        [Parameter()]
        [string]
        $TeamsCardFileName
    )

    begin {

        $moduleInfo = Get-Module $($MyInvocation.MyCommand.ModuleName)

        if ($HtmlReportFileName -and !(Test-Path -Path $HtmlReportFileName)) {
            try {
                $null = New-Item -ItemType File -Path $HtmlReportFileName -Force -ErrorAction Stop
            }
            catch {
                SayError $_
                Continue
            }
        }

        if ($TeamsCardFileName -and !(Test-Path -Path $TeamsCardFileName)) {
            try {
                $null = New-Item -ItemType File -Path $TeamsCardFileName -Force -ErrorAction Stop
            }
            catch {
                SayError $_
                Continue
            }
        }

        ## Get the CSS style
        $css_string = Get-Content (($moduleInfo.ModuleBase.ToString()) + '\source\private\style.css') -Raw

        $issue_collection = [System.Collections.Generic.list[System.Object]]@()

        if (!$Title) {
            $report_title = "[$($OrganizationName)] Microsoft 365 Service Health Report"
        }
        else {
            $report_title = $Title
        }
    }
    process {
        foreach ($item in ($InputObject)) {
            if ($item.psobject.TypeNames[0] -like "*Microsoft.Graph.PowerShell.Models.MicrosoftGraphServiceHealthIssue") {
                $issue_collection.Add($item)
            }
        }
    }
    end {
        if ($issue_collection.Count -gt 0) {
            if ($Format -eq 'Html' -or !$Format) {
                $html_content = [System.Collections.Generic.List[string]]@()
                $html_content.Add("<html><head><meta charset=""UTF-8""><title>$($report_title)</title>")
                $html_content.Add('<style type="text/css">')
                $html_content.Add($css_string)
                $html_content.Add("</style>")
                $html_content.Add("</head><body>")
                $html_content.Add("<hr>")
                $html_content.Add('<table id="section"><tr><th><a name="summary">Summary</a></th></tr></table>')
                $html_content.Add("<hr>")
                $html_content.Add('<table id="data">')
                $html_content.Add("<tr><th>Service</th><th>Event ID</th><th>Classification</th><th>Status</th><th>Title</th></tr>")
                foreach ($item in $issue_collection) {
                    $html_content.Add("<tr><td>$($item.Service)</td>
                        <td>" + '<a href="#' + $($item.ID) + '">' + "$($item.ID)</a></td>
                        <td>$($item.Classification)</td>
                        <td>$($item.Status)</td>
                        <td>$($item.Title)</td></tr>")
                }
                $html_content.Add('</table>')

                foreach ($item in $issue_collection) {
                    $html_content.Add("<hr>")
                    $html_content.Add('<table id="section"><tr><th><a name="' + $item.ID + '">' + $item.ID + '</a> | ' + $item.Service + ' | ' + $item.Title + '</th></tr></table>')
                    $html_content.Add("<hr>")
                    $html_content.Add('<table id="data">')
                    $html_content.Add('<tr><th>Status</th><td><b>' + $item.Status + '</b></td></tr>')
                    $html_content.Add('<tr><th>Organization</th><td>' + $OrganizationName + '</td></tr>')
                    $html_content.Add('<tr><th>Classification</th><td>' + $($item.Classification.substring(0, 1).toupper() + $item.Classification.substring(1)) + '</td></tr>')
                    $html_content.Add('<tr><th>User Impact</th><td>' + $item.ImpactDescription + '</td></tr>')
                    $html_content.Add('<tr><th>Last Updated</th><td>' + "{0:yyyy-MM-dd H:mm}" -f [datetime]$item.lastModifiedDateTime + '</td></tr>')
                    $html_content.Add('<tr><th>Start Time</th><td>' + "{0:yyyy-MM-dd H:mm}" -f [datetime]$item.startDateTime + '</td></tr>')
                    $html_content.Add('<tr><th>End Time</th><td>' + $(
                            if ($item.endDateTime) {
                                "{0:yyyy-MM-dd H:mm}" -f [datetime]$item.endDateTime
                            }
                        ) + '</td></tr>')
                    $latestMessage = ($item.posts[-1].description.content) -replace "`n", "<br />"
                    $html_content.Add('<tr><th>Latest Message</th><td>' + $latestMessage + '</td></tr>')
                    $html_content.Add('</table>')
                    $html_content.Add('<div class="back-to-summary"><a href = "#summary">(back to summary)</a></div>')
                }
                $html_content.Add('<p><font size="2" face="Segoe UI Light"><br />')
                $html_content.Add('<br />')
                $html_content.Add('<a href="' + $moduleInfo.ProjectURI.ToString() + '" target="_blank">' + $moduleInfo.Name.ToString() + ' v' + $moduleInfo.Version.ToString() + ' </a><br></p>')
                $html_content.Add('</body>')
                $html_content.Add('</html>')
                $html_content = $html_content -join "`n" # convert to multiline string

                # Write HTML report to file
                if ($HtmlReportFileName) {
                    $html_report_file = (Resolve-Path $HtmlReportFileName).Path
                    $html_content | Out-File $html_report_file
                    "HTML Report saved @ $($html_report_file)" | SayInfo
                }
            }

            # Create Teams Card
            if ($Format -eq 'TeamsCard' -or !$Format) {
                $teams_card_content = (NewTeamsCardJson -InputObject $issue_collection -Title $report_title)

                if ($TeamsCardFileName) {
                    $teams_card_report_file = (Resolve-Path $TeamsCardFileName).Path
                    $teams_card_content | Out-File $teams_card_report_file
                    "JSON Report saved @ $($teams_card_report_file)" | SayInfo
                }
            }

            # create the result object
            $result = [PSCustomObject]([ordered]@{
                    PSTypeName          = 'M365ServiceHealthReport'
                    OrganizationName    = $OrganizationName
                    Title               = $report_title
                    ReportGeneratedDate = $issue_collection[0].ReportGeneratedDate
                    ReportStartDate     = $issue_collection[0].ReportStartDate
                    Issues              = $issue_collection
                    HtmlFilename        = $(if ($html_report_file) { $html_report_file } else { 'None' })
                    HtmlContent         = $(if ($html_content) { $html_content } else { 'None' })
                    TeamsCardFileName   = $(if ($teams_card_report_file) { $teams_card_report_file } else { 'None' })
                    TeamsCardContent    = $(if ($teams_card_content) { $teams_card_content } else { 'None' })
                })

            # Script method to get issue summary
            $result | Add-Member -MemberType ScriptMethod -Name GetSummary -Value {
                $([PSCustomObject]([ordered]@{
                            Count      = $($this.Issues.Count)
                            Resolved   = $(($this.Issues | Where-Object { $_.IsResolved }).Count)
                            Unresolved = $(($this.Issues | Where-Object { !$_.IsResolved }).Count)
                        }))
            }

            # Script method to get issue summary by service
            $result | Add-Member -MemberType ScriptMethod -Name GetSummaryByService -Value {
                $(foreach ($item in ($this.Issues | Group-Object Service | Sort-Object Count -Descending | Select-Object Count, Name)) {
                        [PSCustomObject]([ordered]@{
                                Service    = $item.Name
                                Count      = $item.Count
                                Resolved   = $(($this.Issues | Where-Object { $_.IsResolved -and $_.Service -eq $item.Name }).Count)
                                Unresolved = $(($this.Issues | Where-Object { !$_.IsResolved -and $_.Service -eq $item.Name }).Count)
                            })
                    })
            }

            $visible_properties = [string[]]@('Title', 'ReportGeneratedDate', 'ReportStartDate', 'Issues', 'HtmlFilename', 'TeamsCardFileName')
            [Management.Automation.PSMemberInfo[]]$default_properties = [System.Management.Automation.PSPropertySet]::new('DefaultDisplayPropertySet', $visible_properties)
            $result | Add-Member -MemberType MemberSet -Name PSStandardMembers -Value $default_properties
            return $result
        }
    }
}