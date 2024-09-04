Function ConvertTo-M365ServiceHealthReportObject {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,
            ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [PSTypeNameAttribute('Microsoft.Graph.PowerShell.Models.MicrosoftGraphServiceHealthIssue')]
        $InputObject,

        [Parameter(Mandatory)]
        [string]$OrganizationName,

        [Parameter()]
        [string]$Title,

        [Parameter()]
        [string]
        $HtmlReportFileName
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

            # $OrganizationName = $issue_collection[0].OrganizationName
            # $report_title = "[$($OrganizationName)] Microsoft 365 Service Health Report"
            if ($HtmlReportFileName) { $report_file = (Resolve-Path $HtmlReportFileName).Path }
            $report_body = [System.Collections.Generic.List[string]]@()
            $report_body.Add("<html><head><meta charset=""UTF-8""><title>$($report_title)</title>")
            $report_body.Add('<style type="text/css">')
            $report_body.Add($css_string)
            $report_body.Add("</style>")
            $report_body.Add("</head><body>")
            $report_body.Add("<hr>")
            $report_body.Add('<table id="section"><tr><th><a name="summary">Summary</a></th></tr></table>')
            $report_body.Add("<hr>")
            $report_body.Add('<table id="data">')
            $report_body.Add("<tr><th>Service</th><th>Event ID</th><th>Classification</th><th>Status</th><th>Title</th></tr>")
            foreach ($item in $issue_collection) {
                $report_body.Add("<tr><td>$($item.Service)</td>
                    <td>" + '<a href="#' + $($item.ID) + '">' + "$($item.ID)</a></td>
                    <td>$($item.Classification)</td>
                    <td>$($item.Status)</td>
                    <td>$($item.Title)</td></tr>")
            }
            $report_body.Add('</table>')

            foreach ($item in $issue_collection) {
                $report_body.Add("<hr>")
                $report_body.Add('<table id="section"><tr><th><a name="' + $item.ID + '">' + $item.ID + '</a> | ' + $item.Service + ' | ' + $item.Title + '</th></tr></table>')
                $report_body.Add("<hr>")
                $report_body.Add('<table id="data">')
                $report_body.Add('<tr><th>Status</th><td><b>' + $item.Status + '</b></td></tr>')
                $report_body.Add('<tr><th>Organization</th><td>' + $OrganizationName + '</td></tr>')
                $report_body.Add('<tr><th>Classification</th><td>' + $($item.Classification.substring(0, 1).toupper() + $item.Classification.substring(1)) + '</td></tr>')
                $report_body.Add('<tr><th>User Impact</th><td>' + $item.ImpactDescription + '</td></tr>')
                $report_body.Add('<tr><th>Last Updated</th><td>' + "{0:yyyy-MM-dd H:mm}" -f [datetime]$item.lastModifiedDateTime + '</td></tr>')
                $report_body.Add('<tr><th>Start Time</th><td>' + "{0:yyyy-MM-dd H:mm}" -f [datetime]$item.startDateTime + '</td></tr>')
                $report_body.Add('<tr><th>End Time</th><td>' + $(
                        if ($item.endDateTime) {
                            "{0:yyyy-MM-dd H:mm}" -f [datetime]$item.endDateTime
                        }
                    ) + '</td></tr>')
                $latestMessage = ($item.posts[-1].description.content) -replace "`n", "<br />"
                $report_body.Add('<tr><th>Latest Message</th><td>' + $latestMessage + '</td></tr>')
                $report_body.Add('</table>')
                # $report_body.Add('<div style="font-family: Tahoma;font-size: 10px"><a href = "#summary">(back to summary)</a></div>')
                $report_body.Add('<div class="back-to-summary"><a href = "#summary">(back to summary)</a></div>')
            }
            $report_body.Add('<p><font size="2" face="Segoe UI Light"><br />')
            $report_body.Add('<br />')
            $report_body.Add('<a href="' + $moduleInfo.ProjectURI.ToString() + '" target="_blank">' + $moduleInfo.Name.ToString() + ' v' + $moduleInfo.Version.ToString() + ' </a><br></p>')
            $report_body.Add('</body>')
            $report_body.Add('</html>')
            $report_body = $report_body -join "`n" # convert to multiline string

            # Write HTML report to file
            if ($report_file) {
                $report_body | Out-File $report_file
                "Report saved @ $($report_file)" | SayInfo
            }

            # create the result object
            $result = [PSCustomObject]([ordered]@{
                    PSTypeName          = 'M365ServiceHealthReport'
                    OrganizationName    = $OrganizationName
                    Title               = $report_title
                    ReportGeneratedDate = $issue_collection[0].ReportGeneratedDate
                    ReportStartDate     = $issue_collection[0].ReportStartDate
                    Issues              = $issue_collection
                    Filename            = $(if ($report_file) { $report_file } else { 'None' })
                    Content             = $report_body
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

            $visible_properties = [string[]]@('Title', 'ReportGeneratedDate', 'ReportStartDate', 'Issues', 'Filename')
            [Management.Automation.PSMemberInfo[]]$default_properties = [System.Management.Automation.PSPropertySet]::new('DefaultDisplayPropertySet', $visible_properties)
            $result | Add-Member -MemberType MemberSet -Name PSStandardMembers -Value $default_properties
            return $result
        }
    }
}