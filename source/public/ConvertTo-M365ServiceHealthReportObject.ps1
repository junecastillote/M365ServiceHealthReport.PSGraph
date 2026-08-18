function ConvertTo-M365ServiceHealthReportObject {
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

        #Region Helper

        # Moved to module_root\source\private\report_object_helpers.ps1

        #EndRegion Helper

        $moduleInfo = Get-Module $($MyInvocation.MyCommand.ModuleName)

        if ($HtmlReportFileName -and !(Test-Path -Path $HtmlReportFileName)) {
            try {
                $null = New-Item -ItemType File -Path $HtmlReportFileName -Force -ErrorAction Stop
            }
            catch {
                SayError $_
                continue
            }
        }

        if ($TeamsCardFileName -and !(Test-Path -Path $TeamsCardFileName)) {
            try {
                $null = New-Item -ItemType File -Path $TeamsCardFileName -Force -ErrorAction Stop
            }
            catch {
                SayError $_
                continue
            }
        }

        # Load the DOT images
        $privatePath = Join-Path -Path $moduleInfo.ModuleBase -ChildPath 'source\private'

        $yellowDotPath = Join-Path -Path $privatePath -ChildPath 'yellow-dot.png'
        $redDotPath = Join-Path -Path $privatePath -ChildPath 'red-dot.png'

        $yellowDotBase64 = Get-ImageBase64String -Path $yellowDotPath
        $redDotBase64 = Get-ImageBase64String -Path $redDotPath

        $yellowDotDataUri = 'data:image/png;base64,' + $yellowDotBase64
        $redDotDataUri = 'data:image/png;base64,' + $redDotBase64

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

                $encodedReportTitle = ConvertTo-HtmlEncodedText -Text $report_title
                $encodedOrganizationName = ConvertTo-HtmlEncodedText -Text $OrganizationName

                $totalIssues = $issue_collection.Count
                $resolvedIssues = ($issue_collection | Where-Object { $_.IsResolved }).Count
                $activeIssues = ($issue_collection | Where-Object { !$_.IsResolved }).Count

                $reportGeneratedDate = Format-ServiceHealthDate -DateTime $issue_collection[0].ReportGeneratedDate

                $reportStartDate = ''

                if ($issue_collection[0].ReportStartDate -ne ([System.DateTime]::MinValue).ToUniversalTime()) {
                    $reportStartDate = Format-ServiceHealthDate -DateTime $issue_collection[0].ReportStartDate
                }

                $html_content.Add('<html><head><meta charset="UTF-8"><title>' + $encodedReportTitle + '</title>')
                $html_content.Add('<style type="text/css">')
                $html_content.Add($css_string)
                $html_content.Add('</style>')
                $html_content.Add('</head><body>')

                $html_content.Add(
                    '<table class="report-header" width="100%" cellpadding="0" cellspacing="0" border="0">' +
                    '<tr>' +
                    '<td>' +
                    '<div class="report-header-title">' + ($encodedReportTitle.Replace("[$($OrganizationName)] ", '')) + '</div>' +
                    '<div class="report-header-meta">' +
                    'Organization: ' + $encodedOrganizationName + '<br />' +
                    'Generated: ' + $reportGeneratedDate + '<br />' +
                    $(if ($reportStartDate) { 'Report Start: ' + $reportStartDate }) +
                    '</div>' +
                    '<div class="report-header-summary">' +
                    '<strong>Total Events:</strong> ' + $totalIssues +
                    ' &nbsp;|&nbsp; <strong>Active:</strong> ' + $activeIssues +
                    ' &nbsp;|&nbsp; <strong>Resolved:</strong> ' + $resolvedIssues +
                    '</div>' +
                    '</td>' +
                    '</tr>' +
                    '</table>'
                )

                # $html_content.Add('<hr>')
                $html_content.Add('<table class="section-table" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><th><a id="summary" name="summary">Summary</a></th></tr></table>')
                $html_content.Add('<hr>')

                $html_content.Add('<table class="data-table" width="100%" cellpadding="0" cellspacing="0" border="0">')

                $html_content.Add(
                    # '<tr>' +
                    '<tr>' +
                    '<th style="background-color:#00B388;color:#FFFFFF;border:1px solid #DDDDDD;padding:7px 8px;text-align:left;">Event ID</th>' +
                    '<th style="background-color:#00B388;color:#FFFFFF;border:1px solid #DDDDDD;padding:7px 8px;text-align:left;">Classification</th>' +
                    '<th style="background-color:#00B388;color:#FFFFFF;border:1px solid #DDDDDD;padding:7px 8px;text-align:left;">Status</th>' +
                    '<th style="background-color:#00B388;color:#FFFFFF;border:1px solid #DDDDDD;padding:7px 8px;text-align:left;white-space:nowrap;">Last Updated</th>' +
                    '<th style="background-color:#00B388;color:#FFFFFF;border:1px solid #DDDDDD;padding:7px 8px;text-align:left;">Title</th>' +
                    '</tr>'
                )

                $itemGroup = $issue_collection |
                Group-Object Service |
                Sort-Object @{ Expression = 'Count'; Descending = $true }, @{ Expression = 'Name'; Ascending = $true }

                foreach ($group in $itemGroup) {

                    $serviceName = ConvertTo-HtmlEncodedText -Text $group.Name

                    $html_content.Add(
                        '<tr class="group-row">' +
                        '<td colspan="5" style="background-color:#EDEDED;color:#242424;font-weight:bold;padding:7px 8px;border:1px solid #DDDDDD;">' +
                        $serviceName + ' (' + $group.Count + ')' +
                        '</td>' +
                        '</tr>'
                    )

                    foreach ($item in ($issue_collection | Where-Object { $_.Service -eq $group.Name } | Sort-Object LastModifiedDateTime -Descending)) {
                        $anchorId = ConvertTo-HtmlAnchorId -Value $item.Id
                        $eventId = ConvertTo-HtmlEncodedText -Text $item.Id
                        $classification = Get-ServiceHealthClassificationHtml `
                            -Classification $item.Classification `
                            -YellowDotSource $yellowDotDataUri `
                            -RedDotSource $redDotDataUri
                        $status = ConvertTo-HtmlEncodedText -Text $item.Status
                        $lastUpdated = ConvertTo-HtmlEncodedText -Text (Format-ServiceHealthDate -DateTime $item.LastModifiedDateTime)
                        $title = ConvertTo-HtmlEncodedText -Text $item.Title

                        $statusColor = if ($item.IsResolved) {
                            '#107C10'
                        }
                        else {
                            '#D13438'
                        }

                        $statusCellStyle = @(
                            'border-top:1px solid #DDDDDD'
                            'border-right:1px solid #DDDDDD'
                            'border-bottom:1px solid #DDDDDD'
                            'border-left:6px solid ' + $statusColor
                            'mso-border-left-alt:6px solid ' + $statusColor
                            'padding:7px 8px'
                        ) -join ';'

                        $html_content.Add(
                            '<tr>' +
                            '<td style="text-align:left;white-space:nowrap;border:1px solid #DDDDDD;padding:5px 8px;">&nbsp;&nbsp;&#8227;<a href="#' + $anchorId + '">' + $eventId + '</a></td>' +
                            '<td style="border:1px solid #DDDDDD;padding:7px 8px;">' + $classification + '</td>' +
                            # '<td style="' + $statusCellStyle + '">' + $status + '</td>' +
                            '<td style="' + $statusCellStyle + '">' + $(
                                if ($item.IsResolved) {
                                    "Resolved | $($status)"
                                }
                                else {
                                    "Active | $($status)"
                                }
                            ) + '</td>' +
                            '<td style="border:1px solid #DDDDDD;padding:7px 8px;white-space:nowrap;">' + $lastUpdated + '</td>' +
                            '<td style="border:1px solid #DDDDDD;padding:7px 8px;">' + $title + '</td>' +
                            '</tr>'
                        )
                    }
                }
                $html_content.Add('</table>')

                $html_content.Add('<table class="section-table" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><th><a id="issues" name="issues">Issues</a></th></tr></table>')
                $html_content.Add('<hr>')
                # Individual issues table
                foreach ($group in $itemGroup) {

                    foreach ($item in ($issue_collection | Where-Object { $_.Service -eq $group.Name } | Sort-Object LastModifiedDateTime -Descending)) {
                        $anchorId = ConvertTo-HtmlAnchorId -Value $item.Id
                        $eventId = ConvertTo-HtmlEncodedText -Text $item.Id
                        $service = ConvertTo-HtmlEncodedText -Text $item.Service
                        $title = ConvertTo-HtmlEncodedText -Text $item.Title
                        $status = ConvertTo-HtmlEncodedText -Text $item.Status
                        $leftColor = if ($item.IsResolved) {
                            '#107C10'
                        }
                        else {
                            '#D13438'
                        }

                        # =====================================
                        # Issue header
                        # =====================================
                        # Outer table
                        $html_content.Add('<table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;border:none;">')
                        $html_content.Add('<tr>')
                        # Left border color
                        $html_content.Add('<td style="background-color:#F3F2F1;border-top:1px solid #DDDDDD;border-right:1px solid #DDDDDD;border-bottom:none;border-left:6px solid ' + $leftColor + ';mso-border-left-alt:6px solid ' + $leftColor + ';padding:10px 12px 10px 12px;">')
                        # Inner table 1
                        $html_content.Add('<table width="100%" cellpadding="0" cellspacing="0" border="0" style="width:100%;border-collapse:collapse;">')

                        # Service row+column
                        $html_content.Add('<tr>')
                        $html_content.Add('<td style="font-family:Aptos,Calibri,''Segoe UI'',Arial,sans-serif;font-size:12px;line-height:16px;color:#666666;padding:0 0 4px 0;mso-line-height-rule:exactly;">' +
                            $service +
                            '</td>')
                        $html_content.Add('</tr>')

                        # EventID row+column
                        $adminCenterUrl = 'https://admin.cloud.microsoft/?#/servicehealth/:/alerts/' + [System.Uri]::EscapeDataString($item.Id)
                        $encodedAdminCenterUrl = ConvertTo-HtmlEncodedText -Text $adminCenterUrl
                        $html_content.Add('<tr>')
                        $html_content.Add('<td style="font-family:Aptos,Calibri,''Segoe UI'',Arial,sans-serif;font-size:18px;line-height:22px;font-weight:bold;color:#242424;padding:0 0 6px 0;mso-line-height-rule:exactly;">' +
                            '<a id="' + $anchorId + '" name="' + $anchorId + '" target="_blank" href="' + $encodedAdminCenterUrl + '">' + $eventId + '</a>' +
                            '</td>')
                        $html_content.Add('</tr>')

                        # Status row+column
                        $html_content.Add('<tr>')
                        $html_content.Add('<td style="font-family:Aptos,Calibri,''Segoe UI'',Arial,sans-serif;font-size:12px;line-height:16px;font-weight:bold;color:#8A5A00;padding:0 0 8px 0;mso-line-height-rule:exactly;">' +
                            $status +
                            '</td>')
                        $html_content.Add('</tr>')

                        # Title row+column
                        $html_content.Add('<tr>')
                        $html_content.Add('<td style="font-family:Aptos,Calibri,''Segoe UI'',Arial,sans-serif;font-size:14px;line-height:18px;color:#242424;padding:0 0 8px 0;mso-line-height-rule:exactly;">' +
                            $title +
                            '</td>')
                        $html_content.Add('</tr>')

                        # Close inner table 1
                        $html_content.Add('</table>')

                        $status = ConvertTo-HtmlEncodedText -Text $item.Status
                        $classification = Get-ServiceHealthClassificationHtml `
                            -Classification $item.Classification `
                            -YellowDotSource $yellowDotDataUri `
                            -RedDotSource $redDotDataUri
                        $impactDescription = ConvertTo-HtmlEncodedText -Text $item.ImpactDescription

                        # Inner table 2
                        $html_content.Add('<table class="data-table" width="100%" cellpadding="0" cellspacing="0" border="0">')
                        $html_content.Add('<tr><th style="width:120px;border-top:none;border-left:none;">Classification</th><td style="border-top:none;border-right:none;">' + $classification + '</td></tr>')
                        $html_content.Add('<tr><th style="width:120px;border-left:none;">User Impact</th><td style="border-right:none;">' + $impactDescription + '</td></tr>')
                        $html_content.Add('<tr><th style="width:120px;border-left:none;">Start Time</th><td style="border-right:none;">' + (Format-ServiceHealthDate -DateTime $item.StartDateTime) + '</td></tr>')
                        if ($item.endDateTime) {
                            $html_content.Add('<tr><th style="width:120px;border-left:none;">End Time</th><td style="border-right:none;">' + $(
                                    (Format-ServiceHealthDate -DateTime $item.EndDateTime)
                                ) + '</td></tr>')
                        }
                        $html_content.Add('<tr><th style="width:120px;border-left:none;">Last Updated</th><td style="border-right:none;">' + (Format-ServiceHealthDate -DateTime $item.LastModifiedDateTime) + '</td></tr>')

                        $latestMessage = Get-ServiceHealthLatestMessageHtml -Issue $item
                        $html_content.Add('<tr><th style="width:120px;border-bottom:none;border-left:none;">Update</th><td style="border-right:none;border-bottom:none;">' + $latestMessage + '</td></tr>')

                        # Close inner table 2
                        $html_content.Add('</table>')

                        # Close outer table cell (containing the inner table 1, 2)
                        $html_content.Add('</td>')
                        $html_content.Add('</tr>')

                        # Close outer table
                        $html_content.Add('</table>')

                        $html_content.Add('<div class="back-to-summary"><a href = "#summary">back to summary</a></div>')
                    }
                }

                $projectUri = ConvertTo-HtmlEncodedText -Text $moduleInfo.ProjectURI.ToString()
                $moduleName = ConvertTo-HtmlEncodedText -Text $moduleInfo.Name.ToString()
                $moduleVersion = ConvertTo-HtmlEncodedText -Text $moduleInfo.Version.ToString()

                $html_content.Add('<p class="report-footer"><br />')
                $html_content.Add('<a href="' + $projectUri + '">' + $moduleName + ' v' + $moduleVersion + '</a><br />')
                $html_content.Add('</p>')

                $html_content.Add('</body>')
                $html_content.Add('</html>')
                $html_content = $html_content -join "`n" # convert to multiline string

                # Write HTML report to file
                if ($HtmlReportFileName) {
                    $html_report_file = (Resolve-Path $HtmlReportFileName).Path
                    $html_content | Out-File $html_report_file -Encoding utf8
                    "HTML Report saved @ $($html_report_file)" | SayInfo
                }
            }

            # Create Teams alert cards
            if ($Format -eq 'TeamsCard' -or !$Format) {
                $teams_card_content = [System.Collections.Generic.List[string]]@()

                foreach ($issue in (
                        $issue_collection |
                        Sort-Object `
                        @{ Expression = { Get-ServiceHealthPriority $_ } ; Ascending = $true },
                        @{ Expression = 'LastModifiedDateTime' ; Descending = $true }
                    )) {
                    $teams_card_content.Add(
                        (New-ServiceHealthAlertCardJson `
                            -Issue $issue `
                            -OrganizationName $OrganizationName)
                    )
                }

                if ($TeamsCardFileName) {
                    $teams_card_report_file = (Resolve-Path $TeamsCardFileName).Path

                    "[" + ($teams_card_content -join ",") + "]" | Out-File $teams_card_report_file -Encoding utf8

                    "Teams alert card JSON saved @ $($teams_card_report_file)" | SayInfo
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