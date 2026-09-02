function ConvertTo-M365ServiceHealthReportObject {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [PSTypeNameAttribute('Microsoft.Graph.PowerShell.Models.MicrosoftGraphServiceHealthIssue')]
        $InputObject,

        [Parameter()]
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
        $TeamsCardFileName,

        # Parameter help description
        [Parameter()]
        [switch]
        $IncludeHealthOverviewHTML,

        [Parameter()]
        [switch]
        $IncludeResolvedInOverviewForTesting
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
                throw "Failed to create the HTML report file [$($HtmlReportFileName)]. $($_.Exception.Message)"
                # SayError $_
                # continue
            }
        }

        if ($TeamsCardFileName -and !(Test-Path -Path $TeamsCardFileName)) {
            try {
                $null = New-Item -ItemType File -Path $TeamsCardFileName -Force -ErrorAction Stop
            }
            catch {
                throw "Failed to create the Teams Adaptive Card JSON file [$($TeamsCardFileName)]. $($_.Exception.Message)"
                # SayError $_
                # continue
            }
        }

        # Load the DOT images
        $privatePath = Join-Path -Path $moduleInfo.ModuleBase -ChildPath 'source\private'

        $activeIconPath = Join-Path -Path $privatePath -ChildPath 'active.png'
        $resolvedIconPath = Join-Path -Path $privatePath -ChildPath 'resolved.png'
        $advisoryIconPath = Join-Path -Path $privatePath -ChildPath 'advisory.png'
        $incidentIconPath = Join-Path -Path $privatePath -ChildPath 'incident.png'

        $activeBase64 = Get-ImageBase64String -Path $activeIconPath
        $resolvedBase64 = Get-ImageBase64String -Path $resolvedIconPath
        $advisoryBase64 = Get-ImageBase64String -Path $advisoryIconPath
        $incidentBase64 = Get-ImageBase64String -Path $incidentIconPath

        $activeDataUri = 'data:image/png;base64,' + $activeBase64
        $resolvedDataUri = 'data:image/png;base64,' + $resolvedBase64
        $advisoryDataUri = 'data:image/png;base64,' + $advisoryBase64
        $incidentDataUri = 'data:image/png;base64,' + $incidentBase64

        ## Get the CSS style
        $css_string = Get-Content (($moduleInfo.ModuleBase.ToString()) + '\source\private\style.css') -Raw

        $issue_collection = [System.Collections.Generic.list[System.Object]]@()

        if (!$Title) {
            $report_title = "[$($OrganizationName)] Microsoft 365 Service Health Report"
        }
        else {
            $report_title = $Title
        }

        $isResolvedColor = @{
            True  = '#107C10'
            False = '#CA5010'
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
                $incidentCount = ($issue_collection | Where-Object {
                        $_.Classification -eq 'Incident'
                    }).Count
                $advisoryCount = ($issue_collection | Where-Object {
                        $_.Classification -eq 'Advisory'
                    }).Count

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


                $classificationIconHtml = @{
                    'Incident' = '<img src="' + $incidentDataUri + '" ' + 'alt="Incident"' + ' width="18" height="18" style="vertical-align:top;">&nbsp;'
                    'Advisory' = '<img src="' + $advisoryDataUri + '" ' + 'alt="Advisory"' + ' width="18" height="18" style="vertical-align:top;">&nbsp;'
                }

                $resolutionStateIconHtml = @{
                    'True'  = '<img src="' + $resolvedDataUri + '" ' + 'alt="Resolved"' + ' width="18" height="18" style="vertical-align:top;">&nbsp;'
                    'False' = '<img src="' + $activeDataUri + '" ' + 'alt="Active"' + ' width="18" height="18" style="vertical-align:top;">&nbsp;'
                }

                $encodedRunId = ConvertTo-HtmlEncodedText -Text $issue_collection[0].RunId.ToString()

                $html_content.Add(
                    '<table class="report-header" width="100%" cellpadding="0" cellspacing="0" border="0">' +
                    '<tr>' +
                    '<td>' +
                    '<div class="report-header-title">' + ($encodedReportTitle.Replace("[$($OrganizationName)] ", '')) + '</div>' +
                    '<div class="report-header-meta">' +
                    'Organization: ' + $encodedOrganizationName + '<br />' +
                    'Run ID: ' + $encodedRunId + '<br />' +
                    'Generated: ' + $reportGeneratedDate + '<br />' +
                    $(if ($reportStartDate) { 'Report Start: ' + $reportStartDate }) +
                    '</div>' +
                    '</td>' +
                    '</tr>' +
                    '</table>'
                )

                $html_content.Add('<hr>')

                if ($IncludeHealthOverviewHTML) {
                    $html_content.Add(
                        (New-ServiceHealthOverviewHtml -IncludeResolvedInOverviewForTesting:$IncludeResolvedInOverviewForTesting)
                    )
                }

                $html_content.Add('<hr>')
                $html_content.Add('<table class="section-table" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><th><a id="summary" name="summary">Summary of Issues</a></th></tr></table>')

                $summaryDescription = Get-ServiceHealthSummaryDescription `
                    -RetrievalContext $issue_collection[0].RetrievalContext

                $html_content.Add(
                    '<div style="' +
                    'font-family:Aptos,Calibri,''Segoe UI'',Arial,sans-serif;' +
                    'font-size:12px;' +
                    'line-height:18px;' +
                    'color:#666666;' +
                    'padding:0 0 10px 0;' +
                    '">' +
                    (ConvertTo-HtmlEncodedText -Text $summaryDescription) +
                    '</div>'
                )

                $html_content.Add(
                    '<div class="report-header-summary">' +
                    '<strong>Total Events:</strong> ' + $totalIssues + '<br>' +
                    '<strong>Status:</strong> ' +
                    ' &nbsp;&nbsp; ' + $resolutionStateIconHtml['False'] + '<strong>Active:</strong> ' + $activeIssues +
                    ' &nbsp;|&nbsp; ' + $resolutionStateIconHtml['True'] + '<strong>Resolved:</strong> ' + $resolvedIssues + '<br>' + '<strong>Classification:</strong> &nbsp;&nbsp; ' +
                    $classificationIconHtml['Incident'] + 'Incidents: ' + $incidentCount +
                    ' &nbsp;|&nbsp; ' +
                    $classificationIconHtml['Advisory'] + 'Advisories: ' + $advisoryCount +
                    '</div>'
                )

                $html_content.Add('<hr>')

                $html_content.Add(
                    '<table width="100%" cellpadding="0" cellspacing="0" border="0" ' +
                    'style="width:100%;border-collapse:collapse;">'
                )

                $html_content.Add('<table class="data-table" width="100%" cellpadding="0" cellspacing="0" border="0">')

                $html_content.Add(
                    '<tr>' +
                    '<th style="background-color:#00B388;color:#FFFFFF;border:1px solid #DDDDDD;padding:7px 8px;text-align:left;">Event ID</th>' +
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
                        '<td colspan="4" style="background-color:#EDEDED;color:#242424;font-weight:bold;padding:7px 8px;border:1px solid #DDDDDD;">' +
                        $serviceName + ' (' + $group.Count + ')' +
                        '</td>' +
                        '</tr>'
                    )

                    foreach ($item in ($issue_collection | Where-Object { $_.Service -eq $group.Name } | Sort-Object LastModifiedDateTime -Descending)) {
                        $anchorId = ConvertTo-HtmlAnchorId -Value $item.Id
                        $eventId = ConvertTo-HtmlEncodedText -Text $item.Id
                        $status = ConvertTo-HtmlEncodedText -Text $item.Status
                        $lastUpdated = ConvertTo-HtmlEncodedText -Text (Format-ServiceHealthDate -DateTime $item.LastModifiedDateTime)
                        $title = ConvertTo-HtmlEncodedText -Text $item.Title

                        $statusColor = $isResolvedColor[($item.IsResolved.ToString())]

                        $statusCellStyle = @(
                            'color:' + $statusColor
                            'padding:7px 8px;'
                        ) -join ';'

                        if (!$item.IsResolved) {
                            $statusCellStyle = "$($statusCellStyle)font-weight:bold;"
                        }

                        $html_content.Add(
                            '<tr>' +
                            '<td style="text-align:left;white-space:nowrap;border:1px solid #DDDDDD;padding:7px 8px;">&nbsp;&nbsp;' + $classificationIconHtml[$item.Classification] + '<a href="#' + $anchorId + '">' + $eventId + '</a></td>' +
                            '<td style="' + $statusCellStyle + '">' + $resolutionStateIconHtml[$item.IsResolved.ToString()] + $status + '</td>' +
                            '<td style="border:1px solid #DDDDDD;padding:7px 8px;white-space:nowrap;">' + $lastUpdated + '</td>' +
                            '<td style="border:1px solid #DDDDDD;padding:7px 8px;">' + $title + '</td>' +
                            '</tr>'
                        )
                    }
                }
                $html_content.Add('</table>')

                $html_content.Add('<hr>')
                $html_content.Add('<table class="section-table" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><th><a id="issues" name="issues">Issue Details</a></th></tr></table>')
                $html_content.Add('<hr>')

                # Individual issues table
                foreach ($group in $itemGroup) {
                    foreach ($item in ($issue_collection | Where-Object { $_.Service -eq $group.Name } | Sort-Object LastModifiedDateTime -Descending)) {
                        $anchorId = ConvertTo-HtmlAnchorId -Value $item.Id
                        $eventId = ConvertTo-HtmlEncodedText -Text $item.Id
                        $service = ConvertTo-HtmlEncodedText -Text $item.Service
                        $title = ConvertTo-HtmlEncodedText -Text $item.Title
                        $status = ConvertTo-HtmlEncodedText -Text $item.Status

                        $statusColor = $isResolvedColor[($item.IsResolved.ToString())]

                        # if ($item.IsResolved) {
                        #     $statusFontSize = '12px'
                        # }
                        # else {
                        #     $statusFontSize = '18px'
                        # }

                        # if ($item.IsResolved) {
                        #     $statusFontSize = '12px'
                        #     $statusLineHeight = '16px'
                        # }
                        # else {
                        #     $statusFontSize = '18px'
                        #     $statusLineHeight = '22px'
                        # }

                        # =====================================
                        # Issue header
                        # =====================================
                        # Outer table
                        $html_content.Add('<table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;border:none;">')
                        $html_content.Add('<tr>')
                        # Left border color
                        $html_content.Add('<td style="background-color:#F3F2F1;border-top:1px solid #DDDDDD;border-right:1px solid #DDDDDD;border-bottom:none;border-left:6px solid ' + $statusColor + ';mso-border-left-alt:6px solid ' + $statusColor + ';padding:10px 12px 10px 12px;">')
                        # Inner table 1
                        $html_content.Add('<table width="100%" cellpadding="0" cellspacing="0" border="0" style="width:100%;border-collapse:collapse;">')

                        # Service row+column
                        $html_content.Add('<tr>')
                        $html_content.Add('<td style="font-family:Aptos,Calibri,''Segoe UI'',Arial,sans-serif;font-size:14px;line-height:16px;color:#666666;padding:0 0 4px 0;mso-line-height-rule:exactly;">' +
                            $service +
                            '</td>')
                        $html_content.Add('</tr>')

                        # EventID row+column
                        $adminCenterUrl = 'https://admin.cloud.microsoft/?#/servicehealth/:/alerts/' + [System.Uri]::EscapeDataString($item.Id)
                        $encodedAdminCenterUrl = ConvertTo-HtmlEncodedText -Text $adminCenterUrl
                        $html_content.Add('<tr>')
                        $html_content.Add('<td style="font-family:Aptos,Calibri,''Segoe UI'',Arial,sans-serif;font-size:18px;line-height:22px;font-weight:bold;color:#242424;padding:0 0 6px 0;mso-line-height-rule:exactly;">' + $classificationIconHtml[$item.Classification] +
                            '<a id="' + $anchorId + '" name="' + $anchorId + '" target="_blank" href="' + $encodedAdminCenterUrl + '">' + $eventId + '</a>' +
                            '</td>')
                        $html_content.Add('</tr>')

                        # Status row+column
                        $html_content.Add('<tr>')
                        # $html_content.Add('<td style="font-family:Aptos,Calibri,''Segoe UI'',Arial,sans-serif;font-size:' + $statusFontSize + ';line-height:' + $statusLineHeight + ';font-weight:bold;color:' + $statusColor + ';padding:0 0 8px 0;mso-line-height-rule:exactly;">' +
                        $html_content.Add('<td style="font-family:Aptos,Calibri,''Segoe UI'',Arial,sans-serif;font-size:18px;line-height:22px;font-weight:bold;color:' + $statusColor + ';padding:0 0 8px 0;mso-line-height-rule:exactly;">' +
                            "$($resolutionStateIconHtml[$item.IsResolved.ToString()])$($status)" +
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

                        $impactDescription = ConvertTo-HtmlEncodedText -Text $item.ImpactDescription

                        # Inner table 2
                        $html_content.Add('<table class="data-table" width="100%" cellpadding="0" cellspacing="0" border="0">')
                        # $html_content.Add('<tr><th style="width:120px;border-top:none;border-left:none;">Classification</th><td style="border-top:none;border-right:none;">' + $classification + '</td></tr>')
                        $html_content.Add('<tr><th style="width:120px;border-left:none;">User Impact</th><td style="border-right:none;">' + $impactDescription + '</td></tr>')
                        if ($item.IsResolved -and $item.EndDateTime) {
                            $duration = New-TimeSpan `
                                -Start $item.StartDateTime `
                                -End $item.EndDateTime

                            $html_content.Add('<tr><th style="width:120px;border-left:none;">Duration</th><td style="border-right:none;">' + (Format-ServiceHealthDuration -TimeSpan $duration) + '</td></tr>')
                        }
                        else {
                            $age = New-TimeSpan `
                                -Start $item.StartDateTime `
                                -End (Get-Date).ToUniversalTime()

                            $html_content.Add('<tr><th style="width:120px;border-left:none;">Age</th><td style="border-right:none;">' + (Format-ServiceHealthDuration -TimeSpan $Age) + '</td></tr>')
                        }
                        $html_content.Add('<tr><th style="width:120px;border-left:none;">Start Time</th><td style="border-right:none;">' + (Format-ServiceHealthDate -DateTime $item.StartDateTime) + '</td></tr>')
                        if ($item.endDateTime) {
                            $html_content.Add('<tr><th style="width:120px;border-left:none;">End Time</th><td style="border-right:none;">' + $(
                                    (Format-ServiceHealthDate -DateTime $item.EndDateTime)
                                ) + '</td></tr>')
                        }

                        $latestMessage = Get-ServiceHealthLatestMessageHtml -Issue $item

                        if ($item.Status -eq 'Post Incident Review Published') {
                            $pirDownloadURL =
                            'https://admin.cloud.microsoft/admin/api/servicehealth/postincidentreport?url=' +
                            'https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/issues(%27' +
                            [System.Uri]::EscapeDataString($item.Id) +
                            '%27)/incidentreport'

                            $latestMessage += '<br /><br />' +
                            '<a href="' + (ConvertTo-HtmlEncodedText -Text $pirDownloadURL) + '" target="_blank">Download post-incident report</a>'
                        }

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
                            -OrganizationName $OrganizationName `
                            -RunId $issue.RunId)
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
                    RunId               = $issue_collection[0].RunId
                    RetrievalContext    = $issue_collection[0].RetrievalContext
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

            $visible_properties = [string[]]@('RunId', 'Title', 'ReportGeneratedDate', 'ReportStartDate', 'Issues', 'HtmlFilename', 'TeamsCardFileName')
            [Management.Automation.PSMemberInfo[]]$default_properties = [System.Management.Automation.PSPropertySet]::new('DefaultDisplayPropertySet', $visible_properties)
            $result | Add-Member -MemberType MemberSet -Name PSStandardMembers -Value $default_properties
            return $result
        }
    }
}