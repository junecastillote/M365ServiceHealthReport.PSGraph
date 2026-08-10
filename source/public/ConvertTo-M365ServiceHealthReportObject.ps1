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
        function Format-ServiceHealthDate {
            [CmdletBinding()]
            param(
                [AllowNull()]
                $DateTime
            )

            if (!$DateTime) {
                return ''
            }

            return '{0:MMM dd, yyyy, hh:mm tt} UTC' -f [datetime]$DateTime
        }

        function ConvertTo-HtmlEncodedText {
            [CmdletBinding()]
            param(
                [AllowNull()]
                [string]$Text
            )

            if ([system.string]::IsNullOrEmpty($Text)) {
                return ''
            }

            return [System.Net.WebUtility]::HtmlEncode($Text)
        }

        function ConvertTo-HtmlAnchorId {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                [ValidateNotNullOrEmpty()]
                [string]$Value
            )

            return ($Value -replace '[^a-zA-Z0-9_-]', '-')
        }

        function Get-ServiceHealthLatestMessageHtml {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                [ValidateNotNull()]
                $Issue
            )

            $latestPost = if ($Issue.Posts -and $Issue.Posts.Count -gt 0) {
                $Issue.Posts[-1]
            }

            if (
                !$latestPost -or
                !$latestPost.Description -or
                [System.String]::IsNullOrWhiteSpace($latestPost.Description.Content)
            ) {
                return 'No latest message available.'
            }

            $message = ConvertTo-HtmlEncodedText -Text $latestPost.Description.Content
            $message = $message -replace "(\r\n|\n|\r)", '<br />'

            return $message
        }

        function Format-ServiceHealthText {
            [CmdletBinding()]
            param(
                [AllowNull()]
                [string]$Text
            )

            if ([System.String]::IsNullOrWhiteSpace($Text)) {
                return ''
            }

            if ($Text.Length -eq 1) {
                return $Text.ToUpperInvariant()
            }

            return $Text.Substring(0, 1).ToUpperInvariant() + $Text.Substring(1)
        }

        function Get-ImageBase64String {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                [ValidateNotNullOrEmpty()]
                [string]$Path
            )

            if (!(Test-Path -Path $Path -PathType Leaf)) {
                throw "Image file not found: $Path"
            }

            $bytes = [System.IO.File]::ReadAllBytes($Path)
            return [System.Convert]::ToBase64String($bytes)
        }

        function Get-ServiceHealthClassificationIconSource {
            [CmdletBinding()]
            param(
                [AllowNull()]
                [string]$Classification,

                [Parameter(Mandatory)]
                [ValidateNotNullOrEmpty()]
                [string]$YellowDotSource,

                [Parameter(Mandatory)]
                [ValidateNotNullOrEmpty()]
                [string]$RedDotSource
            )

            if ([System.String]::IsNullOrWhiteSpace($Classification)) {
                return $YellowDotSource
            }

            switch ($Classification.Trim().ToLowerInvariant()) {
                'incident' { return $RedDotSource }
                default { return $YellowDotSource }
            }
        }

        function Get-ServiceHealthClassificationHtml {
            [CmdletBinding()]
            param(
                [AllowNull()]
                [string]$Classification,

                [Parameter(Mandatory)]
                [ValidateNotNullOrEmpty()]
                [string]$YellowDotSource,

                [Parameter(Mandatory)]
                [ValidateNotNullOrEmpty()]
                [string]$RedDotSource
            )

            $classificationText = Format-ServiceHealthText -Text $Classification
            $encodedClassification = ConvertTo-HtmlEncodedText -Text $classificationText

            $iconSource = Get-ServiceHealthClassificationIconSource `
                -Classification $classificationText `
                -YellowDotSource $YellowDotSource `
                -RedDotSource $RedDotSource

            $altText = if ($classificationText) {
                ConvertTo-HtmlEncodedText -Text $classificationText
            }
            else {
                'Classification'
            }

            return '<img src="' + $iconSource + '" ' + 'alt="' + $altText + '"' + ' width="12" height="12">&nbsp;' + $encodedClassification
        }

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

                $html_content.Add('<hr>')
                $html_content.Add('<table class="section-table" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><th><a id="summary" name="summary">Summary</a></th></tr></table>')
                $html_content.Add('<hr>')

                $html_content.Add('<table class="data-table" width="100%" cellpadding="0" cellspacing="0" border="0">')

                $html_content.Add(
                    '<tr>' +
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
                            'padding:7px 8px'
                        ) -join ';'

                        $html_content.Add(
                            '<tr>' +
                            '<td style="text-align:left;white-space:nowrap;border:1px solid #DDDDDD;padding:5px 8px;">&nbsp;&nbsp;&#8227;<a href="#' + $anchorId + '">' + $eventId + '</a></td>' +
                            '<td style="border:1px solid #DDDDDD;padding:7px 8px;">' + $classification + '</td>' +
                            '<td style="' + $statusCellStyle + '">' + $status + '</td>' +
                            '<td style="border:1px solid #DDDDDD;padding:7px 8px;white-space:nowrap;">' + $lastUpdated + '</td>' +
                            '<td style="border:1px solid #DDDDDD;padding:7px 8px;">' + $title + '</td>' +
                            '</tr>'
                        )
                    }
                }
                $html_content.Add('</table>')

                foreach ($item in $issue_collection) {
                    # foreach ($item in ($issue_collection | Sort-Object LastModifiedDateTime -Descending)) {
                    $anchorId = ConvertTo-HtmlAnchorId -Value $item.Id
                    $eventId = ConvertTo-HtmlEncodedText -Text $item.Id
                    $service = ConvertTo-HtmlEncodedText -Text $item.Service
                    $title = ConvertTo-HtmlEncodedText -Text $item.Title
                    $status = ConvertTo-HtmlEncodedText -Text $item.Status
                    $html_content.Add('<hr>')
                    $leftColor = if ($item.IsResolved) {
                        '#107C10'
                    }
                    else {
                        '#D13438'
                    }
                    $anchorId = ConvertTo-HtmlAnchorId -Value $item.Id
                    $eventId = ConvertTo-HtmlEncodedText -Text $item.Id
                    $service = ConvertTo-HtmlEncodedText -Text $item.Service
                    $title = ConvertTo-HtmlEncodedText -Text $item.Title
                    $status = ConvertTo-HtmlEncodedText -Text $item.Status
                    $html_content.Add(
                        '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;border:none;">' +
                        '<tr>' +
                        '<td style="background-color:#F3F2F1;border-top:1px solid #DDDDDD;border-right:1px solid #DDDDDD;border-bottom:1px solid #DDDDDD;border-left:6px solid ' + $leftColor + ';mso-border-left-alt:6px solid ' + $leftColor + ';padding:10px 12px 10px 12px;">' +
                        '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="width:100%;border-collapse:collapse;">' +
                        '<tr>' +
                        '<td style="font-family:Aptos,Calibri,''Segoe UI'',Arial,sans-serif;font-size:12px;line-height:16px;color:#666666;padding:0 0 4px 0;mso-line-height-rule:exactly;">' +
                        $service +
                        '</td>' +
                        '</tr>' +
                        '<tr>' +
                        '<td style="font-family:Aptos,Calibri,''Segoe UI'',Arial,sans-serif;font-size:18px;line-height:22px;font-weight:bold;color:#242424;padding:0 0 6px 0;mso-line-height-rule:exactly;">' +
                        '<a id="' + $anchorId + '" name="' + $anchorId + '">' + $eventId + '</a>' +
                        '</td>' +
                        '</tr>' +
                        '<tr>' +
                        '<td style="font-family:Aptos,Calibri,''Segoe UI'',Arial,sans-serif;font-size:12px;line-height:16px;font-weight:bold;color:#8A5A00;padding:0 0 8px 0;mso-line-height-rule:exactly;">' +
                        $status +
                        '</td>' +
                        '</tr>' +
                        '<tr>' +
                        '<td style="font-family:Aptos,Calibri,''Segoe UI'',Arial,sans-serif;font-size:14px;line-height:18px;color:#242424;padding:0;mso-line-height-rule:exactly;">' +
                        $title +
                        '</td>' +
                        '</tr>' +
                        '</table>' +
                        '</td>' +
                        '</tr>' +
                        '</table>'
                    )

                    $html_content.Add('<hr>')

                    $html_content.Add('<table class="data-table" width="100%" cellpadding="0" cellspacing="0" border="0">')

                    $status = ConvertTo-HtmlEncodedText -Text $item.Status
                    $classification = Get-ServiceHealthClassificationHtml `
                        -Classification $item.Classification `
                        -YellowDotSource $yellowDotDataUri `
                        -RedDotSource $redDotDataUri
                    $impactDescription = ConvertTo-HtmlEncodedText -Text $item.ImpactDescription

                    $html_content.Add('<tr><th>Classification</th><td>' + $classification + '</td></tr>')
                    $html_content.Add('<tr><th>User Impact</th><td>' + $impactDescription + '</td></tr>')
                    $html_content.Add('<tr><th>Last Updated</th><td>' + (Format-ServiceHealthDate -DateTime $item.LastModifiedDateTime) + '</td></tr>')
                    $html_content.Add('<tr><th>Start Time</th><td>' + (Format-ServiceHealthDate -DateTime $item.StartDateTime) + '</td></tr>')
                    $html_content.Add('<tr><th>End Time</th><td>' + $(
                            if ($item.endDateTime) {
                                (Format-ServiceHealthDate -DateTime $item.EndDateTime)
                            }
                        ) + '</td></tr>')
                    $latestMessage = Get-ServiceHealthLatestMessageHtml -Issue $item
                    $html_content.Add('<tr><th>Latest Message</th><td>' + $latestMessage + '</td></tr>')
                    $html_content.Add('</table>')
                    $html_content.Add('<div class="back-to-summary"><a href = "#summary">back to summary</a></div>')
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

            # Create Teams Card
            if ($Format -eq 'TeamsCard' -or !$Format) {
                $teams_card_content = (NewTeamsCardJson -InputObject $issue_collection -Title $report_title)

                if ($TeamsCardFileName) {
                    $teams_card_report_file = (Resolve-Path $TeamsCardFileName).Path
                    $teams_card_content | Out-File $teams_card_report_file -Encoding utf8
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