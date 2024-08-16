Function New-M365ServiceHealthReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,
            ValueFromPipeline)]
        $InputObject,

        [Parameter(Mandatory)]
        [string]
        $OutputDirectory,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]
        $OrganizationName
    )

    if (@($InputObject)[0].psobject.TypeNames[0] -notlike "*Microsoft.Graph.PowerShell.Models.MicrosoftGraphServiceHealthIssue") {
        Write-Error "The input object's PsTypeName does not match [Microsoft.Graph.PowerShell.Models.MicrosoftGraphServiceHealthIssue]"
        return $null
    }

    if (@($InputObject).Count -lt 1) {
        Write-Error "The input object is empty."
        return $null
    }

    if (!($PSBoundParameters.ContainsKey('OrganizationName'))) {
        $OrganizationName = (Get-MgOrganization).DisplayName
    }

    $moduleInfo = Get-Module $($MyInvocation.MyCommand.ModuleName)

    if (!(Test-Path -Path $OutputDirectory)) {
        try {
            $null = New-Item -ItemType Directory -Path $OutputDirectory -Force -ErrorAction Stop
            $null = New-Item -ItemType File -Path "$($OutputDirectory)\run_history.csv" -Force -ErrorAction Stop
        }
        catch {
            Write-Error $_
            return $null
        }
    }

    ## Get the CSS style
    $css_string = Get-Content (($moduleInfo.ModuleBase.ToString()) + '\source\private\style.css') -Raw

    $report_title = "[$($organizationName)] Microsoft 365 Service Health Report"
    $report_file = "$($OutputDirectory)\service_health_report.html"
    # $event_id_json_file = "$outputDir\consolidated_report.json"
    $report_body = [System.Collections.ArrayList]@()
    $null = $report_body.Add("<html><head><meta charset=""UTF-8""><title>$($report_title)</title>")
    $null = $report_body.Add('<style type="text/css">')
    $null = $report_body.Add($css_string)
    $null = $report_body.Add("</style>")
    $null = $report_body.Add("</head><body>")
    $null = $report_body.Add("<hr>")
    $null = $report_body.Add('<table id="section"><tr><th><a name="summary">Summary</a></th></tr></table>')
    $null = $report_body.Add("<hr>")
    $null = $report_body.Add('<table id="data">')
    $null = $report_body.Add("<tr><th>Workload</th><th>Event ID</th><th>Classification</th><th>Status</th><th>Title</th></tr>")

    foreach ($item in $InputObject) {
        $null = $report_body.Add("<tr><td>$($item.Service)</td>
            <td>" + '<a href="#' + $($item.ID) + '">' + "$($item.ID)</a></td>
            <td>$($item.Classification)</td>
            <td>$($item.Status)</td>
            <td>$($item.Title)</td></tr>")
    }
    $null = $report_body.Add('</table>')

    foreach ($item in $InputObject) {
        $null = $report_body.Add("<hr>")
        $null = $report_body.Add('<table id="section"><tr><th><a name="' + $item.ID + '">' + $item.ID + '</a> | ' + $item.Service + ' | ' + $item.Title + '</th></tr></table>')
        $null = $report_body.Add("<hr>")
        $null = $report_body.Add('<table id="data">')
        $null = $report_body.Add('<tr><th>Status</th><td><b>' + $item.Status + '</b></td></tr>')
        $null = $report_body.Add('<tr><th>Organization</th><td>' + $organizationName + '</td></tr>')
        $null = $report_body.Add('<tr><th>Classification</th><td>' + $($item.Classification.substring(0, 1).toupper() + $item.Classification.substring(1)) + '</td></tr>')
        $null = $report_body.Add('<tr><th>User Impact</th><td>' + $item.ImpactDescription + '</td></tr>')
        $null = $report_body.Add('<tr><th>Last Updated</th><td>' + "{0:yyyy-MM-dd H:mm}" -f [datetime]$item.lastModifiedDateTime + '</td></tr>')
        $null = $report_body.Add('<tr><th>Start Time</th><td>' + "{0:yyyy-MM-dd H:mm}" -f [datetime]$item.startDateTime + '</td></tr>')
        $null = $report_body.Add('<tr><th>End Time</th><td>' + $(
                if ($item.endDateTime) {
                    "{0:yyyy-MM-dd H:mm}" -f [datetime]$item.endDateTime
                }
            ) + '</td></tr>')

        $latestMessage = ($item.posts[-1].description.content) -replace "`n", "<br />"

        $null = $report_body.Add('<tr><th>Latest Message</th><td>' + $latestMessage + '</td></tr>')
        $null = $report_body.Add('</table>')
        # $null = $report_body.Add('<div style="font-family: Tahoma;font-size: 10px"><a href = "#summary">(back to summary)</a></div>')
        $null = $report_body.Add('<div class="back-to-summary"><a href = "#summary">(back to summary)</a></div>')
    }
    $null = $report_body.Add('<p><font size="2" face="Segoe UI Light"><br />')
    $null = $report_body.Add('<br />')
    $null = $report_body.Add('<a href="' + $moduleInfo.ProjectURI.ToString() + '" target="_blank">' + $moduleInfo.Name.ToString() + ' v' + $moduleInfo.Version.ToString() + ' </a><br></p>')
    $null = $report_body.Add('</body>')
    $null = $report_body.Add('</html>')
    $report_body = $report_body -join "`n" # convert to multiline string
    $report_body | Out-File $report_file


    [PSCustomObject]@{
        ReportFile = (Resolve-Path $report_file).Path
    }

}