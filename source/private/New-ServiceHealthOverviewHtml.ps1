function New-ServiceHealthOverviewHtml {
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateNotNull()]
        $HealthOverview,

        # For testing/debugging only to include already resolved issues
        [Parameter()]
        [switch]
        $IncludeResolvedInOverviewForTesting
    )

    if ($null -eq $HealthOverview -or @($HealthOverview).Count -eq 0) {
        try {
            $HealthOverview = Get-MgServiceAnnouncementHealthOverview `
                -ExpandProperty Issues `
                -ErrorAction Stop
        }
        catch {
            throw "Failed to retrieve the Microsoft 365 service health overview. $($_.Exception.Message)"
        }
    }

    if (@($HealthOverview).Count -eq 0) {
        throw 'The Microsoft 365 service health overview returned no services.'
    }

    $html = [System.Collections.Generic.List[string]]::new()

    $columnsPerRow = 6

    $tileTextColors = @{
        Operational = '#5B8E2D'
        Advisory    = '#CA5010'
        Incident    = '#D13438'
    }

    $moduleInfo = Get-Module $MyInvocation.MyCommand.ModuleName

    if (!$moduleInfo) {
        throw "Unable to determine module information for [$($MyInvocation.MyCommand.ModuleName)]."
    }

    $privatePath = Join-Path `
        -Path $moduleInfo.ModuleBase `
        -ChildPath 'source\private'

    $resolvedIconPath = Join-Path `
        -Path $privatePath `
        -ChildPath 'resolved.png'

    $advisoryIconPath = Join-Path `
        -Path $privatePath `
        -ChildPath 'advisory.png'

    $incidentIconPath = Join-Path `
        -Path $privatePath `
        -ChildPath 'incident.png'

    $resolvedDataUri = 'data:image/png;base64,' +
    (Get-ImageBase64String -Path $resolvedIconPath)

    $advisoryDataUri = 'data:image/png;base64,' +
    (Get-ImageBase64String -Path $advisoryIconPath)

    $incidentDataUri = 'data:image/png;base64,' +
    (Get-ImageBase64String -Path $incidentIconPath)

    $imgTagStart = '<' + 'img'
    $imgTagEnd = '>'

    $tileIcons = @{
        Operational = (
            $imgTagStart +
            ' src="' + $resolvedDataUri + '"' +
            ' alt="Operational"' +
            ' width="18"' +
            ' height="18"' +
            ' style="vertical-align:top;"' +
            $imgTagEnd
        )
        Advisory    = (
            $imgTagStart +
            ' src="' + $advisoryDataUri + '"' +
            ' alt="Advisory"' +
            ' width="18"' +
            ' height="18"' +
            ' style="vertical-align:top;"' +
            $imgTagEnd
        )
        Incident    = (
            $imgTagStart +
            ' src="' + $incidentDataUri + '"' +
            ' alt="Incident"' +
            ' width="18"' +
            ' height="18"' +
            ' style="vertical-align:top;"' +
            $imgTagEnd
        )
    }

    $services = foreach ($service in $HealthOverview) {

        if ($IncludeResolvedInOverviewForTesting) {
            $incidentCount = @(
                $service.Issues |
                Where-Object {
                    $_.Classification -eq 'Incident'
                }
            ).Count

            $advisoryCount = @(
                $service.Issues |
                Where-Object {
                    $_.Classification -eq 'Advisory'
                }
            ).Count
        }
        else {
            $incidentCount = @(
                $service.Issues |
                Where-Object {
                    $_.Classification -eq 'Incident' -and
                    -not $_.IsResolved
                }
            ).Count

            $advisoryCount = @(
                $service.Issues |
                Where-Object {
                    $_.Classification -eq 'Advisory' -and
                    -not $_.IsResolved
                }
            ).Count
        }

        if ($incidentCount -gt 0 -or $advisoryCount -gt 0) {
            Write-Debug (
                '{0}: Incident = {1}, Advisory = {2}' -f
                $service.Service,
                $incidentCount,
                $advisoryCount
            )
        }

        $state = if ($incidentCount -gt 0) {
            'Incident'
        }
        elseif ($advisoryCount -gt 0) {
            'Advisory'
        }
        else {
            'Operational'
        }

        $displayStatus = if ([System.String]::IsNullOrWhiteSpace($service.Status)) {
            'Unknown'
        }
        else {
            (
                $service.Status.Substring(0, 1).ToUpperInvariant() +
                $service.Status.Substring(1) -creplace '[^\p{Ll}\s]', ' $&'
            ).Trim()
        }

        [PSCustomObject][ordered]@{
            Service       = $service.Service
            Status        = $service.Status
            DisplayStatus = $displayStatus
            State         = $state
            IncidentCount = $incidentCount
            AdvisoryCount = $advisoryCount
            SortPriority  = switch ($state) {
                'Incident' {
                    0
                }

                'Advisory' {
                    1
                }

                default {
                    2
                }
            }
        }
    }

    $services = @(
        $services |
        Sort-Object `
        @{ Expression = { $_.SortPriority }; Ascending = $true },
        @{ Expression = { $_.Service }; Ascending = $true }
    )

    $html.Add(
        '<table class="section-table" width="100%" cellpadding="0" cellspacing="0" border="0">' +
        '<tr>' +
        '<th>Health Overview</th>' +
        '</tr>' +
        '</table>'
    )

    $overviewDescription = if ($IncludeResolvedInOverviewForTesting) {
        'This section provides an at-a-glance view of Microsoft 365 service health in your tenant. Resolved issues are included for testing and debugging.'
    }
    else {
        'This section provides an at-a-glance view of Microsoft 365 service health in your tenant. Services with active incidents or advisories are highlighted to help identify areas requiring attention.'
    }

    $html.Add(
        '<div style="' +
        'font-family:Aptos,Calibri,''Segoe UI'',Arial,sans-serif;' +
        'font-size:12px;' +
        'line-height:18px;' +
        'color:#666666;' +
        'padding:0 0 10px 0;' +
        '">' +
        (ConvertTo-HtmlEncodedText -Text $overviewDescription) +
        '</div>'
    )

    $html.Add('<hr>')

    $html.Add(
        '<table width="100%" cellpadding="0" cellspacing="0" border="0" ' +
        'style="width:100%;border-collapse:collapse;">'
    )

    $column = 0

    foreach ($service in $services) {
        if ($column -eq 0) {
            $html.Add('<tr>')
        }

        $textColor = $tileTextColors[$service.State]
        $icon = $tileIcons[$service.State]

        $detailLines = [System.Collections.Generic.List[string]]::new()

        $hasOngoingIssue = (
            $service.IncidentCount -gt 0 -or
            $service.AdvisoryCount -gt 0
        )

        if ($hasOngoingIssue) {
            $encodedDisplayStatus = ConvertTo-HtmlEncodedText `
                -Text $service.DisplayStatus

            $detailLines.Add(
                '<strong>' + $encodedDisplayStatus + '</strong>'
            )

            if ($service.IncidentCount -gt 0) {
                $incidentLabel = if ($service.IncidentCount -eq 1) {
                    'Incident'
                }
                else {
                    'Incidents'
                }

                $detailLines.Add(
                    "$($service.IncidentCount) $incidentLabel"
                )
            }

            if ($service.AdvisoryCount -gt 0) {
                $advisoryLabel = if ($service.AdvisoryCount -eq 1) {
                    'Advisory'
                }
                else {
                    'Advisories'
                }

                $detailLines.Add(
                    "$($service.AdvisoryCount) $advisoryLabel"
                )
            }
        }
        else {
            $detailLines.Add(
                '<strong>' +
                (ConvertTo-HtmlEncodedText -Text $service.DisplayStatus) +
                '</strong>'
            )
        }

        $detailText = $detailLines -join '<br />'
        $serviceName = ConvertTo-HtmlEncodedText -Text $service.Service
        $tileWidth = 100 / $columnsPerRow

        $html.Add(
            '<td width="' + $tileWidth + '%" valign="top" ' +
            'style="' +
            'background-color:#F8F8F8;' +
            'border:1px solid #DDDDDD;' +
            'padding:10px;' +
            'height:90px;' +
            '">' +

            '<div style="' +
            'font-family:Aptos,Calibri,''Segoe UI'',Arial,sans-serif;' +
            'font-size:18px;' +
            'line-height:22px;' +
            'height:22px;' +
            '">' +
            $icon +
            '</div>' +

            '<div style="' +
            'font-family:Aptos,Calibri,''Segoe UI'',Arial,sans-serif;' +
            'font-size:15px;' +
            'line-height:20px;' +
            'font-weight:bold;' +
            'color:#242424;' +
            'padding-top:4px;' +
            '">' +
            $serviceName +
            '</div>' +

            '<div style="' +
            'font-family:Aptos,Calibri,''Segoe UI'',Arial,sans-serif;' +
            'font-size:12px;' +
            'line-height:16px;' +
            'color:' + $textColor + ';' +
            'padding-top:8px;' +
            '">' +
            $detailText +
            '</div>' +

            '</td>'
        )

        $column++

        if ($column -eq $columnsPerRow) {
            $html.Add('</tr>')
            $column = 0
        }
    }

    if ($column -gt 0) {
        while ($column -lt $columnsPerRow) {
            $html.Add(
                '<td width="' + (100 / $columnsPerRow) + '%" ' +
                'style="border:none;">&nbsp;</td>'
            )

            $column++
        }

        $html.Add('</tr>')
    }

    $html.Add('</table>')

    return ($html -join "`n")
}