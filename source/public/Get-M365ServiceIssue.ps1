Function Get-M365ServiceIssue {
    [CmdletBinding()]
    param (

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]
        $Id,

        [parameter()]
        [ValidateRange(1, [int]::MaxValue)]
        [int]$PastDays,

        [parameter()]
        [datetime]$LastModifiedDateTime,

        [Parameter()]
        [bool]
        $IsResolved,

        [parameter()]
        [ValidateSet('Advisory', 'Incident')]
        [string]
        $Classification,

        [parameter()]
        [string[]]$Service
    )

    # Initialize the filter (empty)
    $filter = @()

    # Terminate if the -Id parameter is used with other parameters.
    if ($PSBoundParameters.ContainsKey('Id') -and (
            $PSBoundParameters.ContainsKey('PastDays') -or
            $PSBoundParameters.ContainsKey('LastModifiedDateTime') -or
            $PSBoundParameters.ContainsKey('Status') -or
            $PSBoundParameters.ContainsKey('Classification') -or
            $PSBoundParameters.ContainsKey('Service')
        )) {
        Write-Error "The -Id parameter must be used alone."
        return $null
    }

    # Terminate if -LastModifiedDateTime and -PastDays are both in use.
    if ($PSBoundParameters.ContainsKey('LastModifiedDateTime') -and $PSBoundParameters.ContainsKey('PastDays')) {
        Write-Error "The -LastModifiedDateTime and -PastDays parameters cannot be used together."
        return $null
    }

    # Add Id filter
    if ($PSBoundParameters.ContainsKey('Id')) {
        $filter += "Id eq '$($Id)'"
    }

    # Add PastDays filter
    if ($PSBoundParameters.ContainsKey('PastDays')) {
        $start_date = (([datetime]::Today).AddDays(-$PastDays)).ToString('yyyy-MM-ddTHH:mm:ssZ')
        $filter += "LastModifiedDateTime ge $($start_date)"
    }

    # Add LastModifiedDateTime filter
    if ($PSBoundParameters.ContainsKey('LastModifiedDateTime')) {
        $start_date = ($LastModifiedDateTime).ToString('yyyy-MM-ddTHH:mm:ssZ')
    }

    # Add Classification filter
    if ($PSBoundParameters.ContainsKey('Classification')) {
        $filter += "Classification eq '$($Classification)'"
    }

    # Add IsResolved filter
    if ($PSBoundParameters.ContainsKey('IsResolved')) {
        switch ($IsResolved) {
            $true { $filter += 'IsResolved eq true' }
            $false { $filter += 'IsResolved eq false' }
        }
    }

    # Add Service filter
    if ($Service) {
        $isValid = $true
        $valid_service_list = @((Get-MgServiceAnnouncementHealthOverview -All | Sort-Object Service).Service)
        $service_filter = @()
        foreach ($item in $Service) {
            if ($item -notin $valid_service_list) {
                Write-Error "'$item' is not a valid service name."
                $isValid = $false
            }
            else {
                $service_filter += "Service eq '$item'"
            }
        }
        if (!$isValid) {
            Write-Error "Accepted service names: $($valid_service_list -join ";")"
            return $null
        }
        $filter += "($($service_filter -join ' or '))"
    }

    try {
        switch ($true) {
            { $filter } {
                # Get issues with filter
                Write-Verbose ($filter -join ' and ')
                $issue_collection = Get-MgServiceAnnouncementIssue -Filter ($filter -join ' and ') -All
            }
            { !$filter } {
                # Get issues without filter
                $issue_collection = Get-MgServiceAnnouncementIssue -All
            }
        }

        $issue_collection | Add-Member -MemberType NoteProperty -Name LastUpdateContent -Value ''
        $issue_collection | ForEach-Object {
            # Split the Status and convert to title case (ie. 'serviceDegradation' to 'Service Degradation')
            $_.Status = ($_.Status.substring(0, 1).toupper() + $_.Status.substring(1) -creplace '[^\p{Ll}\s]', ' $&').Trim()

            # Capitalize the first letter of the Origin (ie. 'microsoft' to 'Microsoft')
            $_.Origin = ($_.Origin.substring(0, 1).toupper() + $_.Origin.substring(1))

            # Capitalize the first letter of the Classification (ie. 'advisory' to 'Advisory')
            $_.Classification = ($_.Classification.substring(0, 1).toupper() + $_.Classification.substring(1))

            # Bring out the latest message from the Posts collection.
            $_.LastUpdateContent = $_.Posts[-1].Description.Content
        }
        # Return the results
        $issue_collection
    }
    catch {
        Write-Error "Failed to retrieve service health messages. $_"
        return $null
    }
}