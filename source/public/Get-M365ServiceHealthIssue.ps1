Function Get-M365ServiceHealthIssue {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ParameterSetName = 'Id')]
        [ValidateNotNullOrEmpty()]
        [string]
        $Id,

        [Parameter(Mandatory, ParameterSetName = 'PastDays')]
        [ValidateRange(1, [int]::MaxValue)]
        [int]$PastDays,

        [Parameter(Mandatory, ParameterSetName = 'LastModifiedDateTime')]
        [datetime]$LastModifiedDateTime,

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'PastDays')]
        [Parameter(ParameterSetName = 'LastModifiedDateTime')]
        [bool]
        $IsResolved,

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'PastDays')]
        [Parameter(ParameterSetName = 'LastModifiedDateTime')]
        [ValidateSet('Advisory', 'Incident')]
        [string]
        $Classification,

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'PastDays')]
        [Parameter(ParameterSetName = 'LastModifiedDateTime')]
        [string[]]$Service
    )

    $now = ([System.DateTime]::Now)

    # Initialize the filter (empty)
    $filter = @()

    # Add Id filter
    if ($PSBoundParameters.ContainsKey('Id')) {
        $filter += "Id eq '$($Id)'"
    }

    $start_date = ([System.DateTime]::MinValue).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')

    # Add PastDays filter
    if ($PSBoundParameters.ContainsKey('PastDays')) {
        $start_date = (($now).AddDays(-$PastDays)).ToUniversalTime().ToString('yyyy-MM-ddT00:00:00Z')
        $filter += "LastModifiedDateTime ge $($start_date)"
    }

    # Add LastModifiedDateTime filter
    if ($PSBoundParameters.ContainsKey('LastModifiedDateTime')) {
        $start_date = ($LastModifiedDateTime).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        $filter += "LastModifiedDateTime ge $($start_date)"
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
        # Retrieve all valid service names list.
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
            # Terminate if at least one of the service names is not valid.
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

        $issue_collection | Add-Member -MemberType NoteProperty -Name OrganizationName -Value (Get-MgOrganization).DisplayName
        $issue_collection | Add-Member -MemberType NoteProperty -Name ReportStartDate -Value (Get-Date $start_date).ToUniversalTime()
        $issue_collection | Add-Member -MemberType NoteProperty -Name ReportGeneratedDate -Value $now.ToUniversalTime()

        # Return the results
        $issue_collection
    }
    catch {
        Write-Error "Failed to retrieve service health messages. $_"
        return $null
    }
}