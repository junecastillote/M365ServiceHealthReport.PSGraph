Function Get-M365ServiceIssueList {
    [CmdletBinding()]
    param (
        [parameter()]
        [int]$PastDays,

        [parameter()]
        [datetime]$LastUpdatedTime,

        [parameter()]
        [ValidateSet('Resolved', 'Unresolved')]
        [string]$Status,

        [Parameter()]
        [ValidateSet('Advisory', 'Incident')]
        [string]
        $Classification,

        [Parameter()]
        [string[]]$Service
    )

    # Terminate if -LastUpdatedTime and -PastDays are both in use.
    if ($LastUpdatedTime -and $PastDays) {
        Write-Error "The -LastUpdatedTime and -PastDays parameters cannot be used together."
        return $null
    }

    # Initialize the filter (empty)
    $filter = @()

    if ($PastDays) {
        $start_date = (([datetime]::Today).AddDays(-$PastDays)).ToString('yyyy-MM-ddTHH:mm:ssZ')
        $filter += "LastModifiedDateTime ge $($start_date)"
    }

    if ($LastUpdatedTime) {
        $start_date = ($LastUpdatedTime).ToString('yyyy-MM-ddTHH:mm:ssZ')
    }

    if ($Classification) {
        $filter += "Classification eq '$($Classification.ToLower())'"
    }

    switch ($Status) {
        Resolved { $filter += 'IsResolved eq true' }
        Unresolved { $filter += 'IsResolved eq false' }
        Default {}
    }

    if ($Service) {
        $isValid = $true
        $valid_service_list = @((Get-MgServiceAnnouncementHealthOverview -All | Sort-Object Service).Service)
        $service_filter = @()
        foreach ($item in $Service) {
            if ($item -notin $valid_service_list) {
                Write-Error "'$item' is not a valid service name."
                $isValid = $false
                # return $null
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

    if ($filter) {
        Write-Verbose ($filter -join ' and ')
        $issue_collection = Get-MgServiceAnnouncementIssue -Filter ($filter -join ' and ') -All
    }
    else {
        $issue_collection = Get-MgServiceAnnouncementIssue -All
    }

    $issue_collection
}