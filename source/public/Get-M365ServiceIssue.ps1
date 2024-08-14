Function Get-M365ServiceIssue {
    [CmdletBinding()]
    param (

        [Parameter()]
        [string]
        $Id,

        [parameter()]
        [int]$PastDays,

        [parameter()]
        [datetime]$LastUpdatedTime,

        [parameter()]
        [ValidateSet('Resolved', 'Unresolved')]
        [string]$Status,

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
    if ($id -and ($PastDays -or $LastUpdatedTime -or $Status -or $Classification -or $Service)) {
        Write-Error "The -Id parameter must be used alone."
        return $null
    }

    # Terminate if -LastUpdatedTime and -PastDays are both in use.
    if ($LastUpdatedTime -and $PastDays) {
        Write-Error "The -LastUpdatedTime and -PastDays parameters cannot be used together."
        return $null
    }

    if ($Id) {
        $filter += "Id eq '$($Id)'"
    }

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
    }

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

    switch ($true) {
        { $filter } {
            Write-Verbose ($filter -join ' and ')
            $issue_collection = Get-MgServiceAnnouncementIssue -Filter ($filter -join ' and ') -All
        }
        { !$filter } {
            $issue_collection = Get-MgServiceAnnouncementIssue -All
        }
    }

    $issue_collection
}