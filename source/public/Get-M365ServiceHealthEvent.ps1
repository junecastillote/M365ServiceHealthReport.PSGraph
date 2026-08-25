function Get-M365ServiceHealthEvent {
    [CmdletBinding()]
    [Alias('Get-M365ServiceHealthIssue', 'Get-M365ServiceHealthAnnouncement')]
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

        [Parameter(Mandatory, ParameterSetName = 'StartFromLastSuccessfulRun')]
        [switch]$StartFromLastSuccessfulRun,

        [Parameter(Mandatory, ParameterSetName = 'StartFromRunId')]
        [string]$StartFromRunId,

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'PastDays')]
        [Parameter(ParameterSetName = 'LastModifiedDateTime')]
        [Parameter(Mandatory, ParameterSetName = 'StartFromLastSuccessfulRun')]
        [Parameter(Mandatory, ParameterSetName = 'StartFromRunId')]
        [string]$RunHistoryFileName,

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'PastDays')]
        [Parameter(ParameterSetName = 'LastModifiedDateTime')]
        [Parameter(ParameterSetName = 'StartFromLastSuccessfulRun')]
        [ValidateSet('Resolved', 'Unresolved')]
        [string]
        $Status,

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'PastDays')]
        [Parameter(ParameterSetName = 'LastModifiedDateTime')]
        [Parameter(ParameterSetName = 'StartFromLastSuccessfulRun')]
        [ValidateSet('Advisory', 'Incident')]
        [string]
        $Classification,

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'PastDays')]
        [Parameter(ParameterSetName = 'LastModifiedDateTime')]
        [Parameter(ParameterSetName = 'StartFromLastSuccessfulRun')]
        [string[]]$Service
    )

    #Region Helper
    function CreateHistoryFile {
        [CmdletBinding()]
        param ()

        if ($RunHistoryFileName) {
            if (!(Test-Path $RunHistoryFileName) -or !(Import-Csv $RunHistoryFileName -ErrorAction SilentlyContinue)) {

                $null = New-Item -Path $RunHistoryFileName -Force

                [PSCustomObject]@{
                    RunId   = [guid]::NewGuid().Guid
                    RunTime = "{0:yyyy-MM-ddTHH:mm:sszzzz}" -f ([System.DateTime]::Now)
                    Result  = "OK"
                    Note    = "Initial"
                } | Export-Csv -Append -Path $RunHistoryFileName -Force -Confirm:$false
            }
            else {
                $RunHistoryFileName = (Resolve-Path $RunHistoryFileName).Path
                UpgradeHistoryFile
            }
        }
    }

    function GetLastSuccessfulRunTime {
        [CmdletBinding()]
        [OutputType([datetime])]
        param (

        )
        if ($RunHistoryFileName) {
            CreateHistoryFile
            [datetime](@(Import-Csv $RunHistoryFileName | Where-Object { $_.Result -eq 'Ok' })[-1].RunTime)
        }
    }

    function GetRunTimeByRunId {
        [CmdletBinding()]
        [OutputType([datetime])]
        param (

        )

        if (-not $StartFromRunId) {
            Write-Debug "NOT StartFromRunId"
            return
        }

        if ($RunHistoryFileName) {
            CreateHistoryFile
            [datetime](Import-Csv $RunHistoryFileName | Where-Object { $_.RunId -eq $($StartFromRunId) }).RunTime
        }
    }

    function WriteToHistoryFile {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [ValidateSet('OK', 'NotOK', 'NoResult')]
            [string]$Result,
            [Parameter()]
            [string]$Note
        )

        if ($RunHistoryFileName) {
            CreateHistoryFile
            [PSCustomObject]@{
                RunId   = $(if ($CurrentRunId) { $CurrentRunId }else { [guid]::NewGuid().Guid })
                RunTime = "{0:yyyy-MM-ddTHH:mm:sszzzz}" -f ([System.DateTime]::Now)
                Result  = $Result
                Note    = $Note
            } | Export-Csv -Append -Path $RunHistoryFileName -Force -Confirm:$false
        }
    }

    function UpgradeHistoryFile {
        [CmdletBinding()]
        param ()

        if (!$RunHistoryFileName) {
            return
        }

        if (!(Test-Path $RunHistoryFileName)) {
            return
        }

        try {
            $historyData = Import-Csv -Path $RunHistoryFileName -ErrorAction Stop
        }
        catch {
            throw "Failed to read history file [$RunHistoryFileName]. $_"
        }

        # Already upgraded
        if (
            $historyData.Count -gt 0 -and
            ($historyData[0].PSObject.Properties.Name -contains 'RunId')
        ) {
            return
        }

        "Upgrading history file [$RunHistoryFileName] to include RunId column." | Write-Debug

        $upgradedData = foreach ($row in $historyData) {
            [PSCustomObject][ordered]@{
                RunId   = [guid]::NewGuid().Guid
                RunTime = $row.RunTime
                Result  = $row.Result
                Note    = $row.Note
            }
        }

        $upgradedData |
        Export-Csv `
            -Path $RunHistoryFileName `
            -NoTypeInformation `
            -Force
    }
    #EndRegion Helper

    $currentRunId = [guid]::NewGuid().Guid
    Write-Debug "Current RunId: $($currentRunId)"
    SayInfo "RunId: $($currentRunId)"
    $now = ([System.DateTime]::Now)

    # Initialize the filter (empty)
    $filter = @()

    # Add Id filter
    if ($PSBoundParameters.ContainsKey('Id')) {
        $filter += "Id eq '$($Id)'"
        Write-Debug "Filter (Id) - $Id"
    }

    $start_date = ([System.DateTime]::MinValue).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')

    # Add PastDays filter
    if ($PSBoundParameters.ContainsKey('PastDays')) {
        $start_date = (($now).AddDays(-$PastDays)).ToUniversalTime().ToString('yyyy-MM-ddT00:00:00Z')
        $filter += "LastModifiedDateTime ge $($start_date)"
        Write-Debug "Filter (PastDays) - $PastDays, LastModifiedDateTime - $start_date"
    }

    # Add LastModifiedDateTime filter
    if ($PSBoundParameters.ContainsKey('LastModifiedDateTime')) {
        $start_date = ($LastModifiedDateTime).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        $filter += "LastModifiedDateTime ge $($start_date)"
        Write-Debug "Filter (LastModifiedDateTime) - $start_date"
    }

    # Add StartFromLastSuccessfulRun filter
    if ($PSBoundParameters.ContainsKey('StartFromLastSuccessfulRun')) {
        $LastUpdatedTime = GetLastSuccessfulRunTime
        $start_date = ($LastUpdatedTime).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:00Z')
        $filter += "LastModifiedDateTime ge $($start_date)"
        Write-Debug "Filter (StartFromLastSuccessfulRun) - $start_date"
    }

    # Add StartFromRunId filter
    if ($PSBoundParameters.ContainsKey('StartFromRunId')) {
        $StartFromRunId = $StartFromRunId
        # Write-Debug "OldRunId: $($StartFromRunId)"
        Write-Debug "Finding runtime by RunId [$($StartFromRunId)]"
        $LastUpdatedTime = GetRunTimeByRunId
        if (-not $LastUpdatedTime) {
            "RunId [$($StartFromRunId)] does not exist in history." | SayInfo
            return $null
        }
        $start_date = ($LastUpdatedTime).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:00Z')
        $filter += "LastModifiedDateTime ge $($start_date)"
        Write-Debug "Filter (StartFromLastSuccessfulRun) - $start_date"
    }

    # Add Classification filter
    if ($PSBoundParameters.ContainsKey('Classification')) {
        $filter += "Classification eq '$($Classification)'"
        Write-Debug "Filter (Classification) - $Classification"
    }

    # Add IsResolved filter
    if ($PSBoundParameters.ContainsKey('Status')) {
        switch ($Status) {
            'Resolved' { $filter += 'IsResolved eq true' }
            'Unresolved' { $filter += 'IsResolved eq false' }
        }
        Write-Debug "Filter (Status) - $Status"
    }

    # Add Service filter
    if ($Service) {
        $isValid = $true
        # Retrieve all valid service names list.
        $valid_service_list = @((Get-MgServiceAnnouncementHealthOverview -All | Sort-Object Service).Service)
        $invalid_service_list = @()
        $included_service_list = @()
        $service_filter = @()
        foreach ($item in $Service) {
            if ($item -notin $valid_service_list) {
                $isValid = $false
                $invalid_service_list += $item
            }
            else {
                $service_name = ($valid_service_list | Where-Object { $_ -eq $item })
                if (-not ($service_name -ceq $item)) {
                    Write-Debug "Correcting service name case from ($($item)) to ($($service_name))"
                    $item = $service_name
                }
                $included_service_list += $item
                $service_filter += "Service eq '$item'"
            }
        }
        if (!$isValid) {
            # Terminate if at least one of the service names is not valid.
            WriteToHistoryFile NotOK "Invalid Service [$($invalid_service_list -join ";")] specified."
            throw "The service names specified are not valid ($($invalid_service_list -join ", ")). Only the following service names are valid and accepted ($($valid_service_list -join ", "))"
        }

        $filter += "($($service_filter -join ' or '))"
        Write-Debug "Filter (Service) - $($included_service_list -join ",")"
    }

    try {
        switch ($true) {
            { $filter } {
                # Get issues with filter
                Write-Debug "Filter string = $($filter -join ' and ')"
                $issue_collection = @(Get-MgServiceAnnouncementIssue -Filter ($filter -join ' and ') -All -ErrorAction Stop)
            }
            { !$filter } {
                # Get issues without filter
                $issue_collection = @(Get-MgServiceAnnouncementIssue -All -ErrorAction Stop)
            }
        }

        if ($issue_collection) {
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

            # $issue_collection | Add-Member -MemberType NoteProperty -Name OrganizationName -Value (Get-MgOrganization).DisplayName
            $issue_collection | Add-Member -MemberType NoteProperty -Name ReportStartDate -Value (Get-Date $start_date).ToUniversalTime()
            $issue_collection | Add-Member -MemberType NoteProperty -Name ReportGeneratedDate -Value $now.ToUniversalTime()
            $issue_collection | Add-Member -MemberType NoteProperty -Name RunId -Value $currentRunId

            WriteToHistoryFile OK "Count: $($issue_collection.Count)"
            Write-Debug "Event count: $($issue_collection.Count)"
            $issue_collection
        }
        else {
            WriteToHistoryFile NoResult "No result"
        }
    }
    catch {
        WriteToHistoryFile NotOK "Failed to retrieve issues"
        throw "Failed to retrieve issues. $($_.Exception.Message)"
    }
}