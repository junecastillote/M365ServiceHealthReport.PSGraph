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
            RunId   = $(if ($script:CurrentRunId) { $script:CurrentRunId }else { [guid]::NewGuid().Guid })
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