Function CreateHistoryFile {
    [CmdletBinding()]
    param ()

    if ($RunHistoryFileName) {
        if (!(Test-Path $RunHistoryFileName) -or !(Get-Content $RunHistoryFileName -Raw -ErrorAction SilentlyContinue)) {
            $null = New-Item -Path $RunHistoryFileName -Force
            "RunTime,Result,Note" | Set-Content -Path $RunHistoryFileName -Force -Confirm:$false
            # Add initial entry 'OK' (which means successful) dated 7 days ago. This way there will always be a starting point.
            "$("{0:yyyy-MM-dd H:mm}" -f ([System.DateTime]::Now).AddDays(-7)),OK,Initial" | Add-Content -Path $RunHistoryFileName -Force -Confirm:$false
        }
        else {
            $RunHistoryFileName = (Resolve-Path $RunHistoryFileName).Path
        }
    }
}

Function GetLastSuccessfulRunTime {
    [CmdletBinding()]
    param (

    )

    if ($RunHistoryFileName) {
        CreateHistoryFile
        Get-Date (@(Import-Csv $RunHistoryFileName | Where-Object { $_.Result -eq 'Ok' })[-1].RunTime)
    }
}

Function WriteToHistoryFile {
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
        "$("{0:yyyy-MM-dd H:mm}" -f ([System.DateTime]::Now)),$($Result),$($Note)" | Add-Content -Path $RunHistoryFileName -Force -Confirm:$false
    }
}



