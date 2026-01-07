<#
.SYNOPSIS
    Tracks and reports elapsed time between invocations.

.DESCRIPTION
    The ElapsedTime function records a start timestamp in a specified file on the first run.
    On the next invocation, it calculates the elapsed time since that start, returns the result,
    and deletes the stored timestamp file.

    Optionally, it can reset the timer, or output results as a formatted string instead of an object.

.PARAMETER TimeFile
    The file used to store the start time in ISO 8601 UTC format.
    Default: C:\Temp\time.txt

.PARAMETER Reset
    Deletes any existing timestamp file and starts a new timer immediately.

.PARAMETER AsString
    Returns human-readable output instead of an object (e.g., "Elapsed time: 0 days, 1 hours, 32 minutes").

.EXAMPLE
    PS> ElapsedTime
    # First run creates timestamp file.

.EXAMPLE
    PS> Start-Sleep -Seconds 90; ElapsedTime
    Days Hours Minutes
    ---- ----- --------
       0     0        2

.EXAMPLE
    PS> ElapsedTime -AsString
    Elapsed time: 0 days, 0 hours, 2 minutes

.EXAMPLE
    PS> ElapsedTime -Reset
    # Deletes existing time file and starts fresh.

.NOTES
    - Creates parent directory if missing.
    - Timestamp stored in ISO 8601 format (UTC).
    - Removes time file after reporting elapsed duration.
    - Designed for automation scripts needing persistent timing checkpoints.
#>
function ElapsedTime {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$TimeFile = 'C:\Temp\time.txt',

        [switch]$Reset,
        [switch]$AsString
    )

    # ---------------------------------------------------------------------
    # Ensure the directory exists for the timestamp file
    # ---------------------------------------------------------------------
    $dir = Split-Path -Parent $TimeFile
    if (-not (Test-Path -LiteralPath $dir -PathType Container)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    # ---------------------------------------------------------------------
    # Reset mode: start fresh by removing any existing time file
    # ---------------------------------------------------------------------
    if ($Reset) {
        Remove-Item -LiteralPath $TimeFile -Force -ErrorAction SilentlyContinue
    }

    # ---------------------------------------------------------------------
    # If no file exists, create a new one and record start time
    # ---------------------------------------------------------------------
    if (-not (Test-Path -LiteralPath $TimeFile)) {
        (Get-Date).ToUniversalTime().ToString('o') | Set-Content -LiteralPath $TimeFile -Encoding UTF8 -Force
        if ($AsString) {
            Write-Output ("Start time recorded: {0}" -f (Get-Date))
        }
        return
    }

    # ---------------------------------------------------------------------
    # Calculate elapsed time since the stored timestamp
    # ---------------------------------------------------------------------
    try {
        $startRaw = (Get-Content -LiteralPath $TimeFile -ErrorAction Stop | Out-String).Trim()

        # Validate and parse ISO 8601 UTC timestamp
        $startTime = [DateTime]::ParseExact(
            $startRaw,
            'o',
            $null,
            [System.Globalization.DateTimeStyles]::AssumeUniversal
        ).ToLocalTime()
    }
    catch {
        Write-Error "Invalid or unreadable timestamp in $TimeFile. Use -Reset to start a new timer."
        return
    }
    finally {
        # Clean up the timestamp file so the next call starts fresh
        Remove-Item -LiteralPath $TimeFile -Force -ErrorAction SilentlyContinue
    }

    # ---------------------------------------------------------------------
    # Compute elapsed duration
    # ---------------------------------------------------------------------
    $now  = Get-Date
    $span = $now - $startTime

    # Break down the TimeSpan into meaningful units
    $days    = [math]::Floor($span.TotalDays)
    $hours   = $span.Hours
    $minutes = [math]::Round($span.Minutes + ($span.Seconds / 60), 0)

    # Handle rollover edge cases
    if ($minutes -ge 60) {
        $minutes = 0
        $hours  += 1
    }
    if ($hours -ge 24) {
        $hours = 0
        $days += 1
    }

    # Construct output
    $result = [PSCustomObject]@{
        Days    = $days
        Hours   = $hours
        Minutes = $minutes
    }

    if ($AsString) {
        "Elapsed time: {0} days, {1} hours, {2} minutes" -f $days, $hours, $minutes
    }
    else {
        $result
    }
}

Export-ModuleMember -Function ElapsedTime