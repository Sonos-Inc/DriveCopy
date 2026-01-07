<#
.SYNOPSIS
    Sends an email with optional HTML-rendered CSV content and attachments.

.DESCRIPTION
    This script emails CSV-based reports with optional inline HTML previews.
    It supports multiple reports,auto-generates timestamped titles if needed,and handles
    "no report" runs by sending a basic summary email. The subject line and body header
    share a consistent timestamped title. Each CSV report is rendered in a scrollable 
    HTML table. A documentation link is included in the email body.

.PARAMETER ReportPath
    One or more CSV file paths to attach and render in the email.

.PARAMETER Title
    Optional subject/title for the report. If not provided,it's derived from the first CSV file name.

.PARAMETER EmailTo
    Recipient email address(es). Default is 'r-google-admin@sonos.com'.

.PARAMETER NoReport
    If specified,skips report parsing and sends a generic completion email.

.PARAMETER ReportFrom
    Email address to use as the sender. Default: 'google.admin@sonos.com'.

.NOTES
https://sonosinc.atlassian.net/wiki/spaces/PEOPLEPROD/pages/1346830354/Drive+Copy+Automation Architectural
https://sonosinc.atlassian.net/wiki/spaces/ITKB/pages/1348370617/Legacy+Drive+Backup ITKB
https://docs.google.com/spreadsheets/d/1_0WIDlZriHpb1_YqP1nPGs0lcOhliweGt-klJSPg2C4/edit?gid=1966817793#gid=1966817793 Spreadsheet
https://drive.google.com/drive/folders/0ADVPlcaDSgiTUk9PVA Legacy Drive Backup Folder
#>
function Send-EmailReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string[]]$ReportPath,

        [Parameter(Mandatory = $false)]
        [string]$Title,

        [Parameter(Mandatory = $false)]
        [ValidatePattern("^[a-zA-Z0-9](?:[a-zA-Z0-9._%+-]*[a-zA-Z0-9])?(?:@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})?$")]
        [ValidateNotNullOrEmpty()]
        [string]$EmailTo = 'r-google-admin@sonos.com',

        [switch]$NoReport,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$ReportFrom = "google.admin@sonos.com"
    )

    # Generates a formatted title using a provided title or fallback filename
    function Get-FormattedTitle {
        param (
            [datetime]$Now,
            [string]$BaseTitle,
            [string]$FallbackPath
        )

        # Format the timestamp (e.g.,20250623-1:02)
        $timestamp = $Now.ToString("yyyyMMdd-h_mm")

        # Use the provided title,or fall back to the first report file name
        if ([string]::IsNullOrWhiteSpace($BaseTitle)) {
            try {
                $fileName = [System.IO.Path]::GetFileNameWithoutExtension($FallbackPath)
            }
            catch {
                $fileName = "Report"
            }
        }
        else {
            return $BaseTitle
        }

        return "$timestamp $fileName"
    }

    # Builds the full HTML body with header,note,and optional table
    function BuildEmailBody {
        param (
            [string]$ReportTitle,
            [datetime]$DateStamp,
            [string]$HtmlTable = ''
        )

        return @"
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial; }
        .date { text-align: left; font-size: 16px; font-weight: bold; margin-bottom: 10px; }
        h1 { text-align: center; margin-bottom: 10px; }
        .note { text-align: left; font-size: 14px; margin-bottom: 20px; }
        table { border-collapse: collapse; width: 100%; }
        th,td { border: 1px solid black; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        .table-container {
            max-height: 300px;
            overflow-y: auto;
            border: 1px solid #ccc;
            padding: 4px;
            margin-bottom: 25px;
        }
        @media screen and (max-width: 600px) {
            table,th,td { display: block; width: 100%; }
            th { text-align: center; }
        }
    </style>
</head>
<body>
    <div class="date">Generated: $($DateStamp.ToString("yyyy-MM-dd h:mm tt"))</div>
    <h1>$ReportTitle</h1>
    <div class="note">
        For more information please read
        <a href="https://sonosinc.atlassian.net/wiki/spaces/PEOPLEPROD/pages/1346830354/Drive+Copy+Automation">
            Drive Copy Automation
        </a>
        <a href=https://docs.google.com/spreadsheets/d/1_0WIDlZriHpb1_YqP1nPGs0lcOhliweGt-klJSPg2C4">
            Drive Copy Report.
        </a>
    </div>
    <div class="table-container">$HtmlTable</div>
</body>
</html>
"@
    }

    try {
        # Capture the current time once for consistency across the script
        $now = Get-Date

        # Initialize output containers
        $attachments = @()
        $htmlBody = ''

        # Generate a consistent subject/title
        $finalTitle = Get-FormattedTitle -Now $now -BaseTitle $Title -FallbackPath $ReportPath[0]

        # If no report is provided or NoReport flag is set,send a simple completion email
        if ($NoReport -or -not $ReportPath) {
            $htmlBody = BuildEmailBody -ReportTitle $finalTitle -DateStamp $now

            $gamCommand = @(
                'user', 'google.admin@sonos.com',
                'sendemail', 'to', $EmailTo,
                'subject', $finalTitle,
                'textmessage', $htmlBody,
                'html'
            )

            $gamexe = 'gam'

            & $gamexe @gamCommand | Out-Null
            return
        }

        # Process each CSV report path
        foreach ($path in $ReportPath) {
            # Validate the file path exists
            if (-not (Test-Path $path -PathType Leaf)) {
                throw "Invalid file path: $path"
            }

            # Import CSV content
            $csv = Import-Csv -Path $path

            # Verify data exists
            if (-not $csv -or $csv.Count -eq 0) {
                throw "No data in file: $path"
            }

            # Convert CSV to HTML table
            $props = $csv[0].PSObject.Properties.Name
            $htmlTable = $csv | ConvertTo-Html -Property $props -As Table -Fragment

            # Build the body (cumulative if multiple tables)
            $htmlBody += BuildEmailBody -ReportTitle $finalTitle -DateStamp $now -HtmlTable $htmlTable

            # Add attachment
            $attachments += $path
        }

        $gamCommand = @(
            'user', 'google.admin@sonos.com',
            'sendemail', 'to', $EmailTo,
            'subject', $finalTitle,
            'textmessage', $htmlBody,
            'html'
        )

        foreach ($Attachment in $Attachments) {
            $gamCommand += @('attach', $Attachment)
        }

        $gamexe = 'gam'

        & $gamexe @gamCommand | Out-Null
    }
    catch {
        # Log and rethrow any errors
        Write-Error $_.Exception.Message
        throw
    }
}