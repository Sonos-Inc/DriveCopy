<#
.SYNOPSIS
    Sends an email to a manager notifying them about the successful transfer of Google MyDrive data.

.DESCRIPTION
    This function sends an HTML-formatted email to the specified manager,notifying them about the transfer of Google MyDrive data for a user. 
    It includes details such as the user's email,the creation of a new folder in the manager's Google MyDrive,and provides a support ticket link for assistance.

.PARAMETER Manager
    The email address of the manager. This parameter is mandatory and must be a valid email address.

.PARAMETER UserEmail
    The email address of the user whose Google MyDrive data has been transferred. This parameter is mandatory and must be a valid email address.

.PARAMETER WebURL
    The URL where the transferred data can be accessed. This is important and will be prominently displayed in the email.

.NOTES
https://sonosinc.atlassian.net/wiki/spaces/PEOPLEPROD/pages/1346830354/Drive+Copy+Automation Architectural
https://sonosinc.atlassian.net/wiki/spaces/ITKB/pages/1348370617/Legacy+Drive+Backup ITKB
https://docs.google.com/spreadsheets/d/1_0WIDlZriHpb1_YqP1nPGs0lcOhliweGt-klJSPg2C4/edit?gid=1966817793#gid=1966817793 Spreadsheet
https://drive.google.com/drive/folders/0ADVPlcaDSgiTUk9PVA Legacy Drive Backup Folder
#>
function SendManagerEmail {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        [ValidatePattern("^[a-zA-Z0-9](?:[a-zA-Z0-9._%+-]*[a-zA-Z0-9])?(?:@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})?$")]
        [string]$Manager,

        [parameter(Mandatory = $true)]
        [ValidatePattern("^[a-zA-Z0-9](?:[a-zA-Z0-9._%+-]*[a-zA-Z0-9])?(?:@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})?$")]
        [string]$UserEmail,

        [parameter(Mandatory = $true)]
        [string]$WebURL
    )

    try {
        # Get the current date
        $TodayDate = (Get-Date).ToString("yyyy-MM-dd")

        # Extract manager's first name or use the full email as a fallback
        $ManagerFirstName = ($Manager -split '\.')[0] -replace '@.*', ''
        $ManagerFirstName = (Get-Culture).TextInfo.ToTitleCase($ManagerFirstName.Trim())

        # Define backup folder name
        $FolderName = "Backup_From_$UserEmail"

        # Construct email subject
        $Subject = "$TodayDate Transfer of Google MyDrive for $UserEmail"

        # Construct formatted HTML body
        $FormattedBody = @"
<html>
<head>
<style>
body {
    font-family: Arial,sans-serif;
    font-size: 14px;
    color: #333;
}
.container {
    padding: 20px;
    border: 1px solid #ccc;
    border-radius: 5px;
    background-color: #f9f9f9;
}
.signature {
    margin-top: 20px;
    font-style: italic;
    color: #777;
}
a {
    color: #d32f2f;
    font-weight: bold;
    text-decoration: none;
}
a:hover {
    text-decoration: underline;
}
</style>
</head>
<body>
<div class="container">
    <p>Hello $ManagerFirstName,</p>

    <p>You're receiving this message because our records indicate that you are the manager of a former Sonos employee or contractor whose email address was <strong>$UserEmail</strong>.</p>

    <p>As part of Sonos's Google Workspace data lifecycle policy,inactive accounts are archived for up to one year before they are permanently deleted. Before account deletion occurs,any remaining Google Drive files are copied to a centralized Team Shared Drive to preserve business-relevant content.</p>

    <p>A folder named <b>$FolderName</b> has been (or will soon be) created within the Team Shared Drive <strong>Legacydrivebackup</strong>. This folder contains all residual Google Drive content previously owned by <strong>$UserEmail</strong>.</p>

    <p><strong>Please Note:</strong> For detailed information on the drive transfer process,<a href="https://sonosinc.atlassian.net/wiki/spaces/ITKB/pages/1348370617/Legacy+Drive+Backup">refer to this internal knowledge base article</a>.</p>

    <p>If you have any questions or require additional support,please raise a ticket in Jira using this link: <a href="https://jira.sonos.com/plugins/servlet/desk/portal/27">Submit a Support Ticket</a>. Be sure to include relevant details.</p>

    <p><strong>The folder containing <strong>$UserEmail</strong>'s transferred Drive data can be accessed here:</strong></p>
    <ul>
    <li><b><a href='$WebURL'>$WebURL</a></b></li>
    </ul>

    <p class="signature">Sonos Google Admin</p>
</div>
</body>
</html>
"@

        $gamCommand = @(
            'user', 'google.admin@sonos.com',
            'sendemail', 'to', $Manager,
            'subject', $Subject,
            'textmessage', $FormattedBody,
            'html'
        )

        $gamexe = 'gam'

        & $gamexe @gamCommand | Out-Null

        if ($LASTEXITCODE -ne 0) {
            return $false
        }
        return $true
    }
    catch {
        Write-Error "Failed to send manager email. $_"
        return $false
    }
}