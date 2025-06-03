# Function settings
$OU = "ou=,ou=,dc=domain,dc=com" # Your ou with fired accounts
$TargetDatabase = "RecycleDataBase" # Your database for mailboxes of dismissed employees
$LogFile = "C:\ExchangeMigration\RecycleUsers_MoveLog.txt" 

# Mail settings for notifications
$emailSettings = @{
    From       = "exchange@domain.com" # Your Exchange mailing address
    To         = "administrator@domain.com" # Notification recipient's mailing address
    SmtpServer = "mail.domain.com" # Your SmtpServer Name
    Port       = 25
}

# Create a log folder if it doesn't exist
if (!(Test-Path -Path (Split-Path $LogFile))) {
    New-Item -ItemType Directory -Path (Split-Path $LogFile) -Force | Out-Null
}

function Write-Log {
    param(
        [string]$Message
    )
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LogFile -Value "$TimeStamp`t$Message"
}

function Send-Notification {
    param(
        [string]$MailboxName
    )
    $subject = "The mailbox has been sent for migration"
    $body = "Mailbox $MailboxName has been queued for database migration $TargetDatabase."
    
    try {
        Send-MailMessage -From $emailSettings.From -To $emailSettings.To -Subject $subject -Body $body -SmtpServer $emailSettings.SmtpServer -Port $emailSettings.Port
        Write-Log "Notification sent to $MailboxName"
    } catch {
        Write-Log ("Error sending notification to {0}: {1}" -f $MailboxName, $_)
    }
}

# Connecting to Exchange
Write-Log "----- MAILBOX MIGRATION HAS STARTED -----"
Write-Log "Connecting to Exchange..."

try {
    if (-not (Get-PSSession | Where-Object { $_.ConfigurationName -like "Microsoft.Exchange" })) {
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://"your exchange dns name or ip"/PowerShell/ -Authentication Kerberos # Connecting to the Exchange console via URi
        Import-PSSession $Session -DisableNameChecking -ErrorAction Stop
        Write-Log "Connection to Exchange successfully established"
    }
} catch {
    Write-Log "Error connecting to Exchange: $_"
    exit 1
}

Write-Log "OU: $OU"
Write-Log "Target Database: $TargetDatabase"

# Getting users in OU
try {
    $Mailboxes = Get-Mailbox -OrganizationalUnit $OU -ResultSize Unlimited
    Write-Log "Migration boxes found: $($Mailboxes.Count)"
} catch {
    Write-Log "Error while receiving mailboxes: $_"
    exit 1
}

foreach ($Mailbox in $Mailboxes) {
    $User = $Mailbox.Identity
    Write-Log "User processing: $User"

    # Перемещение почтового ящика
    try {
        New-MoveRequest -Identity $User -TargetDatabase $TargetDatabase -ErrorAction Stop
        Write-Log "The movement of the box has been started  $TargetDatabase"
        Send-Notification -MailboxName $User
    } catch {
        Write-Log "Error while moving: $_"
        continue
    }

    # Disabling ActiveSync and OWA
    try {
        Set-CASMailbox -Identity $User -ActiveSyncEnabled $false -OWAEnabled $false
        Write-Log "ActiveSync and Outlook Web App disabled"
    } catch {
        Write-Log "Error disabling services: $_"
    }

    # Hiding from GAL
    try {
        Set-Mailbox -Identity $User -HiddenFromAddressListsEnabled $true
        Write-Log "User hidden from address book"
    } catch {
        Write-Log "Error when hiding from GAL: $_"
    }
}

Write-Log "----- MIGRATION COMPLETED -----"

# Clearing an Exchange Session
if ($Session) {
    Remove-PSSession $Session
    Write-Log "Exchange session closed"
}