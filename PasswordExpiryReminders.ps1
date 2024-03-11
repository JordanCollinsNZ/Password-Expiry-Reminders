# ------------------------------------{ Logging }------------------------------------

# Start logging
Start-Transcript -Path "C:\Logs\ExpiringPasswordsNotification-$ENV:COMPUTERNAME-$(Get-Date -f 'yyyyMMdd-HHmmss').txt" -NoClobber

# Delete log files older than 30 days
Get-ChildItem -Path "C:\Logs" -Recurse | Where-Object {($_.LastWriteTime -lt (Get-Date).AddDays(-30))} | Remove-Item

# ----------------------------------{ Description }----------------------------------

# This script is designed to send upcoming password expiry reminders to users from specific OU's in a domain.
# It will also send a 'catch-all' email to an address with users who have upcoming expiries, but no emails set in AD.

# ------------------------------------{ History }------------------------------------

# Version 1.0 - 11/03/24 - Jordan Collins -  Initial creation

# ----------------------------------{ Requirements }---------------------------------

#Requires -Modules ActiveDirectory

# -----------------------------------{ Variables }-----------------------------------

# How many days out from password expiry to email users
$WarningDays = 20, 10, 5, 4, 3, 2, 1

# Domain name for AD environment
$ADDomainName = "jordancollins.nz"

# OU Path to look for users, accepts multiple OUs
$OUs =
'CN=Users,DC=jordancollins,DC=nz',
'CN=Admins,DC=jordancollins,DC=nz'

# Address to send emails from
$MailSender = "Service Desk <Service.Desk@jordancollins.nz>"

# Hostname/IP address for SMTP server
$SMTPServer = "localhost"

# Email to send accounts to that are expiring passwords but do not have an email in AD
$CatchAllMail = "Service Desk <Service.Desk@jordancollins.nz>"

# Email body in HTML format to send to end users. Current variables are as below:
## VarFirstName - Users GivenName in AD
## VarDays - Number of days until users password expires
## VarPlural - This will use 'day' if expiry date is 1 day in future, and 'days' if not
## VarDate - Date of Password expiry
$EmailBody = @"
Kia Ora VarFirstName,<br>
<br>
Your password will expire in <b>VarDays VarPlural</b> (VarDate). Please change your password before it expires to avoid any disruption in accessing your account.<br>
<br>
Thank you,<br>
Service Desk Team<br>
"@

# Email body in HTML format to send to CatchAllMail address (Note, the closing tags are added just before the email in the final if statement)
$CatchAllEmailBody = @"
<style type="text/css">
.tg  {border-collapse:collapse;border-spacing:0;}
.tg td{border-color:black;border-style:solid;border-width:1px;overflow:hidden;padding:10px 5px;word-break:normal;}
.tg th{border-color:black;border-style:solid;border-width:1px;;overflow:hidden;padding:10px 5px;word-break:normal;}
.tg .tg-sxqf{border-color:inherit;font-family:inherit;text-align:left;vertical-align:top}
.tg .tg-0pky{border-color:inherit;text-align:left;vertical-align:top}
</style>
Kia Ora Service Desk Team,<br>
<br>
There are currently users with upcoming expiring passwords but no email address set in Active Directory to send a notification to.<br>
<br>
<table class="tg">
<thead>
  <tr>
    <th class="tg-sxqf">Name</th>
    <th class="tg-0pky">UPN</th>
    <th class="tg-0pky">Days</th>
    <th class="tg-0pky">Expiry Date</th>
  </tr>
</thead>
<tbody>
<tbody>
"@

# -----------------------------------{ Function }------------------------------------

# Import required modules
Import-Module ActiveDirectory
 
# Define AD server based off domain, always connect to the Infrastructure Master
# Not really required but good when environment has multi domains ¯\_(ツ)_/¯ 
try {
    $ADServer = (Get-ADDomain -Identity "$ADDomainName").InfrastructureMaster
}
catch {
    throw "Cannot get infrastructure master for $ADDomainName."
}
 
# Create warning dates for future warning days
$WarningDates = @()
$ExpiringUsersCount = $null
foreach ($WarningDay in $WarningDays) {
    $WarningDates += (Get-Date).AddDays($WarningDay).ToLongDateString()
}

# Find accounts that are enabled and have expiring passwords in each OU
$Users = foreach ($OU in $OUs) {
  Get-ADUser -Server $ADServer -SearchBase $OU -Filter {
    Enabled -eq $True -and PasswordNeverExpires -eq $False -and PasswordLastSet -gt 0
  }`
  -Properties "givenName", "emailAddress", "msDS-UserPasswordExpiryTimeComputed", "userPrincipalName", "displayName" |
  Select-Object -Property "userPrincipalName", "DisplayName", "givenName", "emailAddress", @{
    Name = "PasswordExpiry"; Expression = {[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed").ToLongDateString()}
  }
}
 
# Check password expiration date and send email on match
foreach ($User in $Users) {
    if ($($User.PasswordExpiry) -in $WarningDates) {
        # Calculate how many days in future again from expiry date, add 1 as hours etc are included in New-TimeSpan
        $PasswordExpiryDays = (New-TimeSpan -End "$($User.PasswordExpiry)").Days
        $PasswordExpiryDays++
        # Determine if should use 'day' or 'days' based on number
        if ($PasswordExpiryDays -eq "1") {$Plural = "Day"} else {$Plural = "Days"}
        # If user has email in AD
        if ($null -ne $($User.EmailAddress)) {

            # Copy email body template
            $NewEmailBody = $EmailBody

            # Replace email variables with user details
            $NewEmailBody = $NewEmailBody.Replace("VarFirstName", $($User.GivenName))
            $NewEmailBody = $NewEmailBody.Replace("VarDays", $PasswordExpiryDays)
            $NewEmailBody = $NewEmailBody.Replace("VarPlural", $Plural)
            $NewEmailBody = $NewEmailBody.Replace("VarDate", $($User.PasswordExpiry))

            # Set subject
            $Subject = "Password Expiry Notification - Your Password Will Expire in $PasswordExpiryDays $Plural"

            # Send email
            Write-Information "Sending email to $($User.DisplayName) - $($User.EmailAddress) - Expiring in $PasswordExpiryDays $Plural."
            Send-MailMessage `
            -To "$($User.DisplayName) <$($User.EmailAddress)>" `
            -From $MailSender `
            -Bcc $MailSender `
            -SmtpServer $SMTPServer `
            -Subject $Subject `
            -Body $NewEmailBody `
            -BodyAsHtml
        }

        #User has no email in AD
        else {
            $ExpiringUsersCount++
            $CatchAllEmailBody += @"
<tr>
  <td class="tg-0pky">$($User.DisplayName)<br></td>
  <td class="tg-0pky">$($User.UserPrincipalName)</td>
  <td class="tg-0pky">$PasswordExpiryDays</td>
  <td class="tg-0pky">$($User.PasswordExpiry)</td>
  </tr>
"@
            Write-Warning "UPN $($User.UserPrincipalName) does not have an email set - Expiring in $PasswordExpiryDays $Plural."
        }
    }
}

# If there are users with expiring passwords and no email set, close the HTML table and send email
if ($null -ne $ExpiringUsersCount) {
    $CatchAllEmailBody += "</tbody></table><br>This is an automated email."
    Write-Information "Sending catch all email to $CatchAllMail for $ExpiringUsersCount user(s)."
    Send-MailMessage `
    -To $CatchAllMail `
    -From $MailSender `
    -Bcc $MailSender `
    -SmtpServer $SMTPServer `
    -Subject "Users with expiring passwords but no email" `
    -Body $CatchAllEmailBody `
    -BodyAsHtml
}

# Stop logging
Stop-Transcript
