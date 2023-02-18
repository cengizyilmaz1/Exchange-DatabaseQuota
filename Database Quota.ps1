<#	
#################################################################################################################
# Yayınlanma Tarihi: 18.02.2023
# Cengiz YILMAZ
# MCT
# https://cozumpark.com/author/cengizyilmaz
# https://cengizyilmaz.net
# cengiz@cengizyilmaz.net
#
##################################################################################################################
#>
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
 
# Set the database, threshold, from, server, port, to, subject and body variables
$Database = "DB01"
$Threshold = "5"
$From = "domain@cengizyilmaz.net"
$Server = "mail.cengizyilmaz.net
$Port = 587
$To = "cengiz@cengizyilmaz.net"
$Subject = "Journal Mailbox Kota Uyarısı!"
$Body = "<html><body><h4>$database içerisinde bulunan hesapların kota durumu %$threshold değerinden fazladır.</h4><table border='3'><tr><th>Display Name</th><th>Email Address</th><th>Quota Usage</th></tr>"


$DirPathCheck = Test-Path -Path $DirPath
If (!($DirPathCheck))
{
        Try
        {
                 
                 New-Item -ItemType Directory $DirPath -Force
        }
        Catch
        {
                 $_ | Out-File ($DirPath + "\" + "Log.txt") -Append
        }
}
#CredObj path
$CredObj = ($DirPath + "\" + "Quota.cred")
#Check if CredObj is present
$CredObjCheck = Test-Path -Path $CredObj
If (!($CredObjCheck))
{
        "$Date - INFO: creating cred object" | Out-File ($DirPath + "\" + "Log.txt") -Append
        #SMTP Info
        $Credential = Get-Credential -Message "Please enter your Mail Server credential that you will use to send e-mail $fromAddress. If you are not using the account $fromAddress make sure this account has 'Send As' rights on $FromEmail."
        #Export cred obj
        $Credential | Export-CliXml -Path $CredObj
}
 
Write-Host "Importing Cred object..." -ForegroundColor Yellow
$Cred = (Import-CliXml -Path $CredObj)

 
# Get all mailbox users in the specified database
$mailboxUsers = Get-Mailbox -Database $Database -ResultSize unlimited | Where-Object {$_.RecipientTypeDetails -eq 'UserMailbox'}
 
# Loop through each mailbox user
foreach ($user in $mailboxUsers) {
    # Get the mailbox statistics
    $mailbox = Get-MailboxStatistics -Identity $user.DistinguishedName
 
    # Calculate the quota usage percentage
    $quotaUsage = ($mailbox.TotalItemSize.Value.ToMB() / $user.ProhibitSendQuota.Value.ToMB()) * 100
 
    # Check if the quota usage exceeds the threshold
    if ($quotaUsage -ge $Threshold) {
        # Create the HTML table row with the user's display name, email address, and quota usage in red text
        $Body += "<tr><td>$($user.DisplayName)</td><td>$($user.PrimarySmtpAddress)</td><td><font color='red'>$($quotaUsage.ToString("0.00"))%</font></td></tr>"
 
# Create the mail message object
$message = New-Object System.Net.Mail.MailMessage
$message.From = $From
$message.To.Add($To)
$message.Subject = $subject
$message.Body = $body
$message.Priority = "High"
$message.IsBodyHtml = $true
    }
}
 
# Close the HTML table and body
$Body += "</table></body></html>"
 

 
# Create the SMTP client object
$client = New-Object System.Net.Mail.SmtpClient
$client.Host = $smtpServer
$client.Port = $smtpPort
$client.Credentials = $credential

 
# Send the email
$client.Send($message)
 
# Dispose of the message object
$message.Dispose()
