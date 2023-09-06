$dir = ".\" # [PATH]
$emailList = ($dir + "DLEmailList.xlsx")
​
#Gets current days
Invoke-WebRequest -URI "[URL TO CSV FILE FROM MICROSOFT FLOW]" -OutFile $emailList
​
Timeout /T 10
​
$SQLCMD="C:\'Program Files'\'Microsoft SQL Server'\'Client SDK'\ODBC\170\Tools\Binn\SQLCMD.EXE"
$SERVER= "" #[SERVER]
$DB = "" #[CREDENTIALS]
$LOGIN= "" #[username]
$PASSWORD= "" #[password]
​
$CMD = "`"SET nocount on;[ADD COMMAND TO PULL DATA HERE]`""
$OUTNAME = ("`"" + $dir + "CSBInvoice.txt" + "`"")
# -W remomves whitespace, -h-1 removes the solumn headers, -s specifies the delimiter, -w specifies the colun width
Invoke-Expression ($SQLCMD + " -S " + $SERVER + " -d " + $DB + " -U " + $LOGIN + " -P " + $PASSWORD + " -Q " + $CMD + " -o " + $OUTNAME + " -W -h-1 -s`";`" -w 200")
​
Invoke-Expression ("python.exe `"" + $dir + "FindDiscrep.py`"")
​
Timeout /T 15
​
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 
​
# Define the sender, recipient, subject, and body of the email
$From = "email@email.com"
$Subject = "List of Invoices Not Sent"
$Body = "Refer to the text file for all the missed invoices."
$file = ($dir + "MissedInvoicesList.txt")
# Define the SMTP server details
$SMTPServer = "smtp.office365.com"
$SMTPPort = 587
$SMTPUsername = "email@email.com"
$SMTPPassword = "" #PASSWORD
 
# Create a new email object
$Email = New-Object System.Net.Mail.MailMessage
$Email.From = $From
$Email.To.Add("person1@email.com")
$Email.To.Add("person2@email.com")
$Email.Subject = $Subject
$Email.Body = $Body
$Email.Attachments.Add($file)
​
# Uncomment below to send HTML formatted email
#$Email.IsBodyHTML = $true

# Create an SMTP client object and send the email
$SMTPClient = New-Object System.Net.Mail.SmtpClient($SMTPServer, $SMTPPort)
$SMTPClient.EnableSsl = $true

$SMTPClient.UseDefaultCredentials = $false
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential($SMTPUsername, $SMTPPassword)
$SMTPClient.Send($Email)

# Output a message indicating that the email was sent successfully
Write-Host "Email sent successfully to $($Email.To.ToString())"
​
Remove-Item $emailList
$date = Get-Date -Format "dd-MM-yyyy"
Move-Item -Path ($dir+"MissedInvoicesList.txt") -Destination ($dir+"PreviousInvoices\MissedInvoicesList" + $date + ".txt")
exit