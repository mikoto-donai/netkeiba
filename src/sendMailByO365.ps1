$o365User = "keiba.keiba@outlook.com";
$o365Pass = "abiek2019"
$from    = "keiba.keiba@outlook.com";
$to      = "michiuchi59@gmail.com";
$mailServer = "smtp.office365.com";
$port ="587"
$subject = "netkeiba";
$body    = "";
$attachment = "README.md"

$msg = New-Object System.Net.Mail.MailMessage($from, $to, $subject, $body);
$msg.Attachments.Add($attachment);

$SMTPClient = New-Object Net.Mail.SmtpClient($mailServer, $port) ;

$SMTPClient.EnableSsl = $true;
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential($o365User, $o365Pass);
$SMTPClient.send($msg);