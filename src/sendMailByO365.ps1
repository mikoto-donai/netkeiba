$o365User = "";
$o365Pass = ""
$from    = "";
$to      = "";
$mailServer = "";
$port =""
$subject = "";
$body    = "";

$msg = New-Object System.Net.Mail.MailMessage($from, $to, $subject, $body);

$SMTPClient = New-Object Net.Mail.SmtpClient($mailServer, $port) ;
$SMTPClient.EnableSsl = $true;
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential($o365User, $o365Pass);
$SMTPClient.send($msg);