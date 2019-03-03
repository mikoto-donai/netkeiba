$mail = @{
    from = "tmichiuc@yahoo-corp.jp";
    to = "tmichiuc@yahoo-corp.jp";
    attachment = "C:\Users\tmichiuc\Desktop\netkeiba\results.pdf";
    smtp_server = "smtp.gmail.com";
    smtp_port = 587;
    user = "tmichiuc";
    password = "uxlbjghnjtqurucn";
}

$password = ConvertTo-SecureString $mail["password"] -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential $mail["user"], $password

Send-MailMessage -To $mail["to"] `
                 -From $mail["from"] `
                 -SmtpServer $mail["smtp_server"] `
                 -Credential $credential `
                 -Attachment $mail["attachment"] `
                 -Subject "ƒeƒXƒg" `
                 -Body "test" `
                 -Encoding ([System.Text.Encoding]::UTF8) `
                 -UseSsl