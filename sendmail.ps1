$mail = @{
    from = "tmichiuc@gmail.com";
    to = "tmichiuc@gmail.com";
    attachment = "C:\Users\tmichiuc\Desktop\netkeiba\results.pdf";
    smtp_server = "smtp.gmail.com";
    smtp_port = 587;
    user = "tmichiuc";
    password = "hgldrjieectinqmg";
}

$password = ConvertTo-SecureString $mail["password"] -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential $mail["user"], $password

Send-MailMessage -To $mail["to"] `
                 -From $mail["from"] `
                 -SmtpServer $mail["smtp_server"] `
                 -Credential $credential `
                 -Attachment $mail["attachment"] `
                 -Subject "netkeiba�f�[�^���H�e�X�g" `
                 -Body "test" `
                 -Encoding ([System.Text.Encoding]::UTF8) `
                 -UseSsl