param(
    $sAMAccount= "",
    $keytabPath = "",
    $password = $null
)
function GetPublicCert{
    $rootDSE = [ADSI]("LDAP://RootDSE") 
    $searcher = New-Object DirectoryServices.DirectorySearcher 
    $searcher.SearchRoot = "GC://" + $RootDSE.rootDomainNamingContext 
    $searcher.SearchScope = "subtree" 
    $searcher.PropertiesToLoad.Add("distinguishedname") | Out-Null 
    $searcher.PropertiesToLoad.Add("mail") | Out-Null 
    $searcher.PropertiesToLoad.Add("usercertificate") | Out-Null 
    $searcher.Filter = ("(&(objectClass=person)(CN=$sAMAccount))") 
    $recipient = $searcher.FindOne()
    $chosenCertificate = $null 
    $now = Get-Date 
    If ($recipient.Properties.usercertificate -ne $null) { 
        ForEach ($userCertificate in $recipient.Properties.usercertificate) { 
            $validSecureEmail = $false 
            $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]$userCertificate 
            $extensions = $certificate.Extensions 
            ForEach ($extension in $extensions) { 
                If ($extension.EnhancedKeyUsages -ne $null) { 
                    ForEach ($enhancedKeyUsage in $extension.EnhancedKeyUsages) { 
                        If ($enhancedKeyUsage.FriendlyName -eq "Secure Email") { 
                            $validSecureEmail = $true 
                            break 
                        } 
                    } 
                    If ($validSecureEmail) { 
                        break 
                    } 
                } 
            } 
            If ($validSecureEmail) { 
                If ($now -gt $certificate.NotBefore.AddMinutes(-5) -and $now -lt $certificate.NotAfter.AddMinutes(5)) { 
                    $chosenCertificate = $certificate 
                } 
            } 
            If ($chosenCertificate -ne $null) { 
                break 
            } 
        } 
    }
    return $chosenCertificate
}

Function SendEmail{

    $keytab = Get-item -Path $keytabPath
    $Cert = $(GetPublicCert)
    $mailServer = "ism.contoso.se"
    $senderEmail = "noreply@contoso.se" 

    Add-Type -assemblyName "System.Security" 
    $mailClient = New-Object System.Net.Mail.SmtpClient $mailServer
    $message = New-Object System.Net.Mail.MailMessage
    $message.To.Add($recipient.properties.mail.item(0)) 
    $message.From = $senderEmail
    $message.Subject = "Test Unencrypted subject of the message"

    $body = "This is the password for your service account: p4ssw09rd23!"

    $MIMEMessage = New-Object system.Text.StringBuilder 
    $MIMEMessage.AppendLine("MIME-Version: 1.0") | Out-Null 
    $MIMEMessage.AppendLine("Content-Type: multipart/mixed; boundary=unique-boundary-1") | Out-Null 
    $MIMEMessage.AppendLine() | Out-Null
    $MIMEMessage.AppendLine("This is a multi-part message in MIME format.") | Out-Null
    $MIMEMessage.AppendLine("--unique-boundary-1") | Out-Null
    $MIMEMessage.AppendLine("Content-Type: text/plain") | Out-Null
    $MIMEMessage.AppendLine("Content-Transfer-Encoding: 7Bit") | Out-Null
    $MIMEMessage.AppendLine()|Out-Null
    $MIMEMessage.AppendLine($body) | Out-Null
    $MIMEMessage.AppendLine() | Out-Null

    $MIMEMessage.AppendLine("--unique-boundary-1") | Out-Null
    $MIMEMessage.AppendLine("Content-Type: application/octet-stream; name="+ $keytab.Name) | Out-Null
    $MIMEMessage.AppendLine("Content-Transfer-Encoding: base64") | Out-Null
    $MIMEMessage.AppendLine("Content-Disposition: attachment; filename="+ $keytab.Name) | Out-Null
    $MIMEMessage.AppendLine() | Out-Null

    [Byte[]] $binaryData = [System.IO.File]::ReadAllBytes($keytab)
    [string] $base64Value = [System.Convert]::ToBase64String($binaryData, 0, $binaryData.Length)
    [int] $position = 0
    while($position -lt $base64Value.Length)
    {
        [int] $chunkSize = 100
        if (($base64Value.Length - ($position + $chunkSize)) -lt 0)
        {
            $chunkSize = $base64Value.Length - $position
        }
    $MIMEMessage.AppendLine($base64Value.Substring($position, $chunkSize))|Out-Null
    $MIMEMessage.AppendLine()|Out-Null
    $position += $chunkSize;
    }
    $MIMEMessage.AppendLine("--unique-boundary-1--") | Out-Null

    [Byte[]] $bodyBytes = [System.Text.Encoding]::ASCII.GetBytes($MIMEMessage.ToString())
    $contentInfo = New-Object System.Security.Cryptography.Pkcs.ContentInfo (,$bodyBytes) 
    $CMSRecipient = New-Object System.Security.Cryptography.Pkcs.CmsRecipient $Cert 
    $envelopedCMS = New-Object System.Security.Cryptography.Pkcs.EnvelopedCms $contentInfo 
    $envelopedCMS.Encrypt($CMSRecipient)

    [Byte[]] $encryptedBytes = $envelopedCMS.Encode() 
    $memoryStream = New-Object System.IO.MemoryStream @(,$encryptedBytes) 
    $alternateView = New-Object System.Net.Mail.AlternateView($memoryStream, "application/pkcs7-mime; smime-type=enveloped-data;name=smime.p7m") 
    $message.AlternateViews.Add($alternateView)
    $MailClient.Send($Message)
}
SendEmail
