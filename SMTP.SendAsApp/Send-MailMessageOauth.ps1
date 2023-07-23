function Send-MailMessageOAuth {
    [CmdletBinding()]
      param (        
        [string]$ServerName="smtp.office365.com",
        $token=$AccessToken,
        [Parameter(Mandatory = $true)]
        [String]$userName,
        [Parameter(Mandatory = $false)]
        [String]$SendingAddress=$UserName,
        [Parameter(Mandatory = $true)]        
        [String]$To,
        [Parameter(Mandatory = $true)]
        [String]$Subject,
        [Parameter(Mandatory = $true)]
        [String]$Body,
        [Parameter(Mandatory = $false)]
        [String]$AttachmentFileName,
        [int]$Port = 587
    )
    Process {       

        $mailMessage = New-Object System.Net.Mail.MailMessage
        $mailMessage.From = New-Object System.Net.Mail.MailAddress($SendingAddress)
        $mailMessage.To.Add($To)
        $mailMessage.Subject = $Subject
        $mailMessage.Body = $Body     
        if(![String]::IsNullOrEmpty($AttachmentFileName)){
            $attachment = New-Object System.Net.Mail.Attachment -ArgumentList $AttachmentFileName
            $mailMessage.Attachments.Add($attachment);
        }        

        $binding = [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic
        $MessageType = $mailMessage.GetType()
        $smtpClient = New-Object System.Net.Mail.SmtpClient
        $scType = $smtpClient.GetType()
        $booleanType = [System.Type]::GetType("System.Boolean")
        $assembly = $scType.Assembly
        $mailWriterType = $assembly.GetType("System.Net.Mail.MailWriter")
        $MemoryStream = New-Object -TypeName "System.IO.MemoryStream"
        $typeArray = ([System.Type]::GetType("System.IO.Stream"))
        $mailWriterConstructor = $mailWriterType.GetConstructor($binding ,$null, $typeArray, $null)
        [System.Array]$paramArray = ($MemoryStream)
        $mailWriter = $mailWriterConstructor.Invoke($paramArray)
        $doubleBool = $true
        $typeArray = ($mailWriter.GetType(),$booleanType,$booleanType)
        $sendMethod = $MessageType.GetMethod("Send", $binding, $null, $typeArray, $null)
        if ($null -eq $sendMethod) {
            $doubleBool = $false
            [System.Array]$typeArray = ($mailWriter.GetType(),$booleanType)
            $sendMethod = $MessageType.GetMethod("Send", $binding, $null, $typeArray, $null)
         }
        [System.Array]$typeArray = @()
        $closeMethod = $mailWriterType.GetMethod("Close", $binding, $null, $typeArray, $null)
        [System.Array]$sendParams = ($mailWriter,$true)
        if ($doubleBool) {
            [System.Array]$sendParams = ($mailWriter,$true,$true)
        }
        $sendMethod.Invoke($mailMessage,$binding,$null,$sendParams,$null)
        [System.Array]$closeParams = @()
        $MessageString = [System.Text.Encoding]::UTF8.GetString($MemoryStream.ToArray());
        $closeMethod.Invoke($mailWriter,$binding,$null,$closeParams,$null)
        [Void]$MemoryStream.Dispose()
        [Void]$mailMessage.Dispose()
        $MessageString = $MessageString.SubString($MessageString.IndexOf("MIME-Version:"))
        $socket = new-object System.Net.Sockets.TcpClient($ServerName, $Port)
        $stream = $socket.GetStream()
        $streamWriter = new-object System.IO.StreamWriter($stream)
        $streamReader = new-object System.IO.StreamReader($stream)
        $streamWriter.AutoFlush = $true
        $sslStream = New-Object System.Net.Security.SslStream($stream)
        $sslStream.ReadTimeout = 30000
        $sslStream.WriteTimeout = 30000        
        $ConnectResponse = $streamReader.ReadLine();
        Write-Host($ConnectResponse)
        if(!$ConnectResponse.StartsWith("220")){
            throw "Error connecting to the SMTP Server"
        }
        $Domain = $SendingAddress.Split('@')[1]
        Write-Host(("helo " + $Domain)) -ForegroundColor Green
        $streamWriter.WriteLine(("helo " + $Domain));
        $ehloResponse = $streamReader.ReadLine();
        Write-Host($ehloResponse)
        if (!$ehloResponse.StartsWith("250")){
            throw "Error in ehelo Response"
        }
        Write-Host("STARTTLS") -ForegroundColor Green
        $streamWriter.WriteLine("STARTTLS");
        $startTLSResponse = $streamReader.ReadLine();
        Write-Host($startTLSResponse)
        $ccCol = New-Object System.Security.Cryptography.X509Certificates.X509CertificateCollection
        $sslStream.AuthenticateAsClient($ServerName,$ccCol,[System.Security.Authentication.SslProtocols]::Tls12,$false);        
        $SSLstreamReader = new-object System.IO.StreamReader($sslStream)
        $SSLstreamWriter = new-object System.IO.StreamWriter($sslStream)
        $SSLstreamWriter.AutoFlush = $true
        $SSLstreamWriter.WriteLine(("helo " + $Domain));
        $ehloResponse = $SSLstreamReader.ReadLine();
        Write-Host($ehloResponse)
        $command = "AUTH XOAUTH2" 
        write-host -foregroundcolor DarkGreen $command
        $SSLstreamWriter.WriteLine($command) 
        $AuthLoginResponse = $SSLstreamReader.ReadLine()
        write-host ($AuthLoginResponse)
        $SALSHeaderBytes = [System.Text.Encoding]::ASCII.GetBytes(("user=" + $userName + [char]1 + "auth=Bearer " + $token + [char]1 + [char]1))
        $Base64AuthSALS = [Convert]::ToBase64String($SALSHeaderBytes)     
        write-host -foregroundcolor DarkGreen $Base64AuthSALS
        $SSLstreamWriter.WriteLine($Base64AuthSALS)        
        $AuthResponse = $SSLstreamReader.ReadLine()
        write-host $AuthResponse
        if($AuthResponse.StartsWith("235")){
            $command = "MAIL FROM: <" + $SendingAddress + ">" 
            write-host -foregroundcolor DarkGreen $command
            $SSLstreamWriter.WriteLine($command) 
            $FromResponse = $SSLstreamReader.ReadLine()
            write-host $FromResponse
            $command = "RCPT TO: <" + $To + ">" 
            write-host -foregroundcolor DarkGreen $command
            $SSLstreamWriter.WriteLine($command) 
            $ToResponse = $SSLstreamReader.ReadLine()
            write-host $ToResponse
            $command = "DATA"
            write-host -foregroundcolor DarkGreen $command
            $SSLstreamWriter.WriteLine($command) 
            $DataResponse = $SSLstreamReader.ReadLine()
            write-host $DataResponse
            write-host -foregroundcolor DarkGreen $MessageString
            $SSLstreamWriter.WriteLine($MessageString) 
            $SSLstreamWriter.WriteLine(".") 
            $DataResponse = $SSLstreamReader.ReadLine()
            write-host $DataResponse
            $command = "QUIT" 
            write-host -foregroundcolor DarkGreen $command
            $SSLstreamWriter.WriteLine($command) 
            # ## Close the streams 
            $stream.Close() 
            $sslStream.Close()
            Write-Host("Done")
        }  

    }


}
