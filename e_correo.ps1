# Function to load .env file variables
function Load-DotEnv {
    param (
        [string]$Path = ".env"
    )

    if (-not (Test-Path $Path)) {
        Write-Warning "'.env' file not found at $Path. Skipping environment variable loading."
        return
    }

    Get-Content $Path | ForEach-Object {
        $line = $_.Trim()
        if ($line -notmatch '^\s*#' -and $line -match '^([A-Za-z_][A-Za-z0-9_]*)=(.*)$') {
            $key = $Matches[1]
            $value = $Matches[2]

            # Remove quotes if present (simplified logic)
            if (($value.StartsWith('"') -and $value.EndsWith('"')) -or ($value.StartsWith("'") -and $value.EndsWith("'"))) {
                $value = $value.Substring(1, $value.Length - 2)
            }

            # Set as a script variable
            Set-Variable -Name $key -Value $value -Scope Script
        }
    }
}

# Load variables from .env
Load-DotEnv -Path "D:\CorreosPowershell\.env"

# --- Configuración de las credenciales del servidor Exchange ---
# Las variables se cargan desde el .env y se acceden directamente por su nombre
$smtpServer = $SMTP_SERVER
$smtpPort = [int]$SMTP_PORT # Convert to integer
$smtpFrom = $FROM_ADDRESS
$smtpTo = $TO_ADDRESS
$loginUser = $LOGIN_USER
$smtpPassword = $SMTP_PASSWORD

$messageSubject = "Prueba de correo PS1 con .env"
$messageBody = "Esto es un envio para probar el servidor desde el exterior usando variables de .env."

# --- Configuración del objeto de mensaje ---
$mailmessage = New-Object system.net.mail.mailmessage
$mailmessage.from = ($smtpFrom)
$mailmessage.To.add($smtpTo)
$mailmessage.Subject = $messageSubject
$mailmessage.Body = $messageBody

# --- Configuración del cliente SMTP ---
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
$smtp = New-Object Net.Mail.SmtpClient($smtpServer)
$smtp.Port = $smtpPort
$smtp.EnableSsl = $true # Habilitar SSL/TLS para el puerto 587

# --- Autenticación ---
# La contraseña se obtiene del .env
$smtp.Credentials = New-Object System.Net.NetworkCredential($loginUser, $smtpPassword)

# --- Envío del correo ---
try {
    $smtp.Send($mailmessage)
    Write-Host "Correo enviado exitosamente a $smtpTo"
}
catch {
    Write-Error "Error al enviar el correo: $($_.Exception.Message)"
}