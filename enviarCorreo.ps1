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

            # Remove quotes if present
            if ($value -match '^"(.+)"$' -or $value -match "^'(.+)'$") {
                $value = $Matches[1]
            }

            # Set as a script variable
            Set-Variable -Name $key -Value $value -Scope Script
        }
    }
}

# Load variables from .env
Load-DotEnv -Path "D:\CorreosPowershell\.env"

# --- Configuración de correo desde .env ---
$MailSubject = $OUTLOOK_SUBJECT
$MailHTMLBody = $OUTLOOK_BODY
$MailTo = $OUTLOOK_TO

for ($i = 1; $i -le 1; $i++) {
    # Crea un objeto COM para Outlook
    $Outlook = New-Object -ComObject Outlook.Application
    # Crea un objeto de espacio de nombres para acceder a las carpetas de Outlook
    $Namespace = $Outlook.GetNamespace('MAPI')
    # Inicia sesión en el espacio de nombres de Outlook
    $Namespace.Logon()
    # Crea un nuevo correo electrónico
    $Mail = $Outlook.CreateItem(0)
    $Mail.Subject = $MailSubject
    # El cuerpo del correo contiene código HTML que muestra la imagen
    $Mail.HTMLBody = $MailHTMLBody    
    # Agrega la imagen al cuerpo del correo
    $Mail.To = $MailTo
    ## Establece la dirección de correo electrónico del remitente
    $Mail.Send()
}