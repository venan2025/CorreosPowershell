# Crea una instancia de la aplicación de Outlook
Write-Host "Creando instancia de Outlook..."
$Outlook = New-Object -ComObject Outlook.Application

# Accede a la bandeja de entrada
Write-Host "Accediendo a la bandeja de entrada..."
$Namespace = $Outlook.GetNamespace("MAPI")
$Inbox = $Namespace.GetDefaultFolder(6) # 6 corresponde a la carpeta de Bandeja de entrada

# Enumera los correos electrónicos en la bandeja de entrada
Write-Host "Enumerando correos electrónicos en la bandeja de entrada..."
$Emails = $Inbox.Items

# Variable para verificar si se encontró un correo con el asunto "prueba correo"
$EncontradoCorreo = $false

# Itera a través de los correos electrónicos
Write-Host "Buscando correo con asunto 'prueba'..."
$Hora = Get-Date
Add-Content -Path "resultado.txt" -Value " Realizado a: $Hora"
foreach ($Email in $Emails) {
    if ($Email.Subject -eq "tu futuro") {
        $EncontradoCorreo = $true
        Write-Host "Asunto: $($Email.Subject)"
        Write-Host "De: $($Email.SenderName)"
        Write-Host "Fecha: $($Email.ReceivedTime)"
        Write-Host "------------------------------------"
        
	# Agregar aquí el código para reproducir un sonido
        [System.Media.SystemSounds]::Hand.Play()
	
	# Escribe el resultado en un archivo .txt
        $Resultado = "$($Email.Subject) - $($Email.SenderName) - $($Email.ReceivedTime)"
        $Hora = Get-Date
        Add-Content -Path "resultado.txt" -Value "$Resultado - $Hora"

        
        
    }
}

# Libera recursos
Write-Host "Liberando recursos..."
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Emails) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Inbox) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Namespace) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null

if ($EncontradoCorreo) {
    Write-Host "Se encontró al menos un correo con asunto 'prueba'."
} else {
    Write-Host "No se encontraron correos con asunto 'prueba'."
}

Write-Host "Tarea completada."
