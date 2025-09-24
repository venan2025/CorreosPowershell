import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import dotenv_values # Importar dotenv_values

# Cargar variables de entorno desde el archivo .env
config = dotenv_values(".env")

# --- Configuración ---
# El usuario para iniciar sesión (el que usas en la web)
login_user = config.get("LOGIN_USER")
# La dirección de correo desde la que quieres enviar
from_address = config.get("FROM_ADDRESS")
# El destinatario
to_address = config.get("TO_ADDRESS")

# Configuración del servidor (IP correcta y puerto STARTTLS)
smtp_server = config.get("SMTP_SERVER")
smtp_port = int(config.get("SMTP_PORT")) # Convertir el puerto a entero
smtp_password = config.get("SMTP_PASSWORD") # Obtener la contraseña del .env

# --- Creación del mensaje ---
msg = MIMEMultipart()
msg['From'] = from_address
msg['To'] = to_address
msg['Subject'] = 'Correo enviado desde servidor Exchange'
body = 'Este correo se envía usando variables de entorno desde el archivo .env.'
msg.attach(MIMEText(body, 'plain', 'utf-8'))

# --- Envío del correo ---
try:
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    
    # Iniciar sesión con el usuario de login
    server.login(login_user, smtp_password)

    # Enviar el correo (desde la dirección de email, al destinatario)
    text = msg.as_string()
    server.sendmail(from_address, to_address, text)
    server.quit()
    
    print('¡ÉXITO! El correo electrónico se ha enviado correctamente.')

except Exception as e:
    print(f'Error al enviar el correo electrónico: {e}')
