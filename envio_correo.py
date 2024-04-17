import os
import smtplib
import pandas as pd
from dotenv import load_dotenv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import datetime
import matplotlib.pyplot as plt

# Cargamos las variables de entorno desde un archivo .env
load_dotenv()

# Obtenemos el remitente y el asunto del archivo .env
remitente = os.getenv("USER")
asunto ='Mensaje Lilly'

# Diccionario que mapea cada valor de "alerta" al archivo HTML correspondiente
archivos_html = {
    'Bienvenido al programa de apoyo': 'PP-AL-CO-0261-COL_Ver Hoja 1.html',
    'Manejo de la diarrea': 'PP-AL-CO-0261-COL_Ver Hoja 2.html',
    'Consejos nutricionales para el manejo de la diarrea': 'PP-AL-CO-0261-COL_Ver Hoja 3.html',
    'Seguimiento de la diarrea': 'PP-AL-CO-0261-COL_Ver Hoja 4.html',
    'Compartenos tu experiencia': 'PP-AL-CO-0261-COL_Ver Hoja 5.html',
    'Prepara tu Proxima cita': 'PP-AL-CO-0261-COL_Ver Hoja 6.html',
    'Recomendaciones de bienestar': 'PP-AL-CO-0261-COL_Ver Hoja 7.html',
    'Recomendaciones para el manejo de la fatiga': 'PP-AL-CO-0261-COL_Ver Hoja 8.html',
    'Ejercicios para mantenerce activo': 'PP-AL-CO-0261-COL_Ver Hoja 9.html',
    'Busca apoyo': 'PP-AL-CO-0261-COL_Ver Hoja 10.html',
    'Cierre del programa': 'PP-AL-CO-0261-COL_Ver Hoja 11.html',
}

# Leemos el archivo Excel con la lista de destinatarios
df = pd.read_excel('C:/Users/david.sanguino/OneDrive - SOULMEDICAL LTDA/LILLY/Correos Lilly/detalle-de-paciente.xlsx')

# Configuramos el servidor SMTP de Gmail
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(remitente, os.getenv('PASS'))

# Obtener fecha y hora actual para el registro de envíos
fecha_envio = datetime.datetime.now()

# Inicializar contadores para seguimiento de correos
correos_enviados = 0
correos_rebotados = 0

# Listas para almacenar datos de correos enviados y rebotados por fecha
fechas = []
correos_enviados_por_fecha = []
correos_rebotados_por_fecha = []

for index, row in df.iterrows():
    destinatario = row['Correo electrónico']
    nombre = row['Nombres'].capitalize()
    alerta = row['Mensaje']

    # Verificamos si la alerta está en el diccionario
    if alerta in archivos_html:
        # Construir la ruta completa al archivo HTML
        ruta_html = f'C:/Users/david.sanguino/OneDrive - SOULMEDICAL LTDA/LILLY/Correos Lilly/Bienvenida/{archivos_html[alerta]}'

        # Verificamos la existencia del archivo antes de intentar abrirlo
        if os.path.exists(ruta_html):
            # Obtener el contenido HTML una vez fuera del bucle
            with open(ruta_html, 'r', encoding='utf-8') as archivo_html:
                html = archivo_html.read()

                # Crear mensaje MIME para el correo
                msg = MIMEMultipart()
                msg['Subject'] = alerta
                msg['From'] = remitente
                msg['To'] = destinatario

                # Leer el contenido HTML personalizado y reemplazar la variable {PruEBA} con el nombre
                html_personalizado = html.replace('{PruEBA}', nombre)
                msg.attach(MIMEText(html_personalizado, 'html', 'utf-8'))

                try:
                    # Intentamos enviar el correo
                    server.sendmail(remitente, destinatario, msg.as_string())
                    correos_enviados += 1
                except Exception as e:
                    # Si ocurre un error al enviar el correo, lo consideramos como rebotado
                    correos_rebotados += 1

                # Guardar datos por fecha DENTRO del bucle
                fechas.append(fecha_envio.date())  # Solo la fecha
                correos_enviados_por_fecha.append(correos_enviados)
                correos_rebotados_por_fecha.append(correos_rebotados)

        else:
            print(f"El archivo HTML especificado para la alerta '{alerta}' no se encuentra en la ruta proporcionada.")
    else:
        print(f"La alerta '{alerta}' no está mapeada a un archivo HTML en el diccionario.")

# Cerrar conexión con el servidor SMTP
server.quit()

# Imprimir resultados
print(f'Fecha de envío: {fecha_envio}')
print(f'Correos Enviados: {correos_enviados}')
print(f'Correos Rebotados: {correos_rebotados}')

# Crear gráfico de torta
plt.figure(figsize=(8, 8))
labels = ['Correos Enviados', 'Correos Rebotados']
sizes = [correos_enviados, correos_rebotados]
colors = ['lightblue', 'lightcoral']

# Etiquetas con número de correos enviados y rebotados
labels_with_count = [f'{label} ({count})' for label, count in zip(labels, sizes)]

plt.pie(sizes, labels=labels_with_count, colors=colors, startangle=140)
plt.title('Correos Enviados y Rebotados', y=1.08)  # Ajustar la posición del título
plt.axis('equal')

# Mostrar el gráfico
plt.show()
