# Mass Email Sender

Este programa envía correos electrónicos a clientes utilizando Outlook, con datos importados desde Excel a SQLite.

## Requisitos

1. Tener Outlook instalado y configurado en Windows
2. Python 3.x instalado
3. Archivo Excel con las columnas:
   - `Agency Code`: Código de la agencia
   - `Report email`: Correo electrónico del destinatario

## Configuración

1. Instalar las dependencias:
```bash
pip install -r requirements.txt
```

2. Preparar los archivos:
   - Archivo Excel `InfoAgencias1099Email.xlsx` con los datos de los clientes
   - Archivos PDF nombrados según el código de agencia (ejemplo: si el Agency Code es "5008", el archivo debe ser "5008.pdf")
   - Colocar los PDFs en la carpeta `uploads`

## Uso de la Interfaz Web

1. Iniciar la aplicación web:
```bash
python app.py
```

2. Abrir el navegador y acceder a:
```
http://localhost:5000
```

3. La interfaz web ofrece las siguientes funciones:
   - **Importar Excel**: Sube el archivo Excel con los datos de los clientes
   - **PDFs Disponibles**: Muestra los PDFs en la carpeta uploads
   - **Envío de Correos**:
     - Control de tiempo de espera entre envíos (1-300 segundos)
     - Botón para enviar todos los correos pendientes
     - Botón para enviar correos individuales
   - **Gestión**:
     - Borrar todos los PDFs
     - Limpiar la base de datos
     - Ver logs del sistema

## Proceso de Envío de Correos

1. **Preparación**:
   - Subir el archivo Excel con los datos
   - Colocar los PDFs en la carpeta `uploads`
   - Verificar que Outlook esté abierto y funcionando

2. **Envío Masivo**:
   - Ajustar el tiempo de espera entre envíos (recomendado: 15 segundos)
   - Hacer clic en "Enviar Todos"
   - El sistema:
     - Verifica los archivos PDF
     - Envía los correos uno por uno
     - Espera el tiempo especificado entre envíos
     - Actualiza el estado en tiempo real
     - Registra cada acción en los logs

3. **Monitoreo**:
   - La interfaz muestra el progreso del envío
   - Se puede ver el historial en la sección de logs
   - Los correos enviados se marcan automáticamente

## Solución de Problemas

1. **Error de Outlook**:
   - Asegurarse de que Outlook esté abierto
   - Verificar que la cuenta de correo esté configurada
   - Reiniciar la aplicación si es necesario

2. **PDFs no encontrados**:
   - Verificar que los nombres coincidan exactamente con los Agency Code
   - Asegurarse de que estén en la carpeta `uploads`

3. **Errores de Excel**:
   - Verificar que el archivo tenga las columnas requeridas
   - Asegurarse de que los correos electrónicos sean válidos

## Mantenimiento

- **Limpiar Base de Datos**: Usar el botón "Limpiar DB" para reiniciar el estado
- **Eliminar PDFs**: Usar "Eliminar Todo" en la sección de PDFs
- **Logs**: Revisar periódicamente y limpiar si es necesario
