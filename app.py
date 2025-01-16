from flask import Flask, render_template, request, jsonify, send_file
import os
from database import DatabaseManager
from mailer import OutlookSender
import pandas as pd
from datetime import datetime
import shutil
import time

app = Flask(__name__)

# Inicializar la base de datos al arrancar
db = DatabaseManager()
db.connect()  # Conectar antes de setup
db.setup_database()
db.close()  # Cerrar después de setup

# Variables globales para el estado del envío
sending_status = {
    'is_sending': False,
    'total': 0,
    'current': 0,
    'current_agency': ''
}

# Crear instancia del manejador de Outlook
mailer = OutlookSender()

# Asegurar que existe la carpeta uploads
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def index():
    # Obtener lista de PDFs en el directorio de uploads
    pdfs = [f for f in os.listdir(UPLOAD_FOLDER) if f.endswith('.pdf')]
    
    # Obtener registros de la base de datos
    sent_emails = db.get_sent_emails()
    pending_emails = db.get_pending_clients()
    
    # Obtener logs
    logs = db.get_logs(limit=50)
    
    return render_template('index.html', 
                         pdfs=pdfs,
                         sent_emails=sent_emails,
                         pending_emails=pending_emails,
                         logs=logs,
                         sending_status=sending_status)

@app.route('/get-status')
def get_status():
    return jsonify(sending_status)

@app.route('/upload-pdf', methods=['POST'])
def upload_pdf():
    try:
        if 'pdf' not in request.files:
            return jsonify({'success': False, 'message': 'No se encontró archivo PDF'})
        
        file = request.files['pdf']
        if file.filename == '':
            return jsonify({'success': False, 'message': 'No se seleccionó archivo'})
        
        if file and file.filename.endswith('.pdf'):
            filename = file.filename
            file.save(os.path.join(UPLOAD_FOLDER, filename))
            return jsonify({'success': True, 'message': f'PDF {filename} subido correctamente'})
            
        return jsonify({'success': False, 'message': 'Tipo de archivo no válido'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/import-excel', methods=['POST'])
def import_excel():
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'message': 'No se encontró archivo Excel'})
            
        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'message': 'No se seleccionó archivo'})
            
        if not file.filename.endswith('.xlsx'):
            return jsonify({'success': False, 'message': 'El archivo debe ser un Excel (.xlsx)'})
            
        # Guardar el archivo temporalmente
        temp_path = 'temp_excel.xlsx'
        file.save(temp_path)
        
        try:
            # Importar a la base de datos
            success, message = db.import_from_excel(temp_path)
            
            # Eliminar archivo temporal
            if os.path.exists(temp_path):
                os.remove(temp_path)
                
            return jsonify({'success': success, 'message': message})
            
        except Exception as e:
            if os.path.exists(temp_path):
                os.remove(temp_path)
            raise e
            
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/send-email/<int:client_id>')
def send_single_email(client_id):
    try:
        # Obtener cliente
        client = db.get_client(client_id)
        if not client:
            return jsonify({'success': False, 'message': 'Cliente no encontrado'})
        
        save_as_draft = request.args.get('draft', default='false').lower() == 'true'
        agency_code = client['Agency Code']
        pdf_path = os.path.join(UPLOAD_FOLDER, f"{agency_code}.pdf")
        
        if not os.path.exists(pdf_path):
            error_msg = f"PDF no encontrado: {pdf_path}"
            db.add_log(agency_code, 'send_email', 'error', error_msg)
            return jsonify({'success': False, 'message': error_msg})
        
        try:
            if mailer.send_email(client, pdf_path, save_as_draft):
                db.mark_email_sent(client['id'])
                status = 'success'
                message = f'Correo {"guardado como borrador" if save_as_draft else "enviado"} exitosamente'
            else:
                status = 'error'
                message = f'Error al {"guardar" if save_as_draft else "enviar"} el correo'
        except Exception as e:
            status = 'error'
            message = str(e)
        
        # Registrar en el log
        action = 'save_draft' if save_as_draft else 'send_email'
        db.add_log(agency_code, action, status, message)
        
        return jsonify({
            'success': status == 'success',
            'message': message
        })
        
    except Exception as e:
        error_msg = f"Error inesperado: {str(e)}"
        db.add_log('SYSTEM', 'send_email', 'error', error_msg)
        return jsonify({'success': False, 'message': error_msg})

@app.route('/send-all')
def send_all_pending():
    global sending_status
    
    try:
        if sending_status['is_sending']:
            return jsonify({'success': False, 'message': 'Ya hay un proceso de envío en marcha'})
        
        delay = request.args.get('delay', default=15, type=int)
        save_as_draft = request.args.get('draft', default='false').lower() == 'true'
        clients = db.get_pending_clients()
        
        sending_status = {
            'is_sending': True,
            'total': len(clients),
            'current': 0,
            'current_agency': ''
        }
        
        results = []
        for i, client in enumerate(clients):
            agency_code = client['Agency Code']
            sending_status['current'] = i + 1
            sending_status['current_agency'] = agency_code
            
            pdf_path = os.path.join(UPLOAD_FOLDER, f"{agency_code}.pdf")
            if os.path.exists(pdf_path):
                if i > 0:  # No esperar antes del primer correo
                    time.sleep(delay)
                    
                try:
                    if mailer.send_email(client, pdf_path, save_as_draft):
                        db.mark_email_sent(client['id'])
                        status = 'success'
                        message = f'{"Guardado como borrador" if save_as_draft else "Enviado"} después de {delay} segundos de espera'
                    else:
                        status = 'error'
                        message = f'Error al {"guardar" if save_as_draft else "enviar"} el correo'
                except Exception as e:
                    status = 'error'
                    message = str(e)
                
                # Registrar en el log
                action = 'save_draft' if save_as_draft else 'send_email'
                db.add_log(agency_code, action, status, message)
                
                results.append({
                    'agency': agency_code,
                    'status': status,
                    'message': message
                })
        
        sending_status['is_sending'] = False
        return jsonify({
            'success': True, 
            'results': results,
            'message': f'Proceso completado: {"guardados como borrador" if save_as_draft else "enviados"} con {delay} segundos de espera entre correos'
        })
    except Exception as e:
        sending_status['is_sending'] = False
        return jsonify({'success': False, 'message': str(e)})

@app.route('/delete-pdf/<filename>')
def delete_pdf(filename):
    try:
        if filename.endswith('.pdf'):
            file_path = os.path.join(UPLOAD_FOLDER, filename)
            if os.path.exists(file_path):
                # Obtener el código de agencia del nombre del archivo
                agency_code = filename.replace('.pdf', '')
                
                # Eliminar el archivo
                os.remove(file_path)
                
                # Eliminar el registro de la base de datos
                db.delete_client(agency_code)
                
                return jsonify({'success': True, 'message': f'PDF {filename} y sus datos eliminados correctamente'})
        return jsonify({'success': False, 'message': 'Archivo no encontrado'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/delete-all-pdfs')
def delete_all_pdfs():
    try:
        pdfs = [f for f in os.listdir(UPLOAD_FOLDER) if f.endswith('.pdf')]
        for pdf in pdfs:
            os.remove(os.path.join(UPLOAD_FOLDER, pdf))
        
        # Eliminar todos los registros de la base de datos
        db.delete_all_clients()
        
        return jsonify({'success': True, 'message': f'Se eliminaron {len(pdfs)} PDFs y todos los datos asociados'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/clear-database')
def clear_database():
    try:
        db.clear_all_records()
        return jsonify({'success': True, 'message': 'Base de datos limpiada correctamente'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/clear-logs')
def clear_logs():
    try:
        db.clear_logs()
        return jsonify({'success': True, 'message': 'Logs limpiados correctamente'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

if __name__ == '__main__':
    app.run(debug=True)
