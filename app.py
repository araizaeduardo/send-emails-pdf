from flask import Flask, render_template, request, jsonify, send_file
import os
from database import DatabaseManager
from mailer import OutlookSender
import pandas as pd
from datetime import datetime
import shutil
import time
import uuid

app = Flask(__name__)

# Configuración de la carpeta de uploads
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

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

@app.route('/')
def index():
    try:
        # Obtener lista de PDFs en el directorio de uploads
        pdfs = [f for f in os.listdir(app.config['UPLOAD_FOLDER']) if f.endswith('.pdf')]
        
        # Obtener registros de la base de datos
        sent_emails = db.get_sent_emails()
        pending_emails = db.get_pending_clients()
        
        # Obtener logs
        logs = db.get_logs()  # Eliminado el parámetro limit
        
        return render_template('index.html', 
                             pdfs=pdfs,
                             sent_emails=sent_emails,
                             pending_emails=pending_emails,
                             logs=logs,
                             sending_status=sending_status)
    except Exception as e:
        return f"Error: {str(e)}"

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
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            return jsonify({'success': True, 'message': f'PDF {filename} subido correctamente'})
            
        return jsonify({'success': False, 'message': 'Tipo de archivo no válido'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/import-excel', methods=['POST'])
def import_excel():
    temp_path = None
    db_instance = None
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'message': 'No se encontró archivo Excel'})
            
        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'message': 'No se seleccionó archivo'})
            
        if not file.filename.endswith('.xlsx'):
            return jsonify({'success': False, 'message': 'El archivo debe ser un Excel (.xlsx)'})
        
        # Crear un nombre único para el archivo temporal
        temp_path = f'temp_excel_{uuid.uuid4()}.xlsx'
        
        # Guardar el archivo temporalmente
        file.save(temp_path)
        
        # Crear una nueva instancia de DatabaseManager para esta operación
        db_instance = DatabaseManager()
        success, message = db_instance.import_from_excel(temp_path)
        
        return jsonify({'success': success, 'message': message})
            
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})
        
    finally:
        # Cerrar la conexión de la base de datos
        if db_instance:
            db_instance.close()
            
        # Asegurarse de eliminar el archivo temporal
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except:
                pass

@app.route('/send-email/<client_id>')
def send_single_email(client_id):
    try:
        client = db.get_client_by_id(client_id)
        if not client:
            return jsonify({'success': False, 'message': 'Cliente no encontrado'})
        
        save_as_draft = request.args.get('draft', default='false').lower() == 'true'
        template_id = request.args.get('template_id')
        template = db.get_template(template_id) if template_id else None
        
        agency_code = client['Agency Code']
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{agency_code}.pdf")
        
        if not os.path.exists(pdf_path):
            error_msg = f"PDF no encontrado: {pdf_path}"
            return jsonify({'success': False, 'message': error_msg})
        
        try:
            mailer.send_email(client, pdf_path, save_as_draft, template)
            
            # Registrar el envío exitoso
            db.add_sent_email(
                agency_code=client['Agency Code'],
                email=client['Report email'],
                status='success',
                message='Correo enviado correctamente' if not save_as_draft else 'Guardado como borrador'
            )
            
            return jsonify({
                'success': True, 
                'message': f"Correo {'guardado como borrador' if save_as_draft else 'enviado'} correctamente"
            })
            
        except Exception as e:
            error_msg = str(e)
            db.add_sent_email(
                agency_code=client['Agency Code'],
                email=client['Report email'],
                status='error',
                message=error_msg
            )
            return jsonify({'success': False, 'message': error_msg})
            
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/send-all-emails')
def send_all_emails():
    try:
        # Obtener la plantilla seleccionada
        template_id = request.args.get('template_id')
        template = db.get_template(template_id) if template_id else None
        
        # Verificar si es borrador
        save_as_draft = request.args.get('draft', 'false').lower() == 'true'
        
        # Obtener todos los clientes pendientes
        pending_clients = db.get_pending_clients()
        if not pending_clients:
            return jsonify({'success': False, 'message': 'No hay correos pendientes para enviar'})
            
        success_count = 0
        error_count = 0
        error_messages = []
        
        # Actualizar estado global
        global sending_status
        sending_status['is_sending'] = True
        sending_status['total'] = len(pending_clients)
        sending_status['current'] = 0
        
        for client in pending_clients:
            try:
                # Convertir los nombres de campos al formato esperado por mailer.py
                mailer_client = {
                    'Agency Code': client['agency_code'],
                    'Report email': client['email']
                }
                
                agency_code = client['agency_code']
                sending_status['current_agency'] = agency_code
                sending_status['current'] += 1
                
                pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{agency_code}.pdf")
                
                if not os.path.exists(pdf_path):
                    error_count += 1
                    error_msg = f"PDF no encontrado para {agency_code}"
                    error_messages.append(error_msg)
                    db.add_sent_email(
                        agency_code=agency_code,
                        email=client['email'],
                        status='error',
                        message=error_msg
                    )
                    continue
                
                # Enviar correo con el formato correcto de cliente
                mailer.send_email(mailer_client, pdf_path, save_as_draft, template)
                
                # Registrar envío exitoso
                status = 'draft' if save_as_draft else 'success'
                message = 'Guardado como borrador' if save_as_draft else 'Correo enviado correctamente'
                db.add_sent_email(
                    agency_code=agency_code,
                    email=client['email'],
                    status=status,
                    message=message
                )
                success_count += 1
                
            except Exception as e:
                error_count += 1
                error_msg = f"Error al procesar {client['agency_code']}: {str(e)}"
                error_messages.append(error_msg)
                db.add_sent_email(
                    agency_code=client['agency_code'],
                    email=client['email'],
                    status='error',
                    message=error_msg
                )
        
        # Actualizar estado final
        sending_status['is_sending'] = False
        sending_status['current_agency'] = ''
        
        # Preparar mensaje de respuesta
        if success_count > 0:
            message = f"{'Borradores guardados' if save_as_draft else 'Correos enviados'}: {success_count}"
            if error_count > 0:
                message += f", Errores: {error_count}"
        else:
            message = "No se pudo procesar ningún correo"
        
        if error_messages:
            message += f"\nErrores encontrados:\n" + "\n".join(error_messages)
        
        return jsonify({
            'success': True,
            'message': message,
            'success_count': success_count,
            'error_count': error_count,
            'errors': error_messages
        })
            
    except Exception as e:
        if 'sending_status' in globals():
            sending_status['is_sending'] = False
            sending_status['current_agency'] = ''
        return jsonify({'success': False, 'message': f"Error en el proceso: {str(e)}"})

@app.route('/send-all')
def send_all_pending():
    try:
        delay = float(request.args.get('delay', default='5'))
        save_as_draft = request.args.get('draft', default='false').lower() == 'true'
        template_id = request.args.get('template_id')
        template = db.get_template(template_id) if template_id else None
        
        # Obtener todos los clientes que no tienen correo enviado
        clients = db.get_all_clients()
        total = len(clients)
        
        if total == 0:
            return jsonify({'success': False, 'message': 'No hay correos pendientes'})
        
        # Inicializar estado del envío
        sending_status['total'] = total
        sending_status['current'] = 0
        sending_status['current_agency'] = ''
        sending_status['errors'] = []
        
        # Enviar correos
        for i, client in enumerate(clients):
            agency_code = client['Agency Code']
            sending_status['current'] = i + 1
            sending_status['current_agency'] = agency_code
            
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{agency_code}.pdf")
            if os.path.exists(pdf_path):
                if i > 0:  # No esperar antes del primer correo
                    time.sleep(delay)
                    
                try:
                    mailer.send_email(client, pdf_path, save_as_draft, template)
                    if not save_as_draft:
                        db.add_sent_email(agency_code, client['Report email'], 'success', 'Correo enviado correctamente')
                except Exception as e:
                    error_msg = f"Error al enviar correo a {client['Report email']}: {str(e)}"
                    db.add_sent_email(agency_code, client['Report email'], 'error', error_msg)
                    sending_status['errors'].append({
                        'agency_code': agency_code,
                        'email': client['Report email'],
                        'error': str(e)
                    })
            else:
                error_msg = f"PDF no encontrado para {agency_code}"
                db.add_sent_email(agency_code, client['Report email'], 'error', error_msg)
                sending_status['errors'].append({
                    'agency_code': agency_code,
                    'email': client['Report email'],
                    'error': error_msg
                })
        
        return jsonify({
            'success': True,
            'message': f"Proceso de envío completado. {total - len(sending_status['errors'])} correos enviados, {len(sending_status['errors'])} errores."
        })
        
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/delete-pdf/<filename>')
def delete_pdf(filename):
    try:
        if filename.endswith('.pdf'):
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
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

@app.route('/delete-all-pdfs', methods=['POST'])
def delete_all_pdfs():
    try:
        # Obtener lista de PDFs
        pdfs = [f for f in os.listdir(app.config['UPLOAD_FOLDER']) if f.endswith('.pdf')]
        
        # Eliminar cada PDF
        for pdf in pdfs:
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf)
            os.remove(pdf_path)
        
        return jsonify({
            'success': True,
            'message': f'Se eliminaron {len(pdfs)} archivos PDF correctamente'
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error al eliminar PDFs: {str(e)}'
        })

@app.route('/clear-database', methods=['POST'])
def clear_database():
    try:
        # Limpiar la tabla de clientes
        db.clear_clients()
        # Registrar la acción en el log
        db.add_log('SYSTEM', 'clear_database', 'success', 'Base de datos limpiada correctamente')
        return jsonify({'success': True, 'message': 'Base de datos limpiada correctamente'})
    except Exception as e:
        db.add_log('SYSTEM', 'clear_database', 'error', str(e))
        return jsonify({'success': False, 'message': str(e)})

@app.route('/clear-logs', methods=['POST'])
def clear_logs():
    try:
        db.clear_logs()
        return jsonify({'success': True, 'message': 'Logs limpiados correctamente'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/preview-email/<int:client_id>')
def preview_email(client_id):
    try:
        db = DatabaseManager()
        client = db.get_client_by_id(client_id)
        if not client:
            return jsonify({'success': False, 'message': 'Cliente no encontrado'})

        # Verificar si existe el PDF
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{client['Agency Code']}.pdf")
        has_pdf = os.path.exists(pdf_path)

        # Generar el cuerpo del correo
        email_body = f"""
        <p>Estimado/a dueño/a de agencia,</p>

        <p>Adjunto a este mensaje encontrará el formulario 1099-NEC correspondiente a las comisiones recibidas durante el período de ventas con Paseo Travel & Tours, Inc.</p>

        <p>El total reflejado puede incluir montos de ventas MCO, las cuales también se reportan. Si necesita un desglose detallado, puede solicitar el reporte de las ventas asociadas a las 890.</p>

        <p>Le pedimos que revise cuidadosamente la información contenida en el formulario. En caso de identificar algún error o si necesita una copia impresa, por favor comuníquese con nosotros a más tardar el 28 de enero al (818) 244-2184 para solicitar la corrección correspondiente.</p>

        <p>Agradecemos su atención y quedamos atentos a cualquier consulta.</p>

        <p>Atentamente,<br>
        Departamento de Contabilidad<br>
        Paseo Travel & Tours, Inc.</p>

        <hr>

        <p><small><strong>Aviso de Confidencialidad:</strong> Este mensaje contiene información confidencial y está dirigido únicamente al destinatario indicado. Si usted no es el destinatario, queda estrictamente prohibida la distribución, copia o divulgación de este correo electrónico. Si lo ha recibido por error, por favor notifíquelo al remitente inmediatamente y elimine el mensaje de su sistema.</p>

        <p>Tenga en cuenta que la transmisión por correo electrónico no puede garantizarse como segura o libre de errores, ya que la información puede ser interceptada, dañada, perdida, llegar incompleta o contener virus. Paseo Travel & Tours, Inc. no asume responsabilidad por errores u omisiones en el contenido de este mensaje. Si necesita verificación, solicite una versión en papel.</small></p>

        <p><small>Paseo Travel & Tours, Inc.<br>
        PO BOX 10060, Glendale, CA 91209<br>
        <a href="http://www.paseotravel.us">www.paseotravel.us</a></small></p>
        """

        preview_data = {
            'success': True,
            'subject': 'Formulario 1099-NEC – Comisiones Recibidas',
            'to': client['Report email'],
            'body': email_body,
            'has_pdf': has_pdf,
            'pdf_name': f"{client['Agency Code']}.pdf" if has_pdf else None
        }

        return jsonify(preview_data)

    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/templates')
def get_templates():
    db = DatabaseManager()
    templates = db.get_all_templates()
    return jsonify(templates)

@app.route('/template/<int:template_id>')
def get_template(template_id):
    db = DatabaseManager()
    template = db.get_template(template_id)
    if template:
        return jsonify({'success': True, 'template': template})
    return jsonify({'success': False, 'message': 'Plantilla no encontrada'})

@app.route('/template/add', methods=['POST'])
def add_template():
    try:
        data = request.json
        name = data.get('name')
        subject = data.get('subject')
        body = data.get('body')

        if not all([name, subject, body]):
            return jsonify({'success': False, 'message': 'Todos los campos son requeridos'})

        db = DatabaseManager()
        if db.add_template(name, subject, body):
            return jsonify({'success': True, 'message': 'Plantilla agregada correctamente'})
        return jsonify({'success': False, 'message': 'Error al agregar la plantilla'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/template/update/<int:template_id>', methods=['POST'])
def update_template(template_id):
    try:
        data = request.json
        name = data.get('name')
        subject = data.get('subject')
        body = data.get('body')

        if not all([name, subject, body]):
            return jsonify({'success': False, 'message': 'Todos los campos son requeridos'})

        db = DatabaseManager()
        if db.update_template(template_id, name, subject, body):
            return jsonify({'success': True, 'message': 'Plantilla actualizada correctamente'})
        return jsonify({'success': False, 'message': 'Error al actualizar la plantilla'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/template/delete/<int:template_id>', methods=['POST'])
def delete_template(template_id):
    try:
        db = DatabaseManager()
        if db.delete_template(template_id):
            return jsonify({'success': True, 'message': 'Plantilla eliminada correctamente'})
        return jsonify({'success': False, 'message': 'No se puede eliminar la plantilla por defecto'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

if __name__ == '__main__':
    app.run(debug=True)
