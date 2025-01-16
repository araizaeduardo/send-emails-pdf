import os
import time
import pythoncom
from typing import Dict
import win32com.client
from database import DatabaseManager

class OutlookSender:
    def __init__(self):
        pass

    def send_email(self, client: Dict, pdf_path: str, save_as_draft: bool = False) -> bool:
        """
        Envía un correo usando Outlook o lo guarda como borrador
        Args:
            client (Dict): Diccionario con datos del cliente
            pdf_path (str): Ruta al archivo PDF a adjuntar
            save_as_draft (bool): Si es True, guarda el correo como borrador en lugar de enviarlo
        Returns:
            bool: True si la operación fue exitosa, False en caso contrario
        """
        # Inicializar COM para este hilo
        pythoncom.CoInitialize()
        
        try:
            if not os.path.exists(pdf_path):
                raise Exception(f"PDF no encontrado: {pdf_path}")

            # Crear una nueva instancia de Outlook
            try:
                outlook = win32com.client.Dispatch('Outlook.Application')
            except Exception as e:
                raise Exception(f"Error al conectar con Outlook: {str(e)}")

            # Crear el mensaje
            try:
                mail = outlook.CreateItem(0)  # 0 = olMailItem
            except Exception as e:
                raise Exception(f"Error al crear mensaje de Outlook: {str(e)}")

            try:
                mail.Subject = "Formulario 1099-NEC – Comisiones Recibidas"
                mail.To = client['Report email']
            except Exception as e:
                raise Exception(f"Error al establecer destinatario o asunto: {str(e)}")
            
            # Cuerpo del mensaje
            try:
                mail.HTMLBody = """
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
            except Exception as e:
                raise Exception(f"Error al establecer el cuerpo del mensaje: {str(e)}")

            # Adjuntar PDF
            try:
                mail.Attachments.Add(pdf_path)
            except Exception as e:
                raise Exception(f"Error al adjuntar PDF: {str(e)}")

            # Enviar el correo o guardarlo como borrador
            try:
                if save_as_draft:
                    mail.Save()
                    print(f"Correo guardado como borrador para {client['Report email']}")
                else:
                    mail.Send()
                    print(f"Correo enviado exitosamente a {client['Report email']}")
                return True
            except Exception as e:
                raise Exception(f"Error al {'guardar' if save_as_draft else 'enviar'} el correo: {str(e)}")

        except Exception as e:
            print(f"Error al procesar correo para {client['Report email']}: {str(e)}")
            raise e
        finally:
            # Liberar COM al terminar
            pythoncom.CoUninitialize()

def main():
    # Inicializar manejador de base de datos y Outlook
    db = DatabaseManager()
    mailer = OutlookSender()

    # Ruta base donde están los PDFs
    base_path = "c:\\Users\\Eduardo\\Documents\\massemails\\uploads"
    
    try:
        # Obtener clientes pendientes de envío
        pending_clients = db.get_pending_clients()
        
        if not pending_clients:
            print("No hay correos pendientes por enviar")
            return

        print(f"Enviando {len(pending_clients)} correos...")
        
        # Enviar correos
        for client in pending_clients:
            try:
                # Construir ruta al PDF
                pdf_filename = f"{client['Agency Code']}.pdf"
                pdf_path = os.path.join(base_path, pdf_filename)
                
                # Enviar correo
                mailer.send_email(client, pdf_path)
                
                # Marcar como enviado en la base de datos
                db.mark_email_sent(client['id'])
                
                # Esperar un poco entre envíos
                time.sleep(1)
                
            except Exception as e:
                print(f"Error procesando cliente {client['Agency Code']}: {str(e)}")
                continue
                
    except Exception as e:
        print(f"Error en el proceso de envío: {str(e)}")
    
    finally:
        print("Proceso finalizado")

if __name__ == "__main__":
    main()
