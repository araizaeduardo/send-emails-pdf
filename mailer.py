import os
import time
import pythoncom
from typing import Dict
import win32com.client
from database import DatabaseManager

class OutlookSender:
    def __init__(self):
        pass

    def send_email(self, client: Dict, pdf_path: str, save_as_draft: bool = False, template: Dict = None) -> bool:
        """
        Envía un correo electrónico usando Outlook
        """
        try:
            # Inicializar COM antes de usar Outlook
            pythoncom.CoInitialize()
            
            outlook = win32com.client.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)  # 0 = olMailItem
            
            mail.To = client['Report email']
            
            # Usar plantilla si está disponible
            if template:
                mail.Subject = template['subject']
                mail.HTMLBody = template['body']
            else:
                mail.Subject = f"New Message - {client['Agency Code']}"
                mail.HTMLBody = f"""
                    <p>Dear Client,</p>
                    <p>Please find attached.</p>
                    <p>Best regards,</p>
                """
            
            # Adjuntar PDF
            if os.path.exists(pdf_path):
                mail.Attachments.Add(pdf_path)
            else:
                raise Exception(f"PDF no encontrado: {pdf_path}")
            
            if save_as_draft:
                mail.Save()
                print(f"Correo guardado como borrador para {client['Report email']}")
            else:
                mail.Send()
                print(f"Correo enviado exitosamente a {client['Report email']}")
                return True
                
        except Exception as e:
            # Capturar errores específicos de Outlook
            if "Outlook" in str(e):
                raise Exception("Error al conectar con Outlook. Asegúrate de que Outlook esté abierto y configurado.")
            elif "Attachments" in str(e):
                raise Exception("Error al adjuntar el archivo PDF. Verifica que el archivo exista y no esté dañado.")
            else:
                raise Exception(f"Error al enviar el correo: {str(e)}")

        finally:
            # Liberar COM al terminar
            pythoncom.CoUninitialize()
            try:
                pythoncom.CoUninitialize()
            except:
                pass

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
