import pandas as pd
import sqlite3
from typing import List, Dict
import warnings

# Ignorar advertencias de openpyxl
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

class DatabaseManager:
    def __init__(self, db_name: str = 'clients.db'):
        self.db_name = db_name
        self.conn = None
        self.cursor = None

    def connect(self):
        """Establece conexión con la base de datos"""
        try:
            if self.conn is None:
                self.conn = sqlite3.connect(self.db_name)
                self.cursor = self.conn.cursor()
        except Exception as e:
            print(f"Error al conectar a la base de datos: {str(e)}")
            raise

    def close(self):
        """Cierra la conexión con la base de datos"""
        try:
            if self.cursor:
                self.cursor.close()
            if self.conn:
                self.conn.commit()
                self.conn.close()
        except Exception as e:
            print(f"Error al cerrar la base de datos: {str(e)}")
        finally:
            self.cursor = None
            self.conn = None

    def ensure_connection(self):
        """Asegura que hay una conexión activa"""
        try:
            if self.conn is None or self.cursor is None:
                self.connect()
            # Verificar que la conexión está activa
            self.cursor.execute("SELECT 1")
        except (sqlite3.Error, AttributeError):
            # Si hay algún error, intentar reconectar
            self.close()
            self.connect()

    def setup_database(self):
        """Configura la base de datos con las tablas necesarias"""
        try:
            self.ensure_connection()
            
            # Tabla de clientes
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS clients (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    "Agency Code" TEXT NOT NULL,
                    "Report email" TEXT NOT NULL,
                    email_sent INTEGER DEFAULT 0,
                    sent_date DATETIME,
                    has_pdf BOOLEAN DEFAULT FALSE
                )
            ''')

            # Tabla de logs
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS activity_logs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
                    agency_code TEXT,
                    action TEXT,
                    status TEXT,
                    message TEXT
                )
            ''')

            # Tabla de correos enviados
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS sent_emails (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    agency_code TEXT NOT NULL,
                    email TEXT NOT NULL,
                    sent_date DATETIME DEFAULT CURRENT_TIMESTAMP,
                    status TEXT,
                    message TEXT
                )
            ''')

            # Tabla de plantillas de correo
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS email_templates (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL,
                    subject TEXT NOT NULL,
                    body TEXT NOT NULL,
                    is_default INTEGER DEFAULT 0
                )
            ''')

            # Tabla de PDFs pendientes
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS pending_pdfs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    agency_code TEXT NOT NULL,
                    upload_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    processed BOOLEAN DEFAULT FALSE,
                    processed_date TIMESTAMP
                )
            ''')

            self.conn.commit()
        except Exception as e:
            print(f"Error en setup_database: {str(e)}")
            raise

    def import_from_excel(self, excel_path: str):
        """
        Importa datos desde Excel a SQLite
        Args:
            excel_path (str): Ruta al archivo Excel
        """
        try:
            # Leer Excel primero para validar
            df = pd.read_excel(excel_path)
            
            # Asegurarse de que existan las columnas necesarias
            required_columns = ['Agency Code', 'Report email']
            if not all(col in df.columns for col in required_columns):
                raise ValueError("El archivo Excel debe contener las columnas 'Agency Code' y 'Report email'")
            
            # Agregar columnas para tracking
            df['email_sent'] = 0
            df['sent_date'] = None
            
            # Asegurar conexión y configuración de la base de datos
            self.ensure_connection()
            self.setup_database()
            
            # Importar datos
            df.to_sql('clients', self.conn, if_exists='replace', index=True, index_label='id')
            self.conn.commit()
            
            # Agregar log
            self.add_log('SYSTEM', 'import_excel', 'success', f'Datos importados exitosamente desde {excel_path}')
            
            return True, "Datos importados exitosamente"
            
        except pd.errors.EmptyDataError:
            return False, "El archivo Excel está vacío"
        except Exception as e:
            error_msg = f"Error al importar datos: {str(e)}"
            try:
                if self.conn and self.cursor:
                    self.add_log('SYSTEM', 'import_excel', 'error', error_msg)
            except:
                pass
            return False, error_msg

    def get_pending_clients(self):
        """
        Obtiene los clientes que aún no han recibido el correo
        """
        try:
            self.ensure_connection()
            self.cursor.execute('''
                SELECT id, "Agency Code" as agency_code, "Report email" as email
                FROM clients c
                WHERE NOT EXISTS (
                    SELECT 1 FROM sent_emails s
                    WHERE s.agency_code = c."Agency Code"
                    AND s.status = 'success'
                )
            ''')
            return [{
                'id': row[0],
                'agency_code': row[1],
                'email': row[2]
            } for row in self.cursor.fetchall()]
        except Exception as e:
            print(f"Error al obtener clientes pendientes: {str(e)}")
            return []
        finally:
            self.close()

    def get_client_by_id(self, client_id):
        """
        Obtiene un cliente por su ID
        """
        try:
            self.ensure_connection()
            self.cursor.execute('''
                SELECT "Agency Code", "Report email"
                FROM clients
                WHERE id = ?
            ''', (client_id,))
            row = self.cursor.fetchone()
            if row:
                return {
                    'Agency Code': row[0],
                    'Report email': row[1]
                }
            return None
        except Exception as e:
            print(f"Error al obtener cliente: {str(e)}")
            return None
        finally:
            self.close()

    def clear_all_records(self):
        """
        Limpia todos los registros de la base de datos
        """
        try:
            self.ensure_connection()
            self.cursor.execute("DELETE FROM clients")
            self.conn.commit()
        finally:
            self.close()

    def reset_email_status(self, agency_code=None):
        """
        Resetea el estado de envío de correos
        Args:
            agency_code (str, optional): Código de agencia específico. Si es None, resetea todos.
        """
        try:
            self.ensure_connection()
            if agency_code:
                self.cursor.execute("""
                    UPDATE clients 
                    SET email_sent = 0,
                        sent_date = NULL
                    WHERE "Agency Code" = ?
                """, (agency_code,))
            else:
                self.cursor.execute("""
                    UPDATE clients 
                    SET email_sent = 0,
                        sent_date = NULL
                """)
            self.conn.commit()
        finally:
            self.close()

    def delete_client(self, agency_code: str):
        """
        Elimina un cliente de la base de datos
        Args:
            agency_code (str): Código de agencia del cliente a eliminar
        """
        try:
            self.ensure_connection()
            self.cursor.execute("""
                DELETE FROM clients 
                WHERE "Agency Code" = ?
            """, (agency_code,))
            self.conn.commit()
        finally:
            self.close()

    def delete_all_clients(self):
        """
        Elimina todos los clientes de la base de datos
        """
        try:
            self.ensure_connection()
            self.cursor.execute("DELETE FROM clients")
            self.conn.commit()
        finally:
            self.close()

    def clear_clients(self):
        """
        Elimina todos los clientes de la base de datos
        """
        try:
            self.ensure_connection()
            self.cursor.execute('DELETE FROM clients')
            self.conn.commit()
            return True
        except Exception as e:
            print(f"Error al limpiar la base de datos: {str(e)}")
            return False
        finally:
            self.close()

    def get_all_templates(self):
        """
        Obtiene todas las plantillas de correo
        """
        try:
            self.ensure_connection()
            self.cursor.execute('SELECT id, name, subject, body, is_default FROM email_templates ORDER BY is_default DESC, name')
            templates = self.cursor.fetchall()
            return [{
                'id': t[0],
                'name': t[1],
                'subject': t[2],
                'body': t[3],
                'is_default': t[4]
            } for t in templates]
        except Exception as e:
            print(f"Error al obtener plantillas: {str(e)}")
            return []
        finally:
            self.close()

    def get_template(self, template_id):
        """
        Obtiene una plantilla específica por ID
        """
        try:
            self.ensure_connection()
            self.cursor.execute('SELECT id, name, subject, body FROM email_templates WHERE id = ?', (template_id,))
            t = self.cursor.fetchone()
            if t:
                return {
                    'id': t[0],
                    'name': t[1],
                    'subject': t[2],
                    'body': t[3]
                }
            return None
        except Exception as e:
            print(f"Error al obtener plantilla: {str(e)}")
            return None
        finally:
            self.close()

    def add_template(self, name, subject, body):
        """
        Agrega una nueva plantilla
        """
        try:
            self.ensure_connection()
            self.cursor.execute('''
                INSERT INTO email_templates (name, subject, body)
                VALUES (?, ?, ?)
            ''', (name, subject, body))
            self.conn.commit()
            return True
        except Exception as e:
            print(f"Error al agregar plantilla: {str(e)}")
            return False
        finally:
            self.close()

    def update_template(self, template_id, name, subject, body):
        """
        Actualiza una plantilla existente
        """
        try:
            self.ensure_connection()
            self.cursor.execute('''
                UPDATE email_templates
                SET name = ?, subject = ?, body = ?
                WHERE id = ?
            ''', (name, subject, body, template_id))
            self.conn.commit()
            return True
        except Exception as e:
            print(f"Error al actualizar plantilla: {str(e)}")
            return False
        finally:
            self.close()

    def delete_template(self, template_id):
        """
        Elimina una plantilla (no permite eliminar la plantilla por defecto)
        """
        try:
            self.ensure_connection()
            self.cursor.execute('DELETE FROM email_templates WHERE id = ? AND is_default = 0', (template_id,))
            self.conn.commit()
            return True
        except Exception as e:
            print(f"Error al eliminar plantilla: {str(e)}")
            return False
        finally:
            self.close()

    def add_log(self, agency_code: str, action: str, status: str, message: str = None):
        """
        Agrega un registro al log de actividades
        """
        try:
            self.ensure_connection()
            self.cursor.execute('''
                INSERT INTO activity_logs (agency_code, action, status, message)
                VALUES (?, ?, ?, ?)
            ''', (agency_code, action, status, message))
            self.conn.commit()
            return True
        except Exception as e:
            print(f"Error al agregar log: {str(e)}")
            return False
        finally:
            self.close()

    def get_logs(self):
        """
        Obtiene todos los registros del log de actividades
        """
        try:
            self.ensure_connection()
            self.cursor.execute('''
                SELECT timestamp, agency_code, action, status, message
                FROM activity_logs
                ORDER BY timestamp DESC
            ''')
            return [{
                'timestamp': row[0],
                'agency_code': row[1],
                'action': row[2],
                'status': row[3],
                'message': row[4]
            } for row in self.cursor.fetchall()]
        except Exception as e:
            print(f"Error al obtener logs: {str(e)}")
            return []
        finally:
            self.close()

    def clear_logs(self):
        """
        Elimina todos los registros del log de actividades
        """
        try:
            self.ensure_connection()
            self.cursor.execute('DELETE FROM activity_logs')
            self.conn.commit()
            return True
        except Exception as e:
            print(f"Error al limpiar logs: {str(e)}")
            return False
        finally:
            self.close()

    def add_sent_email(self, agency_code: str, email: str, status: str, message: str):
        """
        Registra un correo enviado en la base de datos
        """
        try:
            self.ensure_connection()
            self.cursor.execute('''
                INSERT INTO sent_emails (agency_code, email, sent_date, status, message)
                VALUES (?, ?, CURRENT_TIMESTAMP, ?, ?)
            ''', (agency_code, email, status, message))
            
            # También registrar en el log
            self.add_log(agency_code, 'send_email', status, message)
            
            self.conn.commit()
            return True
        except Exception as e:
            print(f"Error al registrar correo enviado: {str(e)}")
            return False
        finally:
            self.close()

    def get_sent_emails(self):
        """
        Obtiene todos los correos enviados
        """
        try:
            self.ensure_connection()
            self.cursor.execute('''
                SELECT agency_code, email, sent_date, status, message
                FROM sent_emails
                ORDER BY sent_date DESC
            ''')
            return [{
                'agency_code': row[0],
                'email': row[1],
                'sent_date': row[2],
                'status': row[3],
                'message': row[4]
            } for row in self.cursor.fetchall()]
        except Exception as e:
            print(f"Error al obtener correos enviados: {str(e)}")
            return []
        finally:
            self.close()

    def get_pending_pdfs(self):
        """Obtiene todos los PDFs pendientes de procesamiento"""
        try:
            self.ensure_connection()
            self.cursor.execute("""
                SELECT id, agency_code, upload_date 
                FROM pending_pdfs 
                WHERE processed = FALSE
            """)
            return [{
                'id': row[0],
                'agency_code': row[1],
                'upload_date': row[2]
            } for row in self.cursor.fetchall()]
        except Exception as e:
            print(f"Error al obtener PDFs pendientes: {str(e)}")
            return []

    def add_pending_pdf(self, agency_code):
        """Agrega un nuevo PDF pendiente"""
        try:
            self.ensure_connection()
            self.cursor.execute("""
                INSERT INTO pending_pdfs (
                    agency_code, 
                    upload_date
                ) VALUES (?, CURRENT_TIMESTAMP)
            """, (agency_code,))
            self.conn.commit()
            self.add_log('DATABASE', 'add_pending_pdf', 'success', 
                        f'PDF pendiente agregado: {agency_code}')
            return True
        except Exception as e:
            self.add_log('DATABASE', 'add_pending_pdf', 'error', str(e))
            return False

    def mark_pdf_as_processed(self, agency_code):
        """Marca un PDF como procesado"""
        try:
            self.ensure_connection()
            self.cursor.execute("""
                UPDATE pending_pdfs 
                SET processed = TRUE,
                    processed_date = CURRENT_TIMESTAMP 
                WHERE agency_code = ? 
                AND processed = FALSE
            """, (agency_code,))
            self.conn.commit()
            self.add_log('DATABASE', 'mark_pdf_as_processed', 'success', 
                        f'PDF marcado como procesado: {agency_code}')
            return True
        except Exception as e:
            self.add_log('DATABASE', 'mark_pdf_as_processed', 'error', str(e))
            return False

    def get_client_by_agency_code(self, agency_code):
        """Obtiene un cliente por su código de agencia"""
        try:
            self.ensure_connection()
            query = """
                SELECT id, "Agency Code", "Report email"
                FROM clients 
                WHERE "Agency Code" = ?
            """
            print(f"Ejecutando query: {query} con código: {agency_code}")  # Debug
            
            self.cursor.execute(query, (agency_code,))
            row = self.cursor.fetchone()
            print(f"Resultado de la consulta: {row}")  # Debug
            
            if row:
                return {
                    'id': row[0],
                    'agency_code': row[1],
                    'email': row[2]
                }
            return None
        except Exception as e:
            print(f"Error en get_client_by_agency_code: {str(e)}")  # Debug
            return None

    def update_client_pdf_status(self, agency_code, has_pdf):
        """Actualiza el estado del PDF de un cliente"""
        try:
            self.ensure_connection()
            query = """
                UPDATE clients 
                SET has_pdf = ? 
                WHERE "Agency Code" = ?
            """
            print(f"Ejecutando update: {query}")  # Debug
            print(f"Parámetros: has_pdf={has_pdf}, agency_code={agency_code}")  # Debug
            
            self.cursor.execute(query, (has_pdf, agency_code))
            
            # Verificar si se actualizó algún registro
            rows_affected = self.cursor.rowcount
            print(f"Filas afectadas: {rows_affected}")  # Debug
            
            if rows_affected > 0:
                self.conn.commit()
                self.add_log('DATABASE', 'update_pdf_status', 'success', 
                            f'Estado de PDF actualizado para {agency_code}')
                return True
            else:
                self.add_log('DATABASE', 'update_pdf_status', 'warning', 
                            f'No se encontró cliente para {agency_code}')
                return False
            
        except Exception as e:
            print(f"Error en update_client_pdf_status: {str(e)}")  # Debug
            self.add_log('DATABASE', 'update_pdf_status', 'error', str(e))
            return False
