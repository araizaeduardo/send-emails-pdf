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
        self.conn = sqlite3.connect(self.db_name)
        self.cursor = self.conn.cursor()

    def close(self):
        """Cierra la conexión con la base de datos"""
        if self.conn:
            self.conn.close()

    def import_from_excel(self, excel_path: str):
        """
        Importa datos desde Excel a SQLite
        Args:
            excel_path (str): Ruta al archivo Excel
        """
        self.connect()  # Conectar al inicio
        try:
            # Leer Excel
            df = pd.read_excel(excel_path)
            
            # Asegurarse de que existan las columnas necesarias
            required_columns = ['Agency Code', 'Report email']
            if not all(col in df.columns for col in required_columns):
                raise ValueError("El archivo Excel debe contener las columnas 'Agency Code' y 'Report email'")
            
            # Agregar columnas para tracking
            df['email_sent'] = 0
            df['sent_date'] = None
            
            # Crear tablas si no existen
            self.setup_database()
            
            # Importar datos
            df.to_sql('clients', self.conn, if_exists='replace', index=True, index_label='id')
            self.conn.commit()
            
            # Agregar log
            self.add_log('SYSTEM', 'import_excel', 'success', f'Datos importados exitosamente desde {excel_path}')
            
            return True, "Datos importados exitosamente"
        except Exception as e:
            error_msg = f"Error al importar datos: {str(e)}"
            if self.conn:  # Solo agregar log si hay conexión
                self.add_log('SYSTEM', 'import_excel', 'error', error_msg)
            return False, error_msg
        finally:
            if self.conn:
                self.close()

    def get_pending_clients(self) -> List[Dict]:
        """
        Obtiene la lista de clientes pendientes por enviar correo
        Returns:
            List[Dict]: Lista de diccionarios con datos de clientes
        """
        try:
            self.connect()
            self.cursor.execute("SELECT * FROM clients WHERE email_sent = 0 OR email_sent IS NULL")
            columns = [description[0] for description in self.cursor.description]
            clients = []
            for row in self.cursor.fetchall():
                clients.append(dict(zip(columns, row)))
            return clients
        finally:
            self.close()

    def mark_email_sent(self, client_id: int):
        """
        Marca un correo como enviado en la base de datos
        Args:
            client_id (int): ID del cliente
        """
        try:
            self.connect()
            self.cursor.execute("""
                UPDATE clients 
                SET email_sent = 1,
                    sent_date = CURRENT_TIMESTAMP
                WHERE id = ?
            """, (client_id,))
            self.conn.commit()
        finally:
            self.close()

    def get_sent_emails(self) -> List[Dict]:
        """
        Obtiene la lista de correos ya enviados
        Returns:
            List[Dict]: Lista de diccionarios con datos de correos enviados
        """
        try:
            self.connect()
            self.cursor.execute("SELECT * FROM clients WHERE email_sent = 1")
            columns = [description[0] for description in self.cursor.description]
            clients = []
            for row in self.cursor.fetchall():
                clients.append(dict(zip(columns, row)))
            return clients
        finally:
            self.close()

    def get_client(self, client_id: int) -> Dict:
        """
        Obtiene los datos de un cliente específico
        Args:
            client_id (int): ID del cliente
        Returns:
            Dict: Datos del cliente
        """
        try:
            self.connect()
            self.cursor.execute("SELECT * FROM clients WHERE id = ?", (client_id,))
            columns = [description[0] for description in self.cursor.description]
            row = self.cursor.fetchone()
            if row:
                return dict(zip(columns, row))
            return None
        finally:
            self.close()

    def clear_all_records(self):
        """
        Limpia todos los registros de la base de datos
        """
        try:
            self.connect()
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
            self.connect()
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
            self.connect()
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
            self.connect()
            self.cursor.execute("DELETE FROM clients")
            self.conn.commit()
        finally:
            self.close()

    def setup_database(self):
        """
        Configura las tablas necesarias en la base de datos
        """
        # No conectamos aquí, asumimos que ya hay una conexión activa
        # Tabla de clientes
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS clients (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                "Agency Code" TEXT,
                "Report email" TEXT,
                email_sent INTEGER DEFAULT 0,
                sent_date TIMESTAMP
            )
        """)
        
        # Tabla de logs
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS email_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                agency_code TEXT,
                action TEXT,
                status TEXT,
                message TEXT,
                timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        self.conn.commit()

    def add_log(self, agency_code: str, action: str, status: str, message: str):
        """
        Agrega un registro al log
        """
        try:
            self.connect()
            self.cursor.execute("""
                INSERT INTO email_logs (agency_code, action, status, message)
                VALUES (?, ?, ?, ?)
            """, (agency_code, action, status, message))
            self.conn.commit()
        finally:
            self.close()

    def get_logs(self, limit: int = 100):
        """
        Obtiene los últimos registros del log
        """
        try:
            self.connect()
            self.cursor.execute("""
                SELECT * FROM email_logs
                ORDER BY timestamp DESC
                LIMIT ?
            """, (limit,))
            columns = [description[0] for description in self.cursor.description]
            logs = []
            for row in self.cursor.fetchall():
                logs.append(dict(zip(columns, row)))
            return logs
        finally:
            self.close()

    def clear_logs(self):
        """
        Limpia todos los logs
        """
        try:
            self.connect()
            self.cursor.execute("DELETE FROM email_logs")
            self.conn.commit()
        finally:
            self.close()
