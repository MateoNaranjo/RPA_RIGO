import pyodbc
import logging
from config.settings import DB_CONFIG
logger = logging.getLogger(__name__)

class Database:
    @staticmethod
    def get_connection():
        """
        Abre conexión bajo demanda.
        El cierre se maneja con 'with'.
        """
        try:
            conn = pyodbc.connect(
                f"DRIVER={{ODBC Driver 17 for SQL Server}};"
                f"SERVER= {DB_CONFIG.get('host')};"
                f"DATABASE={DB_CONFIG.get('database')};"
                f"UID={DB_CONFIG.get('user')};"
                f"PWD={DB_CONFIG.get('password')};"
                "TrustServerCertificate=yes;"
            )
            return conn

        except Exception:
            logger.error("Error conectando a SQL Server", exc_info=True)
            raise