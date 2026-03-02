import pyodbc
import logging
import urllib
from sqlalchemy import create_engine
from Config.Settings import DB_CONFIG

logger = logging.getLogger(__name__)

class Database:
    def __init__(self):
        # Datos desde tu configuración
        self.user = DB_CONFIG.get('user')
        self.password = DB_CONFIG.get('password')
        self.host = DB_CONFIG.get('host')
        self.db = DB_CONFIG.get('Database')
        self._engine = None # Cache para el engine

    def get_connection(self):
        """
        Retorna una conexión pyodbc tradicional. 
        Ideal para: execute("UPDATE..."), call stored procedures.
        """
        try:
            conn_str = (
                f"DRIVER={{ODBC Driver 17 for SQL Server}};"
                f"SERVER={self.host};"
                f"Database={self.db};"
                f"UID={self.user};"
                f"PWD={self.password};"
                "TrustServerCertificate=yes;"
            )
            return pyodbc.connect(conn_str)
        
        except Exception:
            logger.error("Error conectando a SQL Server vía pyodbc", exc_info=True)
            raise

    def get_engine(self):
        """
        Retorna el engine de SQLAlchemy.
        Ideal para: pd.to_sql() y pd.read_sql().
        """
        if self._engine is None:
            try:
                # Escapar la contraseña para evitar errores con caracteres especiales
                safe_password = urllib.parse.quote_plus(self.password)

                # Cadena de conexión estándar para SQL Server
                conn_str = (
                    f"mssql+pyodbc://{self.user}:{safe_password}@{self.host}/{self.db}"
                    "?driver=ODBC+Driver+17+for+SQL+Server"
                )
                
                # fast_executemany=True optimiza las inserciones de Pandas
                self._engine = create_engine(conn_str, fast_executemany=True)
                
                logger.info("Engine de SQLAlchemy creado exitosamente.")
                
            except Exception:
                logger.error("Error creando el engine de SQLAlchemy", exc_info=True)
                raise
        
        return self._engine




# import pyodbc
# import logging
# from sqlalchemy import create_engine, event

# from Config.Settings import DB_CONFIG
# logger = logging.getLogger(__name__)

# class Database:
#     def __init__(self):
#         # Datos desde tu in_config
#         self.user = DB_CONFIG.get('user')
#         self.password = DB_CONFIG.get('password')
#         self.host = DB_CONFIG.get('host')
#         self.db = DB_CONFIG.get('Database')
#     #@staticmethod
#     def get_connection(self):
#         """
#         Abre conexión bajo demanda.
#         El cierre se maneja con 'with'.
#         """
#         try:
#             conn = pyodbc.connect(
#                 f"DRIVER={{ODBC Driver 17 for SQL Server}};"
#                 f"SERVER= {self.host};"
#                 f"Database={self.db};"
#                 f"UID={self.user};"
#                 f"PWD={self.password};"
#                 "TrustServerCertificate=yes;"
#             )
#             return conn

#         except Exception:
#             logger.error("Error conectando a SQL Server", exc_info=True)
#             raise
#     @staticmethod
#     def get_engine(self):        
#         # Cadena de conexión estándar para SQL Server
#         conn_str = f"mssql+pyodbc://{self.user}:{self.password}@{self.host}/{self.db}?driver=ODBC+Driver+17+for+SQL+Server"
#         # fast_executemany=True es el SECRETO para que las inserciones sean 100x más rápidas
#         self.engine = create_engine(conn_str, fast_executemany=True)
#         return self.engine
    
