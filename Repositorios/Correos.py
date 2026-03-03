from turtle import pd

from sqlalchemy import create_engine

from Config.Database import Database
from Config.init_config import in_config

class CorreosRepo:

    
    @staticmethod
    def ObtenerParametrosCorreo(cod_email: int):

        query = """"
            SELECT * FROM PagoArriendos.ParametrosCorreo WHERE CodEmailParamter = ?
        """
        db = Database()

        with db.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, cod_email)
            fila = cursor.fetchone()

            if not fila:
                raise ValueError(f"No existe configuracion de correo para el codigo: {cod_email}")
            
            return {
                "to": fila.TOEmailParameter,
                "cc": fila.CCEmailParameter,
                "bcc": fila.BCCEmailParameter,
                "asunto": fila.AsuntoEmailParameter,
                "body": fila.BodyEmailParameter,
                "is_html": bool(fila.IsHTMLEmailParameter)
            }
        
    # def leer_sql(self, tabla_o_query: str) -> pd.DataFrame:
    #     """
    #     Lee la estructura de correos desde una base de datos SQL.
    #     """
    #     try:
    #         # 1. Configurar la cadena de conexión (Ajusta según tu motor: SQL Server, MySQL, etc.)
    #         # Ejemplo para SQL Server (usando autenticación de Windows):
    #         # conn_str = "mssql+pyodbc://SERVIDOR/BASE_DATOS?driver=ODBC+Driver+17+for+SQL+Server&trusted_connection=yes"
            
    #         # Obtenemos la cadena desde tu configuración para mantener el orden
    #         conn_str = in_config("ConnectionStringSQL") 
    #         engine = create_engine(conn_str)

    #         # 2. Ejecutar la lectura
    #         # Si la entrada empieza con SELECT, lo tratamos como query, si no, como tabla
    #         if "SELECT" in tabla_o_query.upper():
    #             df = pd.read_sql(tabla_o_query, engine)
    #         else:
    #             df = pd.read_sql_table(tabla_o_query, engine)

    #         # 3. Limpieza estándar
    #         df.columns = df.columns.str.strip()
            
    #         WriteLog(mensaje=f"Datos recuperados exitosamente de SQL: {tabla_o_query}", estado="INFO", nombreTarea="EmailSender")
    #         return df

    #     except Exception as e:
    #         WriteLog(mensaje=f"Error al leer desde SQL: {e}", estado="ERROR", nombreTarea="EmailSender")
    #         return None