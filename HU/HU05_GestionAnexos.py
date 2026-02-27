import pandas as pd
from datetime import datetime
from Config.Database import Database

class HU05_CargueSQL:
    
    @staticmethod
    def crear_tabla_nueva():
        """
        Crea la tabla Reporte_Final_Arriendos con las columnas exactas de tu imagen.
        """
        tabla = "PagoArriendos.Reporte_Final_Arriendos"
        
        query = f"""
        IF OBJECT_ID('{tabla}', 'U') IS NOT NULL
            DROP TABLE {tabla};

        CREATE TABLE {tabla} (
            ID INT IDENTITY(1,1) PRIMARY KEY,
            OC VARCHAR(100),
            Estado_FAC VARCHAR(100),
            Tiene_HES VARCHAR(50),
            Diagnostico_Cierre VARCHAR(255),
            Responsable_Accion VARCHAR(255),
            Accion_Sugerida VARCHAR(MAX),
            Fecha_Analisis VARCHAR(50),
            Fecha_Cargue_DB DATETIME DEFAULT GETDATE()
        );
        """
        try:
            with Database.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query)
                conn.commit()
                print(f"[+] Tabla {tabla} configurada con los títulos de la HU03.")
                return True
        except Exception as e:
            print(f"[-] Error creando tabla: {e}")
            return False

    @staticmethod
    def ejecutar_cargue_desde_excel(ruta_excel):
        if not HU05_CargueSQL.crear_tabla_nueva():
            return
        
        try:
            df = pd.read_excel(ruta_excel)
            
            # Query ajustada a los títulos de tu imagen
            query_insert = """
                INSERT INTO PagoArriendos.Reporte_Final_Arriendos (
                    OC, Estado_FAC, Tiene_HES, Diagnostico_Cierre, 
                    Responsable_Accion, Accion_Sugerida, Fecha_Analisis
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """
            
            with Database.get_connection() as conn:
                cursor = conn.cursor()
                
                for _, fila in df.iterrows():
                    # Mapeamos los nombres exactos de las columnas de tu imagen
                    valores = (
                        str(fila.get('OC')),
                        fila.get('Estado_FAC'),
                        fila.get('Tiene HES'),
                        fila.get('Diagnóstico de Cierre'),
                        fila.get('Responsable Acción'),
                        fila.get('Acción Sugerida'),
                        str(fila.get('Fecha_Analisis'))
                    )
                    cursor.execute(query_insert, valores)
                
                conn.commit()
                print(f"[+] Éxito: Se cargaron {len(df)} registros a SQL Server.")

        except Exception as e:
            print(f"[-] Error en el cargue: {e}")