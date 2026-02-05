from Config.Database import Database

class Excel:

    @staticmethod
    def construir_columnas(columnas: dict) -> str:
        return ",\n".join(
            f"{nombre} {tipo} NULL"
            for nombre, tipo in columnas.items()
        )

    # -----------------------------
    # CREAR TABLA TEMPORAL
    # -----------------------------
    @staticmethod
    def crear_tabla_temp(tabla: str, columnas: dict) -> bool:
        tabla_temp = f"{tabla}Temp"
        columnas_sql = Excel.construir_columnas(columnas)

        query = f"""
        IF OBJECT_ID('PagoArriendos.{tabla_temp}', 'U') IS NOT NULL
            DROP TABLE PagoArriendos.{tabla_temp};

        CREATE TABLE PagoArriendos.{tabla_temp} (
            {columnas_sql}
        );
        """

        try:
            with Database.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query)
                conn.commit()
                cursor.close()
            return True

        except Exception as e:
            print(f"Error al crear tabla temporal {tabla_temp}: {e}")
            return False

    # -----------------------------
    # CREAR TABLA FINAL
    # -----------------------------
    @staticmethod
    def crear_tabla_final(tabla: str, columnas: dict) -> bool:
        columnas_sql = Excel.construir_columnas(columnas)

        query = f"""
        IF OBJECT_ID('PagoArriendos.{tabla}', 'U') IS NOT NULL
            DROP TABLE PagoArriendos.{tabla};

        CREATE TABLE PagoArriendos.{tabla} (
            {columnas_sql},
            Estado VARCHAR(100) NOT NULL DEFAULT 'Pendiente'
        );
        """

        try:
            with Database.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query)
                conn.commit()
                cursor.close()
            return True

        except Exception as e:
            print(f"Error al crear tabla final {tabla}: {e}")
            return False

    # -----------------------------
    # BULK + TRANSFERENCIA
    # -----------------------------
    @staticmethod
    def ejecutar_bulk(ruta_txt: str, tabla: str, columnas: dict):

        tabla_temp = f"{tabla}Temp"

        if not Excel.crear_tabla_temp(tabla, columnas):
            return

        bulk_query = f"""
        BULK INSERT PagoArriendos.{tabla_temp}
        FROM '{ruta_txt}'
        WITH (
            FIRSTROW = 2,
            FIELDTERMINATOR = ';',
            ROWTERMINATOR = '0x0a',
            CODEPAGE = '65001',
            TABLOCK
        );
        """

        columnas_sql = ", ".join(columnas.keys())

        insert_query = f"""
        INSERT INTO PagoArriendos.{tabla} ({columnas_sql}, Estado)
        SELECT {columnas_sql}, 'Pendiente'
        FROM PagoArriendos.{tabla_temp};
        """

        drop_temp_query = f"""
        DROP TABLE PagoArriendos.{tabla_temp};
        """

        try:
            with Database.get_connection() as conn:
                conn.autocommit = True
                cursor = conn.cursor()

                Excel.crear_tabla_final(tabla, columnas)
                cursor.execute(bulk_query)
                cursor.execute(insert_query)
                cursor.execute(drop_temp_query)

                cursor.close()

            print("Carga BULK ejecutada correctamente")

        except Exception as e:
            print(f"Error durante la ejecución del BULK: {e}")

    # -----------------------------
    # OBTENER SOLO PENDIENTES
    # -----------------------------
    @staticmethod
    def obtener_valores(tabla: str):

        query = f"""
        SELECT TOP 2 *
        FROM PagoArriendos.{tabla}
        WHERE Estado = 'Pendiente'
        """

        with Database.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query)

            columnas = [col[0] for col in cursor.description]
            rows = cursor.fetchall()
            cursor.close()

            return [dict(zip(columnas, fila)) for fila in rows]
        
    @staticmethod
    def obtener_datos_por_posicion(tabla: str):
        # Corregimos los nombres de las columnas según tu tabla real
        query = f"""
        SELECT TOP 10 *
        FROM PagoArriendos.{tabla}
        WHERE EstadoRegistro = 'Pendiente'
        """

        try:
            with Database.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query)

                columnas = [col[0] for col in cursor.description]
                rows = cursor.fetchall()
                cursor.close()

                # Esto retorna una lista de diccionarios con las llaves correctas
                return [dict(zip(columnas, fila)) for fila in rows]
        except Exception as e:
            print(f"Error al consultar la tabla {tabla}: {e}")
            return []
        

