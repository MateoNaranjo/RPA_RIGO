from Config.Database import Database

class ControlHURepository:

    @staticmethod
    def upsert_control_hu(hu_id: int, nombre_hu: str, estado: int, activa: int, maquina: str):
        """
        Inserta o actualiza la HU
        """
        query = """
        MERGE PagoArriendos.ControlHU as target
        USING (SELECT ? AS HU) AS source
        ON target.hu = source.HU
        WHEN MATCHED THEN
            UPDATE SET
                Estado = ?,
                Activa = ?,
                Maquina = ?,
                FechaActualizacion = SYSDATETIME()
        WHEN NOT MATCHED THEN
            INSERT (HU, NombreHU, Estado, Activa, Maquina)
            VALUES (?, ?, ?, ?, ?);
        """

        with Database.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                query,
                (
                    hu_id,
                    estado,
                    activa,
                    maquina,
                    hu_id,
                    nombre_hu,
                    estado,
                    activa,
                    maquina
                )
            )
            conn.commit()