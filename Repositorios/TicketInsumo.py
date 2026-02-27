from datetime import datetime
from Config.Database import Database

class GestionTicketInsumo:

    def __init__(self, conn):
        self.conn = conn or Database.get_connection()

    def obtener_por_codigo(self, codigo):
        query ="""
            SELECT *
            FROM PagoArriendos.TicketInsumo
            WHERE Codigo = ?
        """
        with self.conn(dictionary=True) as cursor:
            cursor.execute(query, (codigo,))
            cursor.close()
            return cursor.fetchone()
        
    def crear(self, codigo: str, maquina: str):
        query = """
            INSERT INTO PagoArriendos.TicketInsumo
            (Codigo, fechainsercion, estado, numeroreintentos, maquina)
            VALUES (?, ?, ?, ?, ?)
        """
        with self.conn.cursor() as cursor:
            cursor.execute(
                query,
                (codigo, datetime.now(), "PENDIENTE", 0, maquina)
            )

    def actualizar_estado(
        self,
        conn,
        codigo,
        estado,
        observaciones=None,
        incrementar_reintento=False,
        finalizar=False
    ):
        query = """
            UPDATE PagoArriendos.TicketInsumo
            SET estado = %s,
                observaciones = %s,
                numeroreintentos = numeroreintentos + %s,
                fechamodificacion = %s,
                fechafin = %s
            WHERE Codigo = %s
        """

        fechafin = datetime.now() if finalizar else None
        incremento = 1 if incrementar_reintento else 0

        with self.conn.cursor() as cursor:
            cursor.execute(
                query,
                (
                    estado,
                    observaciones,
                    incremento,
                    datetime.now(),
                    fechafin,
                    codigo
                )
            )