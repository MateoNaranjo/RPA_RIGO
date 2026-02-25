from Config.Database import Database

class CorreosRepo:

    
    @staticmethod
    def ObtenerParametrosCorreo(cod_email: int):

        query = """"
            SELECT * FROM PagoArriendos.ParametrosCorreo WHERE CodEmailParamter = ?
        """

        with Database.get_connection() as conn:
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