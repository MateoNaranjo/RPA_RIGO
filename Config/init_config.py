from Repositorios.Parametros import ParametrosRepository

_CONFIG_CACHE = None


def init_config():
    """Carga configuración desde BD una sola vez"""
    global _CONFIG_CACHE

    if _CONFIG_CACHE is not None:
        return

    _CONFIG_CACHE = ParametrosRepository.cargar_parametros()


def in_config(nombre, default=None):
    if _CONFIG_CACHE is None:
        raise RuntimeError(
            "Configuración no inicializada. "
            "Ejecute HU00_DespliegueAmbiente antes de usar in_config()"
        )

    return _CONFIG_CACHE.get(nombre, default)


                
