"""
calendar_builder.py — Construcción del calendario de slots programables.

Toma los parámetros de la corrida y las franjas definidas, y produce
un DataFrame con todas las combinaciones (fecha × franja) del periodo,
indicando cuáles están disponibles para programar y cuáles están bloqueadas.

Cada fila del DataFrame representa un slot: una fecha específica en una
franja específica. El scheduler usará este DataFrame como universo de
slots disponibles para asignar sesiones.
"""

import datetime

import pandas as pd

import config
from models import DiaSemana, Franja, Parametros


# Mapeo de número de weekday de Python al enum DiaSemana.
# datetime.weekday(): 0=lunes, 1=martes, 2=miércoles, 3=jueves, 4=viernes, 5=sábado
NUMERO_DIA_A_DIA_SEMANA = {
    2: DiaSemana.MIERCOLES,
    4: DiaSemana.VIERNES,
    5: DiaSemana.SABADO,
}


def construir_calendario(parametros: Parametros, franjas: list[Franja]) -> pd.DataFrame:
    """
    Construye el DataFrame de slots disponibles para el periodo académico.

    Cada fila representa una combinación (fecha, franja) e indica si ese
    slot está disponible para programar o está bloqueado, junto con el
    motivo del bloqueo si aplica.

    Columnas del DataFrame resultante:
        - fecha: objeto datetime.date
        - dia_semana: nombre del día (ej. 'viernes')
        - nombre_franja: identificador de la franja (ej. 'FRANJA_UNO_VIERNES')
        - hora_inicio: hora de inicio de la franja
        - hora_fin: hora de fin de la franja
        - duracion_mins: duración en minutos
        - es_programable: True si el slot está disponible
        - motivo_bloqueo: descripción del bloqueo, o None si está disponible

    Args:
        parametros: parámetros de la corrida con fechas del periodo.
        franjas: lista de franjas definidas en la hoja 'franjas'.

    Returns:
        DataFrame con una fila por cada slot (fecha × franja) del periodo.

    Raises:
        ValueError: si fin_clases no está definido en los parámetros.
    """
    if parametros.fin_clases is None:
        raise ValueError(
            "El parámetro 'FIN_CLASES' es requerido para construir el calendario "
            "pero no está definido en la hoja 'parametros'."
        )

    fechas_del_periodo = _generar_fechas_del_periodo(
        parametros.inicio_clases,
        parametros.fin_clases,
    )

    fechas_programables = _filtrar_dias_con_clase(fechas_del_periodo)
    franjas_por_dia = _agrupar_franjas_por_dia(franjas)
    fechas_bloqueadas = _construir_conjunto_fechas_bloqueadas(parametros)

    filas = []
    for fecha in fechas_programables:
        dia_semana = NUMERO_DIA_A_DIA_SEMANA[fecha.weekday()]
        franjas_del_dia = franjas_por_dia.get(dia_semana, [])

        es_programable, motivo_bloqueo = _evaluar_disponibilidad(fecha, fechas_bloqueadas)

        for franja in franjas_del_dia:
            fila = _construir_fila(fecha, dia_semana, franja, es_programable, motivo_bloqueo)
            filas.append(fila)

    return pd.DataFrame(filas)


def _generar_fechas_del_periodo(
    inicio: datetime.date,
    fin: datetime.date,
) -> list[datetime.date]:
    """
    Genera la lista de todas las fechas entre inicio y fin, inclusive.

    Args:
        inicio: primera fecha del periodo.
        fin: última fecha del periodo.

    Returns:
        Lista de fechas ordenadas de menor a mayor.
    """
    fechas = []
    fecha_actual = inicio
    while fecha_actual <= fin:
        fechas.append(fecha_actual)
        fecha_actual += datetime.timedelta(days=1)
    return fechas


def _filtrar_dias_con_clase(fechas: list[datetime.date]) -> list[datetime.date]:
    """
    Filtra la lista dejando solo los días en que puede haber clase:
    miércoles, viernes y sábado.

    Args:
        fechas: lista completa de fechas del periodo.

    Returns:
        Lista con solo las fechas que son miércoles, viernes o sábado.
    """
    dias_validos = set(NUMERO_DIA_A_DIA_SEMANA.keys())
    return [fecha for fecha in fechas if fecha.weekday() in dias_validos]


def _agrupar_franjas_por_dia(franjas: list[Franja]) -> dict[DiaSemana, list[Franja]]:
    """
    Agrupa las franjas por día de la semana para facilitar su consulta.

    Args:
        franjas: lista de todas las franjas del sistema.

    Returns:
        Diccionario que mapea cada DiaSemana a su lista de franjas,
        ordenadas por hora de inicio.
    """
    agrupadas: dict[DiaSemana, list[Franja]] = {}

    for franja in franjas:
        dia = franja.dia_semana
        if dia not in agrupadas:
            agrupadas[dia] = []
        agrupadas[dia].append(franja)

    for dia in agrupadas:
        agrupadas[dia].sort(key=lambda franja: franja.hora_inicio)

    return agrupadas


def _construir_conjunto_fechas_bloqueadas(parametros: Parametros) -> dict[datetime.date, str]:
    """
    Construye un diccionario con todas las fechas bloqueadas y su motivo.

    Incluye festivos y todos los días de la semana sin clases.

    Args:
        parametros: parámetros de la corrida.

    Returns:
        Diccionario que mapea fecha bloqueada → motivo de bloqueo.
    """
    fechas_bloqueadas = {}

    for festivo in parametros.festivos:
        fechas_bloqueadas[festivo] = config.MOTIVO_FESTIVO

    fecha_actual = parametros.inicio_semana_sin_clases
    while fecha_actual <= parametros.fin_semana_sin_clases:
        fechas_bloqueadas[fecha_actual] = config.MOTIVO_SEMANA_SIN_CLASES
        fecha_actual += datetime.timedelta(days=1)

    return fechas_bloqueadas


def _evaluar_disponibilidad(
    fecha: datetime.date,
    fechas_bloqueadas: dict[datetime.date, str],
) -> tuple[bool, str | None]:
    """
    Determina si una fecha está disponible para programar.

    Args:
        fecha: fecha a evaluar.
        fechas_bloqueadas: diccionario de fechas bloqueadas con su motivo.

    Returns:
        Tupla (es_programable, motivo_bloqueo).
        Si la fecha está disponible, retorna (True, None).
        Si está bloqueada, retorna (False, motivo).
    """
    if fecha in fechas_bloqueadas:
        return False, fechas_bloqueadas[fecha]

    return True, None


def _construir_fila(
    fecha: datetime.date,
    dia_semana: DiaSemana,
    franja: Franja,
    es_programable: bool,
    motivo_bloqueo: str | None,
) -> dict:
    """
    Construye el diccionario que representa una fila del DataFrame de calendario.

    Args:
        fecha: fecha del slot.
        dia_semana: día de la semana del slot.
        franja: franja horaria del slot.
        es_programable: si el slot está disponible para programar.
        motivo_bloqueo: motivo por el que está bloqueado, o None.

    Returns:
        Diccionario con los datos del slot.
    """
    return {
        config.CAL_FECHA: fecha,
        config.CAL_DIA_SEMANA: dia_semana.value,
        config.CAL_NOMBRE_FRANJA: franja.nombre,
        config.CAL_HORA_INICIO: franja.hora_inicio,
        config.CAL_HORA_FIN: franja.hora_fin,
        config.CAL_DURACION_MINS: franja.duracion_minutos,
        config.CAL_ES_PROGRAMABLE: es_programable,
        config.CAL_MOTIVO_BLOQUEO: motivo_bloqueo,
    }
