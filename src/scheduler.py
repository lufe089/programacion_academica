"""
scheduler.py — Motor de programación académica.

Este módulo es responsable de toda la lógica de asignación de sesiones.
Se construye en fases incrementales:

    Fase 1 (iteración 2): filtrado y preparación de candidatos
        - filtrar_asignaturas_del_semestre
        - construir_candidatos

    Fase 2 (iteración 3): asignación básica de sesiones

    Fase 3 (iteración 4): restricciones de tipo y cruce
        - MismaFranja: asignaturas del mismo tipo van juntas en la misma franja
        - NoCruces: asignaturas de ProcesoDesarrollo no pueden coincidir
          entre sí ni con otros tipos

    Reglas de tipo implementadas:
        - Obligatorio y TemasAvanzados: todas las del mismo tipo se asignan
          a la misma franja y se programan juntas cada semana.
        - ProcesoDesarrollo: cada asignatura ocupa un slot único por semana,
          sin coincidir con ninguna otra asignatura de cualquier tipo.
        - SoloMiercoles: asignatura independiente en día distinto.

El DataFrame de candidatos es el puente entre el calendario de slots
disponibles y el motor de asignación. Cada fila representa una combinación
válida (asignatura × fecha × franja) que el scheduler puede usar.
"""

import datetime
import itertools

import pandas as pd

import config
from models import Asignatura, Franja, Parametros, RestriccionProgramacion, SemestreOferta


# ---------------------------------------------------------------------------
# Constantes para la regla 8.15 (TemasAvanzados y encuentros presenciales)
# ---------------------------------------------------------------------------

# Tipo de asignatura que tiene regla especial de encuentros presenciales
_TIPO_TEMAS_AVANZADOS = "TemasAvanzados"

# Horas máximas para sesión intensiva en semana de encuentro presencial
_HORAS_SESION_INTENSIVA = 6.0

# ---------------------------------------------------------------------------
# Constantes para la regla de relleno (distribución de horas pendientes)
# ---------------------------------------------------------------------------

# Horas máximas preferidas por día para la fase de relleno
_HORAS_MAX_VIERNES = 7.0
_HORAS_MAX_SABADO = 6.0
_HORAS_MAX_MIERCOLES = 2.0  # Miércoles tiene franja más corta

# Máximo de horas de una misma asignatura por fin de semana (regla 8.18)
_HORAS_MAX_ASIGNATURA_POR_SEMANA = 7.0

# Máximo de horas de una misma asignatura en un mismo día (regla 8.19)
# Excepción: TemasAvanzados en semana presencial puede llegar a _HORAS_SESION_INTENSIVA
_HORAS_MAX_ASIGNATURA_POR_DIA = 4.0


def _calcular_lunes_de_semana(fecha: datetime.date) -> datetime.date:
    """Retorna el lunes de la semana que contiene la fecha dada."""
    return fecha - datetime.timedelta(days=fecha.weekday())


def _identificar_semanas_presenciales(parametros: Parametros) -> tuple[datetime.date | None, datetime.date | None]:
    """
    Identifica las semanas (lunes) de los encuentros presenciales.

    Args:
        parametros: parámetros de la corrida con fechas presenciales.

    Returns:
        Tupla (lunes_primer_encuentro, lunes_segundo_encuentro).
        Cada elemento puede ser None si la fecha no está definida.
    """
    lunes_primero = None
    lunes_segundo = None

    # Primer encuentro presencial
    fecha_primero = parametros.viernes_presencial_uno or parametros.sabado_presencial_uno
    if fecha_primero:
        lunes_primero = _calcular_lunes_de_semana(fecha_primero)

    # Segundo encuentro presencial
    fecha_segundo = parametros.viernes_presencial_dos or parametros.sabado_presencial_dos
    if fecha_segundo:
        lunes_segundo = _calcular_lunes_de_semana(fecha_segundo)

    return lunes_primero, lunes_segundo


# ---------------------------------------------------------------------------
# Fase 1 — Filtrado y preparación de candidatos
# ---------------------------------------------------------------------------

def filtrar_asignaturas_del_semestre(
    asignaturas: list[Asignatura],
    semestre: str,
) -> list[Asignatura]:
    """
    Filtra las asignaturas que corresponden a la corrida actual.

    Según la regla de inclusión por semestre de oferta:
    - si el semestre es 'Primero', se incluyen las de oferta Primero y Ambos
    - si el semestre es 'Segundo', se incluyen las de oferta Segundo y Ambos

    Args:
        asignaturas: lista completa de asignaturas del catálogo.
        semestre: semestre de la corrida actual ('Primero' o 'Segundo').

    Returns:
        Lista de asignaturas que deben programarse en esta corrida.

    Raises:
        ValueError: si el semestre recibido no es 'Primero' ni 'Segundo'.
    """
    if semestre == config.TEXTO_SEMESTRE_PRIMERO:
        semestres_incluidos = {SemestreOferta.PRIMERO, SemestreOferta.AMBOS}

    elif semestre == config.TEXTO_SEMESTRE_SEGUNDO:
        semestres_incluidos = {SemestreOferta.SEGUNDO, SemestreOferta.AMBOS}

    else:
        raise ValueError(
            f"Semestre desconocido: '{semestre}'. "
            f"Valores esperados: '{config.TEXTO_SEMESTRE_PRIMERO}' o '{config.TEXTO_SEMESTRE_SEGUNDO}'."
        )

    asignaturas_incluidas = []
    for asignatura in asignaturas:
        if asignatura.semestre_oferta in semestres_incluidos:
            asignaturas_incluidas.append(asignatura)

    return asignaturas_incluidas


def construir_candidatos(
    asignaturas: list[Asignatura],
    calendario: pd.DataFrame,
) -> pd.DataFrame:
    """
    Construye el DataFrame de combinaciones válidas (asignatura × slot).

    Para cada asignatura, cruza sus franjas permitidas con los slots
    disponibles del calendario, produciendo una fila por cada combinación
    válida. Solo se incluyen slots marcados como programables.

    Columnas del DataFrame resultante:
        - codigo_darwin: código interno de la asignatura
        - codigo: código académico
        - asignatura: nombre de la asignatura
        - tipo: tipo de la asignatura
        - restriccion_programacion: restricción de programación
        - horas_totales: horas que debe completar la asignatura
        - fecha: fecha del slot
        - dia_semana: día de la semana del slot
        - nombre_franja: identificador de la franja
        - hora_inicio: hora de inicio de la franja
        - hora_fin: hora de fin de la franja
        - duracion_mins: duración del slot en minutos

    Args:
        asignaturas: lista de asignaturas filtradas para la corrida actual.
        calendario: DataFrame de slots del periodo (salida de construir_calendario).

    Returns:
        DataFrame con una fila por cada combinación válida (asignatura × slot).
    """
    slots_disponibles = calendario[calendario[config.CAL_ES_PROGRAMABLE]].copy()

    filas = []
    for asignatura in asignaturas:
        candidatos_de_asignatura = _cruzar_asignatura_con_slots(asignatura, slots_disponibles)
        filas.extend(candidatos_de_asignatura)

    candidatos_df = pd.DataFrame(filas)
    candidatos_df.to_excel("candidatos.xlsx", index=False)
    return candidatos_df


def construir_candidatos_desde_matriz(
    asignaturas: list[Asignatura],
    fechas_con_horas: dict[str, set[datetime.date]],
    franjas: list[Franja],
) -> pd.DataFrame:
    """
    Construye candidatos SOLO para las combinaciones (asignatura, fecha) de la matriz.

    A diferencia de construir_candidatos (que usa todo el calendario), esta función
    genera candidatos únicamente para las fechas donde la matriz tiene horas asignadas,
    respetando las franjas permitidas de cada asignatura.

    Args:
        asignaturas: lista de asignaturas del semestre.
        fechas_con_horas: diccionario {codigo_asignatura: {fechas con horas > 0}}.
        franjas: lista de franjas definidas en el sistema.

    Returns:
        DataFrame de candidatos con las mismas columnas que construir_candidatos.
    """
    # Agrupar franjas por día de la semana
    franjas_por_dia: dict[str, list[Franja]] = {}
    for franja in franjas:
        dia = franja.dia_semana.value  # 'miercoles', 'viernes', 'sabado'
        if dia not in franjas_por_dia:
            franjas_por_dia[dia] = []
        franjas_por_dia[dia].append(franja)

    # Ordenar franjas por hora de inicio dentro de cada día
    for dia in franjas_por_dia:
        franjas_por_dia[dia].sort(key=lambda f: f.hora_inicio)

    # Mapeo weekday -> nombre día
    dias_map = {2: "miercoles", 4: "viernes", 5: "sabado"}

    filas = []
    for asignatura in asignaturas:
        fechas_asignatura = fechas_con_horas.get(asignatura.codigo, set())

        for fecha in fechas_asignatura:
            dia_nombre = dias_map.get(fecha.weekday())
            if not dia_nombre:
                continue

            franjas_del_dia = franjas_por_dia.get(dia_nombre, [])

            # Filtrar solo las franjas permitidas para esta asignatura
            franjas_validas = [
                f for f in franjas_del_dia
                if f.nombre in asignatura.franjas_permitidas
            ]

            for franja in franjas_validas:
                filas.append({
                    "codigo_darwin": asignatura.codigo_darwin,
                    "codigo": asignatura.codigo,
                    "asignatura": asignatura.nombre,
                    "tipo": asignatura.tipo,
                    "restriccion_programacion": asignatura.restriccion_programacion.value,
                    "horas_totales": asignatura.horas_totales,
                    "fecha": fecha,
                    "dia_semana": dia_nombre,
                    "nombre_franja": franja.nombre,
                    "hora_inicio": franja.hora_inicio,
                    "hora_fin": franja.hora_fin,
                    "duracion_mins": franja.duracion_minutos,
                })

    return pd.DataFrame(filas)


def _cruzar_asignatura_con_slots(
    asignatura: Asignatura,
    slots_disponibles: pd.DataFrame,
) -> list[dict]:
    """
    Genera las filas candidatas para una asignatura específica.

    Filtra los slots cuya franja esté en la lista de franjas permitidas
    de la asignatura y construye una fila por cada slot válido.

    Args:
        asignatura: asignatura a cruzar con los slots.
        slots_disponibles: DataFrame con solo los slots programables.

    Returns:
        Lista de diccionarios, uno por cada combinación válida.
    """
    slots_de_la_asignatura = slots_disponibles[
        slots_disponibles[config.CAL_NOMBRE_FRANJA].isin(asignatura.franjas_permitidas)
    ]

    if asignatura.fechas_bloqueadas:
        slots_de_la_asignatura = slots_de_la_asignatura[
            slots_de_la_asignatura[config.CAL_FECHA].apply(
                lambda fecha: not any(
                    ini <= fecha <= fin for ini, fin in asignatura.fechas_bloqueadas
                )
            )
        ]

    filas = []
    for _, slot in slots_de_la_asignatura.iterrows():
        fila = _construir_fila_candidato(asignatura, slot)
        filas.append(fila)

    return filas


def _construir_fila_candidato(asignatura: Asignatura, slot: pd.Series) -> dict:
    """
    Construye el diccionario que representa una fila del DataFrame de candidatos.

    Args:
        asignatura: asignatura del candidato.
        slot: fila del DataFrame de calendario con los datos del slot.

    Returns:
        Diccionario con los datos combinados de asignatura y slot.
    """
    return {
        "codigo_darwin": asignatura.codigo_darwin,
        "codigo": asignatura.codigo,
        "asignatura": asignatura.nombre,
        "tipo": asignatura.tipo,
        "restriccion_programacion": asignatura.restriccion_programacion.value,
        "horas_totales": asignatura.horas_totales,
        "fecha": slot[config.CAL_FECHA],
        "dia_semana": slot[config.CAL_DIA_SEMANA],
        "nombre_franja": slot[config.CAL_NOMBRE_FRANJA],
        "hora_inicio": slot[config.CAL_HORA_INICIO],
        "hora_fin": slot[config.CAL_HORA_FIN],
        "duracion_mins": slot[config.CAL_DURACION_MINS],
    }


# ---------------------------------------------------------------------------
# Fase 2 — Asignación de sesiones con restricciones de tipo
# ---------------------------------------------------------------------------

# Restricciones que implican franja compartida obligatoria dentro del tipo.
_RESTRICCIONES_MISMA_FRANJA = {
    RestriccionProgramacion.MISMA_FRANJA,
    RestriccionProgramacion.OBLIGATORIOS_MISMA_FRANJA,
}


def asignar_sesiones(
    asignaturas: list[Asignatura],
    candidatos: pd.DataFrame,
    franjas: list[Franja],
    parametros: Parametros | None = None,
) -> pd.DataFrame:
    """
    Asigna sesiones respetando las restricciones de tipo y cruce.

    Flujo de asignación por semana:
    1. Grupos MismaFranja (Obligatorio, TemasAvanzados): todas las asignaturas
       del grupo se asignan a la misma franja y se programan juntas.
    2. Asignaturas NoCruces (ProcesoDesarrollo): cada una ocupa su propio slot,
       sin coincidir con ninguna otra asignatura de cualquier tipo.
    3. SoloMiercoles: asignación independiente en día distinto.

    Los grupos MismaFranja se procesan primero porque son más restrictivos
    (deben compartir franja) y así los NoCruces pueden esquivar sus slots.

    Regla especial para TemasAvanzados (regla 8.15):
    - En la semana del primer encuentro presencial: sesión intensiva (5-6 horas)
    - En la semana del segundo encuentro presencial: no se programa clase

    Args:
        asignaturas: lista de asignaturas de la corrida actual.
        candidatos: DataFrame de slots válidos (salida de construir_candidatos).
        franjas: lista de objetos Franja del sistema.
        parametros: parámetros de la corrida (opcional, necesario para regla 8.15).

    Returns:
        DataFrame con una fila por cada sesión asignada, ordenado por fecha.
    """
    franjas_por_nombre = _construir_mapa_franjas(franjas)
    candidatos_con_semana = _agregar_columna_semana(candidatos)
    semanas_ordenadas = sorted(candidatos_con_semana["lunes_semana"].unique())
    horas_acumuladas: dict[str, float] = {a.codigo: 0.0 for a in asignaturas}
    slots_ocupados: set[tuple] = set()  # (fecha, nombre_franja)

    # Rastrear semana de inicio de cada asignatura (regla 8.16 - continuidad)
    semana_inicio: dict[str, datetime.date | None] = {a.codigo: None for a in asignaturas}

    # Identificar semanas de encuentros presenciales (regla 8.15)
    lunes_presencial_uno, lunes_presencial_dos = (None, None)
    if parametros:
        lunes_presencial_uno, lunes_presencial_dos = _identificar_semanas_presenciales(parametros)

    # Separar asignaturas por tipo de restricción
    grupos_misma_franja = _construir_grupos_misma_franja(asignaturas)
    # Ordenar NoCruces por más restringidas primero (menos slots candidatos disponibles)
    # para que las asignaturas con menos opciones tomen sus slots antes que las flexibles.
    conteo_candidatos = candidatos.groupby("codigo").size().to_dict()
    asignaturas_no_cruces = sorted(
        [a for a in asignaturas if a.restriccion_programacion == RestriccionProgramacion.NO_CRUCES],
        key=lambda a: conteo_candidatos.get(a.codigo, 0),
    )
    asignaturas_solo_miercoles = [
        a for a in asignaturas
        if a.restriccion_programacion == RestriccionProgramacion.SOLO_MIERCOLES
    ]

    # Seleccionar la franja común fija para cada grupo MismaFranja
    franjas_por_grupo: dict[str, list[Franja]] = {}
    for tipo, miembros in grupos_misma_franja.items():
        franjas_por_grupo[tipo] = _seleccionar_franja_comun_grupo(miembros, franjas_por_nombre)

    # Seleccionar franja fija para SoloMiercoles (solo una franja disponible)
    franjas_solo_miercoles: dict[str, list[Franja]] = {}
    for asignatura in asignaturas_solo_miercoles:
        max_minutos = (asignatura.horas_totales * 60) / asignatura.min_semanas_clase
        franjas_obj = [
            franjas_por_nombre[n]
            for n in asignatura.franjas_permitidas
            if n in franjas_por_nombre
        ]
        franjas_solo_miercoles[asignatura.codigo] = _seleccionar_mejor_subconjunto(
            franjas_obj, max_minutos, _HORAS_MAX_ASIGNATURA_POR_DIA * 60
        )

    filas_asignadas = []

    for semana in semanas_ordenadas:
        slots_semana = candidatos_con_semana[candidatos_con_semana["lunes_semana"] == semana]

        # 1. Grupos MismaFranja: van primero para reservar sus slots
        for tipo, miembros in grupos_misma_franja.items():
            # Regla 8.15: TemasAvanzados tiene comportamiento especial en semanas presenciales
            if tipo == _TIPO_TEMAS_AVANZADOS:
                # En semana del segundo encuentro presencial: NO programar clase
                if lunes_presencial_dos and semana == lunes_presencial_dos:
                    continue  # Saltar esta semana para TemasAvanzados

                # En semana del primer encuentro presencial: sesión intensiva (5-6 horas)
                if lunes_presencial_uno and semana == lunes_presencial_uno:
                    franjas_sesion = _seleccionar_franjas_sesion_intensiva(
                        miembros, franjas_por_nombre, _HORAS_SESION_INTENSIVA
                    )
                else:
                    franjas_sesion = franjas_por_grupo[tipo]
            else:
                franjas_sesion = franjas_por_grupo[tipo]

            nuevas_filas = _intentar_asignar_grupo(
                miembros, franjas_sesion, slots_semana, horas_acumuladas, slots_ocupados
            )
            filas_asignadas.extend(nuevas_filas)
            # Registrar semana de inicio para miembros del grupo
            if nuevas_filas:
                for miembro in miembros:
                    if semana_inicio[miembro.codigo] is None:
                        semana_inicio[miembro.codigo] = semana

        # 2. NoCruces: priorizar continuidad (regla 8.16)
        # Las asignaturas ya iniciadas tienen prioridad sobre las no iniciadas.
        # Dentro de cada grupo, priorizar por horas pendientes.
        no_cruces_iniciadas = [
            a for a in asignaturas_no_cruces
            if semana_inicio[a.codigo] is not None and horas_acumuladas[a.codigo] < a.horas_totales
        ]
        no_cruces_no_iniciadas = [
            a for a in asignaturas_no_cruces
            if semana_inicio[a.codigo] is None and horas_acumuladas[a.codigo] < a.horas_totales
        ]

        # Ordenar cada grupo por horas pendientes (más pendientes = mayor prioridad)
        no_cruces_iniciadas = sorted(
            no_cruces_iniciadas,
            key=lambda a: -(a.horas_totales - horas_acumuladas[a.codigo]),
        )
        no_cruces_no_iniciadas = sorted(
            no_cruces_no_iniciadas,
            key=lambda a: -(a.horas_totales - horas_acumuladas[a.codigo]),
        )

        # Primero asignar las ya iniciadas (continuidad), luego las nuevas
        for asignatura in no_cruces_iniciadas + no_cruces_no_iniciadas:
            nuevas_filas = _intentar_asignar_no_cruces(
                asignatura, slots_semana, horas_acumuladas, slots_ocupados, franjas_por_nombre
            )
            filas_asignadas.extend(nuevas_filas)
            # Registrar semana de inicio si es la primera asignación
            if nuevas_filas and semana_inicio[asignatura.codigo] is None:
                semana_inicio[asignatura.codigo] = semana

        # 3. SoloMiercoles: día distinto, no interfiere con los demás
        for asignatura in asignaturas_solo_miercoles:
            franjas_sesion = franjas_solo_miercoles[asignatura.codigo]
            nuevas_filas = _intentar_asignar_grupo(
                [asignatura], franjas_sesion, slots_semana, horas_acumuladas, slots_ocupados
            )
            filas_asignadas.extend(nuevas_filas)
            if nuevas_filas and semana_inicio[asignatura.codigo] is None:
                semana_inicio[asignatura.codigo] = semana

    # --- Fase de relleno (regla 8.17): distribuir horas pendientes ---
    asignaturas_pendientes = [
        a for a in asignaturas
        if horas_acumuladas[a.codigo] < a.horas_totales
    ]

    if asignaturas_pendientes:
        nuevas_filas_relleno = _fase_relleno(
            asignaturas_pendientes,
            candidatos_con_semana,
            horas_acumuladas,
            slots_ocupados,
            franjas_por_nombre,
            filas_asignadas,
        )
        filas_asignadas.extend(nuevas_filas_relleno)

    if not filas_asignadas:
        return pd.DataFrame()

    df_sesiones = pd.DataFrame(filas_asignadas)
    df_sesiones = df_sesiones.sort_values(by=["fecha", "hora_inicio", "codigo"])
    df_sesiones = df_sesiones.reset_index(drop=True)

    return df_sesiones


def seleccionar_sesion_semanal(
    asignatura: Asignatura,
    franjas_por_nombre: dict[str, Franja],
) -> list[Franja]:
    """
    Determina el conjunto de franjas que forma la sesión semanal de una asignatura.

    Usa el criterio de máximo de minutos por semana derivado de min_semanas_clase.
    Delega en _seleccionar_mejor_subconjunto para elegir la mejor combinación.

    Args:
        asignatura: asignatura para la que se selecciona la sesión.
        franjas_por_nombre: diccionario nombre → Franja.

    Returns:
        Lista de objetos Franja que componen la sesión semanal.
    """
    franjas_disponibles = [
        franjas_por_nombre[nombre]
        for nombre in asignatura.franjas_permitidas
        if nombre in franjas_por_nombre
    ]

    if not franjas_disponibles:
        return []

    max_minutos = (asignatura.horas_totales * 60) / asignatura.min_semanas_clase

    return _seleccionar_mejor_subconjunto(
        franjas_disponibles, max_minutos, _HORAS_MAX_ASIGNATURA_POR_DIA * 60
    )


# ---------------------------------------------------------------------------
# Funciones auxiliares — agrupación y selección de franjas
# ---------------------------------------------------------------------------

def _construir_mapa_franjas(franjas: list[Franja]) -> dict[str, Franja]:
    """Construye un diccionario nombre → Franja para resolución rápida."""
    mapa = {}
    for franja in franjas:
        mapa[franja.nombre] = franja
    return mapa


def _agregar_columna_semana(candidatos: pd.DataFrame) -> pd.DataFrame:
    """
    Agrega la columna 'lunes_semana' al DataFrame de candidatos.

    Cada fecha se mapea al lunes de su semana, usado como clave de agrupación.
    """
    resultado = candidatos.copy()
    resultado["lunes_semana"] = resultado[config.CAL_FECHA].apply(
        lambda fecha: fecha - datetime.timedelta(days=fecha.weekday())
    )
    return resultado


def _construir_grupos_misma_franja(
    asignaturas: list[Asignatura],
) -> dict[str, list[Asignatura]]:
    """
    Agrupa las asignaturas con restricción MismaFranja por tipo.

    Cada grupo comparte exactamente la misma franja durante toda la
    programación. Los grupos se keyan por tipo (ej. 'Obligatorio').

    Args:
        asignaturas: lista completa de asignaturas de la corrida.

    Returns:
        Diccionario tipo → lista de asignaturas del grupo.
    """
    grupos: dict[str, list[Asignatura]] = {}
    for asignatura in asignaturas:
        if asignatura.restriccion_programacion in _RESTRICCIONES_MISMA_FRANJA:
            tipo = asignatura.tipo
            if tipo not in grupos:
                grupos[tipo] = []
            grupos[tipo].append(asignatura)
    return grupos


def _seleccionar_franja_comun_grupo(
    miembros: list[Asignatura],
    franjas_por_nombre: dict[str, Franja],
) -> list[Franja]:
    """
    Selecciona el conjunto de franjas compartido por todos los miembros del grupo.

    Busca la intersección de las franjas permitidas de todos los miembros y
    elige el mejor subconjunto usando el criterio del miembro más restrictivo
    (el que tiene menos minutos por semana disponibles).

    Args:
        miembros: asignaturas del grupo que deben compartir franja.
        franjas_por_nombre: diccionario nombre → Franja.

    Returns:
        Lista de Franja que todos los miembros usarán cada semana.
    """
    # Intersección de franjas permitidas de todos los miembros
    conjuntos_franjas = [set(a.franjas_permitidas) for a in miembros]
    nombres_comunes = conjuntos_franjas[0].intersection(*conjuntos_franjas[1:])

    franjas_comunes = [
        franjas_por_nombre[nombre]
        for nombre in nombres_comunes
        if nombre in franjas_por_nombre
    ]

    if not franjas_comunes:
        # Si no hay intersección, usar las franjas del primer miembro
        franjas_comunes = [
            franjas_por_nombre[n]
            for n in miembros[0].franjas_permitidas
            if n in franjas_por_nombre
        ]

    # El miembro más restrictivo define el techo de minutos por semana
    max_minutos = min(
        (a.horas_totales * 60) / a.min_semanas_clase
        for a in miembros
    )

    # Aplicar límite de 6 horas por asignatura por semana (regla 8.18)
    max_minutos = min(max_minutos, _HORAS_MAX_ASIGNATURA_POR_SEMANA * 60)

    return _seleccionar_mejor_subconjunto(
        franjas_comunes, max_minutos, _HORAS_MAX_ASIGNATURA_POR_DIA * 60
    )


def _seleccionar_franjas_sesion_intensiva(
    miembros: list[Asignatura],
    franjas_por_nombre: dict[str, Franja],
    horas_intensivas: float,
) -> list[Franja]:
    """
    Selecciona franjas para una sesión intensiva de TemasAvanzados.

    Similar a _seleccionar_franja_comun_grupo pero con un límite de horas
    mayor (sesión intensiva de 5-6 horas en semana de encuentro presencial).

    Args:
        miembros: asignaturas del grupo que deben compartir franja.
        franjas_por_nombre: diccionario nombre → Franja.
        horas_intensivas: número máximo de horas para la sesión intensiva.

    Returns:
        Lista de Franja para la sesión intensiva.
    """
    # Intersección de franjas permitidas de todos los miembros
    conjuntos_franjas = [set(a.franjas_permitidas) for a in miembros]
    nombres_comunes = conjuntos_franjas[0].intersection(*conjuntos_franjas[1:])

    franjas_comunes = [
        franjas_por_nombre[nombre]
        for nombre in nombres_comunes
        if nombre in franjas_por_nombre
    ]

    if not franjas_comunes:
        franjas_comunes = [
            franjas_por_nombre[n]
            for n in miembros[0].franjas_permitidas
            if n in franjas_por_nombre
        ]

    # Usar el límite de horas intensivas (convertido a minutos)
    max_minutos = horas_intensivas * 60

    return _seleccionar_mejor_subconjunto(franjas_comunes, max_minutos)


def _seleccionar_mejor_subconjunto(
    franjas: list[Franja],
    max_minutos: float,
    max_minutos_por_dia: float | None = None,
) -> list[Franja]:
    """
    Elige el subconjunto de franjas con mayor duración total sin superar max_minutos.

    Evalúa todos los subconjuntos no vacíos. Si ninguno cabe bajo el máximo,
    retorna la franja de menor duración como fallback.

    Args:
        franjas: franjas candidatas a incluir en el subconjunto.
        max_minutos: límite superior de duración total en minutos.
        max_minutos_por_dia: si se especifica, ningún día puede acumular más
            de este número de minutos dentro del subconjunto seleccionado.

    Returns:
        Lista de Franja que forma el mejor subconjunto.
    """
    mejor: list[Franja] = []
    mejor_duracion = 0

    for tamanho in range(1, len(franjas) + 1):
        for combinacion in itertools.combinations(franjas, tamanho):
            duracion = sum(f.duracion_minutos for f in combinacion)
            if duracion > max_minutos:
                continue
            if max_minutos_por_dia is not None:
                mins_por_dia: dict[str, int] = {}
                for f in combinacion:
                    dia = f.dia_semana.value
                    mins_por_dia[dia] = mins_por_dia.get(dia, 0) + f.duracion_minutos
                if any(m > max_minutos_por_dia for m in mins_por_dia.values()):
                    continue
            if duracion > mejor_duracion:
                mejor = list(combinacion)
                mejor_duracion = duracion

    if not mejor and franjas:
        mejor = [min(franjas, key=lambda f: f.duracion_minutos)]

    return mejor


# ---------------------------------------------------------------------------
# Funciones auxiliares — asignación por semana
# ---------------------------------------------------------------------------

def _intentar_asignar_grupo(
    miembros: list[Asignatura],
    franjas_sesion: list[Franja],
    slots_semana: pd.DataFrame,
    horas_acumuladas: dict[str, float],
    slots_ocupados: set[tuple],
) -> list[dict]:
    """
    Intenta asignar la sesión semanal a todos los miembros activos del grupo.

    Todos los miembros del grupo comparten exactamente la misma (fecha, franja).
    Si el slot está ocupado por otro grupo o por cualquier otra asignatura,
    se omite la semana completa para el grupo.

    Solo se asigna a los miembros que aún tienen horas pendientes.
    Si algún miembro activo excedería sus horas con esta sesión, se omite.

    Modifica horas_acumuladas y slots_ocupados si la asignación es exitosa.

    Args:
        miembros: asignaturas del grupo.
        franjas_sesion: franjas fijas del grupo.
        slots_semana: slots disponibles esta semana en el DataFrame de candidatos.
        horas_acumuladas: acumulador de horas, modificado en lugar.
        slots_ocupados: conjunto de (fecha, franja) ocupados, modificado en lugar.

    Returns:
        Lista de filas asignadas, o vacía si no fue posible asignar.
    """
    miembros_activos = [
        m for m in miembros
        if horas_acumuladas[m.codigo] < m.horas_totales
    ]

    if not miembros_activos:
        return []

    # Usar el primer miembro activo para obtener las fechas de los slots
    referencia = miembros_activos[0]
    slots_referencia = slots_semana[slots_semana["codigo"] == referencia.codigo]

    # Verificar que todas las franjas estén disponibles y no ocupadas
    slots_por_franja: dict[str, pd.Series] = {}
    for franja in franjas_sesion:
        filas_franja = slots_referencia[slots_referencia["nombre_franja"] == franja.nombre]
        if filas_franja.empty:
            return []
        slot = filas_franja.iloc[0]
        fecha = slot[config.CAL_FECHA]
        if (fecha, franja.nombre) in slots_ocupados:
            return []
        slots_por_franja[franja.nombre] = slot

    # Verificar que ningún miembro activo excedería sus horas.
    # El miembro con menos horas pendientes define el techo de la sesión.
    horas_sesion_total = sum(f.duracion_minutos for f in franjas_sesion) / 60
    min_pendientes = min(
        m.horas_totales - horas_acumuladas[m.codigo] for m in miembros_activos
    )
    if horas_sesion_total > min_pendientes:
        return []

    # Marcar los slots como ocupados
    for franja in franjas_sesion:
        fecha = slots_por_franja[franja.nombre][config.CAL_FECHA]
        slots_ocupados.add((fecha, franja.nombre))

    # Asignar a cada miembro activo
    filas = []
    for franja in franjas_sesion:
        slot = slots_por_franja[franja.nombre]
        horas_esta_franja = franja.duracion_minutos / 60

        for miembro in miembros_activos:
            horas_acumuladas[miembro.codigo] += horas_esta_franja
            filas.append(_construir_fila_sesion(miembro, slot, franja, horas_esta_franja, horas_acumuladas))

    return filas


def _intentar_asignar_no_cruces(
    asignatura: Asignatura,
    slots_semana: pd.DataFrame,
    horas_acumuladas: dict[str, float],
    slots_ocupados: set[tuple],
    franjas_por_nombre: dict[str, Franja],
) -> list[dict]:
    """
    Intenta asignar una sesión a una asignatura con restricción NoCruces.

    A diferencia de los grupos MismaFranja, selecciona dinámicamente el mejor
    subconjunto de franjas disponibles (no ocupadas) en esta semana. Esto le
    permite adaptarse a los slots que ya tomaron los grupos MismaFranja y otras
    asignaturas NoCruces procesadas antes en la misma semana.

    Modifica horas_acumuladas y slots_ocupados si la asignación es exitosa.

    Args:
        asignatura: asignatura NoCruces a programar.
        slots_semana: slots disponibles esta semana.
        horas_acumuladas: acumulador de horas, modificado en lugar.
        slots_ocupados: conjunto de slots ya tomados, modificado en lugar.
        franjas_por_nombre: diccionario nombre → Franja.

    Returns:
        Lista de filas asignadas, o vacía si no fue posible asignar.
    """
    if horas_acumuladas[asignatura.codigo] >= asignatura.horas_totales:
        return []

    slots_asignatura = slots_semana[slots_semana["codigo"] == asignatura.codigo]

    # Filtrar a las franjas disponibles y no ocupadas esta semana
    franjas_libres: list[Franja] = []
    slots_por_franja: dict[str, pd.Series] = {}

    for _, slot in slots_asignatura.iterrows():
        nombre_franja = slot[config.CAL_NOMBRE_FRANJA]
        fecha = slot[config.CAL_FECHA]
        if (fecha, nombre_franja) not in slots_ocupados:
            franja_obj = franjas_por_nombre.get(nombre_franja)
            if franja_obj:
                franjas_libres.append(franja_obj)
                slots_por_franja[nombre_franja] = slot

    if not franjas_libres:
        return []

    # Seleccionar el mejor subconjunto de franjas libres
    # Aplicar límite de 6 horas por asignatura por semana (regla 8.18)
    # y límite de 4 horas por asignatura por día (regla 8.19)
    max_minutos = (asignatura.horas_totales * 60) / asignatura.min_semanas_clase
    max_minutos = min(max_minutos, _HORAS_MAX_ASIGNATURA_POR_SEMANA * 60)
    franjas_sesion = _seleccionar_mejor_subconjunto(
        franjas_libres, max_minutos, _HORAS_MAX_ASIGNATURA_POR_DIA * 60
    )

    if not franjas_sesion:
        return []

    # Verificar que no excede horas pendientes; si la selección inicial excede,
    # reintentar con el tope ajustado a las horas que quedan.
    horas_sesion = sum(f.duracion_minutos for f in franjas_sesion) / 60
    horas_pendientes = asignatura.horas_totales - horas_acumuladas[asignatura.codigo]

    if horas_sesion > horas_pendientes:
        franjas_sesion = _seleccionar_mejor_subconjunto(
            franjas_libres, horas_pendientes * 60, _HORAS_MAX_ASIGNATURA_POR_DIA * 60
        )
        if not franjas_sesion:
            return []
        horas_sesion = sum(f.duracion_minutos for f in franjas_sesion) / 60
        if horas_sesion > horas_pendientes:
            return []

    # Marcar slots como ocupados y registrar sesiones
    filas = []
    for franja in franjas_sesion:
        slot = slots_por_franja[franja.nombre]
        fecha = slot[config.CAL_FECHA]
        slots_ocupados.add((fecha, franja.nombre))
        horas_esta_franja = franja.duracion_minutos / 60
        horas_acumuladas[asignatura.codigo] += horas_esta_franja
        filas.append(_construir_fila_sesion(asignatura, slot, franja, horas_esta_franja, horas_acumuladas))

    return filas


def _construir_fila_sesion(
    asignatura: Asignatura,
    slot: pd.Series,
    franja: Franja,
    horas_esta_franja: float,
    horas_acumuladas: dict[str, float],
) -> dict:
    """
    Construye el diccionario que representa una sesión asignada.

    Args:
        asignatura: asignatura de la sesión.
        slot: fila del DataFrame de candidatos con los datos del slot.
        franja: franja de la sesión.
        horas_esta_franja: horas que aporta esta franja.
        horas_acumuladas: estado actual del acumulador.

    Returns:
        Diccionario con los datos de la sesión.
    """
    return {
        "codigo": asignatura.codigo,
        "asignatura": asignatura.nombre,
        "tipo": asignatura.tipo,
        "fecha": slot[config.CAL_FECHA],
        "dia_semana": slot[config.CAL_DIA_SEMANA],
        "nombre_franja": franja.nombre,
        "hora_inicio": franja.hora_inicio,
        "hora_fin": franja.hora_fin,
        "duracion_mins": franja.duracion_minutos,
        "horas_sesion": horas_esta_franja,
        "horas_acumuladas": horas_acumuladas[asignatura.codigo],
    }


# ---------------------------------------------------------------------------
# Fase de relleno — distribución de horas pendientes (regla 8.17)
# ---------------------------------------------------------------------------

def _fase_relleno(
    asignaturas_pendientes: list[Asignatura],
    candidatos: pd.DataFrame,
    horas_acumuladas: dict[str, float],
    slots_ocupados: set[tuple],
    franjas_por_nombre: dict[str, Franja],
    sesiones_existentes: list[dict],
) -> list[dict]:
    """
    Distribuye horas pendientes en fechas subutilizadas.

    Busca fechas donde las horas efectivas no superan el máximo
    y asigna sesiones adicionales para completar las horas de
    asignaturas que quedaron incompletas.

    Args:
        asignaturas_pendientes: asignaturas con horas faltantes.
        candidatos: DataFrame de slots candidatos.
        horas_acumuladas: acumulador de horas, modificado en lugar.
        slots_ocupados: conjunto de slots ocupados, modificado en lugar.
        franjas_por_nombre: diccionario nombre → Franja.
        sesiones_existentes: sesiones ya asignadas para calcular carga por fecha.

    Returns:
        Lista de nuevas filas de sesiones asignadas en fase de relleno.
    """
    # Calcular horas efectivas por fecha de las sesiones existentes
    horas_por_fecha = _calcular_horas_efectivas_por_fecha(sesiones_existentes)

    # Calcular horas por asignatura por semana (para regla 8.18 de máx 6h por semana)
    horas_por_asignatura_semana = _calcular_horas_por_asignatura_semana(sesiones_existentes)

    # Calcular horas por asignatura por fecha (para regla 8.19 de máx 4h por día)
    horas_por_asignatura_fecha = _calcular_horas_por_asignatura_fecha(sesiones_existentes)

    # Obtener todas las fechas únicas del calendario ordenadas
    fechas_ordenadas = sorted(candidatos["fecha"].unique())

    filas_nuevas = []

    # Ordenar asignaturas pendientes por más horas faltantes primero
    asignaturas_ordenadas = sorted(
        asignaturas_pendientes,
        key=lambda a: -(a.horas_totales - horas_acumuladas[a.codigo])
    )

    for asignatura in asignaturas_ordenadas:
        horas_pendientes = asignatura.horas_totales - horas_acumuladas[asignatura.codigo]

        if horas_pendientes <= 0:
            continue

        # Buscar fechas subutilizadas donde esta asignatura puede programarse
        slots_asignatura = candidatos[candidatos["codigo"] == asignatura.codigo]

        for fecha in fechas_ordenadas:
            if horas_pendientes <= 0:
                break

            # Obtener slots de esta asignatura en esta fecha
            slots_fecha = slots_asignatura[slots_asignatura["fecha"] == fecha]
            if slots_fecha.empty:
                continue

            # Determinar máximo de horas para este día
            dia_semana = slots_fecha.iloc[0]["dia_semana"]
            max_horas = _obtener_max_horas_dia(dia_semana)

            horas_actuales = horas_por_fecha.get(fecha, 0.0)
            espacio_disponible = max_horas - horas_actuales

            if espacio_disponible <= 0:
                continue

            # Verificar límite de 4 horas por asignatura por día (regla 8.19)
            clave_asig_fecha = (asignatura.codigo, fecha)
            horas_asig_en_fecha = horas_por_asignatura_fecha.get(clave_asig_fecha, 0.0)
            espacio_asignatura_dia = _HORAS_MAX_ASIGNATURA_POR_DIA - horas_asig_en_fecha

            if espacio_asignatura_dia <= 0:
                continue

            # Calcular semana (lunes) para verificar límite de 6 horas por asignatura
            lunes_semana = _calcular_lunes_de_semana(fecha)
            clave_asig_semana = (asignatura.codigo, lunes_semana)
            horas_asig_en_semana = horas_por_asignatura_semana.get(clave_asig_semana, 0.0)
            espacio_asignatura_semana = _HORAS_MAX_ASIGNATURA_POR_SEMANA - horas_asig_en_semana

            if espacio_asignatura_semana <= 0:
                continue

            # Buscar franjas libres para esta asignatura en esta fecha
            for _, slot in slots_fecha.iterrows():
                if (horas_pendientes <= 0 or espacio_disponible <= 0
                        or espacio_asignatura_semana <= 0 or espacio_asignatura_dia <= 0):
                    break

                nombre_franja = slot["nombre_franja"]
                if (fecha, nombre_franja) in slots_ocupados:
                    continue

                franja = franjas_por_nombre.get(nombre_franja)
                if not franja:
                    continue

                horas_franja = franja.duracion_minutos / 60

                # Solo asignar si cabe en todos los límites activos
                if (horas_franja <= espacio_disponible and
                    horas_franja <= horas_pendientes and
                    horas_franja <= espacio_asignatura_semana and
                    horas_franja <= espacio_asignatura_dia):
                    slots_ocupados.add((fecha, nombre_franja))
                    horas_acumuladas[asignatura.codigo] += horas_franja
                    horas_por_fecha[fecha] = horas_por_fecha.get(fecha, 0.0) + horas_franja
                    horas_por_asignatura_semana[clave_asig_semana] = horas_asig_en_semana + horas_franja
                    horas_por_asignatura_fecha[clave_asig_fecha] = horas_asig_en_fecha + horas_franja

                    filas_nuevas.append(_construir_fila_sesion(
                        asignatura, slot, franja, horas_franja, horas_acumuladas
                    ))

                    horas_pendientes -= horas_franja
                    espacio_disponible -= horas_franja
                    espacio_asignatura_semana -= horas_franja
                    espacio_asignatura_dia -= horas_franja
                    horas_asig_en_fecha += horas_franja
                    horas_asig_en_semana += horas_franja

    return filas_nuevas


def _calcular_horas_por_asignatura_fecha(sesiones: list[dict]) -> dict[tuple, float]:
    """
    Calcula las horas por asignatura por fecha.

    Args:
        sesiones: lista de diccionarios de sesiones asignadas.

    Returns:
        Diccionario (codigo, fecha) → horas acumuladas ese día.
    """
    resultado: dict[tuple, float] = {}
    for sesion in sesiones:
        clave = (sesion["codigo"], sesion["fecha"])
        resultado[clave] = resultado.get(clave, 0.0) + sesion["horas_sesion"]
    return resultado


def _calcular_horas_por_asignatura_semana(sesiones: list[dict]) -> dict[tuple, float]:
    """
    Calcula las horas por asignatura por semana.

    Args:
        sesiones: lista de diccionarios de sesiones asignadas.

    Returns:
        Diccionario (codigo, lunes_semana) → horas acumuladas.
    """
    resultado: dict[tuple, float] = {}

    for sesion in sesiones:
        codigo = sesion["codigo"]
        fecha = sesion["fecha"]
        horas = sesion["horas_sesion"]
        lunes = _calcular_lunes_de_semana(fecha)

        clave = (codigo, lunes)
        resultado[clave] = resultado.get(clave, 0.0) + horas

    return resultado


def _calcular_horas_efectivas_por_fecha(sesiones: list[dict]) -> dict[datetime.date, float]:
    """
    Calcula las horas efectivas por fecha de las sesiones existentes.

    Las horas efectivas cuentan cada franja una sola vez, sin importar
    cuántas asignaturas la compartan.

    Args:
        sesiones: lista de diccionarios de sesiones asignadas.

    Returns:
        Diccionario fecha → horas efectivas.
    """
    # Agrupar por (fecha, franja) para evitar contar duplicados
    franjas_vistas: dict[datetime.date, set[str]] = {}
    horas_por_slot: dict[tuple, float] = {}

    for sesion in sesiones:
        fecha = sesion["fecha"]
        franja = sesion["nombre_franja"]
        horas = sesion["horas_sesion"]

        if fecha not in franjas_vistas:
            franjas_vistas[fecha] = set()

        franjas_vistas[fecha].add(franja)
        horas_por_slot[(fecha, franja)] = horas

    # Sumar horas únicas por fecha
    resultado: dict[datetime.date, float] = {}
    for fecha, franjas in franjas_vistas.items():
        total = sum(horas_por_slot.get((fecha, f), 0.0) for f in franjas)
        resultado[fecha] = total

    return resultado


def _obtener_max_horas_dia(dia_semana: str) -> float:
    """
    Retorna el máximo de horas preferido para un día de la semana.

    Args:
        dia_semana: nombre del día ('viernes', 'sabado', 'miercoles').

    Returns:
        Máximo de horas para ese día.
    """
    if dia_semana == "viernes":
        return _HORAS_MAX_VIERNES
    elif dia_semana == "sabado":
        return _HORAS_MAX_SABADO
    elif dia_semana == "miercoles":
        return _HORAS_MAX_MIERCOLES
    else:
        return _HORAS_MAX_VIERNES  # Default


