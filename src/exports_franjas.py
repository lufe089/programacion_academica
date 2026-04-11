"""
exports_franjas.py — Exportación de la versión de franjas de la programación.

Genera un archivo Excel con la hoja 'franjas' que consolida las sesiones
en bloques continuos por asignatura y franja horaria.

Cada fila del resultado representa un bloque continuo (sin semanas vacías de por
medio) en que una asignatura mantuvo la misma franja horaria. Si una asignatura
usa la misma franja en semanas 1-8 y luego vuelve a usarla en semanas 11-14,
se generan dos filas separadas.

Este módulo solo se ocupa de la presentación y formato del output.
No contiene lógica de negocio ni de asignación.
"""

import datetime

import pandas as pd

from exports_hours import calcular_numero_semana
from models import Asignatura


# ---------------------------------------------------------------------------
# Función principal de exportación
# ---------------------------------------------------------------------------

def exportar_version_franjas(
    sesiones: pd.DataFrame,
    asignaturas: list[Asignatura],
    ruta_salida: str,
    inicio_clases: datetime.date,
) -> None:
    """
    Exporta la programación de sesiones al formato de versión de franjas.

    Genera un archivo Excel con la hoja 'franjas' que muestra bloques
    continuos de programación, agrupados por asignatura y franja horaria.

    Args:
        sesiones: DataFrame de sesiones asignadas (salida de asignar_sesiones).
        asignaturas: lista de asignaturas de la corrida actual.
        ruta_salida: ruta completa del archivo Excel a generar.
        inicio_clases: fecha de inicio de clases, usada para calcular
            el número de semana relativo al semestre.
    """
    df_franjas = _construir_hoja_franjas(sesiones, inicio_clases)

    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        df_franjas.to_excel(writer, sheet_name="franjas", index=False)


# ---------------------------------------------------------------------------
# Construcción de la hoja 'franjas'
# ---------------------------------------------------------------------------

def _construir_hoja_franjas(
    sesiones: pd.DataFrame,
    inicio_clases: datetime.date,
) -> pd.DataFrame:
    """
    Construye el DataFrame de la hoja 'franjas' a partir de las sesiones.

    Para cada par (asignatura, franja), agrupa las semanas donde aparece
    en bloques consecutivos y produce una fila por bloque.

    Args:
        sesiones: DataFrame de sesiones asignadas.
        inicio_clases: fecha de inicio de clases para calcular semana relativa.

    Returns:
        DataFrame listo para exportar como hoja 'franjas'.
    """
    if sesiones.empty:
        return pd.DataFrame()

    df = sesiones.copy()
    df["semana"] = df["fecha"].apply(
        lambda fecha: calcular_numero_semana(fecha, inicio_clases)
    )

    bloques = []

    for (codigo, nombre_franja), grupo in df.groupby(["codigo", "nombre_franja"], sort=False):
        grupo_ordenado = grupo.sort_values("semana")
        nuevos_bloques = _extraer_bloques_continuos(grupo_ordenado)
        bloques.extend(nuevos_bloques)

    if not bloques:
        return pd.DataFrame()

    df_resultado = pd.DataFrame(bloques)
    df_resultado = df_resultado.sort_values(
        ["Código", "Semana inicio", "Franja"]
    ).reset_index(drop=True)

    return df_resultado


def _extraer_bloques_continuos(grupo: pd.DataFrame) -> list[dict]:
    """
    Divide las sesiones de un par (asignatura, franja) en bloques de semanas consecutivas.

    Dos semanas son consecutivas si sus números difieren en exactamente 1.
    Cualquier brecha mayor produce un corte y genera un nuevo bloque.

    Args:
        grupo: sesiones de una misma (asignatura, franja), ordenadas por semana.

    Returns:
        Lista de diccionarios, uno por cada bloque continuo encontrado.
    """
    semanas = sorted(grupo["semana"].unique())

    bloques = []
    inicio_bloque = semanas[0]
    fin_bloque = semanas[0]

    for semana in semanas[1:]:
        if semana == fin_bloque + 1:
            fin_bloque = semana
        else:
            sesiones_bloque = grupo[grupo["semana"].between(inicio_bloque, fin_bloque)]
            bloques.append(_construir_fila_bloque(sesiones_bloque, inicio_bloque, fin_bloque))
            inicio_bloque = semana
            fin_bloque = semana

    sesiones_bloque = grupo[grupo["semana"].between(inicio_bloque, fin_bloque)]
    bloques.append(_construir_fila_bloque(sesiones_bloque, inicio_bloque, fin_bloque))

    return bloques


def _construir_fila_bloque(
    sesiones_bloque: pd.DataFrame,
    semana_inicio: int,
    semana_fin: int,
) -> dict:
    """
    Construye el diccionario que representa un bloque continuo de franjas.

    Args:
        sesiones_bloque: sesiones que forman el bloque.
        semana_inicio: número de semana en que comienza el bloque.
        semana_fin: número de semana en que termina el bloque.

    Returns:
        Diccionario con los datos del bloque.
    """
    sesiones_ordenadas = sesiones_bloque.sort_values("fecha")
    primera = sesiones_ordenadas.iloc[0]
    ultima = sesiones_ordenadas.iloc[-1]

    return {
        "Código": primera["codigo"],
        "Asignatura": primera["asignatura"],
        "Tipo": primera["tipo"],
        "Franja": primera["nombre_franja"],
        "Día": primera["dia_semana"],
        "Hora inicio": primera["hora_inicio"],
        "Hora fin": primera["hora_fin"],
        "Semana inicio": semana_inicio,
        "Fecha inicio": primera["fecha"],
        "Semana fin": semana_fin,
        "Fecha fin": ultima["fecha"],
        "Sesiones": len(sesiones_bloque),
        "Horas en bloque": round(float(sesiones_bloque["horas_sesion"].sum()), 1),
    }
