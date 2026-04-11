"""
exports_hours.py — Exportación de la versión de horas de la programación.

Genera un archivo Excel con dos hojas:

    - 'programacion': detalle sesión por sesión, ordenado por fecha y hora.
      Representa la versión de horas de la programación académica.

    - 'resumen': estado de cumplimiento de horas por asignatura.
      Identifica asignaturas completas, con faltante o con exceso.

Este módulo solo se ocupa de la presentación y formato del output.
No contiene lógica de negocio ni de asignación.
"""

import datetime

import pandas as pd

from models import Asignatura


# ---------------------------------------------------------------------------
# Función principal de exportación
# ---------------------------------------------------------------------------

def exportar_version_horas(
    sesiones: pd.DataFrame,
    asignaturas: list[Asignatura],
    ruta_salida: str,
    inicio_clases: datetime.date,
) -> None:
    """
    Exporta la programación de sesiones al formato de versión de horas.

    Genera un archivo Excel con dos hojas: 'programacion' con el detalle
    sesión por sesión, y 'resumen' con el estado de horas por asignatura.

    Args:
        sesiones: DataFrame de sesiones asignadas (salida de asignar_sesiones).
        asignaturas: lista de asignaturas de la corrida actual.
        ruta_salida: ruta completa del archivo Excel a generar.
        inicio_clases: fecha de inicio de clases, usada para calcular
            el número de semana relativo al semestre.
    """
    df_programacion = _construir_hoja_programacion(sesiones, inicio_clases)
    df_resumen = _construir_hoja_resumen(sesiones, asignaturas)

    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        df_programacion.to_excel(writer, sheet_name="programacion", index=False)
        df_resumen.to_excel(writer, sheet_name="resumen", index=False)


# ---------------------------------------------------------------------------
# Construcción de la hoja 'programacion'
# ---------------------------------------------------------------------------

def _construir_hoja_programacion(
    sesiones: pd.DataFrame,
    inicio_clases: datetime.date,
) -> pd.DataFrame:
    """
    Construye el DataFrame de la hoja 'programacion' a partir de las sesiones.

    Ordena por fecha y hora de inicio, agrega el número de semana del semestre
    y renombra las columnas al español para facilitar la lectura.

    Args:
        sesiones: DataFrame de sesiones asignadas.
        inicio_clases: fecha de inicio de clases para calcular semana relativa.

    Returns:
        DataFrame listo para exportar como hoja 'programacion'.
    """
    if sesiones.empty:
        return pd.DataFrame()

    df = sesiones.copy()

    df["semana"] = df["fecha"].apply(
        lambda fecha: calcular_numero_semana(fecha, inicio_clases)
    )

    df = df.sort_values(by=["fecha", "hora_inicio", "codigo"]).reset_index(drop=True)

    df["observaciones"] = ""

    df_exportar = pd.DataFrame({
        "Semana": df["semana"],
        "Fecha": df["fecha"],
        "Día": df["dia_semana"],
        "Asignatura": df["asignatura"],
        "Código": df["codigo"],
        "Franja": df["nombre_franja"],
        "Hora inicio": df["hora_inicio"],
        "Hora fin": df["hora_fin"],
        "Duración (min)": df["duracion_mins"],
        "Horas sesión": df["horas_sesion"],
        "Horas acumuladas": df["horas_acumuladas"],
        "Observaciones": df["observaciones"],
    })

    return df_exportar


def calcular_numero_semana(fecha: datetime.date, inicio_clases: datetime.date) -> int:
    """
    Calcula el número de semana relativo al inicio del semestre.

    La semana 1 es la que contiene la fecha de inicio de clases.

    Args:
        fecha: fecha de la sesión.
        inicio_clases: fecha de inicio de clases del semestre.

    Returns:
        Número de semana (1, 2, 3, ...).
    """
    lunes_inicio = inicio_clases - datetime.timedelta(days=inicio_clases.weekday())
    lunes_fecha = fecha - datetime.timedelta(days=fecha.weekday())
    diferencia_dias = (lunes_fecha - lunes_inicio).days
    return (diferencia_dias // 7) + 1


# ---------------------------------------------------------------------------
# Construcción de la hoja 'resumen'
# ---------------------------------------------------------------------------

def _construir_hoja_resumen(
    sesiones: pd.DataFrame,
    asignaturas: list[Asignatura],
) -> pd.DataFrame:
    """
    Construye el DataFrame de la hoja 'resumen' con el estado por asignatura.

    Para cada asignatura indica cuántas horas se asignaron, cuántas faltan
    o sobran, y cuántas semanas de clase tuvo.

    Args:
        sesiones: DataFrame de sesiones asignadas.
        asignaturas: lista de asignaturas de la corrida actual.

    Returns:
        DataFrame listo para exportar como hoja 'resumen'.
    """
    filas = []

    for asignatura in asignaturas:
        sesiones_asignatura = sesiones[sesiones["codigo"] == asignatura.codigo]

        horas_asignadas = sesiones_asignatura["horas_sesion"].sum()
        semanas_con_clase = sesiones_asignatura["fecha"].nunique()
        horas_faltantes = asignatura.horas_totales - horas_asignadas
        horas_excedentes = max(0.0, horas_asignadas - asignatura.horas_totales)

        if horas_faltantes > 0:
            estado = f"Incompleta — faltan {horas_faltantes:.1f}h"
        elif horas_excedentes > 0:
            estado = f"Excede — sobran {horas_excedentes:.1f}h"
        else:
            estado = "Completa"

        filas.append({
            "Código": asignatura.codigo,
            "Asignatura": asignatura.nombre,
            "Tipo": asignatura.tipo,
            "Horas objetivo": asignatura.horas_totales,
            "Horas asignadas": round(horas_asignadas, 1),
            "Horas faltantes": round(max(0.0, horas_faltantes), 1),
            "Semanas con clase": semanas_con_clase,
            "Min semanas requeridas": asignatura.min_semanas_clase,
            "Estado": estado,
        })

    return pd.DataFrame(filas)
