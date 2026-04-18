"""
models.py — Estructuras de datos del sistema de programación académica.

Contiene los enums y dataclasses que representan el dominio del problema.
Estos modelos son contenedores de datos puros: no contienen lógica de negocio.
"""

from dataclasses import dataclass, field
from datetime import date, time
from enum import Enum


class DiaSemana(Enum):
    """Días de la semana en los que se puede programar clase."""
    MIERCOLES = "miercoles"
    VIERNES = "viernes"
    SABADO = "sabado"


class SemestreOferta(Enum):
    """Indica en qué semestre se ofrece una asignatura."""
    PRIMERO = "Primero"
    SEGUNDO = "Segundo"
    AMBOS = "Ambos"


class RestriccionProgramacion(Enum):
    """
    Tipos de restricción de programación que puede tener una asignatura.

    - MISMA_FRANJA: todas las asignaturas del mismo tipo deben compartir
      exactamente la misma franja horaria.
    - OBLIGATORIOS_MISMA_FRANJA: variante de MISMA_FRANJA aplicada a
      asignaturas obligatorias.
    - NO_CRUCES: la asignatura no puede programarse al mismo tiempo que
      otras asignaturas del mismo grupo.
    - SOLO_MIERCOLES: la asignatura solo puede programarse en franjas
      de miércoles.
    - SIN_RESTRICCION: no tiene restricción de programación específica.
    """
    MISMA_FRANJA = "MismaFranja"
    OBLIGATORIOS_MISMA_FRANJA = "ObligatoriosMismaFranja"
    NO_CRUCES = "NoCruces"
    SOLO_MIERCOLES = "SoloMiercoles"
    SIN_RESTRICCION = "SinRestriccion"


@dataclass
class Franja:
    """
    Representa una franja horaria definida en la hoja 'franjas' del Excel.

    Atributos:
        nombre: identificador único de la franja (ej. 'FRANJA_UNO_VIERNES').
        hora_inicio: hora en que comienza la franja.
        hora_fin: hora en que termina la franja.
        duracion_minutos: duración total de la franja en minutos.
        dia_semana: día de la semana al que pertenece la franja.
    """
    nombre: str
    hora_inicio: time
    hora_fin: time
    duracion_minutos: int
    dia_semana: DiaSemana


@dataclass
class Asignatura:
    """
    Representa una asignatura del catálogo con todas sus restricciones.

    Atributos:
        codigo_darwin: código numérico interno del sistema Darwin.
        codigo: código académico de la asignatura (ej. '400CIS020').
        nombre: nombre completo de la asignatura.
        semestre_oferta: semestre(s) en que se ofrece la asignatura.
        tipo: categoría de la asignatura (ej. 'Obligatorio', 'TemasAvanzados').
        restriccion_programacion: restricción principal de programación.
        franjas_permitidas: lista de nombres de franjas en las que puede
            programarse esta asignatura.
        creditos: número de créditos de la asignatura.
        horas_totales: cantidad total de horas que debe completar en el semestre.
        restricciones_texto: texto libre con restricciones adicionales.
            Puede ser None si no hay restricciones adicionales.
    """
    codigo_darwin: int
    codigo: str
    nombre: str
    semestre_oferta: SemestreOferta
    tipo: str
    restriccion_programacion: RestriccionProgramacion
    franjas_permitidas: list = field(default_factory=list)
    creditos: int = 0
    horas_totales: int = 0
    min_semanas_clase: int = 0
    fechas_bloqueadas: list[tuple[date, date]] = field(default_factory=list)


@dataclass
class Parametros:
    """
    Parámetros generales de una corrida de programación académica.

    Se leen desde la hoja 'parametros' del Excel y definen el periodo
    y las condiciones bajo las cuales se construye el calendario.

    Atributos:
        semestre_programacion: semestre que se está programando ('Primero' o 'Segundo').
        fecha_induccion: fecha de inducción del semestre.
        inicio_clases: fecha en que comienzan las clases.
        inicio_semana_sin_clases: primer día de la semana sin clases (receso).
        fin_semana_sin_clases: último día de la semana sin clases (receso).
        festivos: lista de fechas festivas que deben excluirse del calendario.
        fin_clases: fecha en que terminan las clases. Puede ser None si
            aún no está definida en el Excel.
        viernes_presencial_uno: primera fecha de viernes presencial. Opcional.
        sabado_presencial_uno: primera fecha de sábado presencial. Opcional.
        viernes_presencial_dos: segunda fecha de viernes presencial. Opcional.
        sabado_presencial_dos: segunda fecha de sábado presencial. Opcional.
    """
    semestre_programacion: str
    fecha_induccion: date
    inicio_clases: date
    inicio_semana_sin_clases: date
    fin_semana_sin_clases: date
    festivos: list = field(default_factory=list)
    fin_clases: date | None = None
    viernes_presencial_uno: date | None = None
    sabado_presencial_uno: date | None = None
    viernes_presencial_dos: date | None = None
    sabado_presencial_dos: date | None = None
    semana_inicio: int = 1
