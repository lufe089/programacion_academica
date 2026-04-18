"""
config.py — Constantes de configuración del sistema de programación académica.

Centraliza todos los nombres de hojas, columnas y claves que el sistema
usa para leer el Excel. Si la estructura del archivo cambia, este es
el único lugar que debe modificarse.

No contiene lógica de negocio ni importa módulos del proyecto.
"""

# ---------------------------------------------------------------------------
# Nombres de las hojas del Excel
# ---------------------------------------------------------------------------

NOMBRE_HOJA_CATALOGO = "catalogo"
NOMBRE_HOJA_PARAMETROS = "parametros"
NOMBRE_HOJA_FRANJAS = "franjas"

HOJAS_REQUERIDAS = [
    NOMBRE_HOJA_CATALOGO,
    NOMBRE_HOJA_PARAMETROS,
    NOMBRE_HOJA_FRANJAS,
]

# ---------------------------------------------------------------------------
# Columnas de la hoja 'catalogo'
# ---------------------------------------------------------------------------

COL_CODIGO_DARWIN = "CodigoDarwin"
COL_ASIGNATURA = "Asignatura"
COL_CODIGO = "Codigo"
COL_SEMESTRE_OFERTA = "SemestreOferta"
COL_TIPO = "Tipo"
COL_RESTRICCION_PROGRAMACION = "RestriccionProgramacion"
COL_FRANJAS_PERMITIDAS = "FranjasPermitidas"
COL_CREDITOS = "Creditos"
COL_HORAS = "Horas"
COL_RESTRICCIONES = "Restricciones"
COL_MIN_SEMANAS_CLASE = "MinSemanasClase"

COLUMNAS_REQUERIDAS_CATALOGO = [
    COL_CODIGO_DARWIN,
    COL_ASIGNATURA,
    COL_CODIGO,
    COL_SEMESTRE_OFERTA,
    COL_TIPO,
    COL_RESTRICCION_PROGRAMACION,
    COL_FRANJAS_PERMITIDAS,
    COL_CREDITOS,
    COL_HORAS,
    COL_RESTRICCIONES,
    COL_MIN_SEMANAS_CLASE,
]

# ---------------------------------------------------------------------------
# Columnas de la hoja 'franjas'
# ---------------------------------------------------------------------------

COL_FRANJA_NOMBRE = "NOMBRE"
COL_FRANJA_HORA_INICIO = "HORA_INICIO"
COL_FRANJA_HORA_FIN = "HORA_FIN"
COL_FRANJA_DURACION = "DURACION(MINS)"

COLUMNAS_REQUERIDAS_FRANJAS = [
    COL_FRANJA_NOMBRE,
    COL_FRANJA_HORA_INICIO,
    COL_FRANJA_HORA_FIN,
    COL_FRANJA_DURACION,
]

# ---------------------------------------------------------------------------
# Claves de la hoja 'parametros'
# Las claves se normalizan al leer (strip + espacios → _), por eso
# se definen aquí ya en su forma normalizada.
# ---------------------------------------------------------------------------

CLAVE_SEMESTRE_PROGRAMACION = "SEMESTRE_PROGRAMACION"
CLAVE_FECHA_INDUCCION = "FECHA_INDUCCION"
CLAVE_INICIO_CLASES = "INICIO_CLASES"
CLAVE_INICIO_SEMANA_SIN_CLASES = "INICIO_SEMANA_SIN_CLASES"
CLAVE_FIN_SEMANA_SIN_CLASES = "FIN_SEMANA_SIN_CLASES"
CLAVE_FESTIVOS = "FESTIVOS"
CLAVE_FIN_CLASES = "FIN_CLASES"

# Claves opcionales para fechas de clases presenciales
CLAVE_VIERNES_PRESENCIAL_UNO = "VIERNES_PRESENCIAL_UNO"
CLAVE_SABADO_PRESENCIAL_UNO = "SABADO_PRESENCIAL_UNO"
CLAVE_VIERNES_PRESENCIAL_DOS = "VIERNES_PRESENCIAL_DOS"
CLAVE_SABADO_PRESENCIAL_DOS = "SABADO_PRESENCIAL_DOS"

# Clave opcional para el número de la semana inicial en el calendario base.
# Si no se define en el Excel, la primera semana del calendario se numera como 1.
CLAVE_SEMANA_INICIO = "SEMANA_INICIO"

CLAVES_REQUERIDAS_PARAMETROS = [
    CLAVE_SEMESTRE_PROGRAMACION,
    CLAVE_FECHA_INDUCCION,
    CLAVE_INICIO_CLASES,
    CLAVE_INICIO_SEMANA_SIN_CLASES,
    CLAVE_FIN_SEMANA_SIN_CLASES,
    CLAVE_FESTIVOS,
]

# ---------------------------------------------------------------------------
# Valores de texto esperados en las columnas del catálogo
# Se usan para mapear strings del Excel a enums internos.
# ---------------------------------------------------------------------------

# Valores de la columna 'RestriccionProgramacion'
TEXTO_MISMA_FRANJA = "MismaFranja"
TEXTO_OBLIGATORIOS_MISMA_FRANJA = "ObligatoriosMismaFranja"
TEXTO_NO_CRUCES = "No puede haber cruces - en proceso de desarrollo"
TEXTO_SOLO_MIERCOLES = "SoloMiercoles"
# Caso especial: la asignatura SoloMiercoles tiene el nombre de franja
# directamente en la columna RestriccionProgramacion
TEXTO_SOLO_MIERCOLES_ALTERNATIVO = "FRANJA_UNO_MIERCOLES"

# Valores de la columna 'Semestre oferta'
TEXTO_SEMESTRE_PRIMERO = "Primero"
TEXTO_SEMESTRE_SEGUNDO = "Segundo"
TEXTO_SEMESTRE_AMBOS = "Ambos"

# ---------------------------------------------------------------------------
# Prefijos de nombre de franja para derivar DiaSemana
# El nombre de la franja siempre contiene el día (ej. FRANJA_UNO_VIERNES).
# ---------------------------------------------------------------------------

PREFIJO_DIA_MIERCOLES = "MIERCOLES"
PREFIJO_DIA_VIERNES = "VIERNES"
PREFIJO_DIA_SABADO = "SABADO"

# ---------------------------------------------------------------------------
# Nombres de columnas del DataFrame de calendario
# ---------------------------------------------------------------------------

CAL_FECHA = "fecha"
CAL_DIA_SEMANA = "dia_semana"
CAL_NOMBRE_FRANJA = "nombre_franja"
CAL_HORA_INICIO = "hora_inicio"
CAL_HORA_FIN = "hora_fin"
CAL_DURACION_MINS = "duracion_mins"
CAL_ES_PROGRAMABLE = "es_programable"
CAL_MOTIVO_BLOQUEO = "motivo_bloqueo"

# ---------------------------------------------------------------------------
# Motivos de bloqueo de un slot en el calendario
# ---------------------------------------------------------------------------

MOTIVO_FESTIVO = "festivo"
MOTIVO_SEMANA_SIN_CLASES = "semana sin clases"
