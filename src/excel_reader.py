"""
excel_reader.py — Lectura y parseo del archivo Excel de entrada.

Este módulo es el único punto de contacto del sistema con el archivo Excel.
Su responsabilidad es leer las tres hojas requeridas, validar su estructura
y convertir los datos crudos en objetos del dominio.

Funciones principales:
    leer_excel         → lee el archivo y retorna los DataFrames crudos
    parsear_franjas    → convierte la hoja 'franjas' en lista de Franja
    parsear_parametros → convierte la hoja 'parametros' en Parametros
    parsear_catalogo   → convierte la hoja 'catalogo' en lista de Asignatura
"""

import datetime

import pandas as pd

import config
from models import (
    Asignatura,
    DiaSemana,
    Franja,
    Parametros,
    RestriccionProgramacion,
    SemestreOferta,
)


# ---------------------------------------------------------------------------
# Lectura del archivo Excel
# ---------------------------------------------------------------------------

def leer_excel(ruta_archivo: str) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Abre el archivo Excel y retorna los DataFrames crudos de las tres hojas.

    Valida que el archivo exista y que contenga las hojas requeridas.

    Args:
        ruta_archivo: ruta al archivo Excel de entrada.

    Returns:
        Tupla con (df_catalogo, df_parametros, df_franjas).

    Raises:
        FileNotFoundError: si el archivo no existe en la ruta indicada.
        ValueError: si falta alguna de las hojas requeridas.
    """
    try:
        archivo = pd.ExcelFile(ruta_archivo)
    except FileNotFoundError:
        raise FileNotFoundError(f"No se encontró el archivo Excel en: {ruta_archivo}")

    _validar_hojas(archivo.sheet_names)

    df_catalogo = pd.read_excel(archivo, sheet_name=config.NOMBRE_HOJA_CATALOGO)
    df_parametros = pd.read_excel(archivo, sheet_name=config.NOMBRE_HOJA_PARAMETROS, header=None)
    df_franjas = pd.read_excel(archivo, sheet_name=config.NOMBRE_HOJA_FRANJAS)

    return df_catalogo, df_parametros, df_franjas


def _validar_hojas(hojas_encontradas: list[str]) -> None:
    """
    Verifica que todas las hojas requeridas estén presentes en el Excel.

    Args:
        hojas_encontradas: lista de nombres de hojas en el archivo.

    Raises:
        ValueError: si falta alguna hoja requerida.
    """
    for nombre_hoja in config.HOJAS_REQUERIDAS:
        if nombre_hoja not in hojas_encontradas:
            raise ValueError(
                f"Falta la hoja '{nombre_hoja}' en el Excel. "
                f"Hojas encontradas: {hojas_encontradas}"
            )


# ---------------------------------------------------------------------------
# Parseo de la hoja 'franjas'
# ---------------------------------------------------------------------------

def parsear_franjas(df_franjas: pd.DataFrame) -> list[Franja]:
    """
    Convierte el DataFrame de la hoja 'franjas' en una lista de objetos Franja.

    Valida que existan todas las columnas requeridas y que cada fila tenga
    los datos necesarios para construir una Franja válida.

    Args:
        df_franjas: DataFrame crudo leído desde la hoja 'franjas'.

    Returns:
        Lista de objetos Franja, uno por cada fila del DataFrame.

    Raises:
        ValueError: si falta alguna columna requerida o los datos son inválidos.
    """
    _validar_columnas(df_franjas, config.COLUMNAS_REQUERIDAS_FRANJAS, "franjas")

    franjas = []
    for indice, fila in df_franjas.iterrows():
        franja = _construir_franja(fila, indice)
        franjas.append(franja)

    return franjas


def _construir_franja(fila: pd.Series, indice: int) -> Franja:
    """
    Construye un objeto Franja a partir de una fila del DataFrame de franjas.

    Args:
        fila: fila del DataFrame con los datos de la franja.
        indice: índice de la fila, usado en mensajes de error.

    Returns:
        Objeto Franja construido con los datos de la fila.

    Raises:
        ValueError: si el nombre de la franja no permite determinar el día de la semana.
    """
    nombre = str(fila[config.COL_FRANJA_NOMBRE]).strip()
    hora_inicio = _extraer_hora(fila[config.COL_FRANJA_HORA_INICIO])
    hora_fin = _extraer_hora(fila[config.COL_FRANJA_HORA_FIN])
    duracion_minutos = int(fila[config.COL_FRANJA_DURACION])
    dia_semana = _derivar_dia_semana(nombre, indice)

    return Franja(
        nombre=nombre,
        hora_inicio=hora_inicio,
        hora_fin=hora_fin,
        duracion_minutos=duracion_minutos,
        dia_semana=dia_semana,
    )


def _extraer_hora(valor) -> datetime.time:
    """
    Extrae un objeto time a partir de un valor leído desde Excel.

    Pandas puede leer las horas como strings ('18:30:00'), como objetos
    datetime.time o como datetime.datetime. Esta función normaliza los tres casos.

    Args:
        valor: valor crudo leído desde la celda del Excel.

    Returns:
        Objeto datetime.time correspondiente.

    Raises:
        ValueError: si el valor no puede convertirse a hora.
    """
    if isinstance(valor, datetime.time):
        return valor

    if isinstance(valor, datetime.datetime):
        return valor.time()

    if isinstance(valor, str):
        partes = valor.strip().split(":")
        hora = int(partes[0])
        minuto = int(partes[1])
        return datetime.time(hora, minuto)

    raise ValueError(f"No se puede convertir '{valor}' a hora. Tipo recibido: {type(valor)}")


def _derivar_dia_semana(nombre_franja: str, indice: int) -> DiaSemana:
    """
    Determina el DiaSemana a partir del nombre de la franja.

    Se basa en que el nombre siempre contiene el día como parte del identificador
    (ej. 'FRANJA_UNO_VIERNES' → VIERNES).

    Args:
        nombre_franja: nombre de la franja (ej. 'FRANJA_UNO_VIERNES').
        indice: índice de la fila, usado en mensajes de error.

    Returns:
        Valor del enum DiaSemana correspondiente.

    Raises:
        ValueError: si el nombre no contiene un día reconocido.
    """
    nombre_upper = nombre_franja.upper()

    if config.PREFIJO_DIA_MIERCOLES in nombre_upper:
        return DiaSemana.MIERCOLES

    if config.PREFIJO_DIA_VIERNES in nombre_upper:
        return DiaSemana.VIERNES

    if config.PREFIJO_DIA_SABADO in nombre_upper:
        return DiaSemana.SABADO

    raise ValueError(
        f"No se puede determinar el día de la semana para la franja '{nombre_franja}' "
        f"(fila {indice}). El nombre debe contener 'MIERCOLES', 'VIERNES' o 'SABADO'."
    )


# ---------------------------------------------------------------------------
# Parseo de la hoja 'parametros'
# ---------------------------------------------------------------------------

def parsear_parametros(df_parametros: pd.DataFrame) -> Parametros:
    """
    Convierte el DataFrame de la hoja 'parametros' en un objeto Parametros.

    La hoja es una tabla de pares clave–valor sin encabezados. Las claves
    se normalizan (strip + espacios → _) para tolerar inconsistencias menores
    de digitación en el Excel.

    Args:
        df_parametros: DataFrame crudo leído desde la hoja 'parametros'
            con header=None.

    Returns:
        Objeto Parametros con todos los valores de la corrida actual.

    Raises:
        ValueError: si falta alguna clave requerida.
    """
    mapa_parametros = _construir_mapa_parametros(df_parametros)
    _validar_claves_parametros(mapa_parametros)

    semestre_programacion = str(mapa_parametros[config.CLAVE_SEMESTRE_PROGRAMACION])
    fecha_induccion = _extraer_fecha(mapa_parametros[config.CLAVE_FECHA_INDUCCION])
    inicio_clases = _extraer_fecha(mapa_parametros[config.CLAVE_INICIO_CLASES])
    inicio_semana_sin_clases = _extraer_fecha(mapa_parametros[config.CLAVE_INICIO_SEMANA_SIN_CLASES])
    fin_semana_sin_clases = _extraer_fecha(mapa_parametros[config.CLAVE_FIN_SEMANA_SIN_CLASES])
    festivos = _parsear_festivos(mapa_parametros[config.CLAVE_FESTIVOS])
    fin_clases = _extraer_fecha_opcional(mapa_parametros.get(config.CLAVE_FIN_CLASES))

    # Fechas opcionales de clases presenciales
    viernes_presencial_uno = _extraer_fecha_opcional(
        mapa_parametros.get(config.CLAVE_VIERNES_PRESENCIAL_UNO)
    )
    sabado_presencial_uno = _extraer_fecha_opcional(
        mapa_parametros.get(config.CLAVE_SABADO_PRESENCIAL_UNO)
    )
    viernes_presencial_dos = _extraer_fecha_opcional(
        mapa_parametros.get(config.CLAVE_VIERNES_PRESENCIAL_DOS)
    )
    sabado_presencial_dos = _extraer_fecha_opcional(
        mapa_parametros.get(config.CLAVE_SABADO_PRESENCIAL_DOS)
    )

    # Número de semana inicial para el calendario base (opcional, default 1)
    semana_inicio = _parsear_semana_inicio(
        mapa_parametros.get(config.CLAVE_SEMANA_INICIO)
    )

    return Parametros(
        semestre_programacion=semestre_programacion,
        fecha_induccion=fecha_induccion,
        inicio_clases=inicio_clases,
        inicio_semana_sin_clases=inicio_semana_sin_clases,
        fin_semana_sin_clases=fin_semana_sin_clases,
        festivos=festivos,
        fin_clases=fin_clases,
        viernes_presencial_uno=viernes_presencial_uno,
        sabado_presencial_uno=sabado_presencial_uno,
        viernes_presencial_dos=viernes_presencial_dos,
        sabado_presencial_dos=sabado_presencial_dos,
        semana_inicio=semana_inicio,
    )


def _parsear_semana_inicio(valor) -> int:
    """
    Parsea el valor de SEMANA_INICIO en un entero.

    Si el valor es None, NaN, o no parseable, retorna 1 como valor por defecto.

    Args:
        valor: valor crudo leído desde la celda del Excel (puede ser None o NaN).

    Returns:
        Número entero de la semana inicial (≥ 1).
    """
    if valor is None:
        return 1

    try:
        if pd.isna(valor):
            return 1
    except (TypeError, ValueError):
        pass

    try:
        resultado = int(valor)
        return resultado if resultado >= 1 else 1
    except (ValueError, TypeError):
        return 1


def _construir_mapa_parametros(df_parametros: pd.DataFrame) -> dict:
    """
    Convierte el DataFrame de parámetros en un diccionario clave → valor.

    Normaliza las claves: elimina espacios al inicio/fin y reemplaza
    espacios internos por guiones bajos.

    Args:
        df_parametros: DataFrame con dos columnas (clave, valor) sin encabezados.

    Returns:
        Diccionario con las claves normalizadas y sus valores.
    """
    mapa = {}
    for _, fila in df_parametros.iterrows():
        clave_original = str(fila[0]).strip()
        clave_normalizada = clave_original.replace(" ", "_")
        valor = fila[1]
        mapa[clave_normalizada] = valor
    return mapa


def _validar_claves_parametros(mapa: dict) -> None:
    """
    Verifica que todas las claves requeridas estén presentes en el mapa.

    Args:
        mapa: diccionario de parámetros ya normalizado.

    Raises:
        ValueError: si falta alguna clave requerida.
    """
    for clave in config.CLAVES_REQUERIDAS_PARAMETROS:
        if clave not in mapa:
            raise ValueError(
                f"Falta el parámetro '{clave}' en la hoja '{config.NOMBRE_HOJA_PARAMETROS}'. "
                f"Claves encontradas: {list(mapa.keys())}"
            )


def _extraer_fecha(valor) -> datetime.date:
    """
    Extrae un objeto date a partir de un valor leído desde Excel.

    Pandas lee las fechas del Excel como datetime.datetime. Esta función
    los convierte a datetime.date.

    Args:
        valor: valor crudo leído desde la celda del Excel.

    Returns:
        Objeto datetime.date correspondiente.

    Raises:
        ValueError: si el valor no puede convertirse a fecha.
    """
    if isinstance(valor, datetime.datetime):
        return valor.date()

    if isinstance(valor, datetime.date):
        return valor

    raise ValueError(f"No se puede convertir '{valor}' a fecha. Tipo recibido: {type(valor)}")


def _extraer_fecha_opcional(valor) -> datetime.date | None:
    """
    Extrae una fecha de forma opcional. Retorna None si el valor está ausente.

    Args:
        valor: valor crudo o None.

    Returns:
        Objeto datetime.date o None.
    """
    if valor is None:
        return None

    if pd.isna(valor):
        return None

    return _extraer_fecha(valor)


def _parsear_festivos(valor) -> list[datetime.date]:
    """
    Parsea el valor de la clave FESTIVOS en una lista de fechas.

    El valor puede ser:
    - Un datetime.datetime (un único festivo almacenado como fecha en Excel).
    - Un string con fechas separadas por coma (ej. '07/08/2026, 20/07/2026').

    Args:
        valor: valor crudo leído desde la celda del Excel.

    Returns:
        Lista de objetos datetime.date con los festivos del periodo.
    """
    if isinstance(valor, datetime.datetime):
        return [valor.date()]

    if isinstance(valor, datetime.date):
        return [valor]

    if isinstance(valor, str):
        fechas = []
        for texto_fecha in valor.split(","):
            texto_fecha = texto_fecha.strip()
            if texto_fecha:
                fecha = datetime.datetime.strptime(texto_fecha, "%d/%m/%Y").date()
                fechas.append(fecha)
        return fechas

    return []


# ---------------------------------------------------------------------------
# Parseo de la hoja 'catalogo'
# ---------------------------------------------------------------------------

def parsear_catalogo(
    df_catalogo: pd.DataFrame,
    franjas_validas: list[Franja],
) -> list[Asignatura]:
    """
    Convierte el DataFrame de la hoja 'catalogo' en una lista de objetos Asignatura.

    Valida las columnas requeridas, convierte los tipos de datos y verifica
    que las franjas referenciadas por cada asignatura existan en la definición
    cargada desde la hoja 'franjas'.

    Args:
        df_catalogo: DataFrame crudo leído desde la hoja 'catalogo'.
        franjas_validas: lista de Franja cargadas desde la hoja 'franjas'.
            Se usa para validar que las franjas permitidas existan.

    Returns:
        Lista de objetos Asignatura.

    Raises:
        ValueError: si falta alguna columna requerida o si una asignatura
            referencia una franja que no está definida.
    """
    _validar_columnas(df_catalogo, config.COLUMNAS_REQUERIDAS_CATALOGO, "catalogo")

    nombres_franjas_validas = {franja.nombre for franja in franjas_validas}
    asignaturas = []

    for indice, fila in df_catalogo.iterrows():
        asignatura = _construir_asignatura(fila, indice, nombres_franjas_validas)
        asignaturas.append(asignatura)

    return asignaturas


def _construir_asignatura(
    fila: pd.Series,
    indice: int,
    nombres_franjas_validas: set[str],
) -> Asignatura:
    """
    Construye un objeto Asignatura a partir de una fila del DataFrame del catálogo.

    Args:
        fila: fila del DataFrame con los datos de la asignatura.
        indice: índice de la fila, usado en mensajes de error.
        nombres_franjas_validas: conjunto de nombres de franja válidos para validación.

    Returns:
        Objeto Asignatura construido con los datos de la fila.
    """
    codigo_darwin = int(fila[config.COL_CODIGO_DARWIN])
    codigo = str(fila[config.COL_CODIGO]).strip()
    nombre = str(fila[config.COL_ASIGNATURA]).strip()
    tipo = str(fila[config.COL_TIPO]).strip()
    horas_totales = int(fila[config.COL_HORAS])
    creditos = int(fila[config.COL_CREDITOS])
    min_semanas_clase = int(fila[config.COL_MIN_SEMANAS_CLASE])

    semestre_oferta = _mapear_semestre_oferta(fila[config.COL_SEMESTRE_OFERTA], indice)
    restriccion = _mapear_restriccion_programacion(fila[config.COL_RESTRICCION_PROGRAMACION], indice)
    franjas_permitidas = _parsear_franjas_permitidas(
        fila[config.COL_FRANJAS_PERMITIDAS],
        indice,
        nombres_franjas_validas,
    )
    fechas_bloqueadas = _parsear_fechas_bloqueadas(fila[config.COL_RESTRICCIONES], indice)

    return Asignatura(
        codigo_darwin=codigo_darwin,
        codigo=codigo,
        nombre=nombre,
        semestre_oferta=semestre_oferta,
        tipo=tipo,
        restriccion_programacion=restriccion,
        franjas_permitidas=franjas_permitidas,
        creditos=creditos,
        horas_totales=horas_totales,
        min_semanas_clase=min_semanas_clase,
        fechas_bloqueadas=fechas_bloqueadas,
    )


def _mapear_semestre_oferta(valor, indice: int) -> SemestreOferta:
    """
    Convierte el texto de la columna 'Semestre oferta' al enum SemestreOferta.

    Args:
        valor: texto leído desde la celda.
        indice: índice de la fila, usado en mensajes de error.

    Returns:
        Valor del enum SemestreOferta.

    Raises:
        ValueError: si el texto no corresponde a ningún valor reconocido.
    """
    texto = str(valor).strip()

    if texto == config.TEXTO_SEMESTRE_PRIMERO:
        return SemestreOferta.PRIMERO

    if texto == config.TEXTO_SEMESTRE_SEGUNDO:
        return SemestreOferta.SEGUNDO

    if texto == config.TEXTO_SEMESTRE_AMBOS:
        return SemestreOferta.AMBOS

    raise ValueError(
        f"Valor desconocido en 'Semestre oferta' (fila {indice}): '{texto}'. "
        f"Valores esperados: {config.TEXTO_SEMESTRE_PRIMERO}, "
        f"{config.TEXTO_SEMESTRE_SEGUNDO}, {config.TEXTO_SEMESTRE_AMBOS}."
    )


def _mapear_restriccion_programacion(valor, indice: int) -> RestriccionProgramacion:
    """
    Convierte el texto de la columna 'RestriccionProgramacion' al enum correspondiente.

    Args:
        valor: texto leído desde la celda.
        indice: índice de la fila, usado en mensajes de error.

    Returns:
        Valor del enum RestriccionProgramacion.

    Raises:
        ValueError: si el texto no corresponde a ningún valor reconocido.
    """
    texto = str(valor).strip()

    if texto == config.TEXTO_MISMA_FRANJA:
        return RestriccionProgramacion.MISMA_FRANJA

    if texto == config.TEXTO_OBLIGATORIOS_MISMA_FRANJA:
        return RestriccionProgramacion.OBLIGATORIOS_MISMA_FRANJA

    if texto == config.TEXTO_NO_CRUCES:
        return RestriccionProgramacion.NO_CRUCES

    if texto == config.TEXTO_SOLO_MIERCOLES:
        return RestriccionProgramacion.SOLO_MIERCOLES

    # Caso especial: Proyecto de grado 2 tiene el nombre de la franja
    # directamente en esta columna en lugar del nombre de la restricción.
    if texto == config.TEXTO_SOLO_MIERCOLES_ALTERNATIVO:
        return RestriccionProgramacion.SOLO_MIERCOLES

    raise ValueError(
        f"Valor desconocido en 'RestriccionProgramacion' (fila {indice}): '{texto}'."
    )


def _parsear_franjas_permitidas(
    valor,
    indice: int,
    nombres_franjas_validas: set[str],
) -> list[str]:
    """
    Parsea el contenido de la columna 'Franjas Permitidas' en una lista de nombres.

    La celda puede contener uno o varios nombres de franja separados por saltos
    de línea (\\n).

    Además valida que cada franja referenciada exista en la definición cargada
    desde la hoja 'franjas'.

    Args:
        valor: texto leído desde la celda.
        indice: índice de la fila, usado en mensajes de error.
        nombres_franjas_validas: conjunto de nombres de franja válidos.

    Returns:
        Lista de nombres de franja.

    Raises:
        ValueError: si alguna franja referenciada no existe en la definición.
    """
    if pd.isna(valor):
        return []

    texto = str(valor).strip()
    nombres = [nombre.strip() for nombre in texto.split("\n") if nombre.strip()]

    for nombre in nombres:
        if nombre not in nombres_franjas_validas:
            raise ValueError(
                f"La asignatura en la fila {indice} referencia la franja '{nombre}', "
                f"que no está definida en la hoja 'franjas'. "
                f"Franjas válidas: {sorted(nombres_franjas_validas)}"
            )

    return nombres


def _parsear_fechas_bloqueadas(
    valor,
    indice: int,
) -> list[tuple[datetime.date, datetime.date]]:
    """
    Parsea la columna 'Restricciones' en una lista de rangos de fechas bloqueadas.

    Cada línea del texto tiene el formato 'dd/mm/yyyy - dd/mm/yyyy'.
    Las líneas vacías o mal formadas se ignoran con advertencia.

    Args:
        valor: texto crudo leído desde la celda (puede ser NaN).
        indice: índice de la fila, usado en mensajes de advertencia.

    Returns:
        Lista de tuplas (fecha_inicio, fecha_fin) con los rangos bloqueados.
    """
    if pd.isna(valor):
        return []

    texto = str(valor).strip()
    if not texto:
        return []

    rangos = []
    for linea in texto.split("\n"):
        linea = linea.strip()
        if not linea:
            continue

        partes = linea.split(" - ")
        if len(partes) != 2:
            print(
                f"  ADVERTENCIA: Formato de restricción no reconocido en fila {indice}: '{linea}'. "
                f"Se esperaba 'dd/mm/yyyy - dd/mm/yyyy'."
            )
            continue

        try:
            fecha_inicio = datetime.datetime.strptime(partes[0].strip(), "%d/%m/%Y").date()
            fecha_fin = datetime.datetime.strptime(partes[1].strip(), "%d/%m/%Y").date()
            rangos.append((fecha_inicio, fecha_fin))
        except ValueError:
            print(
                f"  ADVERTENCIA: No se pudo parsear la fecha en fila {indice}: '{linea}'."
            )

    return rangos


# ---------------------------------------------------------------------------
# Utilidades comunes
# ---------------------------------------------------------------------------

def _validar_columnas(
    df: pd.DataFrame,
    columnas_requeridas: list[str],
    nombre_hoja: str,
) -> None:
    """
    Verifica que el DataFrame contenga todas las columnas requeridas.

    Args:
        df: DataFrame a validar.
        columnas_requeridas: lista de nombres de columnas que deben estar presentes.
        nombre_hoja: nombre de la hoja, usado en mensajes de error.

    Raises:
        ValueError: si falta alguna columna requerida.
    """
    columnas_encontradas = set(df.columns)
    for nombre_columna in columnas_requeridas:
        if nombre_columna not in columnas_encontradas:
            raise ValueError(
                f"Falta la columna '{nombre_columna}' en la hoja '{nombre_hoja}'. "
                f"Columnas encontradas: {sorted(columnas_encontradas)}"
            )
