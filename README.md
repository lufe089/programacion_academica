# README — Sistema de apoyo para la programación académica

## 1. Propósito del proyecto

Este proyecto busca construir un software en Python que ayude a realizar la programación académica de asignaturas de la Maestría y la Especialización en Ingeniería de Software, usando como insumo un archivo Excel con:

- catálogo de asignaturas
- parámetros de la corrida de programación
- definición de franjas horarias

El objetivo es disminuir el trabajo manual necesario para:

- construir la programación del semestre
- respetar restricciones horarias y académicas
- considerar fechas especiales y bloqueos del calendario
- generar una versión de horas
- generar una salida consolidada de franjas
- generar posteriormente una versión gráfica

La idea es trabajar este desarrollo de forma incremental, comenzando por una primera versión funcional centrada en la lectura del Excel, construcción del calendario base y generación de una salida tabular de programación.

---

## 2. Alcance de la primera versión

La versión 1 del sistema debe:

1. leer un archivo Excel con la misma estructura del archivo de trabajo actual
2. cargar las asignaturas y sus reglas de programación desde la hoja `catalogo`
3. cargar los parámetros generales de la corrida desde la hoja `parametros`
4. cargar la definición de franjas horarias desde la hoja `franjas`
5. construir el calendario programable del periodo
6. considerar solo las franjas definidas en el Excel
7. excluir fechas no programables
8. preparar la base lógica para una programación por horas
9. generar una salida en formato de horas
10. generar una salida consolidada en formato de franjas
11. dejar preparada una segunda fase para la generación de la versión gráfica

En la corrida actual se trabajará con la programación de **segundo semestre**, aunque el sistema debe quedar preparado para trabajar también con `Primero`.

---

## 3. Enfoque general del problema

El sistema no debe pensarse solamente como un generador de tablas, sino como un motor que integra varias capas:

1. **Entrada estructurada desde Excel**
2. **Lectura de parámetros de la corrida**
3. **Lectura de definición de franjas**
4. **Construcción de calendario académico**
5. **Lectura de restricciones por asignatura**
6. **Aplicación de reglas de programación**
7. **Asignación de sesiones**
8. **Salida tabular de horas**
9. **Salida consolidada de franjas**
10. **Transformación posterior a versión gráfica**

La primera salida relevante del sistema será la **versión de horas**.  
La segunda salida será la **versión de franjas**.  
La **versión gráfica** se desarrollará después, tomando como base la programación lógica generada.

---

## 4. Insumo principal del programa

El sistema siempre recibirá como entrada un archivo Excel con una estructura equivalente al archivo actualmente usado en el proceso manual.

Ese Excel contiene exactamente tres hojas principales:

- una hoja llamada **`catalogo`**
- una hoja llamada **`parametros`**
- una hoja llamada **`franjas`**

Ese archivo se considera la **plantilla base de entrada** para la versión 1.

---

## 5. Supuestos actuales

- El archivo Excel es el insumo oficial del proceso.
- La primera versión trabajará sobre la misma estructura del Excel actual.
- No se construirá inicialmente una interfaz compleja para captura manual.
- La lógica de franjas no debe quedar quemada en el código; debe leerse desde la hoja `franjas`.
- El flujo esperado será:
  1. cargar Excel
  2. leer `catalogo`, `parametros` y `franjas`
  3. construir calendario válido
  4. programar asignaturas
  5. generar salida de horas
  6. generar salida de franjas
  7. generar posteriormente salida gráfica

---

## 5.1 Modos de ejecución

El sistema ofrece un menú interactivo con dos modos de ejecución:

### Opción 1: Generar programación automática

Ejecuta el flujo completo:
1. Lee el archivo `inputs/restricciones.xlsx`
2. Construye el calendario de slots disponibles
3. Asigna sesiones automáticamente según las reglas
4. Genera todos los archivos de salida en `outputs/`

### Opción 2: Regenerar desde matriz ajustada

Permite que el usuario ajuste manualmente la programación:
1. Copiar `outputs/programacion_matriz.xlsx` a `inputs/programacion_matriz.xlsx`
2. Ajustar manualmente las horas en la matriz (agregar, quitar o mover horas entre fechas)
3. Ejecutar la opción 2 del menú
4. El sistema regenera los archivos de salida

#### Flujo técnico de la opción 2

El sistema sigue estos pasos para procesar la matriz ajustada:

1. **Lee restricciones**: Carga parámetros, franjas y catálogo de asignaturas desde `inputs/restricciones.xlsx`
2. **Extrae fechas de la matriz**: Identifica qué fechas tienen horas asignadas para cada asignatura
3. **Construye candidatos específicos**: Genera slots válidos SOLO para las fechas presentes en la matriz, respetando las franjas permitidas de cada asignatura según el catálogo
4. **Distribuye horas en franjas**: Para cada (asignatura, fecha) con horas > 0, asigna las horas a las franjas candidatas en orden de prioridad
5. **Genera archivos de salida**: Exporta `programacion_horas.xlsx`, `programacion_visual.xlsx` y `programacion_franjas.xlsx`

**Importante**: La matriz NO se regenera en este flujo (es la fuente de los ajustes). Si desea regenerarla, ejecute la opción 1 nuevamente.

#### Ventajas de este enfoque

- La matriz puede incluir fechas que originalmente estaban bloqueadas en el calendario (festivos, semana sin clases, etc.)
- El sistema respeta las franjas permitidas de cada asignatura del catálogo
- Se pueden agregar horas en fechas nuevas sin modificar el archivo de restricciones

Este flujo es útil para:
- Corregir asignaciones que el algoritmo no pudo resolver óptimamente
- Ajustar la distribución de horas manualmente
- Mover clases a fechas que originalmente estaban bloqueadas
- Verificar cambios antes de generar las salidas finales

#### Asignaturas adicionales

El sistema conserva las filas adicionales que el usuario agregue manualmente en la matriz, incluso si no existen en el catálogo original (`restricciones.xlsx`). Para estas asignaturas:
- Se lee el nombre de la columna B y el tipo de la columna A
- Se crean sesiones con la información disponible (fecha, horas, tipo)
- Los campos que no se pueden inferir (código, franja, hora inicio/fin) se dejan vacíos
- Estas sesiones se incluyen en los archivos de salida regenerados

**Nota**: La matriz usa fórmulas de Excel en las columnas de totales (Asignadas, Diferencia, Estado) para que el usuario pueda verificar que sus ajustes son correctos antes de regenerar.

---

## 6. Franjas horarias base

Las franjas horarias no deben codificarse manualmente en el sistema como fuente principal.  
La fuente de verdad debe ser la hoja `franjas` del Excel.

Sin embargo, con base en el archivo actual, se espera trabajar con franjas como las siguientes:

### Miércoles
- `FRANJA_UNO_MIERCOLES`

### Viernes
- `FRANJA_UNO_VIERNES`
- `FRANJA_DOS_VIERNES`
- `FRANJA_TRES_VIERNES`

### Sábado
- `FRANJA_UNO_SABADO`
- `FRANJA_DOS_SABADO`
- `FRANJA_TRES_SABADO`
- `FRANJA_CUATRO_SABADO`
- `FRANJA_CINCO_SABADO`

La hora de inicio, hora de fin y duración deben leerse desde el Excel.

---

## 7. Bloques generales

Además de las franjas exactas, el sistema debe manejar la noción de **bloques generales**, porque varias reglas se expresan de esa forma.

Los bloques generales no necesariamente vienen explícitos en el Excel, por lo que el sistema puede derivarlos a partir de la definición de franjas y de su uso institucional.

### Miércoles noche
- franjas de miércoles en horario nocturno
- actualmente: `FRANJA_UNO_MIERCOLES`

### Viernes tarde
- primeras franjas de viernes antes del bloque nocturno
- actualmente:
  - `FRANJA_UNO_VIERNES`
  - `FRANJA_DOS_VIERNES`

### Viernes noche
- franja final de viernes
- actualmente:
  - `FRANJA_TRES_VIERNES`

### Sábado mañana
- primeras franjas del sábado
- actualmente:
  - `FRANJA_UNO_SABADO`
  - `FRANJA_DOS_SABADO`
  - `FRANJA_TRES_SABADO`

### Sábado tarde
- franjas finales del sábado
- actualmente:
  - `FRANJA_CUATRO_SABADO`
  - `FRANJA_CINCO_SABADO`

---

## 8. Reglas del negocio identificadas hasta ahora

Las reglas del sistema pueden dividirse en dos grupos:

### Restricciones duras
Son reglas que no deben violarse, por ejemplo:

- inclusión por semestre de oferta
- `MismaFranja`
- `SoloMiércoles`
- no cruces
- horas objetivo por asignatura
- respeto a las franjas permitidas por cada asignatura

### Reglas de equilibrio o suavización
Son reglas que el sistema debe intentar cumplir para producir una programación más conveniente, por ejemplo:

- carga objetivo de horas por viernes y sábado
- máximo de dos cursos iniciando en la primera semana
- evitar que cursos de distintos bloques inicien o terminen juntos
- evitar bloques parciales cuando sea posible

### 8.1 Regla de inclusión por semestre de oferta
En cada corrida del sistema se programa un único conjunto objetivo: `Primero` o `Segundo`.

Las asignaturas a considerar se determinan así:

- si la corrida es `Primero`, se incluyen asignaturas con `Semestre oferta` igual a `Primero` o `Ambos`
- si la corrida es `Segundo`, se incluyen asignaturas con `Semestre oferta` igual a `Segundo` o `Ambos`

Las asignaturas del otro semestre no deben programarse en esa corrida.

### 8.2 Regla de `MismaFranja`
La regla `MismaFranja` no aplica a todas las asignaturas del archivo en conjunto.

Solo aplica a las asignaturas que cumplan simultáneamente estas condiciones:

- tienen la restricción `MismaFranja`
- pertenecen al mismo valor de la columna `Tipo`
- deben ser consideradas en la corrida actual de programación según la columna `Semestre oferta`

Una vez filtradas las asignaturas que sí entran en la corrida actual, las asignaturas que compartan:

- la restricción `MismaFranja`
- el mismo `Tipo`

deben programarse en el mismo horario exacto.

Esto quiere decir que deben compartir:

- el mismo día
- la misma franja horaria específica
- el mismo patrón de horario dentro del periodo

Ejemplo:  
Si se está programando **Segundo semestre**, y dentro de las asignaturas incluidas hay varias de tipo `TopicoAvanzado` con restricción `MismaFranja`, todas esas asignaturas deben quedar en la misma franja exacta. Si la franja elegida es `FRANJA_UNO_SABADO`, todas las asignaturas de ese grupo deben quedar en esa misma franja.

### 8.3 Regla de `ObligatoriosMismaFranja`
Se entiende como una variante de la lógica de franja compartida aplicada a las asignaturas obligatorias.

En la implementación se trata igual que `MismaFranja`: todas las asignaturas del mismo `Tipo` con esta restricción comparten la misma franja exacta durante todo el periodo.

### 8.4 Regla de no cruces entre tipos (`NoCruces`)
Las asignaturas con restricción `NoCruces` (tipo `ProcesoDesarrollo`) tienen las siguientes prohibiciones:

- **No pueden coincidir entre sí**: dos asignaturas `ProcesoDesarrollo` no pueden programarse en la misma (fecha, franja).
- **No pueden coincidir con ningún otro tipo**: una asignatura `ProcesoDesarrollo` no puede estar en la misma (fecha, franja) que una asignatura `Obligatorio`, `ObligatoriosMismaFranja` ni `TemasAvanzados`.

Esto aplica incluso si pertenecen a diferentes conjuntos académicos, porque los estudiantes deberían poder cursarlas.

### 8.5 Regla de no coincidencia entre tipos distintos
Independientemente de la restricción individual de cada asignatura, **ninguna asignatura de un tipo puede coincidir con una asignatura de otro tipo** en la misma (fecha, franja).

Esto implica:

- `Obligatorio` y `TemasAvanzados` no pueden compartir franja (aunque cada grupo internamente tiene su propia franja fija).
- `ProcesoDesarrollo` no puede coincidir con ningún otro tipo.
- En general, no se permiten cruces entre tipos distintos en el mismo slot.

La implementación logra esto procesando primero los grupos `MismaFranja` (que reservan sus slots) y luego los `NoCruces`, quienes seleccionan dinámicamente franjas libres evitando los slots ya ocupados.

### 8.6 Regla de `SoloMiércoles`
Por ahora, solo **Proyecto de grado 2** usa la franja de miércoles.  
Debe programarse únicamente en una franja de miércoles.

Con el archivo actual, eso corresponde a:

- `FRANJA_UNO_MIERCOLES`

### 8.7 Regla de horas objetivo
Cada asignatura tiene una cantidad total de horas que debe completar.

### 8.8 Regla de franjas permitidas
Cada asignatura trae en la hoja `catalogo` una lista de franjas permitidas.

El sistema solo podrá programar esa asignatura dentro de alguna de esas franjas.

La columna `Franjas Permitidas` es una de las restricciones principales del problema.

### 8.9 Regla de bloques parciales
Se pueden usar bloques parciales, pero se espera evitar este comportamiento siempre que sea posible.

La lógica del sistema debe priorizar soluciones que usen bloques completos y dejar los parciales como recurso de ajuste.

### 8.10 Regla de carga objetivo por día
Como criterio general de programación, se busca que:

- los **viernes** tengan aproximadamente **7 horas de clase**
- los **sábados** tengan aproximadamente **10 horas de clase**

Además, se busca evitar, cuando sea posible, la clase de la última franja del sábado.  
En esos casos es aceptable tener una carga menor, por ejemplo **8 horas**.

Esta regla debe tratarse como una meta de distribución de carga del calendario, no necesariamente como una restricción absoluta en todos los casos.

### 8.11 Regla de arranque controlado en la primera semana
Durante la primera semana de clases deben iniciar como máximo **dos cursos**.

El sistema debe evitar que en la primera semana comiencen demasiadas asignaturas al mismo tiempo, para facilitar un arranque académico más manejable.

### 8.12 Regla de desalineación de inicios y cierres
Las asignaturas que no pertenecen al mismo bloque de horario no deberían iniciar ni terminar en las mismas semanas.

Esta regla busca evitar picos de trabajo para los estudiantes.

En consecuencia, el sistema debe intentar distribuir los inicios y cierres de las asignaturas de forma escalonada.

Ejemplo:  
si una asignatura como **Arquitectura de software** termina en una semana determinada, otras asignaturas de otro bloque, como las de **Tópicos Avanzados**, idealmente no deberían terminar en esa misma semana, sino una o más semanas después.

Esta regla debe tratarse como un criterio de suavización de carga y equilibrio de la programación.


### 8.13 Regla de validación visual de carga
La programación debe poder revisarse en una salida tipo matriz que permita identificar visualmente:

- el inicio y fin de cada asignatura
- el cumplimiento de horas por curso
- la carga efectiva por día
- los días sin clase
- los picos de inicio y cierre de cursos

Esta regla busca que la salida no solo sea correcta desde el punto de vista computacional, sino también útil para revisión humana.

### 8.14 Regla de fechas especiales del calendario

El sistema debe reconocer fechas especiales definidas en la hoja `parametros`, incluyendo:

- fechas de clases presenciales
- fechas sin clase
- fecha de inducción

Estas fechas deben:

- afectar la construcción del calendario
- reflejarse correctamente en la salida de horas
- diferenciarse visualmente en la exportación

### 8.15 Regla de TemasAvanzados y encuentros presenciales

Todas las asignaturas de tipo `TemasAvanzados` que se programen para el semestre tienen una restricción especial relacionada con los encuentros presenciales y el desplazamiento de profesores.

#### Contexto
- Los profesores de estas asignaturas viajan desde otra ciudad para impartir las clases presenciales
- Solo hay presupuesto para **un desplazamiento por semestre** por cada profesor
- Las clases pueden programarse en cualquier semana del semestre (virtuales o presenciales)
- En la semana del encuentro presencial, el profesor aprovecha el viaje para dar una sesión más larga

#### Comportamiento esperado

1. **Sesión intensiva en semana del primer encuentro presencial**: En la semana que tiene el primer encuentro presencial, la asignatura debe programar más horas de lo habitual, entre **5 y 7 horas** en esa semana, para aprovechar el viaje del profesor.

2. **Sin clase en semana del segundo encuentro presencial**: En la semana del segundo encuentro presencial, **no se programa clase** de estas asignaturas. El profesor no viaja dos veces, por lo que no tiene sentido programar clase presencial, y se evita que los estudiantes tengan que conectarse virtualmente en un fin de semana de encuentro presencial.

3. **Clases virtuales en otras semanas**: El curso puede tener clases virtuales antes y después del primer encuentro presencial, según sea necesario para completar las horas.

4. **Posible finalización antes del segundo encuentro**: El curso puede terminar antes de la semana del segundo encuentro presencial, en cuyo caso simplemente no hay clase que bloquear esa semana.

#### Implicaciones para el scheduler

- En la semana del primer encuentro presencial (`VIERNES_PRESENCIAL_UNO`, `SABADO_PRESENCIAL_UNO`): asignar entre 5 y 6 horas (sesión intensiva)
- En la semana del segundo encuentro presencial (`VIERNES_PRESENCIAL_DOS`, `SABADO_PRESENCIAL_DOS`): **bloquear** la programación de estas asignaturas (si el curso aún no ha terminado)
- El resto de semanas: programación normal según las franjas permitidas
- Esta regla aplica a **todas** las asignaturas de tipo `TemasAvanzados` del semestre

#### Identificación de encuentros presenciales

Los encuentros presenciales se identifican a partir de los parámetros:

- `VIERNES_PRESENCIAL_UNO` y `SABADO_PRESENCIAL_UNO` → primer encuentro presencial (sesión intensiva de 5-6 horas)
- `VIERNES_PRESENCIAL_DOS` y `SABADO_PRESENCIAL_DOS` → segundo encuentro presencial (sin clase para TemasAvanzados)

### 8.16 Regla de continuidad de asignaturas

Una vez que una asignatura inicia su programación, debe continuar en semanas consecutivas sin interrupciones.

#### Comportamiento esperado

1. **Continuidad obligatoria**: Si una asignatura tiene su primera sesión en la semana N, debe tener sesiones en las semanas N+1, N+2, etc., hasta completar sus horas.

2. **Excepciones permitidas**: Las únicas interrupciones válidas son:
   - Festivos
   - Semana sin clases
   - Restricciones específicas de la asignatura (columna `Restricciones` del catálogo)
   - TemasAvanzados en semana del segundo encuentro presencial (regla 8.15)

3. **Sin huecos arbitrarios**: El scheduler no debe dejar semanas vacías entre sesiones de una misma asignatura por falta de slots o conveniencia del algoritmo.

#### Implicaciones para el scheduler

- Al asignar una asignatura por primera vez, registrar su semana de inicio
- En semanas posteriores, priorizar la continuidad de asignaturas ya iniciadas
- Solo omitir una semana si hay una excepción válida documentada
- Si no hay slot disponible para continuar, reportar advertencia

### 8.17 Regla de relleno para horas pendientes

Después de la asignación inicial, si quedan asignaturas con horas pendientes, el sistema debe intentar completarlas distribuyendo sesiones adicionales en fechas subutilizadas.

#### Comportamiento esperado

1. **Identificar fechas subutilizadas**: Buscar fechas donde las horas efectivas no superan el máximo preferido:
   - Viernes: máximo 7 horas
   - Sábado: máximo 7 horas
   - Miércoles: máximo 2 horas

2. **Distribuir horas pendientes**: Asignar sesiones adicionales en los slots libres de esas fechas, respetando las restricciones de tipo y cruce.

3. **Prioridad por horas faltantes**: Las asignaturas con más horas pendientes se procesan primero.

4. **Respetar continuidad**: Las sesiones de relleno deben preferir fechas que mantengan la continuidad de la asignatura.

#### Implicaciones para el scheduler

- Ejecutar una fase de relleno después de la asignación principal
- Calcular horas efectivas por fecha (contando franjas únicas, no duplicadas por asignaturas compartidas)
- Solo asignar si hay espacio sin superar el máximo de horas del día
- Actualizar el registro de slots ocupados y horas acumuladas

### 8.18 Regla de máximo 6 horas por asignatura por fin de semana

Ninguna asignatura puede tener más de 6 horas de clase en un mismo fin de semana.

#### Comportamiento esperado

1. **Límite de 6 horas por semana**: Una asignatura no puede acumular más de 6 horas en las sesiones de un mismo fin de semana (viernes + sábado de la misma semana).

2. **Excepción: sesión intensiva presencial**: La única excepción es para asignaturas `TemasAvanzados` en la semana del primer encuentro presencial, donde pueden tener hasta 6 horas (regla 8.15).

3. **Distribución en múltiples semanas**: Si una asignatura necesita más horas, debe distribuirlas en semanas adicionales.

#### Implicaciones para el scheduler

- Al seleccionar franjas para una sesión semanal, verificar que no se superen 6 horas
- La función `_seleccionar_mejor_subconjunto` debe respetar este límite
- Aplica tanto a la asignación principal como a la fase de relleno

### 8.19 Regla de máximo 4 horas por asignatura por día

Ninguna asignatura puede tener más de 4 horas de clase en un mismo día.

#### Comportamiento esperado

1. **Límite de 4 horas diarias**: La selección de franjas para una sesión no puede sumar más de 4 horas en un mismo día de la semana.

2. **Excepción: TemasAvanzados en semana presencial**: Las asignaturas `TemasAvanzados` en la semana del primer encuentro presencial pueden llegar hasta 6 horas en el día de la sesión intensiva (regla 8.15).

3. **Múltiples días en una semana**: Si una asignatura necesita más de 4 horas semanales, las horas adicionales deben distribuirse en otro día de la misma semana (ej. viernes + sábado).

#### Implicaciones para el scheduler

- `_seleccionar_mejor_subconjunto` recibe el parámetro `max_minutos_por_dia` y descarta combinaciones donde un día supere ese tope
- El límite se aplica en la selección de franja común (`_seleccionar_franja_comun_grupo`), en asignaturas NoCruces (`_intentar_asignar_no_cruces`) y en la fase de relleno
- `_seleccionar_franjas_sesion_intensiva` no aplica este límite (excepción presencial)
- La fase de relleno rastrea horas por `(asignatura, fecha)` para respetar el tope diario

---

## 9. Restricciones temporales conocidas

El sistema debe soportar fechas bloqueadas globales y restricciones específicas por asignatura.

Estas restricciones pueden venir desde:

- la hoja `parametros`, para bloqueos globales del periodo
- la hoja `catalogo`, para restricciones específicas por asignatura

Ejemplos de restricciones ya identificadas:

- festivos a considerar
- semana sin clases
- fecha de inducción
- restricciones de disponibilidad de una asignatura
- fechas límite para terminar una asignatura
- fechas puntuales en las que una asignatura no puede programarse

---

## 10. Tipos de asignatura identificados

Los tipos actualmente presentes en el catálogo y sus reglas asociadas son:

| Tipo | Restricción programación | Comportamiento |
|---|---|---|
| `Obligatorio` | `ObligatoriosMismaFranja` | Todas las del grupo comparten la misma franja fija durante todo el semestre |
| `TemasAvanzados` | `MismaFranja` | Todas las del grupo comparten la misma franja fija durante todo el semestre. Además: sesión intensiva (5-6h) en semana del primer encuentro presencial, sin clase en semana del segundo encuentro presencial (ver regla 8.15) |
| `ProcesoDesarrollo` | `NoCruces` | Cada asignatura ocupa su propio slot; no puede coincidir con ninguna otra asignatura de ningún tipo |
| `ProyectoGrado` | `SoloMiercoles` | Programada exclusivamente en la franja de miércoles |

El sistema debe leer el tipo desde Excel y usarlo como apoyo para aplicar reglas de programación.

> Nota: el sistema no asume tipos fijos en el código. Los tipos se leen del Excel y se usan para agrupar las asignaturas con reglas `MismaFranja` y `ObligatoriosMismaFranja`.

---

## 11. Estructura esperada del Excel

La versión 1 asumirá la misma estructura general del archivo Excel actual.

### Hoja `catalogo`
Debe contener la información de las asignaturas y sus restricciones.

Columnas identificadas hasta ahora:

- `Codigo Darwin`
- `Asignatura`
- `Semestre oferta`
- `Tipo`
- `RestriccionProgramacion`
- `Franjas Permitidas`
- `Horas`
- `Restricciones`

### Hoja `parametros`
Debe contener parámetros generales que definen la corrida de programación y el periodo académico a considerar.

En el archivo actual se maneja como pares clave–valor.

Parámetros esperados hasta ahora:

- `SEMESTRE_PROGRAMACION`
- `FECHA_INDUCCION`
- `INICIO_CLASES`
- `FIN_CLASES`
- `INICIO_SEMANA SIN CLASES`
- `FIN_SEMANA_SIN CLASES`
- `FESTIVOS A CONSIDERAR`
- `VIERNES_PRESENCIAL_UNO`
- `SABADO_PRESENCIAL_UNO`
- `VIERNES_PRESENCIAL_DOS`
- `SABADO_PRESENCIAL_DOS`

Estos parámetros pueden crecer en el futuro, por lo que el lector debe ser flexible para aceptar nuevas claves.

### Hoja `franjas`
Debe contener la definición formal de las franjas disponibles.

Columnas esperadas:

- `NOMBRE`
- `HORA_INICIO`
- `HORA_FIN`
- `DURACION(MINS)`

Esta hoja es obligatoria y debe ser usada como fuente de verdad para construir la lógica de horarios.

---

## 12. Reglas de interpretación del Excel

### 12.1 Hoja `catalogo`
La hoja `catalogo` contiene el universo de asignaturas y sus restricciones.

#### `RestriccionProgramacion`
La columna `RestriccionProgramacion` define reglas como:

- `MismaFranja`
- `ObligatoriosMismaFranja`
- `No puede haber cruces - en proceso de desarrollo`

La implementación debe mapear estas expresiones a reglas internas más limpias.

#### `Franjas Permitidas`
La columna `Franjas Permitidas` define explícitamente las franjas en las que una asignatura puede ser programada.

Puede contener una o varias franjas, separadas por saltos de línea.

El sistema debe parsear esa celda y convertirla en una lista interna de franjas válidas para cada asignatura.

#### `Restricciones`
Esta columna contiene uno o más rangos de fechas en los que la asignatura no puede programarse. Cada rango ocupa una línea con el formato:

```
dd/mm/yyyy - dd/mm/yyyy
```

Varias restricciones se separan con saltos de línea. Ejemplo:

```
01/08/2026 - 01/08/2026
06/11/2026 - 06/12/2026
```

El sistema parsea estas líneas en una lista de tuplas `(fecha_inicio, fecha_fin)` almacenada en el campo `fechas_bloqueadas` de la asignatura. Cualquier slot cuya fecha caiga dentro de alguno de esos rangos queda excluido de los candidatos de la asignatura antes de iniciar la asignación.

### 12.2 Hoja `parametros`
La hoja `parametros` define la configuración de la corrida actual y debe leerse antes de construir el calendario.

Ejemplo de valores esperados:

- `SEMESTRE_PROGRAMACION = Segundo`
- `FECHA_INDUCCION = 25/07/2026`
- `INICIO_CLASES = 31/07/2026`
- `INICIO_SEMANA SIN CLASES = 14/09/2026`
- `FIN_SEMANA_SIN CLASES = 19/09/2026`
- `FESTIVOS A CONSIDERAR = 07/08/2026`

El sistema debe usar esta hoja para:

- definir qué conjunto académico se programa en la corrida actual
- identificar la fecha de inducción
- definir la fecha real de inicio de clases
- bloquear una semana completa sin clases cuando aplique
- registrar uno o varios festivos relevantes del periodo

### 12.3 Hoja `franjas`
La hoja `franjas` define la estructura horaria operativa del sistema y debe leerse antes de cualquier asignación.

El sistema debe usar esta hoja para:

- construir los objetos de franja
- conocer hora de inicio y fin de cada franja
- calcular la duración real de cada bloque
- validar que las franjas usadas por las asignaturas existan realmente

---

## 13. Salidas esperadas

### 13.1 Salida principal inicial: versión de horas
La primera salida importante del sistema será una tabla detallada con información como:

- asignatura
- fecha
- día
- franja
- duración de la sesión
- horas acumuladas por asignatura
- observaciones o restricciones aplicadas

Esta salida representa el detalle fino de la programación sesión por sesión.

### 13.2 Salida intermedia: versión de franjas
El sistema también deberá generar una hoja de **franjas**, orientada a resumir períodos en los que una asignatura mantiene el mismo horario.

Esta salida debe permitir identificar entre qué semanas y entre qué fechas una misma franja de horario se mantiene para una asignatura.

Cada registro de esta salida debe consolidar bloques continuos de programación con el mismo patrón horario.

Campos esperados para esta salida:

- asignatura
- semana inicio
- semana fin
- fecha inicio
- día semana de inicio
- fecha fin
- día semana de fin
- hora de inicio
- hora de fin
- cantidad de horas

Ejemplo de interpretación:  
Si una asignatura mantiene la franja de viernes de 2:00 pm a 6:30 pm entre varias semanas consecutivas, la hoja de franjas debe mostrar ese tramo como un único bloque consolidado, en lugar de repetir una fila por cada sesión individual.

Esta salida sirve para entender la continuidad de las franjas y facilitar la revisión administrativa de la programación.


### 13.3 Salida matriz de horas

La salida de **matriz de horas** es la herramienta de trabajo para la asignación y validación de la programación.

Está diseñada para permitir una revisión visual clara de:

- la distribución de sesiones por asignatura
- el balance de carga por día
- el cumplimiento de horas por curso
- el inicio y fin de cada asignatura
- la distribución equilibrada de los cursos en el tiempo

---

### Estructura general

La salida debe construirse como una **matriz principal** con esta lógica:

- **Filas → asignaturas**
- **Columnas → fechas del calendario**

Cada celda representa la programación de una asignatura en una fecha específica.

El contenido visible de cada celda debe ser únicamente:

- **número de horas**

Ejemplos de valores:
- `2`
- `2.5`
- `4`
- vacío si no hay clase para esa asignatura en esa fecha

---

### Propósito de la salida

Esta salida debe permitir:

- validar que los **viernes no superen aproximadamente 7 horas efectivas**
- validar que los **sábados no superen aproximadamente 10 horas efectivas**
- verificar que cada asignatura cumpla su total de horas
- identificar rápidamente días sin clase
- visualizar claramente el inicio y fin de cada asignatura
- detectar picos de carga académica en el tiempo

Esta hoja está orientada principalmente a **uso humano**, para planeación, revisión y ajuste.

---

### Principio de cálculo de carga diaria

La carga diaria debe calcularse por **franja ocupada**, no por número de asignaturas.

#### Regla clave
Si varias asignaturas comparten la misma franja en una fecha, esa franja se cuenta una sola vez en el total del día.

#### Ejemplo
Si tres asignaturas comparten una franja de 2 horas en un sábado, para el balance del día solo se contabilizan **2 horas efectivas**, no 6.

Sin embargo, cada asignatura sí conserva sus propias 2 horas para efectos del cumplimiento de su plan.

---

### Doble nivel de análisis

Esta salida debe permitir analizar simultáneamente:

#### 1. Nivel asignatura
- horas acumuladas por curso
- cumplimiento del total requerido
- continuidad del curso en el tiempo

#### 2. Nivel calendario
- horas efectivas por día
- distribución de carga semanal
- uso real de franjas

---

### Visualización de inicio y fin de cursos

La estructura de filas por asignatura debe permitir identificar claramente:

- fecha de inicio de cada curso
- fecha de finalización
- duración del curso en semanas

Esto debe poder observarse visualmente al ver en qué columna aparece la primera celda con horas y en qué columna aparece la última.

---

### Regla de balance de ciclos académicos

El sistema debe evitar, en la medida de lo posible, concentraciones de inicio y fin de cursos.

#### Restricciones deseadas

- cursos de tipo `ProcesoDesarrollo` no deben iniciar todos en la misma semana
- cursos de tipo `ProcesoDesarrollo` no deben terminar todos en la misma semana
- cursos de tipo `ProcesoDesarrollo` no deberían terminar simultáneamente con:
  - cursos `Obligatorio`
  - cursos `TemasAvanzados`

#### Objetivo
Reducir picos de carga académica para los estudiantes.

---

### Manejo de días sin clase

Los días sin clase deben ser visibles en la tabla:

- como columnas correspondientes a fechas válidas del calendario
- con celdas vacías si no hay programación
- y con formato visual diferencial cuando se trate de fechas bloqueadas o sin clase

En Excel, estos días pueden resaltarse con color.

Ejemplos:
- festivos
- semana sin clases
- fechas bloqueadas por parámetros

---

### Información adicional que debe acompañar la matriz principal

Además de la matriz asignaturas × fechas, la salida debe incluir elementos auxiliares de lectura y validación.

#### 1. Columna final por asignatura
Cada fila de asignatura debe incluir al final, como mínimo:

- **total de horas programadas**
- **horas objetivo**
- **diferencia** entre horas programadas y horas objetivo
- **estado** de cumplimiento, por ejemplo:
  - `Completa`
  - `Incompleta`
  - `Excede`

#### 2. Filas superiores o inferiores de contexto por fecha
La hoja debe incluir filas auxiliares que permitan interpretar mejor cada columna de fecha, por ejemplo:

- número de semana
- día de la semana
- fecha completa
- indicador de si el día es miércoles, viernes o sábado
- indicador visual de día sin clase o fecha bloqueada

#### 3. Fila resumen de horas efectivas por fecha
La hoja debe incluir una fila resumen que muestre, para cada fecha:

- **total de horas efectivas del día**

Este cálculo debe respetar la regla de franja compartida:  
si varias asignaturas usan la misma franja, esa franja cuenta una sola vez.

##### Implementación con fórmulas Excel

La fila de horas efectivas utiliza **fórmulas Excel** en lugar de valores calculados, permitiendo que el usuario pueda ajustar manualmente las horas de cualquier asignatura y ver el resultado recalculado automáticamente.

La fórmula por columna combina:
- **`MAX()`** para grupos de franja compartida (`Obligatorio`, `TemasAvanzados`): como todas las asignaturas del grupo comparten la misma franja, se toma el máximo en lugar de sumar, evitando contar múltiples veces las mismas horas.
- **Referencias directas** para asignaturas individuales (`ProcesoDesarrollo`, etc.): cada una ocupa su propia franja, por lo que se suman directamente.

Ejemplo de fórmula generada:
```
=MAX(C5:C7)+MAX(C11:C13)+C8+C9+C10
```

Donde:
- `C5:C7` son las filas de asignaturas `Obligatorio` (comparten franja)
- `C11:C13` son las filas de asignaturas `TemasAvanzados` (comparten franja)
- `C8`, `C9`, `C10` son asignaturas `ProcesoDesarrollo` (individuales)

#### 4. Indicador de balance por fecha
La hoja debe permitir identificar fácilmente si una fecha está:

- dentro del rango esperado
- subutilizada
- sobrecargada

Esto puede mostrarse con:

- color
- texto breve
- o una fila auxiliar de validación

Ejemplo:
- `OK`
- `Baja carga`
- `Sobrecarga`

#### 5. Totales por tipo de día
La salida debe facilitar revisar la carga de:

- viernes
- sábados
- miércoles

Esto puede lograrse con filas auxiliares, columnas de apoyo o una tabla resumen complementaria.

---

### Formato visual esperado

La hoja debe privilegiar la lectura rápida.

#### Recomendaciones de formato
- celdas con horas centradas
- colores para días sin clase
- colores suaves para resaltar sobrecargas o balances correctos
- encabezados congelados
- filas y columnas fijas para facilitar navegación
- anchos de columna que permitan ver claramente las fechas
- separación visual entre matriz principal y resúmenes

#### Sombreado alternado por semana

Para facilitar la agrupación visual de las columnas por semana, se aplica un sombreado alternado:

- **semanas impares**: fondo blanco
- **semanas pares**: fondo gris claro
- **semanas con encuentro presencial**: fondo verde claro (aplica a toda la semana, tiene prioridad sobre el gris)

Este sombreado se aplica tanto a las filas de encabezado (semana, día, fecha) como a las celdas de datos de asignaturas.

Las fechas especiales individuales (inducción, días sin clase) mantienen su color propio (dorado y rojo respectivamente).

---

### Nivel de precisión requerido

En esta salida:

- no es necesario mostrar la franja horaria exacta en la celda
- sí es obligatorio mostrar la cantidad de horas
- sí debe ser posible reconstruir internamente la franja para el cálculo de carga efectiva diaria
- sí debe ser posible calcular internamente las horas por asignatura y las horas efectivas por fecha

---

### Uso esperado

Esta hoja debe permitir al usuario:

- validar la viabilidad de la programación
- detectar sobrecargas por día
- validar cumplimiento de horas por asignatura
- identificar inicio y cierre de cursos
- visualizar picos de carga académica
- ajustar manualmente o iterativamente la programación

---

### Restricciones clave

- evitar sobrecarga diaria
- evitar duplicar horas por franjas compartidas
- evitar que múltiples cursos críticos inicien al mismo tiempo
- evitar que múltiples cursos críticos terminen al mismo tiempo
- reflejar correctamente días bloqueados

---

### Consideraciones de implementación

- la estructura principal debe ser tipo matriz (asignaturas vs fechas)
- el contenido visible de cada celda debe ser solo el número de horas
- debe ser exportable a Excel
- debe permitir aplicar formato visual
- debe facilitar lectura humana rápida
- debe incluir resúmenes sin perder claridad de la matriz principal
- debe separar claramente:
  - la información por asignatura
  - la información por fecha
  - la validación de carga diaria

### Identificación de fechas especiales (formato visual)

La salida de versión de horas debe permitir identificar visualmente distintos tipos de fechas relevantes dentro del calendario.

Estas fechas deben diferenciarse mediante formato visual (colores en Excel).

---

### Tipos de fechas especiales

El sistema debe reconocer y marcar al menos los siguientes tipos de fechas:

#### 1. Clases presenciales

Definidas en la hoja `parametros` mediante:

- `VIERNES_PRESENCIAL_UNO`
- `SABADO_PRESENCIAL_UNO`
- `VIERNES_PRESENCIAL_DOS`
- `SABADO_PRESENCIAL_DOS`

Estas fechas pueden cambiar en el Excel, por lo que deben leerse dinámicamente.

##### Comportamiento esperado

- deben resaltarse con un color distintivo
- deben ser fácilmente identificables en la matriz
- aplican a las columnas correspondientes a esas fechas

---

#### 2. Días sin clase

Incluyen:

- festivos
- semana sin clases
- cualquier fecha bloqueada definida en `parametros`

##### Comportamiento esperado

- deben resaltarse con un color diferente al de clases presenciales
- deben permitir identificar rápidamente interrupciones del calendario
- las celdas de asignaturas pueden permanecer vacías

---

#### 3. Día de inducción

Definido en:

- `FECHA_INDUCCION`

##### Comportamiento esperado

- debe resaltarse con un color único (diferente a los anteriores)
- permite identificar el inicio institucional del periodo
- puede no tener clases programadas

---

### Reglas de prioridad visual

Si una fecha cumple múltiples condiciones (caso poco probable pero posible), el sistema debe aplicar una jerarquía de visualización clara, por ejemplo:

1. inducción
2. día sin clase
3. clase presencial

(Esta jerarquía puede ajustarse según necesidad.)

---

### Ubicación del formato

El formato visual debe aplicarse principalmente a:

- encabezados de las columnas (fechas)
- filas auxiliares de contexto (si existen)
- opcionalmente a toda la columna de esa fecha

---

### Consideraciones de implementación

- el sistema debe leer estas fechas desde la hoja `parametros`
- no deben estar codificadas en el código
- el exportador a Excel (`exports_matriz.py`) debe encargarse del formato
- la lógica de identificación de fechas especiales debe estar separada de la lógica de asignación de horarios

---

### Objetivo

Esta diferenciación visual permite:

- entender rápidamente la estructura del semestre
- identificar semanas especiales
- revisar la programación con contexto institucional
- mejorar la lectura y validación humana del calendario

### 13.4 Flujo de ajuste manual con la matriz

El ciclo de trabajo típico para ajustar la programación es:

1. Ejecutar **opción 1** (`main.py`) para generar la programación automática y exportar `outputs/programacion_matriz.xlsx`.
2. Copiar `outputs/programacion_matriz.xlsx` a `inputs/programacion_matriz.xlsx`.
3. Editar manualmente las horas en la matriz (mover sesiones, ajustar cargas, corregir excepciones).
4. Ejecutar **opción 2** (`main.py`) para leer la matriz ajustada y regenerar `programacion_horas.xlsx`, `programacion_franjas.xlsx` y `programacion_visual.xlsx`.
5. Revisar los resultados. Si se necesitan más ajustes, repetir desde el paso 3.

La matriz nunca se sobreescribe en la opción 2: siempre es la fuente de verdad del ajuste manual.

### 13.5 Salida posterior: versión gráfica
Después de tener la programación lógica en horas, el sistema deberá poder transformarla en una versión gráfica tipo calendario.

La versión gráfica no es la prioridad inicial, pero el modelo de datos debe diseñarse pensando en soportarla después.

---

## 14. Arquitectura del sistema

El sistema usa una arquitectura modular de archivos planos en Python.
Cada archivo tiene una única responsabilidad. Los modelos no tienen lógica.
La lógica no conoce el Excel. La lectura no construye calendarios.

### `models.py` ✅ iteración 1
Estructuras de datos puras (enums y dataclasses):

- `DiaSemana` — enum: MIERCOLES, VIERNES, SABADO
- `SemestreOferta` — enum: PRIMERO, SEGUNDO, AMBOS
- `RestriccionProgramacion` — enum: MISMA_FRANJA, OBLIGATORIOS_MISMA_FRANJA, NO_CRUCES, SOLO_MIERCOLES, SIN_RESTRICCION
- `Franja` — nombre, hora_inicio, hora_fin, duracion_minutos, dia_semana
- `Asignatura` — todos los campos del catálogo
- `Parametros` — parámetros de la corrida

### `config.py` ✅ iteración 1
Constantes puras (sin lógica, sin imports del proyecto):

- nombres de hojas y columnas del Excel
- claves esperadas en la hoja `parametros` (ya normalizadas)
- textos esperados en las columnas del catálogo
- nombres de columnas del DataFrame de calendario

### `excel_reader.py` ✅ iteración 1
Único punto de contacto con el archivo Excel:

- `leer_excel` — abre el archivo, valida hojas, retorna DataFrames crudos
- `parsear_franjas` — convierte la hoja `franjas` en lista de `Franja`
- `parsear_parametros` — convierte la hoja `parametros` en `Parametros`; normaliza claves (strip + espacios → _)
- `parsear_catalogo` — convierte la hoja `catalogo` en lista de `Asignatura`; valida franjas referenciadas

### `calendar_builder.py` ✅ iteración 1
Construye el universo de slots disponibles como un DataFrame:

- `construir_calendario` — retorna DataFrame con columnas: `fecha`, `dia_semana`, `nombre_franja`, `hora_inicio`, `hora_fin`, `duracion_mins`, `es_programable`, `motivo_bloqueo`
- cada fila representa una combinación (fecha × franja) del periodo
- marca bloqueados los festivos y la semana sin clases

### `scheduler.py` ✅ iteración 4 + reglas 8.15, 8.16, 8.17, 8.18
Motor de asignación de sesiones con restricciones de tipo:

- `filtrar_asignaturas_del_semestre` — filtra por semestre de oferta
- `construir_candidatos` — genera combinaciones válidas (asignatura × slot)
- `asignar_sesiones` — asigna sesiones semana por semana respetando todas las reglas:
  - grupos `MismaFranja` y `ObligatoriosMismaFranja` se asignan primero con franja fija compartida
  - asignaturas `NoCruces` seleccionan dinámicamente la mejor franja libre, priorizando las más rezagadas
  - asignaturas `SoloMiercoles` usan su franja fija independiente
  - se rastrea `slots_ocupados` para evitar cualquier cruce entre tipos
  - permite sesiones parciales al final del ciclo cuando las horas pendientes son menores a una sesión completa
  - **regla 8.15**: para grupos `TemasAvanzados`, sesión intensiva (5-6h) en semana del primer encuentro presencial y bloqueo de semana del segundo encuentro presencial
  - **regla 8.16**: continuidad obligatoria de asignaturas ya iniciadas, sin huecos entre semanas
  - **regla 8.17**: fase de relleno para completar horas pendientes en fechas subutilizadas (máx 7h/día)
  - **regla 8.18**: máximo 6 horas de una misma asignatura por fin de semana

### `validators.py` — iteración futura
Responsable de:

- revisar horas acumuladas
- revisar conflictos de horario
- revisar restricciones temporales por asignatura
- revisar reglas por tipo de curso

### `exports_hours.py` ✅ iteración 5 (parcial)
Genera la salida de la versión de horas en `outputs/programacion_horas.xlsx`:

- hoja `programacion`: detalle sesión por sesión ordenado por fecha y hora,
  con número de semana relativo al inicio del semestre.
- hoja `resumen`: estado de cumplimiento de horas por asignatura
  (completa / incompleta / excede).

### `exports_franjas.py` ✅ iteración 5
Genera la salida consolidada de franjas en `outputs/programacion_franjas.xlsx`:

- hoja `franjas`: un registro por cada bloque continuo de sesiones con la misma franja horaria
- campos: Código, Asignatura, Tipo, Franja, Día, Hora inicio, Hora fin,
  Semana inicio, Fecha inicio, Semana fin, Fecha fin, Sesiones, Horas en bloque
- dos semanas son consecutivas si sus números difieren en exactamente 1;
  cualquier brecha genera un corte y produce un nuevo bloque

### `exports_visual.py` [Por refinar] iteración 6
Genera la vista gráfica en `outputs/programacion_visual.xlsx`:

- hoja `calendario`: cuadrícula con franjas como filas y semanas como columnas
- encabezado de columna: número de semana y fecha del lunes ("S1 / 31-Jul")
- celda: nombre(s) de asignatura programados en esa franja y semana
- color de fondo por tipo de asignatura:
  - Obligatorio → naranja claro
  - TemasAvanzados → purpura claro
  - ProcesoDesarrollo → azul claro
  - ProyectoGrado → verde claro
- paneles congelados en B2 para facilitar la navegación


### `exports_matriz.py` ✅ iteración 7
Genera la matriz de horas en `outputs/programacion_matriz.xlsx`:

- hoja `matriz`: cuadrícula asignaturas × fechas con horas por celda
- columnas iniciales: Tipo y Asignatura (solo nombre, sin código)
- incluye TODAS las fechas del calendario (incluyendo días sin clase)
- filas auxiliares de contexto (semana, día)
- fila resumen de horas efectivas por fecha con **fórmulas Excel**:
  - usa `MAX()` para grupos de franja compartida (Obligatorio, TemasAvanzados)
  - suma directa para asignaturas individuales (ProcesoDesarrollo, etc.)
  - permite ajustes manuales con recálculo automático
- columnas finales de totales y estado por asignatura
- formato visual para fechas especiales:
  - días sin clase (festivos, semana sin clases): fondo rojo claro
  - día de inducción: fondo dorado
- sombreado alternado por semana para facilitar agrupación visual:
  - semanas impares: fondo blanco
  - semanas pares: fondo gris claro
  - semanas con encuentro presencial: fondo verde claro (aplica a toda la semana)
- paneles congelados en C5 para mantener visibles Tipo, Asignatura y encabezados

### `main.py` ✅ iteración 7
Orquestador del flujo principal con menú de dos modos de ejecución:

**Opción 1 — Generar programación automática:**
Lee `inputs/restricciones.xlsx`, ejecuta el scheduler y exporta los cuatro
archivos de salida en `outputs/`.

**Opción 2 — Regenerar desde matriz ajustada:**
Lee `inputs/programacion_matriz.xlsx` (copia ajustada manualmente por el
usuario), reconstruye el DataFrame de sesiones a partir de las horas registradas
en la matriz y regenera `programacion_horas.xlsx`, `programacion_franjas.xlsx`
y `programacion_visual.xlsx`. La matriz en sí no se sobreescribe (es la fuente
del ajuste).

Este segundo modo permite al usuario aplicar correcciones manuales a la
programación automática y luego obtener las demás salidas actualizadas, sin
necesidad de re-ejecutar el scheduler.

**Asignaturas adicionales:** El sistema conserva las filas que el usuario
agregue manualmente en la matriz aunque no existan en el catálogo original.
Para estas asignaturas, se crean sesiones con los campos disponibles (nombre,
tipo, fecha, horas) dejando vacíos los campos que no se pueden inferir (código,
franja, hora inicio/fin).

---

## 15. Orden recomendado de desarrollo

### Iteración 1
- modelos de datos
- lector del Excel
- lectura de hoja `parametros`
- lectura de hoja `franjas`
- constructor de calendario base

### Iteración 2
- selección de asignaturas del conjunto a programar
- lectura de reglas por curso
- validación de franjas permitidas
- salida preliminar de calendario válido

### Iteración 3 ✅
- motor de asignación greedy semana por semana
- selección automática de sesión semanal por asignatura (basada en `MinSemanasClase`)
- acumulación de horas por asignatura
- sesiones que abarcan múltiples franjas en la misma semana (incluso días distintos)

### Iteración 4 ✅
- restricción `MismaFranja` y `ObligatoriosMismaFranja`: asignaturas del mismo tipo comparten franja fija durante todo el semestre
- restricción `NoCruces`: cada asignatura ocupa un slot exclusivo, sin coincidir con ninguna otra asignatura de ningún tipo
- no coincidencia entre tipos distintos en la misma (fecha, franja)
- priorización dinámica semanal de asignaturas `NoCruces` por horas pendientes
- sesiones parciales de fin de ciclo: cuando las horas pendientes son menores a una sesión completa, se asigna el subconjunto de franjas que cabe dentro del remanente

### Iteración 5 ✅
- exportación de la versión de horas a `outputs/programacion_horas.xlsx`
  - hoja `programacion`: detalle sesión por sesión
  - hoja `resumen`: estado de cumplimiento por asignatura
- exportación de la versión de franjas a `outputs/programacion_franjas.xlsx`
  - hoja `franjas`: bloques continuos consolidados por asignatura y franja

### Iteración 6 ✅
- primera version exportación de la versión visual a `outputs/programacion_visual.xlsx`
  - hoja `calendario`: cuadrícula franjas × semanas con colores por tipo

### Iteración 7 ✅ (parcial)
- salida matriz de horas a `outputs/programacion_matriz.xlsx` con formato para revisión humana ✅
- validaciones post-asignación (horas completas, sin cruces, restricciones por asignatura) — pendiente

### Iteración 8 ✅
- implementar regla 8.15 de TemasAvanzados y encuentros presenciales:
  - sesión intensiva (5-6 horas) en semana del primer encuentro presencial
  - bloqueo de programación en semana del segundo encuentro presencial
  - pasar `parametros` a `asignar_sesiones()` para conocer fechas de encuentros presenciales

### Iteración 9 ✅
- implementar regla 8.16 de continuidad de asignaturas:
  - registrar semana de inicio de cada asignatura
  - priorizar continuidad en semanas siguientes (asignaturas ya iniciadas primero)
  - solo permitir interrupciones por excepciones válidas (festivos, semana sin clases, restricciones)
- implementar regla 8.17 de relleno para horas pendientes:
  - fase de relleno después de asignación principal
  - distribuir horas en fechas con menos de 7 horas efectivas
  - priorizar asignaturas con más horas faltantes
- implementar regla 8.18 de máximo 6 horas por asignatura por fin de semana:
  - limitar selección de franjas a máximo 6 horas por semana
  - aplicar límite en asignación principal y fase de relleno

### Iteración 10 ✅
- implementar regla 8.19 de máximo 4 horas por asignatura por día:
  - `_seleccionar_mejor_subconjunto` acepta `max_minutos_por_dia` y descarta combinaciones que superan el tope en cualquier día
  - aplica en selección de franja común, asignaturas NoCruces y fase de relleno
  - excepción: `_seleccionar_franjas_sesion_intensiva` (TemasAvanzados presencial) no aplica el límite diario

### Iteración 11 ✅
- fechas bloqueadas por asignatura (columna `Restricciones` del catálogo):
  - formato: rangos `dd/mm/yyyy - dd/mm/yyyy`, uno por línea
  - `excel_reader` parsea el texto en lista de tuplas `(fecha_inicio, fecha_fin)` almacenada en `Asignatura.fechas_bloqueadas`
  - `scheduler` excluye de los candidatos cualquier slot cuya fecha caiga dentro de un rango bloqueado
  - `models.Asignatura`: campo `restricciones_texto` reemplazado por `fechas_bloqueadas`
- reconstrucción desde matriz ajustada (opción 2) — pendiente de corrección:
  - usa `construir_candidatos` para determinar franjas válidas por asignatura y fecha
  - distribuye horas de la matriz respetando las restricciones de franja y el calendario
  - aún presenta errores en la asignación de franjas

---

## 16. Prioridades de diseño

El sistema debe priorizar:

1. claridad del modelo de datos
2. lectura robusta del Excel
3. trazabilidad de las decisiones de programación
4. facilidad para ajustar reglas futuras
5. separación entre lógica de programación y presentación visual
6. distribución equilibrada de la carga semanal
7. reducción de picos de inicio y cierre de asignaturas
8. alta mantenibilidad y legibilidad del código

---

## 17. Restricciones de implementación recomendadas

- usar Python
- usar `pandas` para lectura de Excel
- usar `dataclasses` para los modelos
- mantener funciones pequeñas y separadas
- evitar acoplar la lógica de programación a una plantilla visual específica
- diseñar la programación lógica primero y la visualización después

### 17.1 Criterio de mantenibilidad del código
El código debe ser especialmente fácil de mantener, entender y modificar por otras personas en el futuro.

Por esta razón, la implementación debe priorizar:

- nombres de variables, funciones y clases claros y descriptivos
- funciones cortas y con una sola responsabilidad
- separación clara entre lectura de datos, reglas de negocio, validaciones y exportación
- comentarios útiles cuando ayuden a entender decisiones importantes
- docstrings en módulos, clases y funciones principales
- lógica explícita en lugar de soluciones excesivamente compactas

El código debe evitar, salvo que exista una razón muy justificada, el uso de construcciones que dificulten la comprensión del sistema, por ejemplo:

- comprensiones demasiado complejas
- expresiones muy condensadas
- encadenamientos difíciles de leer
- uso excesivo de `lambda`
- trucos de Python que hagan el código menos claro
- lógica implícita difícil de rastrear

Se prefiere un estilo de programación más explícito, legible y mantenible, incluso si eso produce un poco más de código.

---

## 18. Decisiones ya tomadas

- El insumo principal será siempre un Excel con la estructura del archivo actual.
- La primera versión usará esa estructura tal como está.
- El Excel de entrada tendrá tres hojas obligatorias: `catalogo`, `parametros` y `franjas`.
- Se iniciará con la programación de segundo semestre.
- La salida prioritaria será la versión de horas.
- Después se generará la salida de franjas.
- La versión gráfica vendrá después.
- Miércoles se usará únicamente para Proyecto de grado 2.
- Se permiten bloques parciales, pero deben evitarse en lo posible.
- La regla `MismaFranja` significa misma franja exacta, no solo mismo bloque general.
- La lógica horaria debe leerse desde la hoja `franjas`.
- No se usará el concepto de `BloqueGeneral` como modelo: es un concepto derivado
  que se calculará como función auxiliar cuando se necesite (exportación).
- El `calendar_builder` retorna un DataFrame plano de slots (fecha × franja),
  sin cruzar con asignaturas. El scheduler es quien asigna asignaturas a slots.
- Cada asignatura puede usar múltiples franjas en la misma semana (incluso días distintos).
  La selección del conjunto de franjas por semana se basa en `MinSemanasClase` y `Horas`.
- La salida parcial se exporta a `outputs/programacion_horas.xlsx` con dos hojas:
  `programacion` (detalle sesión por sesión) y `resumen` (estado por asignatura).
- La clase de parámetros se llama `Parametros` (no `ParametrosCorreida`).
- La hoja `parametros` se lee sin encabezados. Las claves se normalizan
  (strip + espacios → _) para tolerar inconsistencias menores del Excel.
- El sistema genera cuatro archivos de salida en `outputs/`:
  `programacion_horas.xlsx` (detalle sesión a sesión),
  `programacion_franjas.xlsx` (bloques continuos consolidados),
  `programacion_visual.xlsx` (cuadrícula calendario con colores),
  `programacion_matriz.xlsx` (matriz de horas por asignatura y fecha para revisión humana).
- La función `calcular_numero_semana` es pública en `exports_hours.py`
  y es compartida por los tres módulos de exportación.
- Las restricciones de tipo se implementan con un único mecanismo de `slots_ocupados`:
  los grupos `MismaFranja`/`ObligatoriosMismaFranja` procesan primero y reservan sus slots;
  los `NoCruces` luego seleccionan franjas libres. Esto garantiza la no coincidencia entre tipos.
- Los grupos `MismaFranja` seleccionan su franja fija una sola vez al inicio de la corrida
  (intersección de franjas permitidas de todos los miembros, con el tope del miembro más restrictivo).
- Las asignaturas `NoCruces` se repriorizar cada semana según horas pendientes (más rezagada = primera),
  para evitar que las que tienen más opciones acaparen los mejores slots semana tras semana.
- Cuando la sesión seleccionada excede las horas pendientes, el sistema reintenta con el tope
  reducido a las horas que realmente faltan, en lugar de omitir la semana.
- El sistema opera en dos modos: generación automática (opción 1) y regeneración desde
  matriz ajustada (opción 2). La opción 2 lee `inputs/programacion_matriz.xlsx`,
  reconstruye las sesiones a partir de las horas en la matriz y regenera todas las
  salidas excepto la matriz misma (que es la fuente del ajuste manual).
- La reconstrucción de sesiones desde la matriz asigna franjas en orden definición
  (la primera franja del día que cubre las horas indicadas) y recalcula
  `horas_acumuladas` mediante cumsum ordenado por fecha.

---

## 19. Preguntas abiertas o puntos a afinar después

Aunque ya existe una base bastante sólida, hay temas que probablemente se podrán refinar en iteraciones posteriores, por ejemplo:

- cómo desempatar entre varias soluciones válidas
- cómo seleccionar automáticamente la mejor franja compartida
- qué hacer cuando una asignatura no cabe completamente sin usar parciales
- cómo parsear de forma robusta las restricciones específicas en lenguaje natural
- cómo transformar la salida lógica a la plantilla gráfica final

---

## 20. Objetivo inmediato para Claude

A partir de este documento, Claude debe empezar construyendo una primera base técnica que:

1. lea correctamente el Excel
2. valide hojas y columnas
3. construya el calendario académico programable
4. represente cursos, franjas, parámetros y eventos con modelos claros
5. genere la base para una salida detallada de horas
6. genere una salida consolidada de franjas
7. deje lista la base para construir luego el motor de asignación completo y la versión gráfica
8. mantenga un estilo de código explícito, claro, bien documentado y fácil de mantener

## Regla transversal de desarrollo

En este proyecto se debe favorecer la claridad sobre la sofisticación sintáctica.

Es preferible escribir código más largo pero entendible, que código más corto pero difícil de mantener.