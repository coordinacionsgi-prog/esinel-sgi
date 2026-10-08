# I-AMB2 · Generación de Consumos — Script y dashboard

Planilla: **I-AMB2 - Consumos EJ 35**
(`1kjJgtY3Q9WT7-8gFpaQw09JvAtyNPUXxkkUyQj2d7_U`)

Reemplaza el script y el dashboard que había armado Gemini.

## Estado actual (08/10/2026)

Ya aplicado directamente en la planilla, sin script:

- Pestaña **Dashboard** hecha solo con fórmulas y gráficos nativos. Lee cada fila de la ficha AMB2,
  detecta la hoja `HASTA …` desde la fórmula de la ficha y trae resultados, nómina y consumos de las 3 sedes.
  Tiene selector de ejercicio (`Último`, `Todos` o un número).
- Ficha AMB2: agregada la fila Jul–Sep 2026 (los 3 gráficos de la ficha la incluyen). El link "Dashboard" apunta a la pestaña nueva.
- Fechas DESDE/HASTA y textos de trimestre/ejercicio corregidos en todas las hojas `HASTA …`.
- "Total promedio" (`N21`, `AG21`, `AZ21`, `M24`) de las 4 hojas más nuevas: divide por la nómina total en vez de 36 fijo.

Para cada trimestre nuevo alcanza con lo de siempre: duplicar la última hoja `HASTA …`, cargar los datos y agregar
la fila en la ficha AMB2 (antes de "Objetivo"). El dashboard se actualiza solo.

El dashboard en HTML (página web, como el de Gemini) se instala con el script de abajo.

## Instalación (una sola vez, ~2 minutos)

Todo va en **un solo archivo**: [`I-AMB2-completo.gs`](I-AMB2-completo.gs) (script + página HTML del dashboard).

1. Abrí la planilla → **Extensiones → Apps Script**.
2. Borrá los archivos de Gemini (dejá solo `Código.gs`/`Code.gs`) y en ese archivo **reemplazá todo** por el contenido de `I-AMB2-completo.gs`. Guardá (Ctrl+S).
3. **Implementar → Administrar implementaciones → ✏️ (lápiz) → Versión: Nueva versión → Implementar.**
   Así el link del dashboard sigue siendo el mismo que tenía Gemini (celda A60 de la ficha).
   Si no hay ninguna implementación: **Implementar → Nueva implementación → Aplicación web**.
4. Recargá la planilla → menú **Indicador AMB2 → Activar actualización automática** (acepta los permisos).

Listo. El dashboard web se calcula en el momento cada vez que se abre, y cada hora se actualiza la ficha y se
exporta `I-AMB2 Dashboard.html` (con los datos adentro) en la carpeta de la planilla.

`Code.gs` + `Dashboard.html` son las fuentes por separado; `I-AMB2-completo.gs` se genera uniéndolas.

## Qué hace

| Menú | Qué hace |
|---|---|
| Abrir dashboard | Abre el dashboard dentro de la planilla. |
| Actualizar ficha AMB2 | Agrega a la ficha (y a sus 3 gráficos) los trimestres que falten, con las fórmulas apuntando a la hoja correcta. Hoy falta **Jul–Sep 2026**. |
| Corregir encabezados de períodos | Reescribe trimestre, ejercicio, DESDE y HASTA de cada hoja según su nombre. **No toca consumos.** Pide confirmación. |
| Crear hoja del próximo trimestre | Duplica la última hoja `HASTA …`, la nombra con el trimestre siguiente, pone las fechas, deja los consumos en blanco y la agrega a la ficha. |
| Revisar datos | Lista las inconsistencias de la planilla. |
| Exportar HTML ahora | Guarda/actualiza `I-AMB2 Dashboard.html` en la carpeta de la planilla. |
| Activar actualización automática | Cada hora: pone la ficha al día y exporta el HTML. |

El dashboard (`/exec`) se calcula en el momento, leyendo directamente las hojas `HASTA dd-mm-aa`. No hace falta completar nada a mano: alcanza con cargar los consumos en la hoja del trimestre.
`/exec?formato=json` devuelve los mismos datos en JSON.

## Reglas de cálculo

- **Período**: sale del **nombre** de la hoja (`HASTA 30-09-26` = Jul–Sep 2026), no de las celdas DESDE/HASTA.
  Ejercicio de abril a marzo: Ej. 35 = abr-2026 a mar-2027.
- **Resultado por sede**: celdas de resultado `D20` (Remedios), `W20` (Santo Tomé) y `AP20` (Mosconi).
- **Resultado general**: promedio de las 3 sedes **ponderado por nómina**: Σ(resultado × nómina) / Σ nómina.
  La planilla dividía por un **36 fijo** (`M24`/`N21`), que solo daba bien con 22+10+4 empleados.
- **Promedio del ejercicio**: promedio de los trimestres cargados.
- **Objetivo**: la celda a la derecha de "Objetivo" en la ficha (hoy 8).
- Una sede sin consumos cargados en un trimestre no cuenta (antes las celdas vacías daban nota 10).

## Errores que tiene la planilla hoy (los muestra el dashboard)

- `HASTA 30-09-26`, Remedios: DESDE 01/07/2026 y HASTA 30/06/2026 (−1 días).
- `HASTA 30-09-26`, Santo Tomé y Mosconi: DESDE 01/06/2026 (debería ser 01/07).
- `HASTA 30-09-26`, Remedios: **nómina 90** (el trimestre anterior era 22). Si es un error de tipeo, el general pasa de 8,47 a 8,59.
- Ficha AMB2: falta la fila de **Jul–Sep 2026**, así que los gráficos de la ficha no lo muestran.
- Varias hojas de 2024–2026 tienen fechas y "EJERCICIO 33" copiados del trimestre anterior en Santo Tomé y Mosconi.

Las fechas y los textos de ejercicio se corrigen con **Corregir encabezados de períodos**. La nómina hay que revisarla a mano.
