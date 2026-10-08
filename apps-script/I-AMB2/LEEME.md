# I-AMB2 · Generación de Consumos — Script y dashboard

Planilla: **I-AMB2 - Consumos EJ 35**
(`1kjJgtY3Q9WT7-8gFpaQw09JvAtyNPUXxkkUyQj2d7_U`)

Reemplaza el script y el dashboard que había armado Gemini.

## Instalación (una sola vez, ~3 minutos)

1. Abrí la planilla → **Extensiones → Apps Script**.
2. **Borrá todos los archivos** del proyecto (los `.gs` y `.html` de Gemini).
3. Creá `Code.gs` y pegá el contenido de [`Code.gs`](Code.gs).
4. **Archivo + → HTML**, llamalo exactamente `Dashboard` y pegá [`Dashboard.html`](Dashboard.html).
5. Guardá. Volvé a la planilla y recargala: aparece el menú **Indicador AMB2**.
6. Para mantener el **mismo link** del dashboard (celda A59 de la ficha):
   **Implementar → Administrar implementaciones → ✏️ → Versión: Nueva versión → Implementar**.
   No crees una implementación nueva, porque cambia la URL.
7. En el menú **Indicador AMB2 → Activar actualización automática (diaria)** (pide permisos la primera vez).

## Qué hace

| Menú | Qué hace |
|---|---|
| Abrir dashboard | Abre el dashboard dentro de la planilla. |
| Actualizar ficha AMB2 | Agrega a la ficha (y a sus 3 gráficos) los trimestres que falten, con las fórmulas apuntando a la hoja correcta. Hoy falta **Jul–Sep 2026**. |
| Corregir encabezados de períodos | Reescribe trimestre, ejercicio, DESDE y HASTA de cada hoja según su nombre. **No toca consumos.** Pide confirmación. |
| Crear hoja del próximo trimestre | Duplica la última hoja `HASTA …`, la nombra con el trimestre siguiente, pone las fechas, deja los consumos en blanco y la agrega a la ficha. |
| Revisar datos | Lista las inconsistencias de la planilla. |
| Actualización automática | Todos los días a las 7 h pone la ficha al día. |

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
