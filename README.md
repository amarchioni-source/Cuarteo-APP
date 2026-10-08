# Cuarteo final

App web que recibe la **bajada del sistema** (Anexo Cuarteo y Troceo, circular 3510) y genera el
**CUARTEO final** en Excel: formulario con tropas, medias y kg de ingreso, más el bloque de
cálculo por tropa con el peso promedio por media tomado siempre de la bajada.

## Cómo calcula

- **Cantidad** de cada producto = medias de la tropa.
- **Kg** = medias × peso promedio por media.
- **Peso promedio por media** = kg de la categoría en la bajada (todas las tropas) ÷ total de medias del día.
- Cinco categorías: Asado CV, Delantero SF, Entraña Fina, Falda y Pistola.
- La asignación producto → categoría queda a la vista en la hoja **Peso promedio** del archivo generado
  (columna «Categoría final», editable). Los criterios por defecto están en `REGLAS` de `generar_cuarteo.py`:

| Categoría | Productos de la bajada |
|---|---|
| Asado CV | ASADO CV, ASADO SV (banderita), VACIO COLGADO |
| Delantero SF | DELANTERO (completo, musulmán, sin falda), PECHO CB SF |
| Entraña Fina | ENTRAÑA (AA, World Class) |
| Falda | FALDA |
| Pistola | PISTOLA 2C, R&L 2 COST, RUEDA |
| No va | DECOMISO, RECORTE, FALSA ENTRAÑA, VARIOS |

Si aparece un producto que no coincide con ninguna regla, la app lo informa como «producto nuevo»
y queda fuera del final hasta que se lo asigne.

## Uso

1. Abrir la app y subir el archivo de la bajada (`.xlsx` del sistema).
2. Opcional: fecha de faena (si se deja vacía, se toma el día hábil anterior; no contempla feriados)
   y número de serie (si se deja vacío, queda el de la plantilla).
3. **Generar final** y descargar. La pantalla muestra el peso promedio por media, el control por tropa
   contra el total de la bajada y los avisos a revisar.

El archivo se recalcula al abrirlo en Excel (las fórmulas no traen valores guardados).

## Despliegue en Render

1. Crear un repositorio **privado** en GitHub y subir esta carpeta (contiene la plantilla del formulario
   y una bajada de ejemplo para las pruebas).
2. En Render: **New + > Blueprint** y elegir el repositorio. Render lee `render.yaml`.
   Alternativa manual: **New + > Web Service**, runtime Python, Build `pip install -r requirements.txt`,
   Start `gunicorn app:app --workers 2 --threads 2 --timeout 120`, Health check `/salud`.
3. En **Environment** cargar `APP_USER` y `APP_PASSWORD`. Con `APP_PASSWORD` definida la app pide usuario
   y clave; sin ella queda abierta, por lo que conviene definirla siempre.
4. Cada `git push` a la rama principal vuelve a desplegar.

Notas: en el plan gratuito el servicio se suspende por inactividad y la primera carga tarda unos
segundos. Los archivos generados se guardan en disco temporal y se borran a las 2 horas; no se
conserva ninguna bajada.

## Mantenimiento

- **Plantilla del formulario:** `CUARTEO_plantilla.xlsx`. Es el CUARTEO final de un día
  cualquiera; el generador conserva todo lo que está arriba de la fila 18 (encabezado, logo, serie) y
  el formato de las filas. Para cambiarla, reemplazar el archivo manteniendo la estructura.
- **Criterios de categorías:** editar `REGLAS` en `generar_cuarteo.py` y volver a correr las pruebas.

## Desarrollo local

```bash
python -m venv .venv && source .venv/bin/activate      # En Windows: .venv\Scripts\activate
pip install -r requirements-dev.txt
python app.py                                          # http://localhost:5000
pytest                                                 # pruebas
```

Uso por línea de comandos, sin la web:

```bash
python generar_cuarteo.py bajada.xlsx CUARTEO_plantilla.xlsx final.xlsx --serie 2291
```

## Archivos

Todos los archivos van sueltos en la raíz del repositorio (sin carpetas).

```
app.py                      servidor Flask (carga, generación, descarga, acceso)
generar_cuarteo.py          lectura de la bajada, cálculo y armado del Excel
index.html                  interfaz (con los estilos incluidos)
CUARTEO_plantilla.xlsx      formulario base del final
bajada_0610.xlsx            bajada de ejemplo, usada solo por las pruebas
test_app.py, test_generar.py  pruebas
requirements.txt            dependencias de la app
requirements-dev.txt        dependencias para correr las pruebas
render.yaml                 configuración de Render
.gitignore
```
