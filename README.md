**App en vivo:** <https://elluciangradeimporter.streamlit.app/>

Herramienta que convierte un Excel con notas de estudiantes en un archivo
compatible con los requisitos de importación de la plataforma **Banner de
Ellucian**.

Subir notas manualmente en Banner es lento y propenso a errores. Esta
herramienta toma un Excel arbitrario, permite seleccionar la columna del ID
del estudiante, la columna de notas y el componente, y genera un archivo
listo para importar.

## Motivación

En la práctica, cada profesor pierde alrededor de una hora preparando
manualmente el archivo que Banner acepta. Esta herramienta reduce ese
proceso a subir un Excel, elegir dos columnas y descargar el resultado.

## Características

- Interfaz web simple con **Streamlit**.
- Procesamiento de datos con **Polars** (lectura y escritura de Excel).
- Selección dinámica de columnas a partir del archivo subido.
- Generación del archivo compatible con Banner en memoria, listo para
  descargar.
- Procesamiento sin persistencia en disco: el archivo se maneja en memoria.

## Instalación

Requiere Python 3.10+ y [uv](https://docs.astral.sh/uv/) para la gestión de
dependencias.

```bash
git clone https://github.com/rsgenus1/elluciangradeimporter.git
cd elluciangradeimporter
uv sync
