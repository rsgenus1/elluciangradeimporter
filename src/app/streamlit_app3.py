import hashlib
import re
from pathlib import Path

import polars as pl
import streamlit as st

from elluciangradeimporter.converter import (
    converter,
    count_id_coincidences_diff,
    detect_header_row,
    df_to_excel_bytes,
    get_cols,
    get_idc,
    get_sheet_names,
    read_idc_frame,
    read_raw_rows,
    read_sheet,
)

MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# Contacto y enlaces de apoyo
CONTACT_EMAIL = "ellucianimporter@proton.me"
MP_CAFE_URL = "https://mpago.li/2BxMn6E"
MP_MENSUAL_URL = "https://mpago.la/2RsGqp4"

# Alias compatible: pl.String en versiones recientes, pl.Utf8 en todas
STRING = getattr(pl, "String", pl.Utf8)


# RUT chileno: exactamente 9 caracteres, 8 dígitos + dígito verificador (0-9 o K).
ID_ESTUDIANTE_RE = re.compile(r"^\d{8}[\dKk]$")
# Un número decimal (p. ej. "12345678.0") no es un RUT formateado.
DECIMAL_RE = re.compile(r"^\d+\.\d+$")


def _parece_id_estudiante(valor: str) -> bool:
    """True si, tras normalizar el valor, tiene exactamente 9 caracteres de RUT.

    Se quitan puntos, guiones y espacios (igual que hace ``converter`` al leer),
    y luego se exigen 9 caracteres: 8 dígitos + dígito verificador (0-9 o K).
    """
    valor = valor.strip()
    if DECIMAL_RE.match(valor):
        return False
    normalizado = re.sub(r"[.\-\s]", "", valor).upper()
    return bool(ID_ESTUDIANTE_RE.match(normalizado))


def columnas_rut_posibles(df: pl.DataFrame) -> list[str]:
    """Columnas cuyos VALORES parecen IDs de estudiante (no se usa el nombre).

    Se evalúan los valores de cada columna: se considera candidata si al menos
    el 80% de sus valores no vacíos, tras quitar puntos/guiones/espacios, tiene
    exactamente 9 caracteres con forma de RUT. Esto tolera filas de
    totales/promedios al final de la planilla. Las columnas se ordenan de mejor
    a peor coincidencia.
    """
    MINIMO_VALORES = 1
    PROPORCION_MINIMA = 0.8

    candidatas = []
    for name in df.columns:
        try:
            serie = df[name].cast(STRING).str.strip_chars().drop_nulls()
        except Exception:
            continue
        serie = serie.filter(serie != "")
        total = serie.len()
        if total < MINIMO_VALORES:
            continue

        validos = sum(1 for valor in serie.to_list() if _parece_id_estudiante(valor))
        proporcion = validos / total
        if proporcion >= PROPORCION_MINIMA:
            candidatas.append((proporcion, validos, name))

    candidatas.sort(key=lambda item: (-item[0], -item[1]))
    return [name for _, _, name in candidatas]


@st.dialog("⚠️ Posible confusión de planillas")
def dialog_idc_mismatch(coincidencias: int, diferencias: int) -> None:
    st.markdown(
        f"En la planilla de notas, **{diferencias}** estudiantes "
        f"**no coinciden** con la planilla del IDC, y solo "
        f"**{coincidencias}** coinciden."
    )
    st.markdown(
        "Esto suele ocurrir cuando:\n"
        "- Se subieron las **planillas equivocadas** (otro curso / otra sección).\n"
        "- La columna seleccionada como **ID del estudiante** no es la correcta.\n\n"
        "**Verifica los archivos y la selección antes de usar el resultado.**"
    )
    if st.button("Entendido", use_container_width=True):
        st.rerun()


st.title("Convertidor para Banner de Ellucian")
st.markdown(
    """
    Subir notas manualmente en la plataforma **Banner de Ellucian** es lento y
    propenso a errores. Esta herramienta toma un Excel con las notas de los
    estudiantes y genera un archivo compatible con los requisitos de importación
    de Banner.

    **Cómo usarla:**
    1. Sube el Excel con las notas.
    2. Si el archivo tiene varias hojas, elige la que contiene las notas.
    3. Verifica que la fila de nombres de columnas sea la correcta.
    4. Selecciona la columna del ID del estudiante y la columna de notas.
    5. Indica el componente (ej. `C01`, `C02`, ...).
    6. Indica el ID de la componente, si no lo conoce, descarga la planilla desde Banner, y súbela, se detectará automáticamente.
    7. Descarga el archivo generado e impórtalo en Banner.
    """
)


def _reset_result() -> None:
    """Descarta un resultado previo cuando cambia la hoja o el encabezado."""
    for key in (
        "result_df",
        "result_xlsx",
        "result_component",
        "result_filename",
        "idc_mismatch",
    ):
        st.session_state.pop(key, None)


def _preview_row(row: tuple, max_cells: int = 7, max_chars: int = 22) -> str:
    """Resumen corto de una fila para mostrarla en el selector de encabezado."""
    partes = []
    for value in row[:max_cells]:
        text = "" if value is None else str(value).strip()
        if text:
            partes.append(text[:max_chars])
    return " | ".join(partes) if partes else "(fila vacía)"


# Archivos de Banner terminan en algo como "..._C01.xlsx" o "..._C01 (1).xlsx"
COMPONENT_FROM_NAME_RE = re.compile(r"_(C\d{1,2})\s*(?:\(\d+\))?$", re.IGNORECASE)


def _component_from_filename(name: str | None) -> str | None:
    """Extrae el componente (C01, C02, ...) del nombre del archivo, si existe."""
    if not name:
        return None
    match = COMPONENT_FROM_NAME_RE.search(Path(name).stem)
    return match.group(1).upper() if match else None


def _download_filename(idc_file, uploaded) -> str:
    """Nombre de salida: el original (IDC si existe) + ' importable.xlsx'."""
    base = idc_file.name if idc_file is not None else uploaded.name
    return f"{Path(base).stem} importable.xlsx"


uploaded = st.file_uploader("Sube el Excel", type=["xlsx", "xls"])

if uploaded is not None:
    file_bytes = uploaded.getvalue()
    # Identificador estable del contenido para reiniciar selectores al subir otro archivo
    file_id = hashlib.sha1(file_bytes).hexdigest()[:12]

    # 1) Hojas (worksheets) del workbook
    @st.cache_data(show_spinner="Leyendo hojas del archivo...")
    def _get_sheets(data: bytes) -> list[str]:
        return get_sheet_names(data)

    try:
        sheets = _get_sheets(file_bytes)
    except Exception as e:
        st.error(f"No se pudo leer el archivo: {e}")
        _reset_result()
        st.stop()

    if not sheets:
        st.error("El archivo no contiene hojas.")
        _reset_result()
        st.stop()

    # Si hay más de una hoja, el usuario elige cuál contiene las notas
    if len(sheets) > 1:
        sheet_name = st.selectbox(
            "El archivo tiene varias hojas. ¿Cuál contiene las notas?",
            sheets,
            on_change=_reset_result,
            key=f"sheet_select_{file_id}",
        )
    else:
        sheet_name = sheets[0]

    # 2) Fila con los nombres de columnas: detectar y permitir corregir
    @st.cache_data(show_spinner="Analizando filas de encabezado...")
    def _get_raw_rows(data: bytes, sheet: str, _v: int = 1) -> list[tuple]:
        return read_raw_rows(data, sheet)

    @st.cache_data(show_spinner="Detectando fila de encabezado...")
    def _detect_header(data: bytes, sheet: str, _v: int = 1) -> int:
        return detect_header_row(data, sheet)

    raw_rows = _get_raw_rows(file_bytes, sheet_name)
    if not raw_rows:
        st.warning("La hoja seleccionada está vacía.")
        _reset_result()
        st.stop()

    detected_header = _detect_header(file_bytes, sheet_name)
    detected_header = min(max(detected_header, 0), len(raw_rows) - 1)

    def _header_label(index: int) -> str:
        label = f"Fila {index + 1}: {_preview_row(raw_rows[index])}"
        if index == detected_header:
            label += "  ← detectada"
        return label

    header_row = st.selectbox(
        "Fila que contiene los nombres de las columnas",
        list(range(len(raw_rows))),
        index=detected_header,
        format_func=_header_label,
        on_change=_reset_result,
        help=(
            "Se detectó automáticamente la fila "
            f"{detected_header + 1}. Si no es correcta (por ejemplo, si hay un "
            "título con celdas combinadas sobre el encabezado), elige aquí la "
            "fila correcta."
        ),
        key=f"header_select_{file_id}_{sheet_name}",
    )
    st.caption(
        f"Encabezado detectado automáticamente: fila {detected_header + 1}. "
        "Puedes cambiarlo si no corresponde."
    )

    # 3) Columnas de la hoja/fila elegidas
    @st.cache_data(show_spinner="Leyendo columnas...")
    def _get_cols(data: bytes, sheet: str, header: int, _v: int = 1) -> list[str]:
        return get_cols(data, sheet_name=sheet, header_row=header)

    cols = _get_cols(file_bytes, sheet_name, header_row)
    if not cols:
        st.error(
            "No se encontraron columnas en la fila seleccionada. "
            "Prueba con otra fila de encabezado."
        )
        _reset_result()
        st.stop()

    # 4) Detectar candidatas a RUT (9 chars). El _v invalida la caché al cambiar la lógica.
    @st.cache_data(show_spinner="Detectando columna de RUT...")
    def _detectar_rut(data: bytes, sheet: str, header: int, _v: int = 5) -> list[str]:
        df_sample = read_sheet(data, sheet_name=sheet, header_row=header)
        return columnas_rut_posibles(df_sample)

    rut_candidatas = _detectar_rut(file_bytes, sheet_name, header_row)
    rut_default = rut_candidatas[0] if rut_candidatas else cols[0]

    # 5) Intentar extraer idc_value del archivo principal (en cualquier hoja)
    @st.cache_data(show_spinner="Buscando ID de componente...")
    def _get_idc(data: bytes, sheet: str, _v: int = 1) -> str:
        try:
            return get_idc(data, sheet_name=sheet)
        except Exception:
            return ""

    idc_value_default = _get_idc(file_bytes, sheet_name)

    # 6) Uploader opcional (fuera del form) para archivo que contenga el IDC
    idc_file = st.file_uploader(
        "Opcional: sube el Excel que contiene el ID de la componente "
        "(solo si no se detectó automáticamente en el archivo principal)",
        type=["xlsx"],
        key="idc_file",
    )
    if idc_file is not None:
        try:
            idc_value_default = get_idc(idc_file.getvalue())
            st.caption(f"IDC detectado: `{idc_value_default}`")
        except Exception as e:
            st.error(f"No se pudo extraer el IDC del archivo: {e}")
    elif idc_value_default:
        st.caption(f"ID de la componente detectado en el archivo principal: `{idc_value_default}`")

    # Componente sugerido desde el nombre del archivo IDC (ej. ..._C01.xlsx)
    detected_component = (
        _component_from_filename(idc_file.name) if idc_file is not None else None
    ) or _component_from_filename(uploaded.name)
    component_default = detected_component or "C01"
    if idc_file is not None and detected_component:
        st.caption(f"Componente detectado en el nombre del archivo: `{detected_component}`")

    # La clave cambia con el archivo IDC para que el campo se actualice al subirlo
    component_key = f"component_input_{file_id}_{idc_file.name if idc_file else ''}"

    # Fuerza que los selectores de columnas se reinicien (y vuelvan a aplicar la
    # detección de RUT) cuando cambia la hoja o la fila de encabezado.
    selection_key = f"{file_id}_{sheet_name}_{header_row}"

    # 7) Parámetros
    with st.form("params"):
        id_col = st.selectbox(
            "Columna con el ID del estudiante",
            cols,
            index=cols.index(rut_default),
            help=(
                f"Detectada automáticamente: `{rut_default}`"
                if rut_candidatas
                else "No se detectó automáticamente. Elige manualmente."
            ),
            key=f"id_col_{selection_key}",
        )
        grade_col = st.selectbox(
            "Columna con las notas",
            cols,
            key=f"grade_col_{selection_key}",
        )
        idc_value    = st.text_input(
            "Valor IDC",
            value=idc_value_default,
            help="Se detecta automáticamente desde el Excel. Puedes corregirlo aquí.",
        )
        id_component = st.text_input(
            "Componente (ej. C01, C02...)",
            value=component_default,
            help=(
                "Se completa automáticamente si el nombre del archivo termina "
                "en _C01, _C02, etc. Puedes corregirlo aquí."
            ),
            key=component_key,
        )
        reason       = st.text_input("Razón de la nota (Dejar por defecto)", value="OE")
        submitted    = st.form_submit_button("Procesar")

    # 8) Procesar
    if submitted:
        try:
            df = converter(
                file_bytes,
                grade_col=grade_col,
                id_col=id_col,
                idc_value=idc_value,
                id_component=id_component,
                reason=reason,
                sheet_name=sheet_name,
                header_row=header_row,
            )
            st.session_state["result_df"] = df
            st.session_state["result_xlsx"] = df_to_excel_bytes(df)
            st.session_state["result_component"] = id_component
            st.session_state["result_filename"] = _download_filename(idc_file, uploaded)
            print(
                f"[PROCESS] {uploaded.name} -> {df.height} filas "
                f"(componente={id_component}, idc={idc_value})",
                flush=True,
            )

            # Comparación de IDs contra la planilla del IDC (si fue subida)
            if idc_file is not None:
                try:
                    df_grades_orig = read_sheet(
                        file_bytes, sheet_name=sheet_name, header_row=header_row
                    )
                    df_idc = read_idc_frame(idc_file.getvalue())
                    coincidencias, diferencias = count_id_coincidences_diff(
                        df_grades_orig, df_idc, id_col
                    )
                    print(
                        f"[IDC CHECK] coincidencias={coincidencias} "
                        f"diferencias={diferencias}",
                        flush=True,
                    )
                    if diferencias > 2:
                        st.session_state["idc_mismatch"] = (coincidencias, diferencias)
                except Exception as e:
                    # No bloqueamos la descarga si falla la validación
                    print(f"[IDC CHECK ERROR] {type(e).__name__}: {e}", flush=True)

        except Exception as e:
            st.session_state.pop("result_df", None)
            st.session_state.pop("result_xlsx", None)
            st.session_state.pop("result_component", None)
            print(f"[ERROR] {type(e).__name__}: {e}", flush=True)
            st.error(f"Error procesando: {e}")

    # Diálogo de advertencia si los IDs no coinciden mayoritariamente
    # (se hace "pop" para que no reaparezca en reruns posteriores)
    if "idc_mismatch" in st.session_state:
        _coinc, _dif = st.session_state.pop("idc_mismatch")
        dialog_idc_mismatch(_coinc, _dif)

    # 9) Mostrar y descargar
    if "result_df" in st.session_state:
        st.success(f"{st.session_state['result_df'].height} filas procesadas")
        st.dataframe(st.session_state["result_df"])

        def _log_descarga():
            print(
                f"[DOWNLOAD] {st.session_state['result_filename']}",
                flush=True,
            )

        st.download_button(
            label="Descargar Excel",
            data=st.session_state["result_xlsx"],
            file_name=st.session_state["result_filename"],
            mime=MIME_XLSX,
            on_click=_log_descarga,
            key="download_result",
        )

        # 10) Apoyo voluntario y contacto
        st.divider()
        with st.container(border=True):
            st.markdown("##### ¿Te sirvió esta herramienta?")
            st.caption(
                "Es gratuita y la mantengo por mi cuenta. Si quieres retribuir de "
                "forma **totalmente voluntaria**, puedes hacerlo aquí. "
                "No es necesario, y la app seguirá funcionando igual."
            )

            col_cafe, col_mensual = st.columns(2)
            with col_cafe:
                st.link_button(
                    "☕ Invítame un café (única vez)",
                    MP_CAFE_URL,
                    use_container_width=True,
                )
            with col_mensual:
                st.link_button(
                    "🔄 Apoyo mensual voluntario",
                    MP_MENSUAL_URL,
                    use_container_width=True,
                )

            st.caption(
                "🔒 Serás redirigido a MercadoPago para completar el pago de forma segura."
            )

        st.markdown("")
        st.caption(
            f"🐛 ¿Encontraste un problema o tienes una sugerencia? "
            f"Escríbeme a [{CONTACT_EMAIL}](mailto:{CONTACT_EMAIL}"
            "?subject=%5BBUG%5D%20Reporte%20-%20Grade%20Importer"
            "&body=Describe%20el%20problema%20o%20sugerencia%3A%0A%0A"
            "Pasos%20para%20reproducirlo%20(si%20aplica)%3A%0A%0A"
            "Archivo%20que%20estabas%20subiendo%20(si%20aplica)%3A%0A)"
        )
