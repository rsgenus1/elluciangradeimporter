import streamlit as st
from elluciangradeimporter.converter import (
    get_cols, converter, df_to_excel_bytes, get_idc,
)

MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

st.title("Convertidor para Banner de Ellucian")
st.markdown(
    """
    Subir notas manualmente en la plataforma **Banner de Ellucian** es lento y
    propenso a errores. Esta herramienta toma un Excel con las notas de los
    estudiantes y genera un archivo compatible con los requisitos de importación
    de Banner.

    **Cómo usarla:**
    1. Sube el Excel con las notas.
    2. Selecciona la columna del ID del estudiante y la columna de notas.
    3. Indica el componente (ej. `C01`, `C02`, ...).
    4. Indica el ID de la componente, si no lo conoce, descarga la planilla desde Banner, y súbela, se detectará automáticamente.
    5. Descarga el archivo generado e impórtalo en Banner.
    """
)

uploaded = st.file_uploader("Sube el Excel", type=["xlsx"])

if uploaded is not None:
    file_bytes = uploaded.getvalue()   # se extrae una sola vez

    # 1) Columnas (cacheado por contenido del archivo)
    @st.cache_data(show_spinner="Leyendo columnas...")
    def _get_cols(data: bytes) -> list[str]:
        return get_cols(data)

    cols = _get_cols(file_bytes)

    # 2) Intentar extraer idc_value del archivo principal
    idc_value_default = ""
    try:
        idc_value_default = get_idc(file_bytes)
    except Exception:
        pass  # silencioso: puede que el archivo no tenga la columna; se pedirá abajo

    # 3) Uploader opcional (fuera del form) para archivo que contenga el IDC
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

    # 4) Parámetros
    with st.form("params"):
        id_col        = st.selectbox("Columna con el ID del estudiante", cols)
        grade_col     = st.selectbox("Columna con las notas", cols)
        idc_value     = st.text_input(
            "Valor IDC",
            value=idc_value_default,
            help="Se detecta automáticamente desde el Excel. Puedes corregirlo aquí.",
        )
        id_component  = st.text_input("Componente (ej. C01, C02...)", value="C01")
        reason        = st.text_input("Razón de la nota (Dejar por defecto)", value="OE")
        submitted     = st.form_submit_button("Procesar")

    # 5) Procesar
    if submitted:
        try:
            df = converter(
                file_bytes,
                grade_col=grade_col,
                id_col=id_col,
                idc_value=idc_value,
                id_component=id_component,
                reason=reason,
            )
            st.session_state["result_df"] = df
            st.session_state["result_xlsx"] = df_to_excel_bytes(df)
            st.session_state["result_component"] = id_component
            print(
                f"[PROCESS] {uploaded.name} -> {df.height} filas "
                f"(componente={id_component}, idc={idc_value})",
                flush=True,
            )
        except Exception as e:
            st.session_state.pop("result_df", None)
            st.session_state.pop("result_xlsx", None)
            st.session_state.pop("result_component", None)
            print(f"[ERROR] {type(e).__name__}: {e}", flush=True)
            st.error(f"Error procesando: {e}")

    # 6) Mostrar y descargar (persisten entre reruns)
    if "result_df" in st.session_state:
        st.success(f"{st.session_state['result_df'].height} filas procesadas")
        st.dataframe(st.session_state["result_df"])

        def _log_descarga():
            print(
                f"[DOWNLOAD] {st.session_state['result_component']}.xlsx",
                flush=True,
            )

        st.download_button(
            label="Descargar Excel",
            data=st.session_state["result_xlsx"],
            file_name=f"resultado_{st.session_state['result_component']}.xlsx",
            mime=MIME_XLSX,
            on_click=_log_descarga,
            key="download_result",
        )
