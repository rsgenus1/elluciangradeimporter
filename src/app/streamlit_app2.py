import streamlit as st
from elluciangradeimporter.converter import get_cols, converter, df_to_excel_bytes, get_idc

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
    4. Descarga el archivo generado e impórtalo en Banner.
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

    # 2) Parámetros
    with st.form("params"):
        id_col        = st.selectbox("Columna con el ID del estudiante", cols)
        grade_col     = st.selectbox("Columna con las notas", cols)
        id_component  = st.text_input("Componente (ej. C01, C02...)", value="C01")
        reason        = st.text_input("Razón de la nota (Dejar por defecto)", value="OE")
        #aqui falta idc_value que ahora lo NECESITA la fn converter
        # hay una funcion llamada get_idc que devuelve un str con el idc_value que sirve
        submitted     = st.form_submit_button("Procesar")

    # 3) Procesar (guardar resultado en session_state para que sobreviva reruns)
    if submitted:
        try:
            df = converter(file_bytes, grade_col, id_col, id_component, reason)
            st.session_state["result_df"] = df
            st.session_state["result_xlsx"] = df_to_excel_bytes(df)
            st.session_state["result_component"] = id_component
            print(
                f"[PROCESS] {uploaded.name} -> {df.height} filas (componente={id_component})",
                flush=True,
            )
        except Exception as e:
            st.session_state.pop("result_df", None)
            st.session_state.pop("result_xlsx", None)
            st.session_state.pop("result_component", None)
            print(f"[ERROR] {type(e).__name__}: {e}", flush=True)
            st.error(f"Error procesando: {e}")

    # 4) Mostrar y descargar (persisten entre reruns)
    if "result_df" in st.session_state:
        st.success(f"{st.session_state['result_df'].height} filas procesadas")
        st.dataframe(st.session_state["result_df"])

        def _log_descarga():
            # flush=True fuerza a que la línea aparezca ya en los Cloud logs
            print(
                f"[DOWNLOAD] {st.session_state['result_component']}.xlsx",
                flush=True,
            )

        st.download_button(
            label="Descargar Excel",
            data=st.session_state["result_xlsx"],
            file_name=f"resultado_{st.session_state['result_component']}.xlsx",
            mime=MIME_XLSX,
            on_click=_log_descarga,   # solo se ejecuta al hacer clic real
            key="download_result",    # <-- fix: key única
        )
