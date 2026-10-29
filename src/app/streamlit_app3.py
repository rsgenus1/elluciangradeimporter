import io

import polars as pl
import streamlit as st

from elluciangradeimporter.converter import (
    get_cols, converter, df_to_excel_bytes, get_idc,
)

MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# Contacto y enlaces de apoyo
CONTACT_EMAIL = "ellucianimporter@proton.me"
MP_CAFE_URL = "https://mpago.li/2BxMn6E"
MP_MENSUAL_URL = "https://mpago.la/2RsGqp4"

# Alias compatible: pl.String en versiones recientes, pl.Utf8 en todas
STRING = getattr(pl, "String", pl.Utf8)


def columnas_rut_posibles(df: pl.DataFrame) -> list[str]:
    """Columnas cuyos valores no nulos ni vacíos, tras strip, miden todos 9 chars."""
    resultado = []
    for name in df.columns:
        try:
            serie = df[name].cast(STRING).str.strip_chars().drop_nulls()
        except Exception:
            continue
        serie = serie.filter(serie != "")
        if serie.len() == 0:
            continue
        longitudes = serie.str.len_chars()
        if longitudes.min() == 9 and longitudes.max() == 9:
            resultado.append(name)
    return resultado


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
    file_bytes = uploaded.getvalue()

    # 1) Columnas
    @st.cache_data(show_spinner="Leyendo columnas...")
    def _get_cols(data: bytes) -> list[str]:
        return get_cols(data)

    cols = _get_cols(file_bytes)

    # 2) Detectar candidatas a RUT (9 chars). El _v invalida la caché al cambiar la lógica.
    @st.cache_data(show_spinner="Detectando columna de RUT...")
    def _detectar_rut(data: bytes, _v: int = 4) -> list[str]:
        df_sample = pl.read_excel(io.BytesIO(data))
        return columnas_rut_posibles(df_sample)

    rut_candidatas = _detectar_rut(file_bytes)
    rut_default = rut_candidatas[0] if rut_candidatas else cols[0]

    # 3) Intentar extraer idc_value del archivo principal
    idc_value_default = ""
    try:
        idc_value_default = get_idc(file_bytes)
    except Exception:
        pass

    # 4) Uploader opcional (fuera del form) para archivo que contenga el IDC
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

    # 5) Parámetros
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
        )
        grade_col    = st.selectbox("Columna con las notas", cols)
        idc_value    = st.text_input(
            "Valor IDC",
            value=idc_value_default,
            help="Se detecta automáticamente desde el Excel. Puedes corregirlo aquí.",
        )
        id_component = st.text_input("Componente (ej. C01, C02...)", value="C01")
        reason       = st.text_input("Razón de la nota (Dejar por defecto)", value="OE")
        submitted    = st.form_submit_button("Procesar")

    # 6) Procesar
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

    # 7) Mostrar y descargar
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

        # 8) Apoyo voluntario y contacto
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
