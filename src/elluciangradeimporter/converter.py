import polars as pl 
from io import BytesIO
from datetime import date


today_str = date.today().strftime("%d/%m/%Y")  

basic_cols = {
    "id" : "ID de alumno",
    "idc": "ID de componente",
    "comp" : "Nombre de componente",
    "porc" : "Porcentaje",
    "qual" :"Calificación",
    "reas" :"Razón de cambio de calificación",
    "punt" : "Puntaje",
}

grade_number = [ "C01", "C02", "C03", "C04" ]

# we need it to make the user choose which one is the Id one
# and which one is the grade one
def get_cols(file) -> list:
    df = pl.read_excel(
        file
    )

    return df.columns

def converter(data, grade_col: str, id_col: str, id_component: str, idc_value:str, reason = "OE"):
    """ 
    This funcions recieves the file as Bytes
    from an excel file, then process it to get a polars dataframe
    data: the Bytes got from the excel file, this will be asked by streamlit
    grade_col: the name of the column with the names
    id_col: the name of the col that contains the student ID number 
    id_component: Usually "C0x" with x in 1,2,3,4
    reason: (default "OE")
    """
    df_init = pl.read_excel(
        data,
        columns = [grade_col, id_col],
        schema_overrides={id_col: pl.String, grade_col: pl.Float64},
    )
    #normalice the RUT data, removing . , -  and espaces
    # and convert the k to uppercase
    df_init = df_init.with_columns(
        pl.col(id_col)
          .str.replace_many([".", "-", " "], ["","",""])
          .str.to_uppercase()
          .alias(id_col)
    )
    df_init = df_init.with_columns(
    pl.col(grade_col)
        .round(1)
      .cast(pl.String)
      .str.replace(".", ",", literal=True)
      .alias(grade_col)
    )#
    df_final = df_init.with_columns(
        pl.lit(id_component).alias(basic_cols["comp"])
    )
    df_final = df_final.with_columns(
        pl.lit(reason).alias(basic_cols["reas"])
    )
    df_final = df_final.with_columns(
        pl.col(grade_col)
        .cast(pl.String)
        .str.replace(".",",", literal=True)
        .alias(basic_cols["punt"]) #redondear a 1 dec
    )
    df_final = df_final.rename({grade_col: basic_cols["qual"]}) # falta redondear a 1 dec
    df_final = df_final.rename({id_col: basic_cols["id"]})
    df_final = df_final.with_columns(
         pl.lit(today_str).alias("Fecha de envío")
    )
    df_final = df_final.with_columns(
        pl.lit(idc_value).alias(basic_cols["idc"])
    )
    df_final = df_final.with_columns(
    pl.lit("").cast(pl.String).alias("Comentario"),
    pl.lit("").cast(pl.String).alias("Fecha de extensión"),
)
    return df_final

def df_to_excel_bytes(df: pl.DataFrame):
    buffer = BytesIO()
    df.write_excel(buffer)
    return buffer.getvalue()


def get_idc(file) -> str:
    """
    It takes the excel provided by ellucian to extract the unique id for the grade
        """
    df = pl.read_excel(
        file,
        columns = [basic_cols["idc"]]
    )
    return str(df[basic_cols["idc"]].unique().item())
def columnas_rut_posibles(df: pl.DataFrame) -> list[str]:
    """Columnas cuyos valores no nulos, al convertirse a texto, miden todos 9 chars."""
    resultado = []
    for name in df.columns:
        try:
            serie = df[name].cast(pl.String).drop_nulls()
        except Exception:
            continue
        if serie.len() == 0:
            continue
        longitudes = serie.str.len_chars()
        if longitudes.min() == 9 and longitudes.max() == 9:
            resultado.append(name)
    return resultado


def count_id_coincidences_diff(
    df_grades: pl.DataFrame,
    df_idc: pl.DataFrame,
    id_col: str,
    id_col_banner: str = "ID de alumno",
) -> tuple[int, int]:
    """
    Cuenta cuántas filas de `df_grades[id_col]` aparecen en `df_idc[id_col_banner]`
    (coincidencias) y cuántas no (diferencias).

    - semi join: nunca duplica filas de df_grades aunque el ID esté repetido
      en df_idc (p. ej. una fila por componente/evaluación en Banner).
    - cast a Utf8 en ambos lados: si un archivo trae el ID como número y el
      otro como texto, sin el cast el join no matchea nada silenciosamente.
    """
    # IDs únicos del IDC, renombrados al nombre de columna de df_grades
    ids_idc = (
        df_idc
        .select(pl.col(id_col_banner).cast(pl.Utf8).unique().alias(id_col))
    )

    # semi join: devuelve solo las filas de df_grades cuyo ID está en ids_idc
    coincidentes = df_grades.join(ids_idc, on=id_col, how="semi")

    coincidencias = coincidentes.height
    diferencias = df_grades.height - coincidencias
    return (coincidencias, diferencias)

