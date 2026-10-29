import polars as pl 
from io import BytesIO
from datetime import date


today_str = date.today().strftime("%m/%d/%Y")  

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
        .alias(basic_cols["punt"])
    )
    df_final = df_final.rename({grade_col: basic_cols["qual"]})
    df_final = df_final.rename({id_col: basic_cols["id"]})
    df_final = df_final.with_columns(
         pl.lit(today_str).alias("Fecha de envío")
    )
    df_final = df_final.with_columns(
        pl.lit(idc_value).alias(basic_cols["idc"])
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
