from datetime import date
from io import BytesIO

import fastexcel
import polars as pl
import xlsxwriter

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

# Cuántas filas iniciales se revisan para detectar el encabezado y ofrecer
# alternativas al usuario en la app.
HEADER_SCAN_ROWS = 100


def _as_source(file):
    """Normaliza la entrada para que polars/fastexcel la acepten.

    Acepta bytes, BytesIO, archivos abiertos o rutas y devuelve bytes (lo más
    común) o la ruta tal cual.
    """
    if isinstance(file, (bytes, bytearray, memoryview)):
        return bytes(file)
    if hasattr(file, "getvalue"):
        return file.getvalue()
    if hasattr(file, "read"):
        current = file.tell() if hasattr(file, "tell") else None
        content = file.read()
        if current is not None and hasattr(file, "seek"):
            file.seek(current)
        return content
    return file


def get_sheet_names(file) -> list[str]:
    """Nombres de las hojas (worksheets) del workbook, en orden."""
    reader = fastexcel.read_excel(_as_source(file))
    return list(reader.sheet_names)


def read_raw_rows(file, sheet_name: str | None = None, max_rows: int = HEADER_SCAN_ROWS) -> list[tuple]:
    """Lee las primeras filas de una hoja sin tratar ninguna como encabezado.

    Las filas vacías iniciales se conservan para que el índice de cada fila
    coincida con el número de fila real de la hoja (base 0), que es lo que
    espera el parámetro ``header_row``.
    """
    df = pl.read_excel(
        _as_source(file),
        sheet_name=sheet_name,
        has_header=False,
        infer_schema_length=None,
        drop_empty_rows=False,
        drop_empty_cols=False,
        raise_if_empty=False,
        read_options={"header_row": None, "skip_rows": [], "n_rows": max_rows},
    )
    return df.rows()


def _norm_cell(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _looks_numeric(text: str) -> bool:
    try:
        float(text.replace(",", "."))
        return True
    except ValueError:
        return False


def _row_score(stats: tuple[int, float, float]) -> float:
    """Puntaje de una fila como candidata a encabezado (más alto = mejor)."""
    n_values, text_ratio, below_avg = stats
    if n_values == 0:
        return float("-inf")
    followed_by_data = 1.0 if below_avg >= n_values else 0.0
    return n_values + 3.0 * text_ratio + 2.0 * followed_by_data


def detect_header_row(file, sheet_name: str | None = None, max_scan: int = HEADER_SCAN_ROWS) -> int:
    """Detecta la fila (base 0) que contiene los nombres de las columnas.

    Muchos docentes dejan arriba un título combinado (a veces una celda
    combinada con el nombre del curso), filas vacías o datos de la asignatura,
    y recién más abajo empieza el encabezado real.

    La heurística elige la fila con mejor puntaje: cantidad de celdas no
    vacías, proporción de texto (los encabezados suelen ser texto y no números)
    y si hay filas con datos debajo. En caso de empate se prefiere la fila más
    alta.
    """
    rows = read_raw_rows(file, sheet_name, max_rows=max_scan)
    if not rows:
        return 0

    limit = min(len(rows), max_scan)

    def stats(r: int) -> tuple[int, float, float]:
        values = [_norm_cell(c) for c in rows[r]]
        values = [v for v in values if v]
        n_values = len(values)
        if n_values == 0:
            return 0, 0.0, 0.0
        text_ratio = sum(1 for v in values if not _looks_numeric(v)) / n_values

        # Cantidad promedio de celdas no vacías en las primeras filas con datos.
        below_counts = []
        for j in range(r + 1, limit):
            count = sum(1 for c in (_norm_cell(v) for v in rows[j]) if c)
            if count:
                below_counts.append(count)
            if len(below_counts) >= 5:
                break
        below_avg = sum(below_counts) / len(below_counts) if below_counts else 0.0
        return n_values, text_ratio, below_avg

    # Se elige la fila con mejor puntaje. El desempate favorece la fila más
    # alta, que es donde suele estar el encabezado (no una fila de datos).
    best_row = 0
    best_score = None
    for r in range(limit):
        n_values, text_ratio, below_avg = stats(r)
        score = _row_score((n_values, text_ratio, below_avg))
        if score == float("-inf"):
            continue
        score -= r * 1e-6
        if best_score is None or score > best_score:
            best_score = score
            best_row = r

    return best_row


def read_sheet(file, sheet_name: str | None = None, header_row: int = 0, **kwargs) -> pl.DataFrame:
    """Lee una hoja usando ``header_row`` como fila de nombres de columnas.

    El resto de ``kwargs`` se pasa a :func:`polars.read_excel`.
    """
    read_options = dict(kwargs.pop("read_options", None) or {})
    read_options["header_row"] = header_row
    return pl.read_excel(
        _as_source(file),
        sheet_name=sheet_name,
        read_options=read_options,
        **kwargs,
    )


def _find_column(columns: list[str], name: str) -> str | None:
    target = _norm_cell(name).lower()
    for column in columns:
        if _norm_cell(column).lower() == target:
            return column
    return None


def _locate_column(file, column_name: str, prefer_sheet: str | None = None) -> tuple[str, int] | None:
    """Busca (hoja, fila de encabezado) donde aparece ``column_name``."""
    source = _as_source(file)
    sheets = get_sheet_names(source)
    if prefer_sheet in sheets:
        sheets = [prefer_sheet] + [s for s in sheets if s != prefer_sheet]

    target = _norm_cell(column_name).lower()
    for sheet in sheets:
        for idx, row in enumerate(read_raw_rows(source, sheet)):
            if target in (_norm_cell(c).lower() for c in row):
                return sheet, idx
    return None


def read_sheet_with_column(file, column_name: str, prefer_sheet: str | None = None) -> pl.DataFrame:
    """Lee la primera hoja (y fila de encabezado) que contiene ``column_name``."""
    source = _as_source(file)
    located = _locate_column(source, column_name, prefer_sheet=prefer_sheet)
    if located is None:
        raise ValueError(f"No se encontró la columna «{column_name}» en el archivo.")
    sheet, header_row = located
    return read_sheet(source, sheet_name=sheet, header_row=header_row)


# we need it to make the user choose which one is the Id one
# and which one is the grade one
def get_cols(file, sheet_name: str | None = None, header_row: int = 0) -> list:
    df = read_sheet(file, sheet_name=sheet_name, header_row=header_row)

    return df.columns

def converter(
    data,
    grade_col: str,
    id_col: str,
    id_component: str,
    idc_value:str,
    reason = "OE",
    sheet_name: str | None = None,
    header_row: int = 0,
):
    """ 
    This funcions recieves the file as Bytes
    from an excel file, then process it to get a polars dataframe
    data: the Bytes got from the excel file, this will be asked by streamlit
    grade_col: the name of the column with the names
    id_col: the name of the col that contains the student ID number 
    id_component: Usually "C0x" with x in 1,2,3,4
    reason: (default "OE")
    sheet_name: name of the worksheet to read (default: first one)
    header_row: 0-based index of the row with the column names (default 0)
    """
    df_init = read_sheet(
        data,
        sheet_name=sheet_name,
        header_row=header_row,
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

def df_to_excel_bytes(df: pl.DataFrame) -> bytes:
    """Genera el Excel de salida con un formato simple y legible."""
    buffer = BytesIO()

    anchos = {
        "Calificación": 100,
        "ID de alumno": 110,
        "Nombre de componente": 150,
        "Razón de cambio de calificación": 240,
        "Puntaje": 85,
        "Fecha de envío": 110,
        "ID de componente": 130,
        "Comentario": 140,
        "Fecha de extensión": 140,
    }

    workbook = xlsxwriter.Workbook(buffer, {"in_memory": True})
    df.write_excel(
        workbook,
        worksheet="Importable",
        header_format={
            "bold": True,
            "font_color": "#FFFFFF",
            "bg_color": "#1F4E79",
            "align": "center",
            "valign": "vcenter",
            "border": 1,
        },
        column_widths={
            name: ancho for name, ancho in anchos.items() if name in df.columns
        },
        row_heights={0: 24},
        freeze_panes=(1, 0),
        autofilter=False,
    )

    # Bandas muy suaves en las filas de datos, coherentes con la cabecera azul.
    if df.height > 0 and df.width > 0:
        worksheet = workbook.get_worksheet_by_name("Importable")
        banda = workbook.add_format({"bg_color": "#EAF1F8"})
        worksheet.conditional_format(
            1,
            0,
            df.height,
            df.width - 1,
            {
                "type": "formula",
                "criteria": "=MOD(ROW(),2)=0",
                "format": banda,
            },
        )

    workbook.close()
    return buffer.getvalue()


def get_idc(file, sheet_name: str | None = None) -> str:
    """
    It takes the excel provided by ellucian to extract the unique id for the grade.

    Busca la hoja y la fila de encabezado que contengan la columna
    ``ID de componente`` (empezando por ``sheet_name`` si se indica), así
    funciona aunque el archivo tenga varias hojas o filas de título arriba.
    """
    df = read_sheet_with_column(file, basic_cols["idc"], prefer_sheet=sheet_name)
    column = _find_column(df.columns, basic_cols["idc"])
    if column is None:
        raise ValueError(f"No se encontró la columna «{basic_cols['idc']}» en el archivo.")

    serie = df[column].drop_nulls().cast(pl.String).str.strip_chars()
    valores = [v for v in serie.unique().to_list() if v]
    if not valores:
        raise ValueError(f"La columna «{basic_cols['idc']}» está vacía.")
    if len(valores) > 1:
        raise ValueError(
            f"Se encontraron varios valores distintos en «{basic_cols['idc']}»: "
            f"{valores[:5]}."
        )
    return valores[0]


def read_idc_frame(file) -> pl.DataFrame:
    """Lee la hoja de la planilla IDC que contiene la columna «ID de alumno»."""
    return read_sheet_with_column(file, basic_cols["id"])


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

