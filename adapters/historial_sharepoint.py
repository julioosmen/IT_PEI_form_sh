import re
import pandas as pd

    APP_TO_EXCEL_HEADERS = {
        "año": "Año",
        "ng1": "N G 1",
        "ng2": "N G 2",
        "fecha_recepcion": "Fecha de recepción",
        "periodo": "Periodo PEI ",
        "vigencia": "Vigencia",
        "tipo_pei": "Tipo de PEI",
        "estado": "Estado ",
        "responsable_institucional": "Responsable Institucional ",
        "cantidad_revisiones": "Cantidad de revisiones",
        "fecha_derivacion": "Fecha de derivación ",
        "etapa_revision": "Etapas de revisión",
        "comentario": "Comentario adicional/ Emisor de I.T",
        "articulacion": "Articulación ",
        "expediente": "Expediente ",
        "fecha_it": "Fecha de I.T ",
        "numero_it": "Número de I.T",
        "fecha_oficio": "Fecha Oficio",
        "numero_oficio": "Número Oficio",
        "id_sector": "Id_Sector",
        "nombre_sector": "nombre_sector",
        "id_pliego": "Id_Pliego",
        "nombre_pliego": "nombre_pliego",
        "codigo": "Id_UE",
    }


def map_row_to_excel_headers(app_row: dict) -> dict:
    """
    Convierte un dict con claves técnicas del app
    a un dict con encabezados EXACTOS de la tabla Excel (SharePoint).
    """
    out = {}
    for k, v in app_row.items():
        excel_header = APP_TO_EXCEL_HEADERS.get(k)
        if excel_header:
            out[excel_header] = v

    return out


def adaptar_historial_sharepoint(df_raw: pd.DataFrame) -> pd.DataFrame:
    df = df_raw.copy()

    def _norm_col(c: str) -> str:
        c = str(c).strip()
        c = re.sub(r"\s+", " ", c)
        return c

    df.columns = [_norm_col(c) for c in df.columns]
    df = df.rename(columns=APP_TO_EXCEL_HEADERS)

    # convención interna
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.lower()
        .str.replace(" ", "_")
    )

    if "codigo" not in df.columns:
        raise ValueError("Falta columna clave 'codigo' (mapeada desde 'Id_UE').")

    return df
