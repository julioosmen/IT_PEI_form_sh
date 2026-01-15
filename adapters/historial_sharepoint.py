import re
import pandas as pd

MAP_HIST_SP_TO_STD = {
    "Id_UE": "codigo",
    "Año": "año",
    "Responsable Institucional": "responsable_institucional",
    "Fecha de recepción": "fecha_recepcion",
    "Periodo PEI": "periodo",
    "Vigencia": "vigencia",
    "Tipo de PEI": "tipo_pei",
    "Estado": "estado",
    "Cantidad de revisiones": "cantidad_revisiones",
    "Fecha de derivación": "fecha_derivacion",
    "Etapas de revisión": "etapa_revision",
    "Comentario adicional/ Emisor de I.T": "comentario",
    "Articulación": "articulacion",
    "Expediente": "expediente",
    "Fecha de I.T": "fecha_it",
    "Número de I.T": "numero_it",
    "Fecha Oficio": "fecha_oficio",
    "Número Oficio": "numero_oficio",
}

def adaptar_historial_sharepoint(df_raw: pd.DataFrame) -> pd.DataFrame:
    df = df_raw.copy()

    def _norm_col(c: str) -> str:
        c = str(c).strip()
        c = re.sub(r"\s+", " ", c)
        return c

    df.columns = [_norm_col(c) for c in df.columns]
    df = df.rename(columns=MAP_HIST_SP_TO_STD)

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
