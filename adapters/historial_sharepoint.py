import re
import pandas as pd
import pandas as pd
from sharepoint_excel import norm_key

# Mapeo desde encabezados (normalizados) del Excel -> columnas estándar de la app
EXCEL_NORM_TO_APP = {
    "id_ue": "codigo",
    "nombre_pliego": "nombre",

    "ano": "año",
    "anio": "año",

    "n_g_1": "ng1",
    "n_g_2": "ng2",

    "fecha_de_recepcion": "fecha_recepcion",
    "periodo_pei": "periodo",
    "vigencia": "vigencia",
    "tipo_de_pei": "tipo_pei",
    "estado": "estado",
    "responsable_institucional": "responsable_institucional",
    "cantidad_de_revisiones": "cantidad_revisiones",
    "fecha_de_derivacion": "fecha_derivacion",
    "etapas_de_revision": "etapa_revision",
    "comentario_adicional_emisor_de_i_t": "comentario",
    "articulacion": "articulacion",
    "expediente": "expediente",
    "fecha_de_i_t": "fecha_it",
    "numero_de_i_t": "numero_it",

    "fecha_oficio": "fecha_oficio",
    "numero_oficio": "numero_oficio",

    "id_sector": "id_sector",
    "nombre_sector": "nombre_sector",
    "id_pliego": "id_pliego",
    "nombre_unidad_ejecutora": "nombre_unidad_ejecutora",

    "id_departamento": "id_departamento",
    "nombre_departamento": "nombre_departamento",
    "id_provincia": "id_provincia",
    "nombre_provincia": "nombre_provincia",
    "id_4distrito": "id_4distrito",
    "nombre_distrito": "nombre_distrito",

    "idregistro": "id_registro",
    "lastupdated": "last_updated",
    "updatedby": "updated_by",
}

def adaptar_historial_sharepoint(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # 1) Limpia headers (quita espacios laterales) y elimina columnas Unnamed
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.str.startswith("Unnamed")]

    # 2) Construye mapa normalizado -> nombre real
    norm_to_real = {norm_key(c): c for c in df.columns}

    # 3) Prepara renombrado usando equivalencias
    rename_map = {}
    for excel_norm, app_col in EXCEL_NORM_TO_APP.items():
        if excel_norm in norm_to_real:
            rename_map[norm_to_real[excel_norm]] = app_col

    df = df.rename(columns=rename_map)

    # 4) Validación columna clave
    if "codigo" not in df.columns:
        raise ValueError(
            "Falta columna clave 'codigo' (mapeada desde 'Id_UE'). "
            f"Columnas detectadas: {df.columns.tolist()}"
        )

    # 5) Limpieza de código
    df["codigo"] = df["codigo"].astype(str).str.strip()

    return df
