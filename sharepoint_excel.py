import io
import re
import unicodedata
import requests
import msal
import pandas as pd

def norm_key(s: str) -> str:
    s = "" if s is None else str(s).strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    s = re.sub(r"[^a-z0-9]+", "_", s)
    return s.strip("_")

def _graph_get_token(sp: dict) -> str:
    authority = f"https://login.microsoftonline.com/{sp['tenant_id']}"
    app = msal.ConfidentialClientApplication(
        client_id=sp["client_id"],
        authority=authority,
        client_credential=sp["client_secret"],
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(f"No se pudo obtener token Graph: {result}")
    return result["access_token"]

def _graph_get_site_id(token: str, site_hostname: str, site_path: str) -> str:
    url = f"https://graph.microsoft.com/v1.0/sites/{site_hostname}:{site_path}"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    r.raise_for_status()
    return r.json()["id"]

def _graph_get_drive_item_id(token: str, site_id: str, file_path: str) -> str:
    # Obtiene metadata del item (incluye id)
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:{file_path}"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    r.raise_for_status()
    return r.json()["id"]

def _excel_get_table_header_names(token: str, site_id: str, item_id: str, table_name: str) -> list[str]:
    # Devuelve los nombres de columnas de la tabla (en orden)
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}/workbook/tables/{table_name}/columns"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    r.raise_for_status()
    cols = r.json().get("value", [])
    # Cada columna trae { "name": "..." }
    return [c.get("name", "").strip() for c in cols]

def _excel_table_add_row(token: str, site_id: str, item_id: str, table_name: str, row_values_in_order: list) -> None:
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}/workbook/tables/{table_name}/rows/add"
    body = {"values": [row_values_in_order]}
    r = requests.post(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json=body,
        timeout=60,
    )
    r.raise_for_status()

def _graph_download_file(token: str, site_id: str, file_path: str) -> bytes:
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:{file_path}:/content"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=120)
    r.raise_for_status()
    return r.content


def _graph_upload_file(token: str, site_id: str, file_path: str, content: bytes) -> None:
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:{file_path}:/content"
    r = requests.put(url, headers={"Authorization": f"Bearer {token}"}, data=content, timeout=120)
    r.raise_for_status()

def read_table_from_sharepoint_as_df(
    secrets,
    table_name: str | None = None,
    table_name_key_in_secrets: str = "table_name",
) -> pd.DataFrame:
    sp = secrets["sharepoint"]
    token = _graph_get_token(sp)
    site_id = _graph_get_site_id(token, sp["site_hostname"], sp["site_path"])
    item_id = _graph_get_drive_item_id(token, site_id, sp["file_path"])

    tn = table_name or sp.get(table_name_key_in_secrets)
    if not tn:
        raise ValueError(f"No se indicó table_name y secrets['sharepoint'].{table_name_key_in_secrets} no existe.")

    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}/workbook/tables/{tn}/range"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    r.raise_for_status()

    values = r.json().get("values", [])
    if not values:
        return pd.DataFrame()

    headers = [str(x).strip() for x in values[0]]
    rows = values[1:]
    df = pd.DataFrame(rows, columns=headers)

    # limpia columnas tipo "Unnamed"
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]

    return df

def read_table_from_sharepoint_as_df_with_ids(
    token: str,
    site_id: str,
    item_id: str,
    table_name: str,
) -> pd.DataFrame:

    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}/workbook/tables/{table_name}/range"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    r.raise_for_status()

    values = r.json().get("values", [])
    if not values:
        return pd.DataFrame()

    headers = [str(x).strip() for x in values[0]]
    rows = values[1:]
    return pd.DataFrame(rows, columns=headers)

def append_row_to_sharepoint_excel(secrets, row_by_app_key: dict, table_name_key="table_name_hist") -> None:
    """
    Inserta una fila en la TABLA del Excel (SharePoint) usando headers reales.
    Acepta claves técnicas del app (snake_case) y las traduce a los headers de Excel.
    """
    sp = secrets["sharepoint"]
    token = _graph_get_token(sp)
    site_id = _graph_get_site_id(token, sp["site_hostname"], sp["site_path"])
    item_id = _graph_get_drive_item_id(token, site_id, sp["file_path"])

    table_name = sp.get(table_name_key)
    if not table_name:
        raise ValueError("Falta secrets['sharepoint'].table_name")

    # 1) Headers reales de la tabla (en orden)
    table_headers = _excel_get_table_header_names(token, site_id, item_id, table_name)
    if not table_headers:
        raise RuntimeError(f"No se pudieron leer columnas de la tabla '{table_name}'.")

    # 2) Normaliza headers reales del Excel
    excel_norm_headers = [norm_key(h) for h in table_headers]

    # 3) Alias: claves técnicas del app -> claves normalizadas del Excel
    #    (Esto resuelve fecha_recepcion vs fecha_de_recepcion, etc.)
    APPKEY_TO_EXCELNORM = {
        "codigo": "id_ue",
        "nombre": "nombre_unidad_ejecutora",
        "año": "ano",              # norm_key("Año") => "ano"
        "anio": "ano",
        "ng1": "n_g_1",
        "ng2": "n_g_2",

        "fecha_recepcion": "fecha_de_recepcion",
        "periodo": "periodo_pei",
        "vigencia": "vigencia",
        "tipo_pei": "tipo_de_pei",
        "estado": "estado",
        "responsable_institucional": "responsable_institucional",
        "cantidad_revisiones": "cantidad_de_revisiones",
        "fecha_derivacion": "fecha_de_derivacion",
        "etapa_revision": "etapas_de_revision",
        "comentario": "comentario_adicional_emisor_de_i_t",
        "articulacion": "articulacion",
        "expediente": "expediente",
        "fecha_it": "fecha_de_i_t",
        "numero_it": "numero_de_i_t",
        "fecha_oficio": "fecha_oficio",
        "numero_oficio": "numero_oficio",

        "id_sector": "id_sector",
        "nombre_sector": "nombre_sector",
        "id_pliego": "id_pliego",
        "nombre_pliego": "nombre_pliego",

        "id_departamento": "id_departamento",
        "nombre_departamento": "nombre_departamento",
        "id_provincia": "id_provincia",
        "nombre_provincia": "nombre_provincia",
        "id_4distrito": "id_4distrito",
        "nombre_distrito": "nombre_distrito",

        "id_registro": "idregistro",
        "last_updated": "lastupdated",
        "updated_by": "updatedby",
    }

    # 4) Construye diccionario normalizado con alias
    data_norm = {}
    for k, v in row_by_app_key.items():
        k0 = norm_key(k)  # por si acaso llega con espacios/tildes
        excel_norm = APPKEY_TO_EXCELNORM.get(k0, k0)  # si no hay alias, usa k0
        data_norm[excel_norm] = v

    # 5) Arma la fila en el mismo orden de la tabla
    new_row = [data_norm.get(h_norm, "") for h_norm in excel_norm_headers]

    # 6) Inserta fila por Graph Excel API
    _excel_table_add_row(token, site_id, item_id, table_name, new_row)

def _excel_table_get_all_values(token: str, site_id: str, item_id: str, table_name: str) -> list[list]:
    # Devuelve matriz: [ [fila1...], [fila2...] ... ] (sin headers)
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}/workbook/tables/{table_name}/range"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    r.raise_for_status()
    return r.json().get("values", [])

def update_row_in_table_by_idregistro(
    secrets,
    updates_by_app_key: dict,
    id_registro: str,
    appkey_to_excelnorm: dict,
    table_name_key="table_name_hist", 
) -> None:
    """
    Actualiza un registro existente en la tabla (SharePoint Excel) buscando por IdRegistro.
    - updates_by_app_key: dict con claves técnicas del app (estado, comentario, etc.)
    - id_registro: valor exacto de la columna IdRegistro de esa fila
    - appkey_to_excelnorm: el mismo alias que ya usas para insertar (Opción A)
    """
    sp = secrets["sharepoint"]
    table_name = sp.get(table_name_key)
    if not table_name:
        raise ValueError("Falta secrets['sharepoint'].table_name")

    token = _graph_get_token(sp)
    site_id = _graph_get_site_id(token, sp["site_hostname"], sp["site_path"])
    item_id = _graph_get_drive_item_id(token, site_id, sp["file_path"])

    headers = _excel_get_table_header_names(token, site_id, item_id, table_name)
    headers_norm = [norm_key(h) for h in headers]

    # datos actuales (matriz sin headers)
    values = _excel_table_get_all_values(token, site_id, item_id, table_name)

    # ubicar columna IdRegistro
    if "idregistro" not in headers_norm:
        raise ValueError("La tabla no tiene columna 'IdRegistro' (requerida para actualizar).")

    id_col = headers_norm.index("idregistro")

    # ubicar fila por IdRegistro
    target_idx = None
    for i, row in enumerate(values):
        if i < len(values) and id_col < len(row) and str(row[id_col]).strip() == str(id_registro).strip():
            target_idx = i
            break

    if target_idx is None:
        raise ValueError(f"No se encontró IdRegistro={id_registro} en la tabla.")

    # construir dict normalizado de updates
    updates_norm = {}
    for k, v in updates_by_app_key.items():
        k0 = norm_key(k)
        excel_norm = appkey_to_excelnorm.get(k0, k0)
        updates_norm[excel_norm] = v

    # armar nueva fila completa preservando lo existente
    current = list(values[target_idx])
    # asegurar largo correcto (por si excel devuelve menos columnas)
    if len(current) < len(headers):
        current += [""] * (len(headers) - len(current))

    # aplicar updates a índices
    for hn, v in updates_norm.items():
        if hn in headers_norm:
            current[headers_norm.index(hn)] = v

    # escribir la fila completa usando rows/itemAt(index)/range
    url = (
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}"
        f"/workbook/tables/{table_name}/rows/itemAt(index={target_idx})/range"
    )
    r = requests.patch(
        url,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"values": [current]},
        timeout=60
    )
    r.raise_for_status()
