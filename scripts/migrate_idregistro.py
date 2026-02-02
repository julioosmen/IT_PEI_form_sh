import time
from uuid import uuid4
import requests
import msal
import tomllib

# -----------------------------
# Helpers Graph auth + ids
# -----------------------------
def norm_str(x):
    return "" if x is None else str(x).strip()

def graph_get_token(sp: dict) -> str:
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

def graph_get_site_id(token: str, site_hostname: str, site_path: str) -> str:
    url = f"https://graph.microsoft.com/v1.0/sites/{site_hostname}:{site_path}"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    r.raise_for_status()
    return r.json()["id"]

def graph_get_drive_item_id_by_path(token: str, site_id: str, file_path: str) -> str:
    # /drive/root:{path}
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:{file_path}"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    r.raise_for_status()
    return r.json()["id"]

# -----------------------------
# Excel table helpers
# -----------------------------
def excel_table_get_headers(token: str, site_id: str, item_id: str, table_name: str) -> list[str]:
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}/workbook/tables/{table_name}/columns"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    r.raise_for_status()
    cols = r.json().get("value", [])
    # Cada item tiene "name"
    return [c.get("name", "") for c in cols]

def excel_table_get_all_values(token: str, site_id: str, item_id: str, table_name: str) -> list[list]:
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}/workbook/tables/{table_name}/range"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    r.raise_for_status()
    return r.json().get("values", [])

def excel_table_patch_row_full(token: str, site_id: str, item_id: str, table_name: str, row_index_0: int, row_values: list):
    # Actualiza toda la fila (dentro de la tabla) por índice 0-based
    url = (
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}"
        f"/workbook/tables/{table_name}/rows/itemAt(index={row_index_0})/range"
    )
    r = requests.patch(
        url,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"values": [row_values]},
        timeout=60
    )
    r.raise_for_status()

# -----------------------------
# Migración
# -----------------------------
def migrate_fill_idregistro(sp: dict, secrets_path: str, dry_run: bool = True, sleep_s: float = 0.15):
    token = graph_get_token(sp)
    site_id = graph_get_site_id(token, sp["site_hostname"], sp["site_path"])
    item_id = graph_get_drive_item_id_by_path(token, site_id, sp["file_path"])
    table_name = sp["table_name"]

    headers = excel_table_get_headers(token, site_id, item_id, table_name)
    headers = [norm_str(h) for h in headers]
    header_to_idx = {h: i for i, h in enumerate(headers)}

    # Validación columnas
    if "IdRegistro" not in header_to_idx:
        raise RuntimeError("La tabla no tiene la columna 'IdRegistro'. Agrega esa columna en Excel y vuelve a ejecutar.")

    id_col = header_to_idx["IdRegistro"]
    lastupdated_col = header_to_idx.get("LastUpdated")
    updatedby_col = header_to_idx.get("UpdatedBy")

    values = excel_table_get_all_values(token, site_id, item_id, table_name)

    total = len(values)
    to_update = []

    for i, row in enumerate(values):
        # Asegura largo
        row = list(row)
        if len(row) < len(headers):
            row += [""] * (len(headers) - len(row))

        if norm_str(row[id_col]) == "":
            to_update.append(i)

    print(f"Total filas en tabla: {total}")
    print(f"Filas sin IdRegistro: {len(to_update)}")
    print(f"Modo: {'DRY RUN (no escribe)' if dry_run else 'WRITE (escribe cambios)'}")

    updated = 0
    for idx in to_update:
        row = list(values[idx])
        if len(row) < len(headers):
            row += [""] * (len(headers) - len(row))

        new_id = str(uuid4())
        row[id_col] = new_id

        if lastupdated_col is not None:
            row[lastupdated_col] = time.strftime("%Y-%m-%d %H:%M:%S")

        if updatedby_col is not None:
            row[updatedby_col] = "migration_script"

        if dry_run:
            print(f"[DRY] row={idx} set IdRegistro={new_id}")
        else:
            excel_table_patch_row_full(token, site_id, item_id, table_name, idx, row)
            updated += 1
            if updated % 20 == 0:
                print(f"Actualizadas {updated}/{len(to_update)} filas...")

            time.sleep(sleep_s)  # evita throttling

    print("FIN.")
    if not dry_run:
        print(f"Filas actualizadas: {updated}")

def load_secrets(path: str) -> dict:
    with open(path, "rb") as f:
        data = tomllib.load(f)
    if "sharepoint" not in data:
        raise RuntimeError("El archivo secrets no tiene bloque [sharepoint].")
    return data["sharepoint"]

if __name__ == "__main__":
    # Cambia este path si deseas otro nombre
    secrets_file = "secrets.local.toml"
    sp = load_secrets(secrets_file)

    # 1) Primero ejecuta en DRY RUN
    migrate_fill_idregistro(sp, secrets_file, dry_run=True)

    # 2) Si el DRY RUN se ve bien, cambia a False:
    # migrate_fill_idregistro(sp, secrets_file, dry_run=False)
