import re

PERIODO_REGEX = re.compile(r"^\d{4}-\d{4}$")

def validar_formulario(datos: dict) -> list[str]:
    errores = []

    # Periodo PEI obligatorio y con formato YYYY-YYYY
    periodo = str(datos.get("periodo", "")).strip()
    if not periodo:
        errores.append("Periodo PEI es obligatorio.")
    elif not PERIODO_REGEX.match(periodo):
        errores.append("Periodo PEI debe tener el formato YYYY-YYYY (ej. 2028-2033).")

    # Regla que ya tienes: si Estado=Emitido, exige campos
    if datos.get("estado") == "Emitido":
        if not str(datos.get("expediente", "")).strip():
            errores.append("Para Estado=Emitido debes completar Expediente (SGD).")
        if not str(datos.get("numero_it", "")).strip():
            errores.append("Para Estado=Emitido debes completar Número de I.T.")
        if not str(datos.get("fecha_it", "")).strip():
            errores.append("Para Estado=Emitido debes completar Fecha de I.T.")

    return errores
