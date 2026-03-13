"""
Script de análisis automatizado de reportes RVTools.

Este módulo procesa reportes de RVTools (vHealth y vPartition) para detectar
anomalías en la infraestructura de virtualización. Compara el estado actual
con el estado anterior para identificar nuevos problemas y genera reportes
en Excel y Markdown.

Características principales:
- Carga datos de hojas vHealth y vPartition desde archivos Excel
- Detecta automáticamente anomalías (bajo espacio en disco, Tools desactualizados, etc.)
- Compara estado actual con estado previo para identificar nuevos problemas
- Exporta resultados en Excel y Markdown
- Mantiene historial de estado para comparaciones futuras
"""

import pandas as pd
import json
import os
from datetime import datetime

# Directorio de salida para reportes
OUTPUT_DIR = "rvtools_reports"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Ruta del archivo de reporte Excel generado con fecha actual
OUTPUT_FILE = os.path.join(
    OUTPUT_DIR, f"rvtools_report_{datetime.now().strftime('%Y%m%d')}.xlsx"
)

# Archivo que guarda el estado anterior para comparación
PREVIOUS_STATE = os.path.join(OUTPUT_DIR, "rvtools_previous_state.json")

# Configuración de datos del reporte
ISSUE_CONFIG = {
    "title": "RVTools Infrastructure Issues Report",
    "generated_by": "RVTools automated check",
}

# -------------------------


def load_sheet(file, sheet):
    """
    Carga una hoja específica de un archivo Excel.

    Args:
        file (str): Ruta del archivo Excel
        sheet (str): Nombre de la hoja a cargar

    Returns:
        DataFrame: Datos de la hoja o None si ocurre un error
    """
    try:
        return pd.read_excel(file, sheet_name=sheet)
    except Exception:
        return None


def load_previous():
    """
    Carga el estado anterior guardado desde archivo JSON.

    Returns:
        dict: Estado anterior o diccionario vacío si no existe el archivo
    """
    if os.path.exists(PREVIOUS_STATE):
        with open(PREVIOUS_STATE, "r") as f:
            return json.load(f)
    return {}


def save_current(state):
    """
    Guarda el estado actual en archivo JSON para futuras comparaciones.

    Args:
        state (dict): Diccionario con el estado actual de anomalías
    """
    with open(PREVIOUS_STATE, "w") as f:
        json.dump(state, f, indent=2)


# --------- Inicialización de estado y anomalías ---------

prev_state = load_previous()
anomalies = {}

# --------- Detección de anomalías en vHealth ---------


def check_vhealth(CURRENT_FILE, load_sheet, anomalies):
    """
    Analiza la hoja vHealth del reporte RVTools para detectar anomalías.

    Detecta:
    - Tipos de mensajes de problemas (excepto ciertos tipos ignorados)
    - CDROMs conectados
    - VMware Tools desactualizados o no instalados
    - VMs zombies

    Args:
        CURRENT_FILE (str): Ruta del archivo RVTools
        load_sheet (function): Función para cargar hojas Excel
        anomalies (dict): Diccionario para guardar anomalías encontradas
    """
    vhealth = load_sheet(CURRENT_FILE, "vHealth")

    ignore_message_types = [
        "FOLDERNAME",
        "PERFORMANCE TIP",
        "SECURITY",
        "STORAGE",
        "USB",
    ]

    if vhealth is not None:
        vhealth["Message type"] = vhealth["Message type"].astype(str)

        # Detectar todos los tipos de problemas automáticamente
        problem_rows = vhealth[vhealth["Message type"].notna()]

        for msg_type, group in problem_rows.groupby("Message type"):
            if msg_type.strip().upper() in ignore_message_types:
                continue
            key = msg_type.lower().replace(" ", "_")

            anomalies[key] = group[["Name", "Message type"]].to_dict("records")

        # Detectar CDROMs conectados
        if "CDROM" in vhealth.columns:
            cdrom = vhealth[
                vhealth["CDROM"].str.contains("connected", case=False, na=False)
            ]
            anomalies["cdrom_connected"] = cdrom[["Name", "CDROM"]].to_dict("records")

        # Detectar VMware Tools desactualizados o no instalados
        if "Tools" in vhealth.columns:
            tools = vhealth[
                vhealth["Tools"].str.contains("old|not installed", case=False, na=False)
            ]
            anomalies["vmtools_issue"] = tools[["Name", "Tools"]].to_dict("records")

        # Detectar VMs zombies
        if "Zombie" in vhealth.columns:
            zombies = vhealth[vhealth["Zombie"].astype(str) != "0"]
            anomalies["zombies"] = zombies[["Name", "Zombie"]].to_dict("records")


# --------- Detección de anomalías en vPartition ---------


def check_vpartition(CURRENT_FILE, load_sheet, anomalies):
    """
    Analiza la hoja vPartition del reporte RVTools para detectar bajo espacio en disco.

    Detecta particiones con menos del 10% de espacio libre.

    Args:
        CURRENT_FILE (str): Ruta del archivo RVTools
        load_sheet (function): Función para cargar hojas Excel
        anomalies (dict): Diccionario para guardar anomalías encontradas
    """
    vpartition = load_sheet(CURRENT_FILE, "vPartition")

    if vpartition is not None:
        vpartition["Free %"] = pd.to_numeric(vpartition["Free %"], errors="coerce")

        # Filtrar particiones con menos del 10% de espacio libre
        low = vpartition[vpartition["Free %"] < 10]

        anomalies["low_disk_space"] = low[["VM", "Free %"]].to_dict("records")


# --------- Comparación con estado previo ---------


def compare_previous(prev_state, anomalies):
    """
    Compara anomalías actuales con el estado anterior para identificar nuevos problemas.

    Solo devuelve anomalías que no estaban presentes en el estado anterior.

    Args:
        prev_state (dict): Estado anterior cargado desde JSON
        anomalies (dict): Anomalías actuales detectadas

    Returns:
        dict: Diccionario con solo las anomalías nuevas
    """

    def filter_new(key):
        prev = prev_state.get(key, [])
        curr = anomalies.get(key, [])

        # Comparar usando JSON para detectar diferencias exactas
        prev_set = {json.dumps(x, sort_keys=True) for x in prev}

        new_items = [x for x in curr if json.dumps(x, sort_keys=True) not in prev_set]

        return new_items

    new_anomalies = {k: filter_new(k) for k in anomalies}
    return new_anomalies


# --------- Exportación en Excel ---------


def export_xls(OUTPUT_FILE, new_anomalies):
    """
    Exporta anomalías a un archivo Excel con una hoja por tipo de anomalía.

    Args:
        OUTPUT_FILE (str): Ruta del archivo Excel de salida
        new_anomalies (dict): Diccionario con anomalías nuevas a exportar
    """
    writer = pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl")

    for k, v in new_anomalies.items():
        if v:
            df = pd.DataFrame(v)
        else:
            df = pd.DataFrame({"info": ["No new issues"]})

        # Limitar nombre de hoja a 31 caracteres (límite de Excel)
        sheet = k[:31]
        df.to_excel(writer, sheet_name=sheet, index=False)

    writer.close()


# --------- Exportación en Markdown ---------


def export_markdown(output_file, anomalies):
    """
    Genera un reporte completo en formato Markdown con todas las anomalías.

    Incluye resumen y detalle de cada tipo de anomalía encontrada.

    Args:
        output_file (str): Ruta del archivo Markdown a generar
        anomalies (dict): Diccionario con todas las anomalías detectadas
    """
    md = f"# {ISSUE_CONFIG['title']}\n\n"
    md += f"**Generado:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
    md += "**Tipo de Reporte:** Todas las anomalías actuales\n\n"

    total_issues = sum(len(v) for v in anomalies.values())
    md += "## Resumen\n\n"
    md += f"**Total de anomalías detectadas:** {total_issues}\n\n"

    # Resumen por tipo
    for issue_type, items in anomalies.items():
        md += f"- {issue_type.upper()}: {len(items)}\n"

    md += "\n---\n\n"
    md += "## Detalle de Anomalías\n\n"

    # Detalle de cada anomalía
    for issue_type, items in anomalies.items():
        md += format_issue_for_markdown(issue_type, items)
        md += "---\n\n"

    with open(output_file, "w") as f:
        f.write(md)


def format_issue_for_markdown(issue_type, items):
    """
    Formatea un tipo de anomalía en formato Markdown.

    Args:
        issue_type (str): Tipo de anomalía
        items (list): Lista de anomalías de este tipo

    Returns:
        str: Texto en formato Markdown
    """
    if not items:
        return f"### {issue_type.upper()}\n\nNo se detectaron anomalías.\n\n"

    md = f"### {issue_type.upper()} ({len(items)})\n\n"

    for idx, item in enumerate(items, 1):
        # Listar los detalles de cada anomalía
        for key, value in item.items():
            md += f"- {key}: {value}\n"
        md += "\n"

    return md


def main():
    """
    Función principal que ejecuta el flujo completo del análisis.

    Proceso:
    1. Solicita ruta del archivo RVTools (o usa predeterminado)
    2. Carga datos y detecta anomalías
    3. Compara con estado anterior para identificar nuevas anomalías
    4. Exporta resultados a Excel y Markdown
    5. Guarda estado actual para futuras comparaciones
    6. Muestra resumen en consola
    """
    # --------- Solicitar entrada del usuario ---------

    CURRENT_FILE = str(
        input("Ingrese la ruta del archivo RVTools (defecto: RVTools_export.xlsx): ")
        or "RVTools_export.xlsx"
    )

    more_info = (
        input(
            "¿Desea ver detalles adicionales de las anomalías en consola? (s/n, defecto: n): "
        ).lower()
        or "n"
    )

    # --------- Cargar datos y detectar anomalías ---------
    check_vhealth(CURRENT_FILE, load_sheet, anomalies)
    check_vpartition(CURRENT_FILE, load_sheet, anomalies)

    # --------- Comparar con estado previo y exportar ---------
    new_anomalies = compare_previous(prev_state, anomalies)
    export_xls(OUTPUT_FILE, new_anomalies)

    # --------- Exportar reporte en Markdown ---------
    MD_REPORT = os.path.join(OUTPUT_DIR, "current_issues.md")
    export_markdown(MD_REPORT, anomalies)
    print(f"Reporte Markdown: {MD_REPORT}")

    # --------- Guardar estado actual para futuras comparaciones ---------

    if os.path.exists(OUTPUT_FILE):
        overwrite = input(
            f"El archivo {OUTPUT_FILE} ya existe. ¿Desea sobrescribirlo? (s/n): "
        ).lower()
        if overwrite == "s":
            save_current(anomalies)
    else:
        save_current(anomalies)

    # --------- Mostrar resumen por consola ---------

    print("\n===== RESUMEN DE ANOMALÍAS RVTools =====\n")
    total_anomalies = sum(len(v) for v in anomalies.values())
    print(f"Total de anomalías detectadas: {total_anomalies}\n")
    print("Anomalías por tipo:")
    for k, v in anomalies.items():
        print(f"{k.upper():20} : {len(v)}")

    if more_info == "s" or more_info == "y":
        print("\n===== ANOMALÍAS DETALLADAS =====\n")
        for k, v in anomalies.items():
            print(f"--- {k.upper()} ---")
            if v:
                for item in v:
                    print(item)
            else:
                print("No se detectaron anomalías")
            print()

    # --------- Mostrar nuevas anomalías ---------
    print("\n===== NUEVAS ANOMALÍAS DETECTADAS =====\n")
    total = 0
    for k, v in new_anomalies.items():
        print(f"{k.upper():20} : {len(v)}")
        total += len(v)

    print(f"\nTOTAL DE NUEVAS ANOMALÍAS: {total}")
    print(f"Reporte Excel: {OUTPUT_FILE}")

    if more_info == "s" or more_info == "y":
        print("\n===== NUEVAS ANOMALÍAS DETALLADAS =====\n")
        for k, v in new_anomalies.items():
            print(f"--- {k.upper()} ---")
            if v:
                for item in v:
                    print(item)
            else:
                print("No hay nuevas anomalías")
            print()


if __name__ == "__main__":
    # Ejecutar función principal
    main()
