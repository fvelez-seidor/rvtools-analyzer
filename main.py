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

import argparse
import pandas as pd
import json
import os
import sys
from datetime import datetime
import smtplib
from email.message import EmailMessage

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


def save_current(state, state_file=PREVIOUS_STATE):
    """
    Guarda el estado actual en archivo JSON para futuras comparaciones.

    Args:
        state (dict): Diccionario con el estado actual de anomalías
        state_file (str): Ruta del archivo JSON de estado.
    """
    with open(state_file, "w") as f:
        json.dump(state, f, indent=2)


def resolve_latest_rvtools_file(path):
    """Resolver archivo de entrada si path es directorio."""
    if not os.path.isdir(path):
        return path

    entries = [
        os.path.join(path, f)
        for f in os.listdir(path)
        if f.lower().endswith(".xlsx") and os.path.isfile(os.path.join(path, f))
    ]

    if not entries:
        raise FileNotFoundError(
            f"No se encontraron archivos .xlsx en el directorio {path}"
        )

    import re
    from datetime import datetime

    pattern = re.compile(
        r"^rvtools_(?P<prefix>.+?)_(?P<year>\d{4})-(?P<mon>\d{2})-(?P<day>\d{2})_(?P<hour>\d{2})\.(?P<min>\d{2})\.xlsx$",
        re.IGNORECASE,
    )

    matched = []
    for p in entries:
        name = os.path.basename(p)
        m = pattern.match(name)
        if m:
            try:
                ts = datetime(
                    int(m.group("year")),
                    int(m.group("mon")),
                    int(m.group("day")),
                    int(m.group("hour")),
                    int(m.group("min")),
                )
                matched.append((ts, p))
            except ValueError:
                continue

    if matched:
        matched.sort(key=lambda x: x[0], reverse=True)
        chosen = matched[0][1]
        print(f"Carpeta detectada, archivo elegido por nombre cronológico: {chosen}")
        return chosen

    entries.sort(key=lambda f: os.path.getmtime(f), reverse=True)
    chosen = entries[0]
    print(f"Carpeta detectada, archivo elegido por fecha de modificación: {chosen}")
    return chosen


# --------- Inicialización de estado y anomalías ---------

prev_state = load_previous()
anomalies = {}

# --------- Detección de anomalías en vHealth ---------


# Helper
def get_col(*names, sheet=None):
    for n in names:
        if n in sheet.columns:
            return sheet[n].astype(str)
    return None


def check_vhealth(CURRENT_FILE, load_sheet, anomalies):
    """
    Analiza la hoja vHealth del reporte RVTools para detectar anomalías.
    """

    vhealth = load_sheet(CURRENT_FILE, "vHealth")
    if vhealth is None or vhealth.empty:
        return

    vhealth.columns = vhealth.columns.str.strip().str.lower().str.replace(" ", "_")

    ignore_message_types = {
        "FOLDERNAME",
        "PERFORMANCE TIP",
        "SECURITY",
        "STORAGE",
        "USB",
        "CPU",
    }

    message_col = get_col("message", sheet=vhealth)
    type_col = get_col("message_type", sheet=vhealth)

    if message_col is not None:
        vhealth["message"] = message_col
        has_message = True
    else:
        vhealth["message"] = ""
        has_message = False

    if type_col is not None:
        vhealth["message_type"] = type_col
        has_type = True
    else:
        vhealth["message_type"] = ""
        has_type = False

    if has_type:
        problem_rows = vhealth[vhealth["message_type"].notna()]

        for msg_type, group in problem_rows.groupby("message_type"):
            if str(msg_type).strip().upper() in ignore_message_types:
                continue

            key = str(msg_type).lower().replace(" ", "_")

            cols = ["name", "message_type"]
            if has_message:
                cols.insert(1, "message")

            anomalies[key] = group[[c for c in cols if c in group.columns]].to_dict(
                "records"
            )

    if "cdrom" in vhealth.columns:
        cdrom = vhealth[
            vhealth["cdrom"].astype(str).str.contains("connected", case=False, na=False)
        ]

        anomalies["cdrom_connected"] = cdrom[
            [c for c in ["name", "message", "cdrom"] if c in cdrom.columns]
        ].to_dict("records")

    if "tools" in vhealth.columns:
        tools = vhealth[
            vhealth["tools"]
            .astype(str)
            .str.contains("old|not installed", case=False, na=False)
        ]

        anomalies["vmtools_issue"] = tools[
            [c for c in ["name", "message", "tools"] if c in tools.columns]
        ].to_dict("records")

    if "zombie" in vhealth.columns:
        zombies = vhealth[vhealth["zombie"].astype(str) != "0"]

        anomalies["zombies"] = zombies[
            [c for c in ["name", "message", "zombie"] if c in zombies.columns]
        ].to_dict("records")


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
    if vpartition is None or vpartition.empty:
        return

    vpartition.columns = (
        vpartition.columns.str.strip().str.lower().str.replace(" ", "_")
    )

    def has_col(name):
        return name in vpartition.columns

    free_col = None

    if has_col("free_%"):
        free_col = "free_%"
    elif has_col("free_percent"):
        free_col = "free_percent"
    elif has_col("free"):
        free_col = "free"
    else:
        return  # no usable data

    vpartition[free_col] = pd.to_numeric(vpartition[free_col], errors="coerce")

    low = vpartition[vpartition[free_col] < 10]

    if low.empty:
        return

    possible_cols = ["vm", "disk", free_col, "annotation"]
    available_cols = [c for c in possible_cols if c in low.columns]

    anomalies["low_disk_space"] = low[available_cols].to_dict("records")


def check_vdatastore(CURRENT_FILE, load_sheet, anomalies):
    """
    Analiza la hoja vDatastore del reporte RVTools para detectar bajo espacio en disco.

    Detecta particiones con menos del 10% de espacio libre.

    Args:
        CURRENT_FILE (str): Ruta del archivo RVTools
        load_sheet (function): Función para cargar hojas Excel
        anomalies (dict): Diccionario para guardar anomalías encontradas
    """
    vdatastore = load_sheet(CURRENT_FILE, "vDatastore")
    if vdatastore is None or vdatastore.empty:
        return

    vdatastore.columns = (
        vdatastore.columns.str.strip().str.lower().str.replace(" ", "_")
    )

    def has_col(name):
        return name in vdatastore.columns

    free_col = None

    if has_col("free_%"):
        free_col = "free_%"
    elif has_col("free_percent"):
        free_col = "free_percent"
    elif has_col("free"):
        free_col = "free"
    else:
        return  # no usable data

    vdatastore[free_col] = pd.to_numeric(vdatastore[free_col], errors="coerce")

    low = vdatastore[vdatastore[free_col] < 10]

    if low.empty:
        return

    possible_cols = ["vm", "disk", free_col, "annotation"]
    available_cols = [c for c in possible_cols if c in low.columns]

    anomalies["low_disk_space"] = low[available_cols].to_dict("records")


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


def export_xls(output_file, anomalies):
    """
    Exporta anomalías a un archivo Excel con una hoja por tipo de anomalía.

    Args:
        output_file (str): Ruta del archivo Excel de salida
        anomalies (dict): Diccionario con anomalías a exportar
    """
    writer = pd.ExcelWriter(output_file, engine="openpyxl")

    if not anomalies:
        pd.DataFrame({"info": ["No issues"]}).to_excel(
            writer, sheet_name="summary", index=False
        )
    else:
        for k, v in anomalies.items():
            if v:
                df = pd.DataFrame(v)
            else:
                df = pd.DataFrame({"info": ["No issues"]})

            sheet = k[:31]
            df.to_excel(writer, sheet_name=sheet, index=False)

    writer.close()


# --------- Exportación en Markdown ---------


def export_html(output_file, anomalies):
    """
    Genera un reporte completo en formato HTML con todas las anomalías.

    Args:
        output_file (str): Ruta del archivo HTML a generar
        anomalies (dict): Diccionario con todas las anomalías detectadas
    """
    total_issues = sum(len(v) for v in anomalies.values())

    html = """
    <!DOCTYPE html>
    <html lang='es'>
    <head>
      <meta charset='UTF-8'>
      <title>{title}</title>
      <style>
        body {{ font-family: Arial, sans-serif; line-height: 1.6; margin: 24px; }}
        h1,h2,h3,h4 {{ color: #2a4365; }}
        table {{ border-collapse: collapse; width: 100%; margin-bottom: 22px; }}
        th,td {{ border: 1px solid #ccc; padding: 8px; text-align: left; }}
        th {{ background: #f2f7ff; }}
        tr:nth-child(even) {{ background: #f8faff; }}
      </style>
    </head>
    <body>
      <h1>{title}</h1>
      <p><strong>Fecha de Generación:</strong> {date}</p>
      <p><strong>Generado por:</strong> {generated_by}</p>
      <h2>Resumen de Anomalías</h2>
      <p>Total: <strong>{total}</strong></p>
      <table>
        <thead><tr><th>Categoría</th><th>Cantidad</th></tr></thead>
        <tbody>
    """.format(
        title=ISSUE_CONFIG["title"],
        date=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        generated_by=ISSUE_CONFIG["generated_by"],
        total=total_issues,
    )

    for issue_type, items in anomalies.items():
        html += f"<tr><td>{issue_type.upper()}</td><td>{len(items)}</td></tr>\n"

    html += "</tbody></table>\n"

    html += "<h2>Detalles</h2>\n"

    for issue_type, items in anomalies.items():
        html += f"<h3>{issue_type.upper()} ({len(items)})</h3>\n"
        if not items:
            html += "<p>Sin anomalías detectadas.</p>\n"
            continue

        columns = sorted({k for item in items for k in item.keys()})

        html += "<table><thead><tr>"
        for col in columns:
            html += f"<th>{col}</th>"
        html += "</tr></thead><tbody>\n"

        for item in items:
            html += "<tr>"
            for col in columns:
                value = str(item.get(col, "")).replace("\n", " ")
                html += f"<td>{value}</td>"
            html += "</tr>\n"

        html += "</tbody></table>\n"

    html += "</body></html>"

    with open(output_file, "w", encoding="utf-8") as f:
        f.write(html)


def export_markdown(output_file, anomalies):
    """
    Genera un reporte completo en formato Markdown con todas las anomalías.

    Incluye resumen y detalle de cada tipo de anomalía encontrada en tablas.

    Args:
        output_file (str): Ruta del archivo Markdown a generar
        anomalies (dict): Diccionario con todas las anomalías detectadas
    """
    md = f"# {ISSUE_CONFIG['title']}\n\n"
    md += f"**Fecha de Generación:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
    md += f"**Generado por:** {ISSUE_CONFIG['generated_by']}\n\n"

    total_issues = sum(len(v) for v in anomalies.values())

    md += "## Resumen de Anomalías\n\n"
    md += "| Categoría | Cantidad |\n"
    md += "| --- | --- |\n"

    # Resumen por tipo en tabla
    for issue_type, items in anomalies.items():
        md += f"| {issue_type.upper()} | {len(items)} |\n"

    md += f"\n**Total General:** {total_issues} anomalías detectadas\n\n"
    md += "---\n\n"

    md += "## Detalle de Anomalías\n\n"

    # Detalle de cada anomalía
    for issue_type, items in anomalies.items():
        md += format_issue_for_markdown(issue_type, items)

    with open(output_file, "w") as f:
        f.write(md)


def format_issue_for_markdown(issue_type, items):
    """
    Formatea un tipo de anomalía en formato Markdown usando tabla.

    Args:
        issue_type (str): Tipo de anomalía
        items (list): Lista de anomalías de este tipo

    Returns:
        str: Texto en formato Markdown
    """
    if not items:
        return f"#### {issue_type.upper()}\n\nSin anomalías detectadas.\n\n"

    md = f"#### {issue_type.upper()} ({len(items)})\n\n"

    if items:
        # Obtener todas las columnas de los items
        all_columns = set()
        for item in items:
            all_columns.update(item.keys())

        # Reorganizar columnas: primero las básicas, luego Message si existe, luego Annotation si existe
        priority_order = [
            "Name",
            "VM",
            "Disk",
            "Free %",
            "Tools",
            "Zombie",
            "CDROM",
            "Message type",
            "Message",
            "Annotation",
        ]
        columns = []

        for col in priority_order:
            if col in all_columns:
                columns.append(col)
                all_columns.remove(col)

        # Agregar cualquier columna restante al final
        columns.extend(sorted(all_columns))

        # Crear encabezado de tabla
        md += "| " + " | ".join(columns) + " |\n"
        md += "| " + " | ".join(["---"] * len(columns)) + " |\n"

        # Agregar datos en filas
        for item in items:
            values = [str(item.get(col, "")).replace("\n", " ") for col in columns]
            md += "| " + " | ".join(values) + " |\n"

        md += "\n"

    return md


def send_email_via_relay(
    relay_host, sender, receiver, body, files=None, html_body=None
):
    # 1. Configuración del Relay
    relay_port = 25

    # 2. Crear el mensaje
    msg = EmailMessage()
    msg["Subject"] = "RVTools: Resumen de anomalías y archivo de reporte"
    msg["From"] = sender
    msg["To"] = receiver

    # Establecer contenido en texto y HTML (si no se proporciona, se genera a partir del texto)
    msg.set_content(body)
    if html_body:
        msg.add_alternative(html_body, subtype="html")
    else:
        fallback_html = "<html><body>"
        fallback_html += f"<p>{body.replace(chr(10), '<br/>')}</p>"
        fallback_html += "</body></html>"
        msg.add_alternative(fallback_html, subtype="html")

    if files is not None:
        for file in files:
            with open(file, "rb") as f:
                file_data = f.read()
                file_name = f.name
                msg.add_attachment(
                    file_data,
                    maintype="application",
                    subtype="octet-stream",
                    filename=file_name,
                )

    # 3. Enviar sin login
    try:
        # Conexión directa al host y puerto especificados
        with smtplib.SMTP(relay_host, relay_port) as server:
            # Nota: No llamamos a server.login() porque el relay confía en tu IP
            server.send_message(msg)
        print("Correo enviado correctamente a través del relay.")
    except Exception as e:
        print(f"Error al conectar con el relay: {e}")


def main():
    """
    Función principal que ejecuta el flujo completo del análisis.

    Proceso:
    1. Toma ruta del archivo RVTools desde argumentos de línea de comandos (o usa predeterminado)
    2. Carga datos y detecta anomalías
    3. Compara con estado anterior para identificar nuevas anomalías
    4. Exporta resultados a Excel y Markdown
    5. Guarda estado actual para futuras comparaciones
    6. Muestra resumen en consola
    """
    parser = argparse.ArgumentParser(
        description="Analiza reportes RVTools y detecta anomalías"
    )

    parser.add_argument(
        "file",
        nargs="?",
        default="RVTools_export.xlsx",
        help="Ruta al archivo RVTools (XLSX) o directorio que contiene archivos.",
    )

    parser.add_argument(
        "--more-info",
        "-i",
        action="store_true",
        help="Mostrar detalles completos de anomalías en consola.",
    )

    parser.add_argument(
        "--overwrite",
        "-y",
        action="store_true",
        help="Sobrescribir estado previo sin pedir confirmación.",
    )

    parser.add_argument(
        "--output",
        "-o",
        default=OUTPUT_FILE,
        help="Ruta de salida del reporte Excel.",
    )

    parser.add_argument(
        "--state",
        default=PREVIOUS_STATE,
        help="Ruta del archivo de estado previo JSON.",
    )

    parser.add_argument(
        "--markdown",
        default=os.path.join(OUTPUT_DIR, "current_issues.md"),
        help="Ruta del reporte Markdown generado.",
    )

    parser.add_argument(
        "--html",
        default=os.path.join(OUTPUT_DIR, "current_issues.html"),
        help="Ruta del reporte HTML generado.",
    )

    parser.add_argument(
        "--report-format",
        choices=["md", "html"],
        default="html",
        help="Formato del reporte de texto (md o html).",
    )

    parser.add_argument(
        "--full-output",
        default=os.path.join(OUTPUT_DIR, "rvtools_report_full.xlsx"),
        help="Ruta del reporte XLSX completo con todas las anomalías.",
    )

    parser.add_argument(
        "--auto-folder",
        action="store_true",
        help="Si se pasa un directorio, usa el último archivo RVTools disponible.",
    )

    parser.add_argument(
        "--send-email",
        action="store_true",
        help="Enviar notificación por email a través de relay interno.",
    )

    parser.add_argument(
        "--email-relay",
        default="192.168.1.100",
        help="IP o hostname del relay SMTP para enviar email (usado con --send-email).",
    )

    parser.add_argument(
        "--email-sender",
        default="servidor@tudominio.com",
        help="Dirección de correo del remitente (usado con --send-email).",
    )

    parser.add_argument(
        "--email-receiver",
        default="destino@ejemplo.com",
        help="Dirección de correo del destinatario (usado con --send-email). Se pueden definir varios destinatarios separados por comas.",
    )

    args = parser.parse_args()

    if len(sys.argv) == 1:
        parser.print_help()
        sys.exit(0)

    if os.path.isdir(args.file):
        if args.auto_folder:
            CURRENT_FILE = resolve_latest_rvtools_file(args.file)
        else:
            parser.error(
                "El argumento 'file' es un directorio. Use '--auto-folder' para seleccionar el último archivo dentro del directorio."
            )
    else:
        CURRENT_FILE = args.file

    more_info = args.more_info
    output_file = args.output
    previous_state_file = args.state

    check_vhealth(CURRENT_FILE, load_sheet, anomalies)
    check_vpartition(CURRENT_FILE, load_sheet, anomalies)
    check_vdatastore(CURRENT_FILE, load_sheet, anomalies)

    # Leer estado previo según argumento si existe
    global prev_state
    if previous_state_file != PREVIOUS_STATE:
        # support custom state path without updating global constant permanently
        if os.path.exists(previous_state_file):
            with open(previous_state_file, "r") as f:
                prev_state = json.load(f)
        else:
            prev_state = {}

    new_anomalies = compare_previous(prev_state, anomalies)
    export_xls(output_file, new_anomalies)
    export_xls(args.full_output, anomalies)

    report_format = args.report_format.lower()
    if report_format == "html":
        export_html(args.html, anomalies)
        report_file = args.html
        print(f"Reporte HTML: {report_file}")
    else:
        export_markdown(args.markdown, anomalies)
        report_file = args.markdown
        print(f"Reporte Markdown: {report_file}")

    # Guardar estado actual
    if os.path.exists(previous_state_file) and not args.overwrite:
        print(
            f"El archivo {previous_state_file} ya existe. Use --overwrite para sobrescribir."
        )
    else:
        save_current(anomalies, previous_state_file)

    print("\n===== RESUMEN DE ANOMALÍAS RVTools =====\n")
    total_anomalies = sum(len(v) for v in anomalies.values())
    print(f"Total de anomalías detectadas: {total_anomalies}\n")
    print("Anomalías por tipo:")
    for k, v in anomalies.items():
        print(f"{k.upper():20} : {len(v)}")

    if more_info:
        print("\n===== ANOMALÍAS DETALLADAS =====\n")
        for k, v in anomalies.items():
            print(f"--- {k.upper()} ---")
            if v:
                for item in v:
                    print(item)
            else:
                print("No se detectaron anomalías")
            print()

    print("\n===== NUEVAS ANOMALÍAS DETECTADAS =====\n")
    total = 0
    for k, v in new_anomalies.items():
        print(f"{k.upper():20} : {len(v)}")
        total += len(v)

    print(f"\nTOTAL DE NUEVAS ANOMALÍAS: {total}")
    print(f"Reporte Excel: {output_file}")

    if more_info:
        print("\n===== NUEVAS ANOMALÍAS DETALLADAS =====\n")
        for k, v in new_anomalies.items():
            print(f"--- {k.upper()} ---")
            if v:
                for item in v:
                    print(item)
            else:
                print("No hay nuevas anomalías")
            print()

    if args.send_email:
        files = [output_file, args.full_output, report_file]
        print("\nEnviando notificación por email a través del relay...")

        # Resumen de estadísticas para el cuerpo del correo
        total_current = sum(len(v) for v in anomalies.values())
        new_total = sum(len(v) for v in new_anomalies.values())

        header_plain = "RVTools - Resumen de Anomalías\n"
        header_plain += "========================================\n"
        header_plain += f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
        header_plain += f"Total anomalías detectadas: {total_current}\n"
        header_plain += f"Nuevas anomalías respecto al estado previo: {new_total}\n\n"

        body_plain = header_plain
        body_plain += "Resumen por categoría:\n"
        for key, items in anomalies.items():
            body_plain += f"- {key.upper():20}: {len(items)}\n"

        body_plain += "\nDetalles completos en los archivos adjuntos:\n"
        body_plain += f"- Ruta (nuevas anomalías): {output_file}\n"
        body_plain += f"- Ruta (todas las anomalías): {args.full_output}\n"
        body_plain += f"- Ruta (reporte legible): {report_file}\n"

        body_plain += "\nPor favor revise primero el archivo de resumen en HTML y luego el XLSX completo para auditoría y seguimiento."

        # Cuerpo HTML con tabla de resumen
        rows_summary = ""
        for key, items in anomalies.items():
            rows_summary += f"<tr><td>{key.upper()}</td><td>{len(items)}</td></tr>"

        html_body = f"""
        <!DOCTYPE html>
        <html lang='es'>
        <head>
          <meta charset='UTF-8'>
          <style>
            body {{ font-family: Arial, sans-serif; line-height: 1.5; margin: 16px; }}
            h1 {{ color: #1f4e79; }}
            table {{ border-collapse: collapse; width: 100%; max-width: 800px; }}
            th, td {{ border: 1px solid #a8c4e5; padding: 8px; }}
            th {{ background-color: #dde9f7; text-align: left; }}
            tr:nth-child(even) {{ background-color: #f6f9ff; }}
          </style>
        </head>
        <body>
          <h1>RVTools - Resumen de Anomalías</h1>
          <p><strong>Fecha:</strong> {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>
          <p><strong>Total anomalías detectadas:</strong> {total_current}</p>
          <p><strong>Nuevas anomalías:</strong> {new_total}</p>
          <h2>Resumen por categoría</h2>
          <table>
            <thead><tr><th>Categoría</th><th>Cantidad</th></tr></thead>
            <tbody>
              {rows_summary}
            </tbody>
          </table>
          <p>Detalles completos adjuntos en archivo XLSX y reporte legible.</p>
          <ul>
            <li>Excel nuevas: {output_file}</li>
            <li>Excel completo: {args.full_output}</li>
            <li>Reporte legible: {report_file}</li>
          </ul>
          <p>Por favor revisar primero el resumen y luego el detalle completo para auditoría.</p>
        </body>
        </html>
        """

        send_email_via_relay(
            args.email_relay,
            args.email_sender,
            args.email_receiver,
            body_plain,
            files=files,
            html_body=html_body,
        )


if __name__ == "__main__":
    # Ejecutar función principal
    main()
