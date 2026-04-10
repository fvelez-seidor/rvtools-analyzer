"""
RVTools Infrastructure Analysis Tool.

Analyzes RVTools reports (vHealth and vPartition) to detect anomalies
in virtualization infrastructure. Compares current state with previous state
to identify new issues and generates reports in Excel and Markdown.

Main features:
- Automatically detects low disk space, outdated VMware Tools, etc.
- Compares current state with previous state to identify new issues
- Exports results in Excel and Markdown formats
- Maintains state history for future comparisons
"""

import argparse
import pandas as pd
import json
import os
import sys
from datetime import datetime
from pathlib import Path
from jinja2 import Environment, FileSystemLoader
from email_utils.email_handler import EmailHandler
from email_utils.config.email_config import ATTACHMENTS_TO_INCLUDE

# Output directory for reports
OUTPUT_DIR = "rvtools_reports"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Excel report file path with current date
OUTPUT_FILE = os.path.join(
    OUTPUT_DIR, f"rvtools_report_{datetime.now().strftime('%Y%m%d')}.xlsx"
)

# JSON file storing previous state for comparison
PREVIOUS_STATE = os.path.join(OUTPUT_DIR, "rvtools_previous_state.json")

# Report configuration
ISSUE_CONFIG = {
    "title": "RVTools Infrastructure Issues Report",
    "generated_by": "RVTools automated check",
}


def load_sheet(file, sheet):
    """Load and return a specific sheet from an Excel file.

    Args:
        file: Path to the Excel file
        sheet: Sheet name to load

    Returns:
        DataFrame with sheet data or None if error occurs
    """
    try:
        return pd.read_excel(file, sheet_name=sheet)
    except Exception:
        return None


def load_previous():
    """Load previous state from JSON file.

    Returns:
        dict: Previous state or empty dict if file doesn't exist
    """
    if os.path.exists(PREVIOUS_STATE):
        with open(PREVIOUS_STATE, "r") as f:
            return json.load(f)
    return {}


def save_current(state, state_file=PREVIOUS_STATE):
    """Save current state to JSON file for future comparisons.

    Args:
        state: Dictionary with current anomalies state
        state_file: Path to the JSON state file
    """
    with open(state_file, "w") as f:
        json.dump(state, f, indent=2)


def resolve_latest_rvtools_file(path):
    """Resolve input file if path is a directory.

    Searches for .xlsx files and selects the newest by timestamp pattern
    or modification time.
    """
    if not os.path.isdir(path):
        return path

    entries = [
        os.path.join(path, f)
        for f in os.listdir(path)
        if f.lower().endswith(".xlsx") and os.path.isfile(os.path.join(path, f))
    ]

    if not entries:
        raise FileNotFoundError(f"No .xlsx files found in directory {path}")

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
        print(f"Directory detected, file selected by chronological name: {chosen}")
        return chosen

    entries.sort(key=lambda f: os.path.getmtime(f), reverse=True)
    chosen = entries[0]
    print(f"Directory detected, file selected by modification date: {chosen}")
    return chosen


def resolve_recursive_rvtools_files(root_path):
    """
    Recursively search for .xlsx files organized by hostname subdirectories.
    
    Scans the root_path for subdirectories (each named after a hostname) and
    collects all .xlsx files found within each subdirectory, returning a
    dictionary mapping hostname → list of file paths.
    
    Args:
        root_path: Root directory to search (rvtools_export_output_dir)
        
    Returns:
        dict: Mapping of hostname -> list of latest xlsx files from each subdirectory
              Returns single files per hostname (the newest by timestamp)
    """
    if not os.path.isdir(root_path):
        raise NotADirectoryError(f"Path {root_path} is not a directory")
    
    hosts_files = {}
    
    # Iterate through subdirectories (each should be a hostname)
    for entry in os.listdir(root_path):
        entry_path = os.path.join(root_path, entry)
        
        # Skip if not a directory
        if not os.path.isdir(entry_path):
            continue
        
        # Find all xlsx files in this subdirectory
        xlsx_files = []
        for root, dirs, files in os.walk(entry_path):
            for f in files:
                if f.lower().endswith(".xlsx"):
                    full_path = os.path.join(root, f)
                    xlsx_files.append(full_path)
        
        if xlsx_files:
            # Select the newest file by timestamp pattern or modification time
            latest_file = _select_latest_xlsx(xlsx_files)
            hosts_files[entry] = latest_file
    
    if not hosts_files:
        raise FileNotFoundError(f"No .xlsx files found in any subdirectory of {root_path}")
    
    return hosts_files


def _select_latest_xlsx(files_list):
    """
    Select the newest xlsx file from a list by timestamp pattern or modification time.
    
    Args:
        files_list: List of file paths to evaluate
        
    Returns:
        str: Path to the newest file
    """
    import re
    from datetime import datetime
    
    pattern = re.compile(
        r"^rvtools_(?P<prefix>.+?)_(?P<year>\d{4})-(?P<mon>\d{2})-(?P<day>\d{2})_(?P<hour>\d{2})\.(?P<min>\d{2})\.xlsx$",
        re.IGNORECASE,
    )
    
    matched = []
    for p in files_list:
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
        return matched[0][1]
    
    # Fall back to modification time if no pattern matches
    return max(files_list, key=lambda f: os.path.getmtime(f))


# State initialization and anomaly detection

prev_state = load_previous()
anomalies = {}


# vHealth anomaly detection
def get_col(*names, sheet=None):
    for n in names:
        if n in sheet.columns:
            return sheet[n].astype(str)
    return None


def check_vhealth(CURRENT_FILE, load_sheet, anomalies):
    """Analyze vHealth sheet from RVTools report for anomalies."""

    vhealth = load_sheet(CURRENT_FILE, "vHealth")
    if vhealth is None or vhealth.empty:
        return

    vhealth.columns = vhealth.columns.str.strip().str.lower().str.replace(" ", "_")

    ignore_message_types = {
        "FOLDERNAME",
        "PERFORMANCE TIP",
        "SECURITY",
        "STORAGE",
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

    if "snapshot" in vhealth.columns:
        snapshots = vhealth[
            vhealth["snapshot"]
            .astype(str)
            .str.contains("present", case=False, na=False)
        ]

        anomalies["snapshot_present"] = snapshots[
            [
                c
                for c in ["name", "message", "snapshot", "age"]
                if c in snapshots.columns
            ]
        ].to_dict("records")

    if "cdrom" in vhealth.columns:
        cdrom = vhealth[
            vhealth["cdrom"].astype(str).str.contains("connected", case=False, na=False)
        ]

        anomalies["cdrom_connected"] = cdrom[
            [c for c in ["name", "message", "cdrom"] if c in cdrom.columns]
        ].to_dict("records")

    if "usb" in vhealth.columns:
        usb = vhealth[
            vhealth["usb"].astype(str).str.contains("connected", case=False, na=False)
        ]

        anomalies["usb_connected"] = usb[
            [c for c in ["name", "message", "usb"] if c in usb.columns]
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


# vPartition anomaly detection


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

    anomalies["low_disk_space_partition"] = low[available_cols].to_dict("records")


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

    possible_cols = ["name", "hosts", free_col, "address"]
    available_cols = [c for c in possible_cols if c in low.columns]

    anomalies["low_disk_space_datastore"] = low[available_cols].to_dict("records")


def check_report_age(file_path, anomalies, max_days=60):
    """
    Verifica la antigüedad del reporte y advierte si es mayor a max_days.

    Args:
        file_path (str): Ruta del archivo de reporte
        anomalies (dict): Diccionario de anomalías
        max_days (int): Número máximo de días permitidos (por defecto 60)
    """
    # Get the modification date of the current file
    if not os.path.exists(file_path):
        return

    file_mtime = os.path.getmtime(file_path)
    file_date = datetime.fromtimestamp(file_mtime)
    age_days = (datetime.now() - file_date).days

    if age_days > max_days:
        anomalies["old_report_warning"] = [
            {
                "file": os.path.basename(file_path),
                "age_days": age_days,
                "last_modified": file_date.strftime("%Y-%m-%d %H:%M:%S"),
                "warning": f"Report is {age_days} days old (threshold: {max_days} days)",
            }
        ]


# Compare with previous state


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


# Export to Excel


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


# Export to Markdown


def export_html(output_file, anomalies, new_anomalies=None):
    """
    Genera un reporte HTML usando el template Jinja2.

    Args:
        output_file (str): Ruta del archivo HTML a generar
        anomalies (dict): Diccionario con todas las anomalías detectadas
        new_anomalies (dict, optional): Diccionario con nuevas anomalías
    """
    if new_anomalies is None:
        new_anomalies = {}

    # Categorize issues
    critical_cats = ["zombie", "low_disk_space_datastore", "low_disk_space_partition"]
    critical_items = []
    warning_items = []
    new_items = []

    for key, items in anomalies.items():
        if key in critical_cats and items:
            critical_items.append(f"{key.replace('_', ' ').title()}: {len(items)}")
        elif key != "old_report_warning" and items:
            warning_items.append(f"{key.replace('_', ' ').title()}: {len(items)}")

    for key, items in new_anomalies.items():
        if items and key != "old_report_warning":
            new_items.append(f"{key.replace('_', ' ').title()}: {len(items)} new")

    total_issues = sum(len(v) for v in anomalies.values())

    # Prepare template data
    template_data = {
        "total_issues": total_issues,
        "critical_items": critical_items,
        "warning_items": warning_items,
        "new_items": new_items,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }

    # Setup Jinja2 environment to load from email_utils directory
    template_dir = Path(__file__).parent / "email_utils"
    env = Environment(loader=FileSystemLoader(str(template_dir)))

    try:
        template = env.get_template("email_template.j2")
        html_content = template.render(**template_data)

        with open(output_file, "w", encoding="utf-8") as f:
            f.write(html_content)
    except Exception as e:
        print(f"Error rendering HTML template: {e}")
        # Fallback to simple HTML generation
        _export_html_fallback(output_file, anomalies)


def _export_html_fallback(output_file, anomalies):
    """
    Fallback function to generate simple HTML if template rendering fails.

    Args:
        output_file (str): Path to output HTML file
        anomalies (dict): Dictionary with all detected anomalies
    """
    total_issues = sum(len(v) for v in anomalies.values())

    html = """<!DOCTYPE html>
<html>
<head>
  <meta charset='UTF-8'>
  <title>RVTools Issues Report</title>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif; color: #333; line-height: 1.6; margin: 24px; background: #f5f5f5; }
    .container { max-width: 900px; margin: 0 auto; background: white; padding: 32px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
    h1 { color: #4158D9; border-bottom: 2px solid #5079F2; padding-bottom: 12px; }
    h2 { color: #4158D9; margin-top: 24px; }
    table { border-collapse: collapse; width: 100%; margin: 16px 0; }
    th { background: #F2F2F2; border: 1px solid #ddd; padding: 12px; text-align: left; font-weight: 600; color: #4158D9; }
    td { border: 1px solid #ddd; padding: 10px; }
    tr:nth-child(even) { background: #fafafa; }
    .warning-row { background-color: #fffbf6 !important; }
    .critical-row { background-color: #fef9f9 !important; }
    .warning-box { background-color: #fffaf0; border-left: 4px solid #f39c12; padding: 12px; margin: 16px 0; border-radius: 4px; }
  </style>
</head>
<body>
  <div class="container">
    <h1>📊 RVTools Issues Report</h1>
    <p><strong>Generated:</strong> {date}</p>
""".format(date=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    html += f"    <h2>Summary</h2>\n    <p>Total Issues: <strong>{total_issues}</strong></p>\n"
    html += "    <table><thead><tr><th>Category</th><th>Count</th></tr></thead><tbody>\n"

    for issue_type, items in anomalies.items():
        html += f"      <tr><td>{issue_type.upper()}</td><td>{len(items)}</td></tr>\n"

    html += "    </tbody></table>\n"

    for issue_type, items in anomalies.items():
        html += f"    <h2>{issue_type.upper()}</h2>\n"
        if not items:
            html += "    <p>No anomalies detected.</p>\n"
            continue

        html += "    <table><thead><tr>"
        columns = sorted({k for item in items for k in item.keys()})
        for col in columns:
            html += f"<th>{col}</th>"
        html += "</tr></thead><tbody>\n"

        for item in items:
            html += "      <tr>"
            for col in columns:
                value = str(item.get(col, "")).replace("\n", " ")
                html += f"<td>{value}</td>"
            html += "</tr>\n"

        html += "    </tbody></table>\n"

    html += "  </div>\n</body>\n</html>"

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

    # Show warning box if old report
    if "old_report_warning" in anomalies and anomalies["old_report_warning"]:
        warning_item = anomalies["old_report_warning"][0]
        md += f"⚠️ **WARNING: {warning_item.get('warning', 'Old report detected')}**\n"
        md += f"Last modified: {warning_item.get('last_modified', 'N/A')}\n\n"
        md += "---\n\n"

    md += "## Anomalies Summary\n\n"
    md += "| Categoría | Cantidad |\n"
    md += "| --- | --- |\n"

    # Resumen por tipo en tabla
    for issue_type, items in anomalies.items():
        md += f"| {issue_type.upper()} | {len(items)} |\n"

    md += f"\n**Total General:** {total_issues} anomalías detectadas\n\n"
    md += "---\n\n"

    md += "## Anomalies Details\n\n"

    # Iterate through each anomaly type
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
        return f"#### {issue_type.upper()}\n\nNo anomalies detected.\n\n"

    md = f"#### {issue_type.upper()} ({len(items)})\n\n"

    if items:
        # Obtener todas las columnas de los items
        all_columns = set()
        for item in items:
            all_columns.update(item.keys())

        # Reorganize columns: basic first, then Message if exists, then Annotation if exists
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


def generate_email(anomalies: dict, new_anomalies: dict) -> bool:
    """
    Generate email template and metadata for ansible-email dispatcher.

    Args:
        anomalies: All detected anomalies
        new_anomalies: New anomalies since last run

    Returns:
        True if successful, False otherwise
    """
    handler = EmailHandler("rvtools-analyzer")

    # Categorize issues
    critical_cats = ["zombie", "low_disk_space_datastore", "low_disk_space_partition"]
    critical_items = []
    warning_items = []
    new_items = []

    for key, items in anomalies.items():
        if key in critical_cats and items:
            critical_items.append(f"{key.replace('_', ' ').title()}: {len(items)}")
        elif key != "old_report_warning" and items:
            warning_items.append(f"{key.replace('_', ' ').title()}: {len(items)}")

    for key, items in new_anomalies.items():
        if items and key != "old_report_warning":
            new_items.append(f"{key.replace('_', ' ').title()}: {len(items)} new")

    total_issues = sum(len(v) for v in anomalies.values())

    # Prepare template data
    template_data = {
        "total_issues": total_issues,
        "critical_items": critical_items,
        "warning_items": warning_items,
        "new_items": new_items,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }

    # Generate email with attachments
    return handler.generate(
        template_name="email_template.j2",
        template_data=template_data,
        attachments=ATTACHMENTS_TO_INCLUDE,
    )


def export_html_multi_host(output_file, hosts_results):
    """
    Generate a multi-host HTML report with sections for each hostname.
    
    Args:
        output_file: Path to output HTML file
        hosts_results: dict mapping hostname -> (anomalies, new_anomalies)
    """
    # Aggregate data across all hosts
    hosts_data = []
    total_all_issues = 0
    total_all_new = 0
    
    for hostname in sorted(hosts_results.keys()):
        anomalies, new_anomalies = hosts_results[hostname]
        
        # Categorize issues for this host
        critical_cats = ["zombie", "low_disk_space_datastore", "low_disk_space_partition"]
        critical_items = []
        warning_items = []
        new_items = []
        
        for key, items in anomalies.items():
            if key in critical_cats and items:
                critical_items.append(f"{key.replace('_', ' ').title()}: {len(items)}")
            elif key != "old_report_warning" and items:
                warning_items.append(f"{key.replace('_', ' ').title()}: {len(items)}")
        
        for key, items in new_anomalies.items():
            if items and key != "old_report_warning":
                new_items.append(f"{key.replace('_', ' ').title()}: {len(items)} new")
        
        total_host_issues = sum(len(v) for v in anomalies.values())
        total_host_new = sum(len(v) for v in new_anomalies.values())
        
        total_all_issues += total_host_issues
        total_all_new += total_host_new
        
        hosts_data.append({
            "hostname": hostname,
            "total_issues": total_host_issues,
            "total_new": total_host_new,
            "critical_items": critical_items,
            "warning_items": warning_items,
            "new_items": new_items,
        })
    
    # Prepare template data
    template_data = {
        "total_issues": total_all_issues,
        "total_new": total_all_new,
        "hosts": hosts_data,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    
    # Setup Jinja2 environment
    template_dir = Path(__file__).parent / "email_utils"
    env = Environment(loader=FileSystemLoader(str(template_dir)))
    
    try:
        template = env.get_template("email_template.j2")
        html_content = template.render(**template_data)
        
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(html_content)
    except Exception as e:
        print(f"Error rendering HTML template: {e}")
        _export_html_multi_host_fallback(output_file, hosts_results)


def _export_html_multi_host_fallback(output_file, hosts_results):
    """Fallback HTML generation for multi-host reports."""
    html = """<!DOCTYPE html>
<html>
<head>
  <meta charset='UTF-8'>
  <title>RVTools Multi-Host Report</title>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif; color: #333; line-height: 1.6; margin: 24px; background: #f5f5f5; }
    .container { max-width: 900px; margin: 0 auto; background: white; padding: 32px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
    h1 { color: #4158D9; border-bottom: 2px solid #5079F2; padding-bottom: 12px; }
    h2 { color: #4158D9; margin-top: 24px; font-size: 18px; }
    .host-section { margin: 24px 0; padding: 16px; background: #f9f9f9; border-left: 4px solid #4158D9; }
    table { border-collapse: collapse; width: 100%; margin: 16px 0; }
    th { background: #F2F2F2; border: 1px solid #ddd; padding: 12px; text-align: left; font-weight: 600; color: #4158D9; }
    td { border: 1px solid #ddd; padding: 10px; }
    tr:nth-child(even) { background: #fafafa; }
  </style>
</head>
<body>
  <div class="container">
    <h1>📊 RVTools Multi-Host Analysis Report</h1>
    <p><strong>Generated:</strong> """ + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + """</p>
"""
    
    total_all_issues = sum(
        sum(len(v) for v in anomalies.values())
        for anomalies, _ in hosts_results.values()
    )
    html += f"    <p><strong>Total Issues Across All Hosts:</strong> {total_all_issues}</p>\n"
    
    for hostname in sorted(hosts_results.keys()):
        anomalies, new_anomalies = hosts_results[hostname]
        total_host = sum(len(v) for v in anomalies.values())
        total_new = sum(len(v) for v in new_anomalies.values())
        
        html += f"""    <div class="host-section">
      <h2>🖥️ {hostname}</h2>
      <p>Total Issues: <strong>{total_host}</strong> | New Issues: <strong>{total_new}</strong></p>
      <table>
        <thead><tr><th>Category</th><th>Count</th></tr></thead>
        <tbody>
"""
        
        for issue_type, items in anomalies.items():
            new_count = len(new_anomalies.get(issue_type, []))
            badge = f" <span style='background: #374ACB; color: white; padding: 2px 6px; border-radius: 3px; font-size: 11px;'>+{new_count} new</span>" if new_count > 0 else ""
            html += f"        <tr><td>{issue_type.upper()}</td><td>{len(items)}{badge}</td></tr>\n"
        
        html += "        </tbody></table>\n"
        html += "      </div>\n"
    
    html += "  </div>\n</body>\n</html>"
    
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(html)


def export_markdown_multi_host(output_file, hosts_results):
    """Generate a multi-host Markdown report."""
    md = "# RVTools Multi-Host Analysis Report\n\n"
    md += f"**Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
    
    total_all_issues = sum(
        sum(len(v) for v in anomalies.values())
        for anomalies, _ in hosts_results.values()
    )
    
    md += f"**Total Issues Across All Hosts:** {total_all_issues}\n\n"
    md += "## Host Summary\n\n"
    
    for hostname in sorted(hosts_results.keys()):
        anomalies, new_anomalies = hosts_results[hostname]
        total_host = sum(len(v) for v in anomalies.values())
        total_new = sum(len(v) for v in new_anomalies.values())
        md += f"- **{hostname}**: {total_host} total issues, {total_new} new\n"
    
    md += "\n---\n\n"
    
    # Detailed per-host sections
    for hostname in sorted(hosts_results.keys()):
        anomalies, new_anomalies = hosts_results[hostname]
        md += f"## {hostname}\n\n"
        
        total_host = sum(len(v) for v in anomalies.values())
        total_new = sum(len(v) for v in new_anomalies.values())
        md += f"- **Total Issues:** {total_host}\n"
        md += f"- **New This Run:** {total_new}\n\n"
        
        for issue_type, items in anomalies.items():
            if not items:
                continue
            
            md += f"### {issue_type.upper()}\n\n"
            new_count = len(new_anomalies.get(issue_type, []))
            if new_count > 0:
                md += f"⚠️ **{new_count} NEW ISSUES** | {len(items)} total\n\n"
            else:
                md += f"{len(items)} issues\n\n"
            
            md += "| Issue Details |\n"
            md += "|---|\n"
            
            for item in items:
                if isinstance(item, dict):
                    item_str = " | ".join(str(v) for v in item.values())
                else:
                    item_str = str(item)
                md += f"| {item_str} |\n"
            
            md += "\n"
        
        md += "---\n\n"
    
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(md)



    """
    Process a single RVTools file and detect anomalies.
    
    Args:
        current_file: Path to the xlsx file
        host_name: Hostname identifier
        state_dir: Directory for storing per-host state files
        
    Returns:
        tuple: (anomalies_dict, new_anomalies_dict)
    """
    host_anomalies = {}
    host_prev_state = {}
    
    # Check all sheets for anomalies
    check_vhealth(current_file, load_sheet, host_anomalies)
    check_vpartition(current_file, load_sheet, host_anomalies)
    check_vdatastore(current_file, load_sheet, host_anomalies)
    check_report_age(current_file, host_anomalies)
    
    # Load host-specific state
    os.makedirs(state_dir, exist_ok=True)
    host_state_file = os.path.join(state_dir, f"{host_name}_state.json")
    if os.path.exists(host_state_file):
        with open(host_state_file, "r") as f:
            host_prev_state = json.load(f)
    
    # Compare with previous state
    host_new_anomalies = compare_previous(host_prev_state, host_anomalies)
    
    # Save current state
    with open(host_state_file, "w") as f:
        json.dump(host_anomalies, f, indent=2, default=str)
    
    return host_anomalies, host_new_anomalies


def process_single_host(current_file, host_name, state_dir):
    """
    Process a single RVTools file and detect anomalies.
    
    Args:
        current_file: Path to the xlsx file
        host_name: Hostname identifier
        state_dir: Directory for storing per-host state files
        
    Returns:
        tuple: (anomalies_dict, new_anomalies_dict)
    """
    host_anomalies = {}
    host_prev_state = {}
    
    # Check all sheets for anomalies
    check_vhealth(current_file, load_sheet, host_anomalies)
    check_vpartition(current_file, load_sheet, host_anomalies)
    check_vdatastore(current_file, load_sheet, host_anomalies)
    check_report_age(current_file, host_anomalies)
    
    # Load host-specific state
    os.makedirs(state_dir, exist_ok=True)
    host_state_file = os.path.join(state_dir, f"{host_name}_state.json")
    if os.path.exists(host_state_file):
        with open(host_state_file, "r") as f:
            host_prev_state = json.load(f)
    
    # Compare with previous state
    host_new_anomalies = compare_previous(host_prev_state, host_anomalies)
    
    # Save current state
    with open(host_state_file, "w") as f:
        json.dump(host_anomalies, f, indent=2, default=str)
    
    return host_anomalies, host_new_anomalies


def process_multi_host(root_path, load_sheet_func):
    """
    Process multiple RVTools files organized by hostname subdirectories.
    
    Args:
        root_path: Root directory containing hostname subdirectories
        load_sheet_func: Function to load Excel sheets
        
    Returns:
        dict: Mapping of hostname -> (anomalies, new_anomalies)
    """
    hosts_data = resolve_recursive_rvtools_files(root_path)
    state_dir = os.path.join(OUTPUT_DIR, "host_states")
    
    results = {}
    for hostname, xlsx_file in sorted(hosts_data.items()):
        print(f"\nProcessing host: {hostname}")
        print(f"  File: {xlsx_file}")
        
        try:
            anomalies, new_anomalies = process_single_host(xlsx_file, hostname, state_dir)
            results[hostname] = (anomalies, new_anomalies)
            total = sum(len(v) for v in anomalies.values())
            print(f"  Anomalies found: {total}")
        except Exception as e:
            print(f"  Error processing {hostname}: {e}")
            continue
    
    return results


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
        default=os.path.join(OUTPUT_DIR, f"current_issues_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.md"),
        help="Ruta del reporte Markdown generado.",
    )

    parser.add_argument(
        "--html",
        default=os.path.join(OUTPUT_DIR, f"current_issues_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.html"),
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
        default=os.path.join(OUTPUT_DIR, f"rvtools_report_full_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"),
        help="Ruta del reporte XLSX completo con todas las anomalías.",
    )

    parser.add_argument(
        "--auto-folder",
        action="store_true",
        help="Si se pasa un directorio, usa el último archivo RVTools disponible.",
    )

    parser.add_argument(
        "--generate-email",
        action="store_true",
        help="Generate email template and metadata for ansible-email dispatcher",
    )

    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Search recursively for xlsx files by hostname subdirectories. Generates per-host reports.",
    )

    args = parser.parse_args()

    if len(sys.argv) == 1:
        parser.print_help()
        sys.exit(0)

    more_info = args.more_info
    output_file = args.output
    previous_state_file = args.state

    # Handle recursive multi-host mode
    if args.recursive:
        if not os.path.isdir(args.file):
            parser.error("--recursive requires a directory path")
        
        hosts_results = process_multi_host(args.file, load_sheet)
        
        if not hosts_results:
            print("No hosts found with xlsx files")
            sys.exit(1)
        
        print(f"\n✓ Processed {len(hosts_results)} hosts")
        
        # Export multi-host HTML report
        report_format = args.report_format.lower()
        if report_format == "html":
            export_html_multi_host(args.html, hosts_results)
            print(f"Reporte HTML multi-host: {args.html}")
        else:
            export_markdown_multi_host(args.markdown, hosts_results)
            print(f"Reporte Markdown multi-host: {args.markdown}")
        
        # Summary
        print("\n===== RESUMEN DE ANOMALÍAS POR HOST =====\n")
        for hostname, (host_anomalies, host_new_anomalies) in sorted(hosts_results.items()):
            total = sum(len(v) for v in host_anomalies.values())
            new_total = sum(len(v) for v in host_new_anomalies.values())
            print(f"{hostname:20} : {total} total, {new_total} new")
        
        sys.exit(0)
    
    # Single-host mode (original logic)
    if os.path.isdir(args.file):
        if args.auto_folder:
            CURRENT_FILE = resolve_latest_rvtools_file(args.file)
        else:
            parser.error(
                "El argumento 'file' es un directorio. Use '--auto-folder' para seleccionar el último archivo dentro del directorio."
            )
    else:
        CURRENT_FILE = args.file

    check_vhealth(CURRENT_FILE, load_sheet, anomalies)
    check_vpartition(CURRENT_FILE, load_sheet, anomalies)
    check_vdatastore(CURRENT_FILE, load_sheet, anomalies)

    # Check report age
    check_report_age(CURRENT_FILE, anomalies)

    # Load previous state from argument if it exists
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
        export_html(args.html, anomalies, new_anomalies)
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

    # Generate email if requested
    if args.generate_email:
        print("\nGenerating email...")
        if generate_email(anomalies, new_anomalies):
            print("✅ Email preparation complete")
        else:
            print("❌ Email generation failed")


if __name__ == "__main__":
    # Run main function
    main()
