"""
RVTools Infrastructure Analysis Tool.

Analyzes RVTools reports (vHealth and vPartition) to detect anomalies
in virtualization infrastructure. Compares current state with previous state
to identify new issues and generates reports in Excel and HTML.
"""

import argparse
import pandas as pd
import json
import os
import sys
import re
from datetime import datetime
from pathlib import Path
from jinja2 import Environment, FileSystemLoader
from email_utils.email_handler import EmailHandler
from email_utils.config.email_config import ATTACHMENTS_TO_INCLUDE

# Constants
OUTPUT_DIR = "rvtools_reports"
PREVIOUS_STATE_FILE = os.path.join(OUTPUT_DIR, "rvtools_previous_state.json")

class RVToolsAnalyzer:
    def __init__(self, output_dir=OUTPUT_DIR):
        self.output_dir = output_dir
        os.makedirs(self.output_dir, exist_ok=True)
        self.anomalies = {}
        self.new_anomalies = {}
        self.prev_state = {}

    def load_sheet(self, file, sheet):
        """Load and return a specific sheet from an Excel file."""
        try:
            return pd.read_excel(file, sheet_name=sheet)
        except Exception:
            return None

    def resolve_latest_file(self, path):
        """Resolve input file if path is a directory."""
        if not os.path.isdir(path):
            return path

        entries = [
            os.path.join(path, f)
            for f in os.listdir(path)
            if f.lower().endswith(".xlsx") and os.path.isfile(os.path.join(path, f))
        ]

        if not entries:
            raise FileNotFoundError(f"No .xlsx files found in directory {path}")

        pattern1 = re.compile(r"^rvtools_.+?_(?P<year>\d{4})-(?P<mon>\d{2})-(?P<day>\d{2})_(?P<hour>\d{2})\.(?P<min>\d{2})\.xlsx$", re.IGNORECASE)
        pattern2 = re.compile(r"^rvtools_export_(?P<year>\d{4})(?P<mon>\d{2})(?P<day>\d{2})_(?P<hour>\d{2})(?P<min>\d{2})(?P<sec>\d{2})\.xlsx$", re.IGNORECASE)

        matched = []
        for p in entries:
            name = os.path.basename(p)
            for pattern in [pattern1, pattern2]:
                m = pattern.match(name)
                if m:
                    try:
                        dt_parts = {k: int(v) for k, v in m.groupdict().items()}
                        ts = datetime(dt_parts['year'], dt_parts['mon'], dt_parts['day'], 
                                     dt_parts['hour'], dt_parts['min'], dt_parts.get('sec', 0))
                        matched.append((ts, p))
                        break
                    except ValueError: continue

        if matched:
            return sorted(matched, key=lambda x: x[0], reverse=True)[0][1]

        entries.sort(key=lambda f: os.path.getmtime(f), reverse=True)
        return entries[0]

    def get_col(self, *names, sheet=None):
        for n in names:
            if n in sheet.columns:
                return sheet[n].astype(str)
        return None

    def check_vhealth(self, file):
        vhealth = self.load_sheet(file, "vHealth")
        if vhealth is None or vhealth.empty: return

        vhealth.columns = vhealth.columns.str.strip().str.lower().str.replace(" ", "_")
        ignore_types = {"FOLDERNAME", "PERFORMANCE TIP", "SECURITY", "STORAGE", "CPU"}

        # vHealth generic messages
        if "message_type" in vhealth.columns:
            msg_col = "message" if "message" in vhealth.columns else ""
            problem_rows = vhealth[vhealth["message_type"].notna()]
            
            for msg_type, group in problem_rows.groupby("message_type"):
                if str(msg_type).strip().upper() in ignore_types: continue
                key = str(msg_type).lower().replace(" ", "_")
                cols = ["name", "message_type"]
                if msg_col: cols.insert(1, "message")
                self.anomalies[key] = group[[c for c in cols if c in group.columns]].to_dict("records")

        # Specific checks
        checks = {
            "snapshot_present": ("snapshot", "present", ["name", "message", "snapshot", "age"]),
            "cdrom_connected": ("cdrom", "connected", ["name", "message", "cdrom"]),
            "usb_connected": ("usb", "connected", ["name", "message", "usb"]),
            "vmtools_issue": ("tools", "old|not installed", ["name", "message", "tools"]),
        }

        for key, (col, pattern, fields) in checks.items():
            if col in vhealth.columns:
                matches = vhealth[vhealth[col].astype(str).str.contains(pattern, case=False, na=False)]
                if not matches.empty:
                    self.anomalies[key] = matches[[f for f in fields if f in matches.columns]].to_dict("records")

        if "zombie" in vhealth.columns:
            zombies = vhealth[vhealth["zombie"].astype(str) != "0"]
            if not zombies.empty:
                self.anomalies["zombies"] = zombies[[c for c in ["name", "message", "zombie"] if c in zombies.columns]].to_dict("records")

    def check_vpartition(self, file):
        vpartition = self.load_sheet(file, "vPartition")
        if vpartition is None or vpartition.empty: return
        vpartition.columns = vpartition.columns.str.strip().str.lower().str.replace(" ", "_")

        free_col = next((c for c in ["free_%", "free_percent", "free"] if c in vpartition.columns), None)
        if not free_col: return

        vpartition[free_col] = pd.to_numeric(vpartition[free_col], errors="coerce")
        low = vpartition[vpartition[free_col] < 10]
        if not low.empty:
            cols = [c for c in ["vm", "disk", free_col, "annotation"] if c in low.columns]
            self.anomalies["low_disk_space_partition"] = low[cols].to_dict("records")

    def check_vdatastore(self, file):
        vds = self.load_sheet(file, "vDatastore")
        if vds is None or vds.empty: return
        vds.columns = vds.columns.str.strip().str.lower().str.replace(" ", "_")

        free_col = next((c for c in ["free_%", "free_percent", "free"] if c in vds.columns), None)
        if not free_col: return

        vds[free_col] = pd.to_numeric(vds[free_col], errors="coerce")
        low = vds[vds[free_col] < 10]
        if not low.empty:
            cols = [c for c in ["name", "hosts", free_col, "address"] if c in low.columns]
            self.anomalies["low_disk_space_datastore"] = low[cols].to_dict("records")

    def check_report_age(self, file, max_days=60):
        if not os.path.exists(file): return
        age_days = (datetime.now() - datetime.fromtimestamp(os.path.getmtime(file))).days
        if age_days > max_days:
            self.anomalies["old_report_warning"] = [{
                "file": os.path.basename(file), "age_days": age_days,
                "warning": f"Report is {age_days} days old"
            }]

    def get_item_identity(self, item, category):
        """Stable identity for comparison, ignoring dynamic values."""
        if category == "low_disk_space_partition":
            return f"{item.get('vm')}:{item.get('disk')}"
        if category == "low_disk_space_datastore":
            return f"{item.get('name')}"
        if category == "snapshot_present":
            return f"{item.get('name')}"
        
        # Default identifiers
        for field in ["name", "vm", "id"]:
            if field in item: return str(item[field])
            
        return json.dumps(item, sort_keys=True)

    def compare_previous(self, prev_state):
        self.prev_state = prev_state
        self.new_anomalies = {}
        
        for cat, items in self.anomalies.items():
            if cat == "old_report_warning": continue
            prev_items = self.prev_state.get(cat, [])
            prev_ids = {self.get_item_identity(x, cat) for x in prev_items}
            new_items = [x for x in items if self.get_item_identity(x, cat) not in prev_ids]
            if new_items:
                self.new_anomalies[cat] = new_items
        return self.new_anomalies

    def export_xls(self, filename, data):
        writer = pd.ExcelWriter(filename, engine="openpyxl")
        if not data:
            pd.DataFrame({"info": ["No issues"]}).to_excel(writer, sheet_name="summary", index=False)
        else:
            for k, v in data.items():
                pd.DataFrame(v).to_excel(writer, sheet_name=k[:31], index=False)
        writer.close()

    def prepare_template_data(self):
        critical_cats = ["zombies", "low_disk_space_datastore", "low_disk_space_partition"]
        
        critical_summary = []
        warning_summary = []
        new_summary = []
        
        for cat, items in self.anomalies.items():
            if cat == "old_report_warning": continue
            label = cat.replace('_', ' ').title()
            if cat in critical_cats:
                critical_summary.append({"label": label, "count": len(items), "issue_list": items[:5]})
            else:
                warning_summary.append({"label": label, "count": len(items), "issue_list": items[:5]})

        for cat, items in self.new_anomalies.items():
            new_summary.append({"label": cat.replace('_', ' ').title(), "count": len(items), "issue_list": items})

        return {
            "total_issues": sum(len(v) for v in self.anomalies.values()),
            "critical_summary": critical_summary,
            "warning_summary": warning_summary,
            "new_summary": new_summary,
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }

    def export_html(self, output_file):
        template_dir = Path(__file__).parent / "email_utils"
        env = Environment(loader=FileSystemLoader(str(template_dir)))
        try:
            template = env.get_template("email_template.j2")
            html_content = template.render(**self.prepare_template_data())
            with open(output_file, "w", encoding="utf-8") as f:
                f.write(html_content)
        except Exception as e:
            print(f"Error rendering HTML: {e}")

    def generate_email(self, args):
        path = args.email_dir
        handler = EmailHandler("rvtools-analyzer", output_base=path)
        return handler.generate(
            template_name="email_template.j2",
            template_data=self.prepare_template_data(),
            attachments=ATTACHMENTS_TO_INCLUDE,
        )

def main():
    parser = argparse.ArgumentParser(description="Analiza reportes RVTools")
    parser.add_argument("file", nargs="?", default="RVTools_export.xlsx")
    parser.add_argument("--state", default=PREVIOUS_STATE_FILE)
    parser.add_argument("--overwrite", action="store_true")
    parser.add_argument("--generate-email", action="store_true")
    parser.add_argument("--email-dir")
    args = parser.parse_args()

    analyzer = RVToolsAnalyzer()
    
    try:
        current_file = analyzer.resolve_latest_file(args.file)
    except Exception as e:
        print(f"Error: {e}")
        return

    analyzer.check_vhealth(current_file)
    analyzer.check_vpartition(current_file)
    analyzer.check_vdatastore(current_file)
    analyzer.check_report_age(current_file)

    prev_state = {}
    if os.path.exists(args.state):
        with open(args.state, "r") as f:
            prev_state = json.load(f)

    analyzer.compare_previous(prev_state)
    
    # Exports
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    html_report = os.path.join(OUTPUT_DIR, f"current_issues_{timestamp}.html")
    full_xlsx = os.path.join(OUTPUT_DIR, f"rvtools_report_{timestamp}.xlsx")
    new_xlsx = os.path.join(OUTPUT_DIR, f"new_anomalies_{timestamp}.xlsx")
    
    analyzer.export_xls(new_xlsx, analyzer.new_anomalies)
    analyzer.export_xls(full_xlsx, analyzer.anomalies)
    analyzer.export_html(html_report)

    if not os.path.exists(args.state) or args.overwrite:
        with open(args.state, "w") as f:
            json.dump(analyzer.anomalies, f, indent=2)

    if args.generate_email:
        analyzer.generate_email(args)

    print(f"Análisis completado. Total: {sum(len(v) for v in analyzer.anomalies.values())}, Nuevos: {sum(len(v) for v in analyzer.new_anomalies.values())}")

if __name__ == "__main__":
    main()
