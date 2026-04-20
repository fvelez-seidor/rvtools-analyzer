#!/usr/bin/env python
"""Quick test to demonstrate the new HTML template styling."""

from pathlib import Path
from jinja2 import Environment, FileSystemLoader
from datetime import datetime

# Setup Jinja2 environment
template_dir = Path(__file__).parent / "email_utils"
env = Environment(loader=FileSystemLoader(str(template_dir)))

# Sample data
template_data = {
    "total_issues": 12,
    "critical_items": [
        "Zombie Processes: 2",
        "Low Disk Space (Datastore): 3",
        "Low Disk Space (Partition): 1",
    ],
    "warning_items": ["CDRom Connected: 4", "VMtools Issue: 2"],
    "new_items": ["Snapshot Present: 1 new", "USB Connected: 2 new"],
    "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
}

# Render template
template = env.get_template("email_template.j2")
html_content = template.render(**template_data)

# Save to test output
output_file = Path(__file__).parent / "test_report.html"
with open(output_file, "w", encoding="utf-8") as f:
    f.write(html_content)

print(f"✅ Test HTML generated: {output_file}")
print(f"\nTemplate rendering successful!")
print(f"File size: {output_file.stat().st_size} bytes")
