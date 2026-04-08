"""Email configuration for rvtools-analyzer."""

# Define attachments to include in email
#
# Path patterns:
# - Literal:     "path/to/file.json"
# - Keywords:    "report_{latest}.xlsx"    → finds report_*.xlsx, uses newest
#                "backup_{oldest}.json"    → finds backup_*.json, uses oldest
#                "logs_{largest}.txt"      → finds logs_*.txt, uses biggest
#                "{smallest}_report.xlsx"  → finds *_report.xlsx, uses smallest
# - Glob:        "reports/*.xlsx"          → any xlsx file in reports/
#
# Keywords: {latest}, {oldest}, {largest}, {smallest}
# Just one keyword per attachment, placed anywhere in the filename

ATTACHMENTS_TO_INCLUDE = {
    "current_issues": {
        "source": "rvtools_reports/current_issues_{latest}.html",
        "optional": False,
    },
    "rvtools_report": {
        # Find latest rvtools_report_*.xlsx file in rvtools_reports/
        "source": "rvtools_reports/rvtools_report_{latest}.xlsx",
        "optional": False,
    },
    "previous_state": {
        "source": "/var/lib/semaphore/exports/rvtools_analyzer/rvtools_previous_state.json",
        "optional": False,
    },
}
