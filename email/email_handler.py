"""
Universal email handler for preparing emails for ansible-email dispatcher.

Renders Jinja2 templates, copies attachments, and generates email metadata.

## Attachment Path Patterns

Supports three types of source paths:

1. **Literal paths** (static):
   "source": "path/to/file.json"

2. **Keyword patterns** (dynamic file selection):
   - {latest}  - newest by modification time
   - {oldest}  - oldest by modification time
   - {largest} - biggest file
   - {smallest} - smallest file

   Examples:
   "source": "reports/report_{latest}.xlsx"    → finds report_*.xlsx, uses newest
   "source": "backups/{oldest}.json"            → finds *.json, uses oldest
   "source": "logs/{largest}_output.log"        → finds *_output.log, uses largest

3. **Glob patterns** (wildcard):
   "source": "reports/*.xlsx"                   → any xlsx file in reports/

Pattern resolution order:
1. Try literal path → if exists, use it
2. Try keyword/glob pattern → find matches, select based on keyword/time
3. If not found: fail if optional=False, warn if optional=True
"""

import glob
import json
import re
import shutil
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Optional
from jinja2 import Environment, FileSystemLoader


class EmailHandler:
    """Handles email template rendering and metadata generation."""

    # Supported keywords for dynamic file selection
    KEYWORDS = {
        "latest": lambda files: max(files, key=lambda p: Path(p).stat().st_mtime),
        "oldest": lambda files: min(files, key=lambda p: Path(p).stat().st_mtime),
        "largest": lambda files: max(files, key=lambda p: Path(p).stat().st_size),
        "smallest": lambda files: min(files, key=lambda p: Path(p).stat().st_size),
    }

    def __init__(self, project_name: str, output_base: Optional[str] = None):
        """
        Initialize email handler.

        Args:
            project_name: Name of the project (used for output directory)
            output_base: Base directory for email output (default: ~/.email-templates)
        """
        if output_base is None:
            output_base = str(Path.home() / ".email-templates")

        self.project_name = project_name
        self.output_dir = Path(output_base) / project_name
        self.output_dir.mkdir(parents=True, exist_ok=True)

        # Setup Jinja2 environment
        template_dir = Path(__file__).parent
        self.env = Environment(loader=FileSystemLoader(template_dir))

    def generate(
        self,
        template_name: str,
        template_data: dict[str, Any],
        attachments: Optional[dict[str, dict[str, Any]]] = None,
    ) -> bool:
        """
        Generate email template, copy attachments, and create metadata.

        Args:
            template_name: Name of the Jinja2 template file (e.g., 'email_template.j2')
            template_data: Context data for template rendering
            attachments: Dict of attachment configurations
                {
                    "report": {
                        "source": "report_{latest}.json",
                        "optional": False
                    }
                }

        Returns:
            True if successful, False otherwise
        """
        try:
            # Render HTML template
            template = self.env.get_template(template_name)
            html_content = template.render(**template_data)

            # Save HTML
            html_file = self.output_dir / "email_template.html"
            html_file.write_text(html_content, encoding="utf-8")
            print(f"✓ Email template rendered: {html_file}")

            # Handle attachments
            attachment_paths = []
            if attachments:
                attachment_paths = self._copy_attachments(attachments)

            # Generate metadata
            metadata = {
                "template_type": self.project_name,
                "html_file": "email_template.html",
                "generated_at": datetime.now(timezone.utc).isoformat(),
                "should_send": True,
                "attachments": attachment_paths,
            }

            metadata_file = self.output_dir / "email_metadata.json"
            with metadata_file.open("w", encoding="utf-8") as f:
                json.dump(metadata, f, indent=2)

            print(f"✓ Email metadata generated: {metadata_file}")
            return True

        except Exception as e:
            print(f"✗ Error generating email: {e}")
            return False

    def _copy_attachments(self, attachments: dict[str, dict[str, Any]]) -> list[str]:
        """
        Copy attachment files to output directory.

        Supports literal paths, keyword patterns, and glob patterns.

        Args:
            attachments: Dict of attachments to copy

        Returns:
            List of relative paths to copied files
        """
        copied_files = []

        for name, config in attachments.items():
            source = config.get("source")
            optional = config.get("optional", False)

            if not source:
                continue

            # Resolve relative paths
            source_path = Path(source)
            if not source_path.is_absolute():
                source_path = Path.cwd() / source_path

            # Try literal path first
            if source_path.exists():
                self._copy_file(source_path, copied_files, optional, source)
                continue

            # Try keyword/glob pattern
            resolved_path = self._resolve_pattern(source_path, source)
            if resolved_path:
                self._copy_file(resolved_path, copied_files, optional, source)
            elif not optional:
                raise FileNotFoundError(f"Required attachment not found: {source}")
            else:
                print(f"⚠ Optional attachment not found: {source}")

        return copied_files

    def _resolve_pattern(self, source_path: Path, source_desc: str) -> Optional[Path]:
        """
        Resolve keyword or glob patterns in source path.

        Args:
            source_path: Path object with potential keywords
            source_desc: Original source description for error messages

        Returns:
            Path to resolved file, or None if not found
        """
        path_str = str(source_path)

        # Check for keywords {latest}, {oldest}, {largest}, {smallest}
        keyword_match = re.search(r"\{(latest|oldest|largest|smallest)\}", path_str)
        if keyword_match:
            keyword = keyword_match.group(1)
            # Replace keyword with * to create glob pattern
            glob_pattern = path_str.replace(f"{{{keyword}}}", "*")
            matches = glob.glob(glob_pattern)

            if matches:
                # Select file based on keyword
                selector = self.KEYWORDS[keyword]
                return Path(selector(matches))
            return None

        # Try as plain glob pattern
        matches = glob.glob(path_str)
        if matches:
            # If multiple matches, use latest by default
            return Path(max(matches, key=lambda p: Path(p).stat().st_mtime))

        return None

    def _copy_file(
        self, source_path: Path, copied_files: list, optional: bool, source_desc: str
    ) -> None:
        """Helper to copy a single file."""
        try:
            dest_path = self.output_dir / source_path.name
            shutil.copy2(source_path, dest_path)
            print(f"✓ Attachment copied: {dest_path.name}")
            copied_files.append(dest_path.name)
        except Exception as e:
            if optional:
                print(f"⚠ Failed to copy attachment {source_desc}: {e}")
            else:
                raise
