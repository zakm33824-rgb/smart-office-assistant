from __future__ import annotations

from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from ppt_template_library.manager import bootstrap_library


def main() -> None:
    manifest = bootstrap_library()
    print(f"Initialized template library")
    print(f"Sources: {manifest.get('source_count', 0)}")
    print(f"Template blueprints: {manifest.get('template_count', 0)}")
    print(f"Slides: {manifest.get('slide_count', 0)}")
    print(f"Components: {manifest.get('component_count', 0)}")
    print(f"Premium templates: {manifest.get('premium_templates', 0)}")
    print(f"Premium slides: {manifest.get('premium_slides', 0)}")
    print(f"Root: {PROJECT_ROOT / 'ppt_template_library'}")
    print(f"Metadata: {(PROJECT_ROOT / 'ppt_template_library' / 'metadata').resolve()}")
    print(f"Database: {(PROJECT_ROOT / 'ppt_template_library' / 'database' / 'ppt_template_library.sqlite').resolve()}")


if __name__ == "__main__":
    main()
