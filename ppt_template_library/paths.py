from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class LibraryPaths:
    root: Path
    templates: Path
    components: Path
    preview: Path
    metadata: Path
    database: Path
    commercial_allowed: Path
    personal_use_only: Path
    license_uncertain: Path
    premium_quality: Path
    industry: Path
    scenario: Path
    style: Path
    color: Path
    layout: Path
    logs: Path
    source_records: Path


def library_paths(root: str | Path = "ppt_template_library") -> LibraryPaths:
    root_path = Path(root)
    return LibraryPaths(
        root=root_path,
        templates=root_path / "templates",
        components=root_path / "components",
        preview=root_path / "preview",
        metadata=root_path / "metadata",
        database=root_path / "database",
        commercial_allowed=root_path / "commercial_allowed",
        personal_use_only=root_path / "personal_use_only",
        license_uncertain=root_path / "license_uncertain",
        premium_quality=root_path / "premium_quality",
        industry=root_path / "industry",
        scenario=root_path / "scenario",
        style=root_path / "style",
        color=root_path / "color",
        layout=root_path / "layout",
        logs=root_path / "logs",
        source_records=root_path / "source_records",
    )


def ensure_library_structure(root: str | Path = "ppt_template_library") -> LibraryPaths:
    paths = library_paths(root)
    for path in paths.__dict__.values():
        path.mkdir(parents=True, exist_ok=True)
    return paths

