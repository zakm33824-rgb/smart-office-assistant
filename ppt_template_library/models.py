from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Any


TEMPLATE_METADATA_FIELDS: list[str] = [
    "template_id",
    "template_name",
    "source_url",
    "source_website",
    "author",
    "license",
    "commercial_use",
    "modification_allowed",
    "download_date",
    "file_format",
    "file_size",
    "slide_count",
    "aspect_ratio",
    "language",
    "industry",
    "scenario",
    "style",
    "primary_color",
    "secondary_color",
    "dark_or_light",
    "contains_chart",
    "contains_table",
    "contains_timeline",
    "contains_process",
    "contains_map",
    "contains_infographic",
    "contains_images",
    "contains_icons",
    "contains_animations",
    "quality_score",
    "design_score",
    "usability_score",
    "duplicate_hash",
    "preview_image",
    "local_path",
]


@dataclass(frozen=True)
class SourceEntry:
    source_id: str
    name: str
    source_url: str
    source_website: str
    source_type: str
    license_hint: str
    commercial_use_hint: str
    modification_allowed_hint: str
    download_method: str
    programmatic_access: str
    quality_hint: int
    language: str
    tags: tuple[str, ...] = field(default_factory=tuple)
    notes: str = ""
    score: int = 0

    def to_row(self) -> dict[str, Any]:
        return {
            "source_id": self.source_id,
            "name": self.name,
            "source_url": self.source_url,
            "source_website": self.source_website,
            "source_type": self.source_type,
            "license_hint": self.license_hint,
            "commercial_use_hint": self.commercial_use_hint,
            "modification_allowed_hint": self.modification_allowed_hint,
            "download_method": self.download_method,
            "programmatic_access": self.programmatic_access,
            "quality_hint": self.quality_hint,
            "language": self.language,
            "tags_json": list(self.tags),
            "notes": self.notes,
            "score": self.score,
        }


def default_template_record(template_id: str, template_name: str, source_url: str, source_website: str) -> dict[str, Any]:
    return {
        "template_id": template_id,
        "template_name": template_name,
        "source_url": source_url,
        "source_website": source_website,
        "author": "",
        "license": "unknown",
        "commercial_use": "uncertain",
        "modification_allowed": "uncertain",
        "download_date": date.today().isoformat(),
        "file_format": "",
        "file_size": None,
        "slide_count": None,
        "aspect_ratio": "",
        "language": "unknown",
        "industry": "",
        "scenario": "",
        "style": "",
        "primary_color": "",
        "secondary_color": "",
        "dark_or_light": "",
        "contains_chart": False,
        "contains_table": False,
        "contains_timeline": False,
        "contains_process": False,
        "contains_map": False,
        "contains_infographic": False,
        "contains_images": False,
        "contains_icons": False,
        "contains_animations": False,
        "quality_score": 0,
        "design_score": 0,
        "usability_score": 0,
        "duplicate_hash": "",
        "preview_image": "",
        "local_path": "",
        "metadata_json": {},
        "status": "new",
    }


@dataclass
class SearchFilters:
    query: str = ""
    industries: tuple[str, ...] = ()
    scenarios: tuple[str, ...] = ()
    styles: tuple[str, ...] = ()
    colors: tuple[str, ...] = ()
    languages: tuple[str, ...] = ()
    source_types: tuple[str, ...] = ()
    licenses: tuple[str, ...] = ()
    min_score: int = 0
    commercial_only: bool = False
    exclude_personal_only: bool = False

