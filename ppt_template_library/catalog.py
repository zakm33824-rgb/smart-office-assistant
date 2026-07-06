from __future__ import annotations

import json
from dataclasses import asdict
from pathlib import Path
from typing import Any

import pandas as pd

from .models import SearchFilters, SourceEntry
from .paths import ensure_library_structure
from .seeds import build_seed_sources


def source_dataframe() -> pd.DataFrame:
    return pd.DataFrame([entry.to_row() for entry in build_seed_sources()])


def save_source_catalog_json(path: str | Path) -> Path:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    payload = [asdict(entry) for entry in build_seed_sources()]
    target.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return target


def build_source_catalog() -> list[SourceEntry]:
    return build_seed_sources()


def build_template_catalog() -> list[dict[str, str]]:
    from .seeds import build_template_catalog as _build

    return _build()


def filter_source_dataframe(df: pd.DataFrame, filters: SearchFilters) -> pd.DataFrame:
    result = df.copy()
    if filters.query:
        terms = [part for part in filters.query.lower().split() if part]
        if terms:
            mask = pd.Series(True, index=result.index)
            text_columns = ["name", "source_url", "source_website", "notes", "source_type", "license_hint"]
            for term in terms:
                term_mask = pd.Series(False, index=result.index)
                for column in text_columns:
                    term_mask |= result[column].astype(str).str.lower().str.contains(term, na=False)
                mask &= term_mask
            result = result[mask]
    if filters.industries:
        result = result[result["tags_json"].astype(str).str.contains("|".join(filters.industries), case=False, na=False)]
    if filters.scenarios:
        result = result[result["tags_json"].astype(str).str.contains("|".join(filters.scenarios), case=False, na=False)]
    if filters.styles:
        result = result[result["tags_json"].astype(str).str.contains("|".join(filters.styles), case=False, na=False)]
    if filters.colors:
        result = result[result["tags_json"].astype(str).str.contains("|".join(filters.colors), case=False, na=False)]
    if filters.languages:
        result = result[result["language"].astype(str).str.lower().isin([value.lower() for value in filters.languages])]
    if filters.source_types:
        result = result[result["source_type"].astype(str).str.lower().isin([value.lower() for value in filters.source_types])]
    if filters.licenses:
        result = result[result["license_hint"].astype(str).str.lower().isin([value.lower() for value in filters.licenses])]
    result = result[result["score"] >= filters.min_score]
    if filters.commercial_only:
        result = result[result["commercial_use_hint"].astype(str).str.lower().eq("allowed")]
    if filters.exclude_personal_only:
        result = result[~result["commercial_use_hint"].astype(str).str.lower().eq("restricted")]
    return result.sort_values(["score", "quality_hint", "name"], ascending=[False, False, True])


def build_library_summary() -> dict[str, Any]:
    df = source_dataframe()
    return {
        "source_count": int(len(df)),
        "repository_count": int((df["source_type"] == "repository").sum()),
        "site_count": int((df["source_type"].isin(["category_page", "collection", "engine_home"])).sum()),
        "topic_count": int((df["source_type"] == "topic_index").sum()),
        "explicit_license_count": int((df["license_hint"].isin(["MIT", "Apache-2.0", "CC0-1.0", "LGPL-2.1", "GPL-3.0", "CC-BY-SA-4.0", "CC-BY-SA", "CC-BY", "Public-Domain"])).sum()),
        "commercial_allowed_count": int((df["commercial_use_hint"] == "allowed").sum()),
        "uncertain_count": int((df["commercial_use_hint"] == "uncertain").sum()),
        "average_score": round(float(df["score"].mean()), 1) if len(df) else 0.0,
        "top_score": int(df["score"].max()) if len(df) else 0,
    }

