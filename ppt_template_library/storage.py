from __future__ import annotations

import json
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Iterable

import pandas as pd

from .catalog import build_source_catalog, source_dataframe
from .models import SourceEntry
from .paths import ensure_library_structure, library_paths


DB_NAME = "ppt_template_library.sqlite"


def db_path(root: str | Path = "ppt_template_library") -> Path:
    paths = library_paths(root)
    return paths.database / DB_NAME


def _connect(root: str | Path = "ppt_template_library") -> sqlite3.Connection:
    paths = ensure_library_structure(root)
    db_file = paths.database / DB_NAME
    conn = sqlite3.connect(db_file)
    conn.row_factory = sqlite3.Row
    return conn


def initialize_storage(root: str | Path = "ppt_template_library") -> Path:
    paths = ensure_library_structure(root)
    conn = _connect(root)
    with conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS source_catalog (
                source_id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                source_url TEXT NOT NULL,
                source_website TEXT NOT NULL,
                source_type TEXT NOT NULL,
                license_hint TEXT NOT NULL,
                commercial_use_hint TEXT NOT NULL,
                modification_allowed_hint TEXT NOT NULL,
                download_method TEXT NOT NULL,
                programmatic_access TEXT NOT NULL,
                quality_hint INTEGER NOT NULL,
                score INTEGER NOT NULL,
                language TEXT NOT NULL,
                tags_json TEXT NOT NULL,
                notes TEXT NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS template_catalog (
                template_id TEXT PRIMARY KEY,
                template_name TEXT NOT NULL,
                source_url TEXT NOT NULL,
                source_website TEXT NOT NULL,
                author TEXT,
                license TEXT,
                commercial_use TEXT,
                modification_allowed TEXT,
                download_date TEXT,
                file_format TEXT,
                file_size INTEGER,
                slide_count INTEGER,
                aspect_ratio TEXT,
                language TEXT,
                industry TEXT,
                scenario TEXT,
                style TEXT,
                primary_color TEXT,
                secondary_color TEXT,
                dark_or_light TEXT,
                contains_chart INTEGER,
                contains_table INTEGER,
                contains_timeline INTEGER,
                contains_process INTEGER,
                contains_map INTEGER,
                contains_infographic INTEGER,
                contains_images INTEGER,
                contains_icons INTEGER,
                contains_animations INTEGER,
                quality_score INTEGER,
                design_score INTEGER,
                usability_score INTEGER,
                duplicate_hash TEXT,
                preview_image TEXT,
                local_path TEXT,
                metadata_json TEXT,
                status TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS tag_catalog (
                entity_type TEXT NOT NULL,
                entity_id TEXT NOT NULL,
                tag_type TEXT NOT NULL,
                tag_value TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_source_type_score
            ON source_catalog(source_type, score DESC)
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_template_score
            ON template_catalog(quality_score DESC)
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_tag_entity
            ON tag_catalog(entity_type, entity_id)
            """
        )
    conn.close()
    return paths.root


def save_source_catalog(entries: Iterable[SourceEntry], root: str | Path = "ppt_template_library") -> Path:
    initialize_storage(root)
    conn = _connect(root)
    now = datetime.utcnow().isoformat(timespec="seconds")
    rows = []
    for entry in entries:
        rows.append(
            (
                entry.source_id,
                entry.name,
                entry.source_url,
                entry.source_website,
                entry.source_type,
                entry.license_hint,
                entry.commercial_use_hint,
                entry.modification_allowed_hint,
                entry.download_method,
                entry.programmatic_access,
                int(entry.quality_hint),
                int(entry.score),
                entry.language,
                json.dumps(list(entry.tags), ensure_ascii=False),
                entry.notes,
                now,
                now,
            )
        )
    with conn:
        conn.executemany(
            """
            INSERT INTO source_catalog (
                source_id, name, source_url, source_website, source_type, license_hint,
                commercial_use_hint, modification_allowed_hint, download_method, programmatic_access,
                quality_hint, score, language, tags_json, notes, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(source_id) DO UPDATE SET
                name=excluded.name,
                source_url=excluded.source_url,
                source_website=excluded.source_website,
                source_type=excluded.source_type,
                license_hint=excluded.license_hint,
                commercial_use_hint=excluded.commercial_use_hint,
                modification_allowed_hint=excluded.modification_allowed_hint,
                download_method=excluded.download_method,
                programmatic_access=excluded.programmatic_access,
                quality_hint=excluded.quality_hint,
                score=excluded.score,
                language=excluded.language,
                tags_json=excluded.tags_json,
                notes=excluded.notes,
                updated_at=excluded.updated_at
            """,
            rows,
        )
    conn.close()
    return db_path(root)


def load_source_catalog(root: str | Path = "ppt_template_library") -> pd.DataFrame:
    initialize_storage(root)
    conn = _connect(root)
    df = pd.read_sql_query("SELECT * FROM source_catalog ORDER BY score DESC, name ASC", conn)
    conn.close()
    return df


def save_template_catalog(records: Iterable[dict], root: str | Path = "ppt_template_library") -> None:
    initialize_storage(root)
    conn = _connect(root)
    now = datetime.utcnow().isoformat(timespec="seconds")
    rows = []
    for record in records:
        rows.append(
            (
                record.get("template_id", ""),
                record.get("template_name", ""),
                record.get("source_url", ""),
                record.get("source_website", ""),
                record.get("author", ""),
                record.get("license", ""),
                record.get("commercial_use", ""),
                record.get("modification_allowed", ""),
                record.get("download_date", ""),
                record.get("file_format", ""),
                record.get("file_size"),
                record.get("slide_count"),
                record.get("aspect_ratio", ""),
                record.get("language", ""),
                record.get("industry", ""),
                record.get("scenario", ""),
                record.get("style", ""),
                record.get("primary_color", ""),
                record.get("secondary_color", ""),
                record.get("dark_or_light", ""),
                int(bool(record.get("contains_chart"))),
                int(bool(record.get("contains_table"))),
                int(bool(record.get("contains_timeline"))),
                int(bool(record.get("contains_process"))),
                int(bool(record.get("contains_map"))),
                int(bool(record.get("contains_infographic"))),
                int(bool(record.get("contains_images"))),
                int(bool(record.get("contains_icons"))),
                int(bool(record.get("contains_animations"))),
                int(record.get("quality_score", 0)),
                int(record.get("design_score", 0)),
                int(record.get("usability_score", 0)),
                record.get("duplicate_hash", ""),
                record.get("preview_image", ""),
                record.get("local_path", ""),
                json.dumps(record.get("metadata_json", {}), ensure_ascii=False),
                record.get("status", "new"),
                now,
                now,
            )
        )
    with conn:
        conn.executemany(
            """
            INSERT INTO template_catalog (
                template_id, template_name, source_url, source_website, author, license,
                commercial_use, modification_allowed, download_date, file_format, file_size,
                slide_count, aspect_ratio, language, industry, scenario, style,
                primary_color, secondary_color, dark_or_light, contains_chart, contains_table,
                contains_timeline, contains_process, contains_map, contains_infographic,
                contains_images, contains_icons, contains_animations, quality_score,
                design_score, usability_score, duplicate_hash, preview_image, local_path,
                metadata_json, status, created_at, updated_at
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(template_id) DO UPDATE SET
                template_name=excluded.template_name,
                source_url=excluded.source_url,
                source_website=excluded.source_website,
                author=excluded.author,
                license=excluded.license,
                commercial_use=excluded.commercial_use,
                modification_allowed=excluded.modification_allowed,
                download_date=excluded.download_date,
                file_format=excluded.file_format,
                file_size=excluded.file_size,
                slide_count=excluded.slide_count,
                aspect_ratio=excluded.aspect_ratio,
                language=excluded.language,
                industry=excluded.industry,
                scenario=excluded.scenario,
                style=excluded.style,
                primary_color=excluded.primary_color,
                secondary_color=excluded.secondary_color,
                dark_or_light=excluded.dark_or_light,
                contains_chart=excluded.contains_chart,
                contains_table=excluded.contains_table,
                contains_timeline=excluded.contains_timeline,
                contains_process=excluded.contains_process,
                contains_map=excluded.contains_map,
                contains_infographic=excluded.contains_infographic,
                contains_images=excluded.contains_images,
                contains_icons=excluded.contains_icons,
                contains_animations=excluded.contains_animations,
                quality_score=excluded.quality_score,
                design_score=excluded.design_score,
                usability_score=excluded.usability_score,
                duplicate_hash=excluded.duplicate_hash,
                preview_image=excluded.preview_image,
                local_path=excluded.local_path,
                metadata_json=excluded.metadata_json,
                status=excluded.status,
                updated_at=excluded.updated_at
            """,
            rows,
        )
    conn.close()


def load_template_catalog(root: str | Path = "ppt_template_library") -> pd.DataFrame:
    initialize_storage(root)
    conn = _connect(root)
    df = pd.read_sql_query("SELECT * FROM template_catalog ORDER BY quality_score DESC, template_name ASC", conn)
    conn.close()
    return df


def load_catalog_summary(root: str | Path = "ppt_template_library") -> dict[str, int]:
    initialize_storage(root)
    conn = _connect(root)
    summary = {
        "sources": conn.execute("SELECT COUNT(*) FROM source_catalog").fetchone()[0],
        "templates": conn.execute("SELECT COUNT(*) FROM template_catalog").fetchone()[0],
        "premium": conn.execute("SELECT COUNT(*) FROM template_catalog WHERE quality_score >= 90").fetchone()[0],
        "commercial_allowed": conn.execute("SELECT COUNT(*) FROM source_catalog WHERE commercial_use_hint = 'allowed'").fetchone()[0],
    }
    conn.close()
    return summary

