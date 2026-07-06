from __future__ import annotations

import json
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

import pandas as pd

from .page_models import ComponentRecord, PageRecord
from .paths import ensure_library_structure, library_paths

DB_NAME = 'ppt_template_library.sqlite'


def db_path(root: str | Path = 'ppt_template_library') -> Path:
    paths = library_paths(root)
    return paths.database / DB_NAME


def _connect(root: str | Path = 'ppt_template_library') -> sqlite3.Connection:
    paths = ensure_library_structure(root)
    conn = sqlite3.connect(paths.database / DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn


def initialize_slide_storage(root: str | Path = 'ppt_template_library') -> Path:
    paths = ensure_library_structure(root)
    conn = _connect(root)
    with conn:
        conn.execute(
            '''
            CREATE TABLE IF NOT EXISTS slide_catalog (
                slide_id TEXT PRIMARY KEY,
                source_template_id TEXT NOT NULL,
                source_file TEXT NOT NULL,
                source_url TEXT NOT NULL,
                slide_number INTEGER NOT NULL,
                slide_type TEXT NOT NULL,
                slide_subtype TEXT NOT NULL,
                industry TEXT NOT NULL,
                scenario TEXT NOT NULL,
                style TEXT NOT NULL,
                layout_type TEXT NOT NULL,
                primary_color TEXT NOT NULL,
                secondary_color TEXT NOT NULL,
                background_color TEXT NOT NULL,
                dark_or_light TEXT NOT NULL,
                aspect_ratio TEXT NOT NULL,
                text_density REAL NOT NULL,
                image_density REAL NOT NULL,
                chart_count INTEGER NOT NULL,
                table_count INTEGER NOT NULL,
                shape_count INTEGER NOT NULL,
                icon_count INTEGER NOT NULL,
                image_count INTEGER NOT NULL,
                text_box_count INTEGER NOT NULL,
                has_chart INTEGER NOT NULL,
                has_table INTEGER NOT NULL,
                has_timeline INTEGER NOT NULL,
                has_process INTEGER NOT NULL,
                has_map INTEGER NOT NULL,
                has_people INTEGER NOT NULL,
                has_infographic INTEGER NOT NULL,
                has_animation INTEGER NOT NULL,
                editable_level TEXT NOT NULL,
                design_score INTEGER NOT NULL,
                layout_score INTEGER NOT NULL,
                color_score INTEGER NOT NULL,
                usability_score INTEGER NOT NULL,
                modern_score INTEGER NOT NULL,
                overall_quality_score INTEGER NOT NULL,
                preview_image TEXT,
                thumbnail_path TEXT,
                slide_file_path TEXT,
                embedding_vector TEXT,
                metadata_json TEXT,
                status TEXT NOT NULL,
                tags_json TEXT NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            '''
        )
        conn.execute(
            '''
            CREATE TABLE IF NOT EXISTS component_catalog (
                component_id TEXT PRIMARY KEY,
                component_type TEXT NOT NULL,
                component_subtype TEXT NOT NULL,
                source_slide_id TEXT NOT NULL,
                layout_type TEXT NOT NULL,
                bounding_box TEXT NOT NULL,
                style_token TEXT NOT NULL,
                color_token TEXT NOT NULL,
                width_ratio REAL NOT NULL,
                height_ratio REAL NOT NULL,
                preview_image TEXT,
                metadata_json TEXT,
                status TEXT NOT NULL,
                tags_json TEXT NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            '''
        )
        conn.execute(
            '''
            CREATE TABLE IF NOT EXISTS assembly_runs (
                run_id TEXT PRIMARY KEY,
                request_json TEXT NOT NULL,
                plan_json TEXT NOT NULL,
                output_path TEXT NOT NULL,
                slide_count INTEGER NOT NULL,
                layout_summary TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
            '''
        )
        conn.execute(
            '''
            CREATE TABLE IF NOT EXISTS feedback_events (
                event_id TEXT PRIMARY KEY,
                event_type TEXT NOT NULL,
                entity_type TEXT NOT NULL,
                entity_id TEXT NOT NULL,
                payload_json TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
            '''
        )
        conn.execute('CREATE INDEX IF NOT EXISTS idx_slide_layout_score ON slide_catalog(layout_type, overall_quality_score DESC)')
        conn.execute('CREATE INDEX IF NOT EXISTS idx_slide_theme ON slide_catalog(industry, scenario, style)')
        conn.execute('CREATE INDEX IF NOT EXISTS idx_component_type ON component_catalog(component_type, component_subtype)')
        conn.execute('CREATE INDEX IF NOT EXISTS idx_feedback_entity ON feedback_events(entity_type, entity_id)')
    conn.close()
    return paths.root


def _row_from_record(record: PageRecord | dict[str, Any]) -> dict[str, Any]:
    if isinstance(record, PageRecord):
        row = record.to_row()
    else:
        row = dict(record)
        row.setdefault('tags_json', [])
        row.setdefault('metadata_json', {})
    row['tags_json'] = json.dumps(list(row.get('tags_json', [])), ensure_ascii=False)
    row['metadata_json'] = json.dumps(row.get('metadata_json', {}), ensure_ascii=False)
    return row


def save_slide_catalog(records: Iterable[PageRecord | dict[str, Any]], root: str | Path = 'ppt_template_library') -> Path:
    initialize_slide_storage(root)
    conn = _connect(root)
    now = datetime.utcnow().isoformat(timespec='seconds')
    rows: list[tuple[Any, ...]] = []
    for record in records:
        row = _row_from_record(record)
        rows.append(
            (
                row.get('slide_id', ''),
                row.get('source_template_id', ''),
                row.get('source_file', ''),
                row.get('source_url', ''),
                int(row.get('slide_number', 0)),
                row.get('slide_type', ''),
                row.get('slide_subtype', ''),
                row.get('industry', ''),
                row.get('scenario', ''),
                row.get('style', ''),
                row.get('layout_type', ''),
                row.get('primary_color', ''),
                row.get('secondary_color', ''),
                row.get('background_color', ''),
                row.get('dark_or_light', ''),
                row.get('aspect_ratio', ''),
                float(row.get('text_density', 0.0)),
                float(row.get('image_density', 0.0)),
                int(row.get('chart_count', 0)),
                int(row.get('table_count', 0)),
                int(row.get('shape_count', 0)),
                int(row.get('icon_count', 0)),
                int(row.get('image_count', 0)),
                int(row.get('text_box_count', 0)),
                int(bool(row.get('has_chart', False))),
                int(bool(row.get('has_table', False))),
                int(bool(row.get('has_timeline', False))),
                int(bool(row.get('has_process', False))),
                int(bool(row.get('has_map', False))),
                int(bool(row.get('has_people', False))),
                int(bool(row.get('has_infographic', False))),
                int(bool(row.get('has_animation', False))),
                row.get('editable_level', ''),
                int(row.get('design_score', 0)),
                int(row.get('layout_score', 0)),
                int(row.get('color_score', 0)),
                int(row.get('usability_score', 0)),
                int(row.get('modern_score', 0)),
                int(row.get('overall_quality_score', 0)),
                row.get('preview_image', ''),
                row.get('thumbnail_path', ''),
                row.get('slide_file_path', ''),
                row.get('embedding_vector', ''),
                row.get('metadata_json', '{}'),
                row.get('status', 'seed'),
                row.get('tags_json', '[]'),
                now,
                now,
            )
        )
    with conn:
        conn.executemany(
            '''
            INSERT INTO slide_catalog (
                slide_id, source_template_id, source_file, source_url, slide_number,
                slide_type, slide_subtype, industry, scenario, style, layout_type,
                primary_color, secondary_color, background_color, dark_or_light,
                aspect_ratio, text_density, image_density, chart_count, table_count,
                shape_count, icon_count, image_count, text_box_count, has_chart,
                has_table, has_timeline, has_process, has_map, has_people,
                has_infographic, has_animation, editable_level, design_score,
                layout_score, color_score, usability_score, modern_score,
                overall_quality_score, preview_image, thumbnail_path, slide_file_path,
                embedding_vector, metadata_json, status, tags_json, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(slide_id) DO UPDATE SET
                source_template_id=excluded.source_template_id,
                source_file=excluded.source_file,
                source_url=excluded.source_url,
                slide_number=excluded.slide_number,
                slide_type=excluded.slide_type,
                slide_subtype=excluded.slide_subtype,
                industry=excluded.industry,
                scenario=excluded.scenario,
                style=excluded.style,
                layout_type=excluded.layout_type,
                primary_color=excluded.primary_color,
                secondary_color=excluded.secondary_color,
                background_color=excluded.background_color,
                dark_or_light=excluded.dark_or_light,
                aspect_ratio=excluded.aspect_ratio,
                text_density=excluded.text_density,
                image_density=excluded.image_density,
                chart_count=excluded.chart_count,
                table_count=excluded.table_count,
                shape_count=excluded.shape_count,
                icon_count=excluded.icon_count,
                image_count=excluded.image_count,
                text_box_count=excluded.text_box_count,
                has_chart=excluded.has_chart,
                has_table=excluded.has_table,
                has_timeline=excluded.has_timeline,
                has_process=excluded.has_process,
                has_map=excluded.has_map,
                has_people=excluded.has_people,
                has_infographic=excluded.has_infographic,
                has_animation=excluded.has_animation,
                editable_level=excluded.editable_level,
                design_score=excluded.design_score,
                layout_score=excluded.layout_score,
                color_score=excluded.color_score,
                usability_score=excluded.usability_score,
                modern_score=excluded.modern_score,
                overall_quality_score=excluded.overall_quality_score,
                preview_image=excluded.preview_image,
                thumbnail_path=excluded.thumbnail_path,
                slide_file_path=excluded.slide_file_path,
                embedding_vector=excluded.embedding_vector,
                metadata_json=excluded.metadata_json,
                status=excluded.status,
                tags_json=excluded.tags_json,
                updated_at=excluded.updated_at
            ''' ,
            rows,
        )
    conn.close()
    return db_path(root)


def load_slide_catalog(root: str | Path = 'ppt_template_library') -> pd.DataFrame:
    initialize_slide_storage(root)
    conn = _connect(root)
    df = pd.read_sql_query('SELECT * FROM slide_catalog ORDER BY overall_quality_score DESC, slide_id ASC', conn)
    conn.close()
    return df


def _component_row_from_record(record: ComponentRecord | dict[str, Any]) -> dict[str, Any]:
    if isinstance(record, ComponentRecord):
        row = record.to_row()
    else:
        row = dict(record)
        row.setdefault('tags_json', [])
        row.setdefault('metadata_json', {})
    row['tags_json'] = json.dumps(list(row.get('tags_json', [])), ensure_ascii=False)
    row['metadata_json'] = json.dumps(row.get('metadata_json', {}), ensure_ascii=False)
    return row


def save_component_catalog(records: Iterable[ComponentRecord | dict[str, Any]], root: str | Path = 'ppt_template_library') -> Path:
    initialize_slide_storage(root)
    conn = _connect(root)
    now = datetime.utcnow().isoformat(timespec='seconds')
    rows: list[tuple[Any, ...]] = []
    for record in records:
        row = _component_row_from_record(record)
        rows.append(
            (
                row.get('component_id', ''),
                row.get('component_type', ''),
                row.get('component_subtype', ''),
                row.get('source_slide_id', ''),
                row.get('layout_type', ''),
                row.get('bounding_box', ''),
                row.get('style_token', ''),
                row.get('color_token', ''),
                float(row.get('width_ratio', 0.0)),
                float(row.get('height_ratio', 0.0)),
                row.get('preview_image', ''),
                row.get('metadata_json', '{}'),
                row.get('status', 'seed'),
                row.get('tags_json', '[]'),
                now,
                now,
            )
        )
    with conn:
        conn.executemany(
            '''
            INSERT INTO component_catalog (
                component_id, component_type, component_subtype, source_slide_id,
                layout_type, bounding_box, style_token, color_token, width_ratio,
                height_ratio, preview_image, metadata_json, status, tags_json,
                created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(component_id) DO UPDATE SET
                component_type=excluded.component_type,
                component_subtype=excluded.component_subtype,
                source_slide_id=excluded.source_slide_id,
                layout_type=excluded.layout_type,
                bounding_box=excluded.bounding_box,
                style_token=excluded.style_token,
                color_token=excluded.color_token,
                width_ratio=excluded.width_ratio,
                height_ratio=excluded.height_ratio,
                preview_image=excluded.preview_image,
                metadata_json=excluded.metadata_json,
                status=excluded.status,
                tags_json=excluded.tags_json,
                updated_at=excluded.updated_at
            ''' ,
            rows,
        )
    conn.close()
    return db_path(root)


def load_component_catalog(root: str | Path = 'ppt_template_library') -> pd.DataFrame:
    initialize_slide_storage(root)
    conn = _connect(root)
    df = pd.read_sql_query('SELECT * FROM component_catalog ORDER BY component_type ASC, component_id ASC', conn)
    conn.close()
    return df


def save_assembly_run(record: dict[str, Any], root: str | Path = 'ppt_template_library') -> None:
    initialize_slide_storage(root)
    conn = _connect(root)
    now = datetime.utcnow().isoformat(timespec='seconds')
    run_id = record.get('run_id', f"run-{now}")
    with conn:
        conn.execute(
            '''
            INSERT INTO assembly_runs (run_id, request_json, plan_json, output_path, slide_count, layout_summary, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(run_id) DO UPDATE SET
                request_json=excluded.request_json,
                plan_json=excluded.plan_json,
                output_path=excluded.output_path,
                slide_count=excluded.slide_count,
                layout_summary=excluded.layout_summary
            ''' ,
            (
                run_id,
                json.dumps(record.get('request_json', {}), ensure_ascii=False),
                json.dumps(record.get('plan_json', {}), ensure_ascii=False),
                record.get('output_path', ''),
                int(record.get('slide_count', 0)),
                json.dumps(record.get('layout_summary', {}), ensure_ascii=False),
                now,
            ),
        )
    conn.close()


def save_feedback_event(record: dict[str, Any], root: str | Path = 'ppt_template_library') -> None:
    initialize_slide_storage(root)
    conn = _connect(root)
    now = datetime.utcnow().isoformat(timespec='seconds')
    event_id = record.get('event_id', f"evt-{now}")
    with conn:
        conn.execute(
            '''
            INSERT INTO feedback_events (event_id, event_type, entity_type, entity_id, payload_json, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(event_id) DO UPDATE SET
                event_type=excluded.event_type,
                entity_type=excluded.entity_type,
                entity_id=excluded.entity_id,
                payload_json=excluded.payload_json
            ''',
            (
                event_id,
                record.get('event_type', 'select'),
                record.get('entity_type', 'slide'),
                record.get('entity_id', ''),
                json.dumps(record.get('payload_json', {}), ensure_ascii=False),
                now,
            ),
        )
    conn.close()


def load_slide_summary(root: str | Path = 'ppt_template_library') -> dict[str, Any]:
    df = load_slide_catalog(root)
    if df.empty:
        return {
            'slide_count': 0,
            'component_count': int(load_component_catalog(root).shape[0]),
            'layout_types': 0,
            'average_score': 0.0,
            'premium_count': 0,
        }
    return {
        'slide_count': int(len(df)),
        'component_count': int(load_component_catalog(root).shape[0]),
        'layout_types': int(df['layout_type'].nunique()),
        'average_score': round(float(df['overall_quality_score'].mean()), 1),
        'premium_count': int((df['overall_quality_score'] >= 90).sum()),
        'chart_pages': int(df['has_chart'].astype(int).sum()),
        'table_pages': int(df['has_table'].astype(int).sum()),
        'timeline_pages': int(df['has_timeline'].astype(int).sum()),
        'process_pages': int(df['has_process'].astype(int).sum()),
        'map_pages': int(df['has_map'].astype(int).sum()),
    }
