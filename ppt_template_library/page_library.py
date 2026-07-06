from __future__ import annotations

import hashlib
from dataclasses import asdict
from itertools import cycle
from pathlib import Path
from typing import Any, Iterable

from .page_models import ComponentRecord, PageRecord
from .preview import create_component_preview, create_page_preview


LAYOUT_BLUEPRINTS: dict[str, dict[str, Any]] = {
    'cover_slide': {
        'slide_subtypes': ['business_cover', 'technology_cover', 'minimal_cover', 'academic_cover', 'government_cover', 'product_cover'],
        'tags': ('cover', 'title', 'hero'),
        'slide_type': 'cover',
        'editable_level': 'high',
        'base_score': 95,
        'has_chart': False,
        'has_table': False,
        'has_timeline': False,
        'has_process': False,
        'has_map': False,
        'has_people': False,
        'has_infographic': True,
        'component_types': ('title_block', 'subtitle_block', 'image_block', 'hero_block'),
    },
    'agenda_slide': {
        'slide_subtypes': ['normal_agenda', 'timeline_agenda', 'card_agenda', 'icon_agenda', 'numbered_agenda'],
        'tags': ('agenda', 'table_of_contents', 'sequence'),
        'slide_type': 'agenda',
        'editable_level': 'high',
        'base_score': 90,
        'has_chart': False,
        'has_table': False,
        'has_timeline': True,
        'has_process': False,
        'has_map': False,
        'has_people': False,
        'has_infographic': True,
        'component_types': ('title_block', 'timeline_block', 'agenda_block', 'icon_block'),
    },
    'section_slide': {
        'slide_subtypes': ['numeric_section', 'image_section', 'minimal_section', 'gradient_section', 'dark_section'],
        'tags': ('section', 'chapter', 'divider'),
        'slide_type': 'section',
        'editable_level': 'high',
        'base_score': 91,
        'has_chart': False,
        'has_table': False,
        'has_timeline': False,
        'has_process': False,
        'has_map': False,
        'has_people': False,
        'has_infographic': True,
        'component_types': ('title_block', 'number_card', 'image_block', 'section_banner'),
    },
    'content_slide': {
        'slide_subtypes': ['two_column', 'three_column', 'card_grid', 'text_image', 'story_card', 'quote_card'],
        'tags': ('content', 'text', 'cards'),
        'slide_type': 'content',
        'editable_level': 'high',
        'base_score': 86,
        'has_chart': False,
        'has_table': False,
        'has_timeline': False,
        'has_process': False,
        'has_map': False,
        'has_people': False,
        'has_infographic': True,
        'component_types': ('title_block', 'text_block', 'card_block', 'image_block'),
    },
    'data_analysis_slide': {
        'slide_subtypes': ['kpi_overview', 'trend_review', 'financial_review', 'market_review', 'growth_review', 'ops_review'],
        'tags': ('data', 'analysis', 'insight'),
        'slide_type': 'data_analysis',
        'editable_level': 'high',
        'base_score': 94,
        'has_chart': True,
        'has_table': True,
        'has_timeline': False,
        'has_process': False,
        'has_map': False,
        'has_people': False,
        'has_infographic': True,
        'component_types': ('kpi_card', 'chart_block', 'table_block', 'insight_block'),
    },
    'chart_slide': {
        'slide_subtypes': ['bar_chart', 'line_chart', 'area_chart', 'pie_chart', 'combo_chart', 'radar_chart'],
        'tags': ('chart', 'visualization', 'analytics'),
        'slide_type': 'chart',
        'editable_level': 'high',
        'base_score': 92,
        'has_chart': True,
        'has_table': False,
        'has_timeline': False,
        'has_process': False,
        'has_map': False,
        'has_people': False,
        'has_infographic': True,
        'component_types': ('chart_block', 'legend_block', 'metric_strip', 'annotation_block'),
    },
    'table_slide': {
        'slide_subtypes': ['data_table', 'financial_table', 'comparison_table', 'ranking_table', 'plan_table', 'quotation_table'],
        'tags': ('table', 'grid', 'data'),
        'slide_type': 'table',
        'editable_level': 'high',
        'base_score': 88,
        'has_chart': False,
        'has_table': True,
        'has_timeline': False,
        'has_process': False,
        'has_map': False,
        'has_people': False,
        'has_infographic': True,
        'component_types': ('table_block', 'kpi_card', 'summary_block', 'note_block'),
    },
    'timeline_slide': {
        'slide_subtypes': ['horizontal_timeline', 'vertical_timeline', 'roadmap', 'history', 'milestone', 'year_summary'],
        'tags': ('timeline', 'roadmap', 'history'),
        'slide_type': 'timeline',
        'editable_level': 'high',
        'base_score': 90,
        'has_chart': False,
        'has_table': False,
        'has_timeline': True,
        'has_process': False,
        'has_map': False,
        'has_people': False,
        'has_infographic': True,
        'component_types': ('timeline_block', 'milestone_block', 'date_block', 'connector_block'),
    },
    'process_slide': {
        'slide_subtypes': ['flow_process', 'approval_process', 'cycle_process', 'business_process', 'project_process', 'service_process'],
        'tags': ('process', 'workflow', 'steps'),
        'slide_type': 'process',
        'editable_level': 'high',
        'base_score': 91,
        'has_chart': False,
        'has_table': False,
        'has_timeline': False,
        'has_process': True,
        'has_map': False,
        'has_people': False,
        'has_infographic': True,
        'component_types': ('process_block', 'step_block', 'arrow_block', 'summary_block'),
    },
    'comparison_slide': {
        'slide_subtypes': ['before_after', 'two_column_compare', 'product_compare', 'pros_cons', 'competitive_compare', 'feature_compare'],
        'tags': ('comparison', 'before_after', 'decision'),
        'slide_type': 'comparison',
        'editable_level': 'high',
        'base_score': 88,
        'has_chart': False,
        'has_table': False,
        'has_timeline': False,
        'has_process': False,
        'has_map': False,
        'has_people': False,
        'has_infographic': True,
        'component_types': ('comparison_block', 'card_block', 'divider_block', 'conclusion_block'),
    },
    'relation_slide': {
        'slide_subtypes': ['org_chart', 'network_map', 'ecosystem_map', 'structure_map', 'relationship_map', 'hierarchy_map'],
        'tags': ('relation', 'org_chart', 'network'),
        'slide_type': 'relation',
        'editable_level': 'medium',
        'base_score': 87,
        'has_chart': False,
        'has_table': False,
        'has_timeline': False,
        'has_process': False,
        'has_map': False,
        'has_people': True,
        'has_infographic': True,
        'component_types': ('org_chart_block', 'node_block', 'connector_block', 'caption_block'),
    },
    'strategy_slide': {
        'slide_subtypes': ['swot', 'pest', 'porter', 'business_model', 'value_chain', 'positioning_map'],
        'tags': ('strategy', 'swot', 'planning'),
        'slide_type': 'strategy',
        'editable_level': 'high',
        'base_score': 93,
        'has_chart': False,
        'has_table': False,
        'has_timeline': False,
        'has_process': False,
        'has_map': False,
        'has_people': False,
        'has_infographic': True,
        'component_types': ('swot_block', 'matrix_block', 'strategy_block', 'decision_block'),
    },
    'planning_slide': {
        'slide_subtypes': ['gantt', 'annual_plan', 'quarterly_plan', 'task_plan', 'roadmap_plan', 'milestone_plan'],
        'tags': ('plan', 'gantt', 'milestone'),
        'slide_type': 'plan',
        'editable_level': 'high',
        'base_score': 91,
        'has_chart': False,
        'has_table': True,
        'has_timeline': True,
        'has_process': True,
        'has_map': False,
        'has_people': False,
        'has_infographic': True,
        'component_types': ('gantt_block', 'timeline_block', 'task_card', 'milestone_block'),
    },
    'people_slide': {
        'slide_subtypes': ['team_intro', 'leader_intro', 'expert_intro', 'staff_intro', 'speaker_intro', 'member_cards'],
        'tags': ('people', 'team', 'profile'),
        'slide_type': 'people',
        'editable_level': 'high',
        'base_score': 87,
        'has_chart': False,
        'has_table': False,
        'has_timeline': False,
        'has_process': False,
        'has_map': False,
        'has_people': True,
        'has_infographic': True,
        'component_types': ('person_card', 'avatar_block', 'bio_block', 'contact_block'),
    },
    'product_slide': {
        'slide_subtypes': ['product_intro', 'feature_matrix', 'product_roadmap', 'parameter_sheet', 'advantage_map', 'release_plan'],
        'tags': ('product', 'feature', 'launch'),
        'slide_type': 'product',
        'editable_level': 'high',
        'base_score': 92,
        'has_chart': True,
        'has_table': True,
        'has_timeline': True,
        'has_process': False,
        'has_map': False,
        'has_people': False,
        'has_infographic': True,
        'component_types': ('product_card', 'feature_block', 'roadmap_block', 'spec_block'),
    },
    'gallery_slide': {
        'slide_subtypes': ['single_image', 'dual_image', 'triple_image', 'portfolio', 'photography', 'case_gallery'],
        'tags': ('gallery', 'image', 'portfolio'),
        'slide_type': 'gallery',
        'editable_level': 'high',
        'base_score': 86,
        'has_chart': False,
        'has_table': False,
        'has_timeline': False,
        'has_process': False,
        'has_map': False,
        'has_people': False,
        'has_infographic': True,
        'component_types': ('gallery_block', 'image_block', 'caption_block', 'footer_block'),
    },
    'map_slide': {
        'slide_subtypes': ['world_map', 'china_map', 'region_map', 'distribution_map', 'logistics_map', 'market_map'],
        'tags': ('map', 'location', 'distribution'),
        'slide_type': 'map',
        'editable_level': 'high',
        'base_score': 89,
        'has_chart': False,
        'has_table': False,
        'has_timeline': False,
        'has_process': False,
        'has_map': True,
        'has_people': False,
        'has_infographic': True,
        'component_types': ('map_block', 'location_card', 'legend_block', 'callout_block'),
    },
    'dashboard_slide': {
        'slide_subtypes': ['executive_dashboard', 'sales_dashboard', 'ops_dashboard', 'finance_dashboard', 'kpi_wall', 'monitoring_panel'],
        'tags': ('dashboard', 'kpi', 'monitoring'),
        'slide_type': 'dashboard',
        'editable_level': 'high',
        'base_score': 95,
        'has_chart': True,
        'has_table': True,
        'has_timeline': False,
        'has_process': False,
        'has_map': False,
        'has_people': False,
        'has_infographic': True,
        'component_types': ('dashboard_block', 'kpi_card', 'chart_block', 'table_block'),
    },
    'kpi_slide': {
        'slide_subtypes': ['kpi_cards', 'kpi_overview', 'kpi_trend', 'metric_board', 'metrics_detail', 'performance_wall'],
        'tags': ('kpi', 'metrics', 'performance'),
        'slide_type': 'kpi',
        'editable_level': 'high',
        'base_score': 94,
        'has_chart': True,
        'has_table': False,
        'has_timeline': False,
        'has_process': False,
        'has_map': False,
        'has_people': False,
        'has_infographic': True,
        'component_types': ('kpi_card', 'metric_card', 'trend_block', 'sparkline_block'),
    },
    'ending_slide': {
        'slide_subtypes': ['thanks', 'summary', 'q_and_a', 'contact', 'action_plan', 'outlook'],
        'tags': ('ending', 'thanks', 'contact'),
        'slide_type': 'ending',
        'editable_level': 'high',
        'base_score': 90,
        'has_chart': False,
        'has_table': False,
        'has_timeline': False,
        'has_process': False,
        'has_map': False,
        'has_people': False,
        'has_infographic': True,
        'component_types': ('title_block', 'contact_block', 'thankyou_block', 'footer_block'),
    },
}

THEME_VARIANTS: list[dict[str, Any]] = [
    {
        'style': 'business',
        'industry': 'finance',
        'scenario': 'annual_report',
        'primary_color': '#1B3756',
        'secondary_color': '#245A8D',
        'background_color': '#F7F9FC',
        'dark_or_light': 'light',
        'accent_color': '#0F766E',
        'tags': ('business', 'blue', 'corporate'),
    },
    {
        'style': 'technology',
        'industry': 'technology',
        'scenario': 'product_launch',
        'primary_color': '#102A43',
        'secondary_color': '#2563EB',
        'background_color': '#F8FAFF',
        'dark_or_light': 'light',
        'accent_color': '#06B6D4',
        'tags': ('technology', 'blue', 'digital'),
    },
    {
        'style': 'academic',
        'industry': 'education',
        'scenario': 'research_report',
        'primary_color': '#4A427C',
        'secondary_color': '#9E6F2C',
        'background_color': '#FAFAF6',
        'dark_or_light': 'light',
        'accent_color': '#7C3AED',
        'tags': ('academic', 'purple', 'report'),
    },
    {
        'style': 'creative',
        'industry': 'marketing',
        'scenario': 'campaign_plan',
        'primary_color': '#164E63',
        'secondary_color': '#F97316',
        'background_color': '#FFFDF8',
        'dark_or_light': 'light',
        'accent_color': '#EC4899',
        'tags': ('creative', 'orange', 'editorial'),
    },
    {
        'style': 'government',
        'industry': 'government',
        'scenario': 'policy_report',
        'primary_color': '#7F1D1D',
        'secondary_color': '#B91C1C',
        'background_color': '#FFF8F8',
        'dark_or_light': 'light',
        'accent_color': '#14532D',
        'tags': ('government', 'red', 'formal'),
    },
    {
        'style': 'finance',
        'industry': 'finance',
        'scenario': 'financial_report',
        'primary_color': '#0F5132',
        'secondary_color': '#10B981',
        'background_color': '#F6FBF8',
        'dark_or_light': 'light',
        'accent_color': '#1D4ED8',
        'tags': ('finance', 'green', 'business'),
    },
    {
        'style': 'minimal',
        'industry': 'consulting',
        'scenario': 'project_report',
        'primary_color': '#334155',
        'secondary_color': '#64748B',
        'background_color': '#FBFCFD',
        'dark_or_light': 'light',
        'accent_color': '#0EA5E9',
        'tags': ('minimal', 'gray', 'clean'),
    },
    {
        'style': 'dark',
        'industry': 'ai',
        'scenario': 'strategy_deck',
        'primary_color': '#E5E7EB',
        'secondary_color': '#F59E0B',
        'background_color': '#0F172A',
        'dark_or_light': 'dark',
        'accent_color': '#22D3EE',
        'tags': ('dark', 'black_gold', 'premium'),
    },
]

COMPONENT_TYPES = [
    'title_block',
    'subtitle_block',
    'text_block',
    'number_card',
    'kpi_card',
    'chart_block',
    'table_block',
    'timeline_block',
    'process_block',
    'comparison_block',
    'matrix_block',
    'org_chart_block',
    'swot_block',
    'funnel_block',
    'pyramid_block',
    'people_card',
    'product_card',
    'gallery_block',
    'map_block',
    'dashboard_block',
    'header_block',
    'footer_block',
    'page_number_block',
    'icon_block',
]


def _slugify(value: str) -> str:
    text = value.lower()
    out = []
    previous_dash = False
    for ch in text:
        if ch.isalnum():
            out.append(ch)
            previous_dash = False
        else:
            if not previous_dash:
                out.append('-')
                previous_dash = True
    slug = ''.join(out).strip('-')
    return slug or 'slide'


def _score_slide(layout_type: str, variant_index: int, theme: dict[str, Any]) -> tuple[int, int, int, int, int, int]:
    blueprint = LAYOUT_BLUEPRINTS[layout_type]
    base = int(blueprint['base_score'])
    style_bonus = 4 if theme['style'] in {'business', 'technology', 'finance', 'academic'} else 2
    color_bonus = 3 if theme['dark_or_light'] == 'dark' else 2
    layout_score = min(100, base + (variant_index % 3) * 2)
    design_score = min(100, base - 2 + style_bonus)
    color_score = min(100, base - 3 + color_bonus)
    usability_score = min(100, base - 1 + (1 if blueprint['editable_level'] == 'high' else 0))
    modern_score = min(100, base - 1 + (2 if layout_type in {'dashboard_slide', 'data_analysis_slide', 'chart_slide', 'cover_slide'} else 0))
    overall = round(layout_score * 0.18 + design_score * 0.22 + color_score * 0.12 + usability_score * 0.22 + modern_score * 0.26)
    return layout_score, design_score, color_score, usability_score, modern_score, overall


def _theme_for(layout_index: int, variant_index: int) -> dict[str, Any]:
    theme = THEME_VARIANTS[(layout_index + variant_index) % len(THEME_VARIANTS)].copy()
    theme['theme_id'] = f"{theme['style']}-{variant_index + 1}"
    return theme


def build_seed_slide_catalog() -> list[PageRecord]:
    records: list[PageRecord] = []
    for layout_index, (layout_type, blueprint) in enumerate(LAYOUT_BLUEPRINTS.items()):
        subtype_cycle = cycle(blueprint['slide_subtypes'])
        for variant_index in range(8):
            theme = _theme_for(layout_index, variant_index)
            subtype = next(subtype_cycle)
            slide_id = f"{layout_type}-{variant_index + 1:02d}-{theme['style']}"
            layout_score, design_score, color_score, usability_score, modern_score, overall = _score_slide(layout_type, variant_index, theme)
            tags = tuple(dict.fromkeys((layout_type.replace('_slide', ''), subtype, *blueprint['tags'], *theme['tags'])))
            preview_path = str(Path('ppt_template_library') / 'preview' / 'slides' / f'{slide_id}.png')
            record = PageRecord(
                slide_id=slide_id,
                source_template_id=f"seed-template-{layout_index + 1:02d}",
                source_file=f"synthetic/{layout_type}/{slide_id}.pptx",
                source_url=f"local://ppt-template-library/{slide_id}",
                slide_number=variant_index + 1,
                slide_type=blueprint['slide_type'],
                slide_subtype=subtype,
                industry=theme['industry'],
                scenario=theme['scenario'],
                style=theme['style'],
                layout_type=layout_type.replace('_slide', ''),
                primary_color=theme['primary_color'],
                secondary_color=theme['secondary_color'],
                background_color=theme['background_color'],
                dark_or_light=theme['dark_or_light'],
                aspect_ratio='16:9' if variant_index % 5 else '4:3',
                text_density=round(0.32 + (variant_index % 4) * 0.12, 2),
                image_density=round(0.18 + (variant_index % 3) * 0.1, 2),
                chart_count=1 if blueprint['has_chart'] else 0,
                table_count=1 if blueprint['has_table'] else 0,
                shape_count=8 + variant_index,
                icon_count=3 + (variant_index % 4),
                image_count=1 if layout_type in {'cover_slide', 'gallery_slide', 'product_slide'} else 0,
                text_box_count=4 + (variant_index % 5),
                has_chart=bool(blueprint['has_chart']),
                has_table=bool(blueprint['has_table']),
                has_timeline=bool(blueprint['has_timeline']),
                has_process=bool(blueprint['has_process']),
                has_map=bool(blueprint['has_map']),
                has_people=bool(blueprint['has_people']),
                has_infographic=bool(blueprint['has_infographic']),
                has_animation=False,
                editable_level=blueprint['editable_level'],
                design_score=design_score,
                layout_score=layout_score,
                color_score=color_score,
                usability_score=usability_score,
                modern_score=modern_score,
                overall_quality_score=overall,
                preview_image=preview_path,
                thumbnail_path=preview_path,
                slide_file_path=f"templates/{slide_id}.pptx",
                embedding_vector='',
                metadata_json={
                    'layout_goal': layout_type.replace('_slide', ''),
                    'theme_id': theme['theme_id'],
                    'tags': list(tags),
                    'suggested_components': list(blueprint['component_types']),
                    'render_hint': layout_type,
                },
                status='seed',
                tags=tags,
            )
            records.append(record)
    return records


def build_seed_component_catalog(slide_records: Iterable[PageRecord] | None = None) -> list[ComponentRecord]:
    slides = list(slide_records) if slide_records is not None else build_seed_slide_catalog()
    by_layout: dict[str, list[PageRecord]] = {}
    for slide in slides:
        by_layout.setdefault(slide.layout_type, []).append(slide)

    records: list[ComponentRecord] = []
    for comp_index, component_type in enumerate(COMPONENT_TYPES):
        for variant_index in range(3):
            source_slide = slides[(comp_index * 3 + variant_index) % len(slides)]
            layout_type = source_slide.layout_type
            component_id = f"{component_type}-{variant_index + 1:02d}-{_slugify(layout_type)}"
            tags = tuple(dict.fromkeys((component_type.replace('_block', ''), layout_type, source_slide.style, source_slide.industry)))
            preview_path = str(Path('ppt_template_library') / 'preview' / 'components' / f'{component_id}.png')
            records.append(
                ComponentRecord(
                    component_id=component_id,
                    component_type=component_type.replace('_block', ''),
                    component_subtype=f"{component_type}_{variant_index + 1}",
                    source_slide_id=source_slide.slide_id,
                    layout_type=layout_type,
                    bounding_box=f"x:{0.08 + variant_index * 0.02:.2f},y:{0.10 + variant_index * 0.03:.2f},w:{0.68 - variant_index * 0.05:.2f},h:{0.22 + variant_index * 0.04:.2f}",
                    style_token=source_slide.style,
                    color_token=source_slide.primary_color,
                    width_ratio=round(0.72 - variant_index * 0.08, 2),
                    height_ratio=round(0.24 + variant_index * 0.05, 2),
                    preview_image=preview_path,
                    metadata_json={
                        'tags': list(tags),
                        'recommended_usage': component_type,
                        'source_layout_type': layout_type,
                        'source_slide_id': source_slide.slide_id,
                    },
                    status='seed',
                    tags=tags,
                )
            )
    return records


def generate_slide_previews(records: Iterable[PageRecord], limit: int | None = None) -> list[Path]:
    preview_paths: list[Path] = []
    for index, record in enumerate(records):
        if limit is not None and index >= limit:
            break
        preview_paths.append(
            create_page_preview(
                title=record.slide_type.replace('_', ' ').title(),
                subtitle=f"{record.style} · {record.industry} · {record.scenario}",
                layout_type=record.layout_type,
                score=record.overall_quality_score,
                tags=record.tags,
                out_path=record.preview_image,
            )
        )
    return preview_paths


def generate_component_previews(records: Iterable[ComponentRecord], limit: int | None = None) -> list[Path]:
    preview_paths: list[Path] = []
    for index, record in enumerate(records):
        if limit is not None and index >= limit:
            break
        preview_paths.append(
            create_component_preview(
                title=record.component_type.replace('_', ' ').title(),
                subtitle=f"{record.style_token} · {record.layout_type}",
                component_type=record.component_type,
                score=82 + (index % 9),
                tags=record.tags,
                out_path=record.preview_image,
            )
        )
    return preview_paths


def build_layout_summary(records: Iterable[PageRecord]) -> dict[str, Any]:
    slides = list(records)
    summary: dict[str, Any] = {
        'slide_count': len(slides),
        'layout_types': len({record.layout_type for record in slides}),
        'styles': len({record.style for record in slides}),
        'industries': len({record.industry for record in slides}),
        'premium_quality': sum(1 for record in slides if record.overall_quality_score >= 90),
        'chart_pages': sum(1 for record in slides if record.has_chart),
        'table_pages': sum(1 for record in slides if record.has_table),
        'timeline_pages': sum(1 for record in slides if record.has_timeline),
        'process_pages': sum(1 for record in slides if record.has_process),
        'map_pages': sum(1 for record in slides if record.has_map),
        'people_pages': sum(1 for record in slides if record.has_people),
    }
    return summary


def build_component_summary(records: Iterable[ComponentRecord]) -> dict[str, Any]:
    components = list(records)
    return {
        'component_count': len(components),
        'component_types': len({record.component_type for record in components}),
        'layout_types': len({record.layout_type for record in components}),
    }
