from __future__ import annotations

from dataclasses import dataclass, field, asdict
from typing import Any


PAGE_METADATA_FIELDS = [
    'slide_id',
    'source_template_id',
    'source_file',
    'source_url',
    'slide_number',
    'slide_type',
    'slide_subtype',
    'industry',
    'scenario',
    'style',
    'layout_type',
    'primary_color',
    'secondary_color',
    'background_color',
    'dark_or_light',
    'aspect_ratio',
    'text_density',
    'image_density',
    'chart_count',
    'table_count',
    'shape_count',
    'icon_count',
    'image_count',
    'text_box_count',
    'has_chart',
    'has_table',
    'has_timeline',
    'has_process',
    'has_map',
    'has_people',
    'has_infographic',
    'has_animation',
    'editable_level',
    'design_score',
    'layout_score',
    'color_score',
    'usability_score',
    'modern_score',
    'overall_quality_score',
    'preview_image',
    'thumbnail_path',
    'slide_file_path',
    'embedding_vector',
    'metadata_json',
    'status',
    'tags_json',
]

COMPONENT_METADATA_FIELDS = [
    'component_id',
    'component_type',
    'component_subtype',
    'source_slide_id',
    'layout_type',
    'bounding_box',
    'style_token',
    'color_token',
    'width_ratio',
    'height_ratio',
    'preview_image',
    'metadata_json',
    'status',
    'tags_json',
]


@dataclass(frozen=True)
class DesignSystem:
    primary_color: str
    secondary_color: str
    accent_color: str
    background_color: str
    title_font: str
    body_font: str
    number_font: str
    title_size: int
    subtitle_size: int
    body_size: int
    caption_size: int
    border_radius: int
    spacing_scale: str
    chart_palette: tuple[str, ...] = field(default_factory=tuple)
    icon_style: str = 'line'
    image_style: str = 'clean'
    dark_or_light: str = 'light'

    def to_row(self) -> dict[str, Any]:
        data = asdict(self)
        data['chart_palette'] = list(self.chart_palette)
        return data


@dataclass(frozen=True)
class ExcelProfile:
    rows: int = 0
    columns: int = 0
    numeric_columns: tuple[str, ...] = field(default_factory=tuple)
    text_columns: tuple[str, ...] = field(default_factory=tuple)
    date_columns: tuple[str, ...] = field(default_factory=tuple)
    category_columns: tuple[str, ...] = field(default_factory=tuple)
    region_columns: tuple[str, ...] = field(default_factory=tuple)
    month_columns: tuple[str, ...] = field(default_factory=tuple)
    metric_columns: tuple[str, ...] = field(default_factory=tuple)
    has_comparison: bool = False
    has_trend: bool = False
    has_percent: bool = False
    has_map: bool = False
    has_table: bool = False
    suggested_visuals: tuple[str, ...] = field(default_factory=tuple)
    summary: str = ''
    field_roles: dict[str, str] = field(default_factory=dict)

    def to_row(self) -> dict[str, Any]:
        data = asdict(self)
        for key in ('numeric_columns', 'text_columns', 'date_columns', 'category_columns', 'region_columns', 'month_columns', 'metric_columns', 'suggested_visuals'):
            data[key] = list(data[key])
        return data


@dataclass(frozen=True)
class AssemblyRequest:
    prompt: str
    title: str
    subtitle: str
    style: str
    page_range: tuple[int, int]
    industry: str = ''
    scenario: str = ''
    audience: str = ''
    tone: str = ''
    colors: tuple[str, ...] = field(default_factory=tuple)
    topics: tuple[str, ...] = field(default_factory=tuple)
    preferred_layouts: tuple[str, ...] = field(default_factory=tuple)
    data_profile: ExcelProfile | None = None

    def to_row(self) -> dict[str, Any]:
        data = asdict(self)
        data['page_range'] = list(self.page_range)
        data['colors'] = list(self.colors)
        data['topics'] = list(self.topics)
        data['preferred_layouts'] = list(self.preferred_layouts)
        if self.data_profile is not None:
            data['data_profile'] = self.data_profile.to_row()
        return data


@dataclass
class PageSearchFilters:
    query: str = ''
    layout_types: tuple[str, ...] = ()
    slide_types: tuple[str, ...] = ()
    industries: tuple[str, ...] = ()
    scenarios: tuple[str, ...] = ()
    styles: tuple[str, ...] = ()
    colors: tuple[str, ...] = ()
    tags: tuple[str, ...] = ()
    min_score: int = 0
    has_chart: bool | None = None
    has_table: bool | None = None
    has_timeline: bool | None = None
    has_process: bool | None = None
    has_map: bool | None = None
    has_people: bool | None = None
    has_infographic: bool | None = None
    dark_or_light: str = ''
    min_quality: int = 0


@dataclass(frozen=True)
class PageRecord:
    slide_id: str
    source_template_id: str
    source_file: str
    source_url: str
    slide_number: int
    slide_type: str
    slide_subtype: str
    industry: str
    scenario: str
    style: str
    layout_type: str
    primary_color: str
    secondary_color: str
    background_color: str
    dark_or_light: str
    aspect_ratio: str
    text_density: float
    image_density: float
    chart_count: int
    table_count: int
    shape_count: int
    icon_count: int
    image_count: int
    text_box_count: int
    has_chart: bool
    has_table: bool
    has_timeline: bool
    has_process: bool
    has_map: bool
    has_people: bool
    has_infographic: bool
    has_animation: bool
    editable_level: str
    design_score: int
    layout_score: int
    color_score: int
    usability_score: int
    modern_score: int
    overall_quality_score: int
    preview_image: str = ''
    thumbnail_path: str = ''
    slide_file_path: str = ''
    embedding_vector: str = ''
    metadata_json: dict[str, Any] = field(default_factory=dict)
    status: str = 'seed'
    tags: tuple[str, ...] = field(default_factory=tuple)

    def to_row(self) -> dict[str, Any]:
        data = asdict(self)
        data['tags_json'] = list(self.tags)
        data['metadata_json'] = self.metadata_json or {}
        return data


@dataclass(frozen=True)
class ComponentRecord:
    component_id: str
    component_type: str
    component_subtype: str
    source_slide_id: str
    layout_type: str
    bounding_box: str
    style_token: str
    color_token: str
    width_ratio: float
    height_ratio: float
    preview_image: str = ''
    metadata_json: dict[str, Any] = field(default_factory=dict)
    status: str = 'seed'
    tags: tuple[str, ...] = field(default_factory=tuple)

    def to_row(self) -> dict[str, Any]:
        data = asdict(self)
        data['tags_json'] = list(self.tags)
        data['metadata_json'] = self.metadata_json or {}
        return data


def normalize_tokens(values: str | tuple[str, ...] | list[str]) -> tuple[str, ...]:
    if isinstance(values, str):
        raw = values.split(',')
    else:
        raw = list(values)
    return tuple(token.strip().lower().replace(' ', '_') for token in raw if str(token).strip())
