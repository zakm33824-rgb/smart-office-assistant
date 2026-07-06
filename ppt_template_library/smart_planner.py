from __future__ import annotations

import re
from dataclasses import asdict
from datetime import date
from typing import Any, Iterable

import pandas as pd

from .page_models import AssemblyRequest, DesignSystem, ExcelProfile, PageRecord
from .page_search import filter_slide_dataframe


STYLE_THEME_MAP = {
    '????': ('business', 'blue', 'light'),
    '????': ('academic', 'purple', 'light'),
    '????': ('creative', 'orange', 'light'),
    '???': ('technology', 'blue', 'light'),
    '???': ('finance', 'green', 'light'),
    '???': ('government', 'red', 'light'),
    '????': ('dark', 'black_gold', 'dark'),
    'business': ('business', 'blue', 'light'),
    'academic': ('academic', 'purple', 'light'),
    'creative': ('creative', 'orange', 'light'),
    'technology': ('technology', 'blue', 'light'),
    'finance': ('finance', 'green', 'light'),
    'government': ('government', 'red', 'light'),
    'minimal': ('minimal', 'gray', 'light'),
    'dark': ('dark', 'black_gold', 'dark'),
}

INDUSTRY_PATTERNS = [
    (r'??|??|??|??|??|finance|bank|insurance|securities|investment', 'finance'),
    (r'???|??|??|??|??|automotive|auto|vehicle|car|energy|renewable', 'new_energy'),
    (r'???|??|????|AI|???|???|????|internet|software|ai|cloud|security', 'technology'),
    (r'??|??|??|??|medical|healthcare|pharma|biotech', 'medical'),
    (r'??|??|??|??|??|education|school|university|research', 'education'),
    (r'??|??|??|????|government|public|official', 'government'),
    (r'???|??|??|real estate|construction|engineering', 'construction'),
    (r'??|??|??|manufacturing|industrial|machinery', 'manufacturing'),
    (r'??|???|??|??|logistics|supply chain|transport|aviation', 'logistics'),
    (r'??|??|??|??|??|retail|e-commerce|commerce|marketing|advertising|media', 'commerce'),
    (r'??|??|??|??|travel|hotel|restaurant|food|tourism', 'service'),
]

SCENARIO_PATTERNS = [
    (r'??|??|????|annual report|annual summary|yearly report', 'annual_report'),
    (r'??|Q[1-4]|quarterly', 'quarterly_report'),
    (r'??|??|monthly', 'monthly_report'),
    (r'??|weekly', 'weekly_report'),
    (r'??|daily', 'daily_report'),
    (r'????|????|project report|project review', 'project_report'),
    (r'????|?????|business plan', 'business_plan'),
    (r'??|??|BP|pitch deck|fundraising', 'pitch_deck'),
    (r'????|??|????|product launch|product intro', 'product_launch'),
    (r'????|????|??|????|market analysis|industry analysis|competitive analysis', 'market_analysis'),
    (r'SWOT|PEST|??|strategy analysis|strategic analysis', 'strategy_analysis'),
    (r'??|??|??|training|courseware|teaching', 'training'),
    (r'??|??|defense|thesis|dissertation', 'academic_defense'),
    (r'??|???|??|??|??|??|resume|portfolio|photography|travel|wedding|birthday', 'personal_showcase'),
]

LAYOUT_KEYWORDS = [
    (r'??|??|??|??|cover|title page|hero', 'cover_slide'),
    (r'??|??|agenda', 'agenda_slide'),
    (r'??|??|??|section|divider', 'section_slide'),
    (r'KPI|??|???|??|dashboard|metric', 'kpi_slide'),
    (r'??|??|??|??|??|??|??|data|analysis|report', 'data_analysis_slide'),
    (r'??|??|??|??|??|??|chart', 'chart_slide'),
    (r'??|??|??|table|list', 'table_slide'),
    (r'???|??|???|???|roadmap|timeline|milestone', 'timeline_slide'),
    (r'??|??|??|??|???|process|workflow', 'process_slide'),
    (r'??|??|??|before|after|comparison', 'comparison_slide'),
    (r'??|??|??|??|??|??|org|team|people', 'relation_slide'),
    (r'SWOT|PEST|??|??|??|????|strategy', 'strategy_slide'),
    (r'??|??|??|??|??|??|plan|gantt', 'planning_slide'),
    (r'??|??|??|??|??|product', 'product_slide'),
    (r'??|??|??|??|??|map|region|distribution', 'map_slide'),
    (r'??|??|??|??|??|??|gallery|portfolio|case', 'gallery_slide'),
    (r'??|??|Q&A|????|??|thanks|end', 'ending_slide'),
]

COLOR_PATTERNS = [
    (r'蓝|blue|科技', 'blue'),
    (r'绿|green|环保', 'green'),
    (r'红|red|政务', 'red'),
    (r'紫|purple', 'purple'),
    (r'黑|dark|深色|高端', 'dark'),
    (r'金|gold|奢华', 'black_gold'),
]


def _normalize(text: str) -> str:
    return re.sub(r'\s+', ' ', (text or '').strip())


def infer_title(prompt: str) -> str:
    text = _normalize(prompt)
    head = re.split(r'包含|包括|分为|围绕|需要|要求|，|。', text, maxsplit=1)[0]
    head = re.sub(r'^(请|帮我|麻烦|制作|生成|创建|输出)?(一份|一个|1份)?', '', head).strip()
    head = head.replace('PPT', '').replace('ppt', '').replace('幻灯片', '').replace('演示文稿', '')
    head = head.strip(' -—,，.。') or '智能汇报'
    if '汇报' not in head and '报告' not in head and '复盘' not in head and '方案' not in head:
        head += '汇报'
    return head


def infer_sections(prompt: str) -> list[str]:
    text = _normalize(prompt)
    match = re.search(r'(?:包含|包括|分为|围绕)(.+?)(?:\d+\s*个?章节|章节|部分|模块|$)', text)
    if match:
        raw = match.group(1)
        sections = [item.strip(' 。,，;；:：') for item in re.split(r'[、,，;；/]+', raw) if item.strip(' 。,，;；:：')]
        sections = [item for item in sections if item and len(item) <= 20]
        if sections:
            return sections
    keyword_sections = []
    keyword_map = [
        (r'业绩|营收|销售|经营|财务', '业绩概况'),
        (r'用户|增长|流量|留存|转化', '用户增长'),
        (r'问题|风险|挑战|瓶颈', '问题分析'),
        (r'规划|计划|策略|下一步|下月|下季度', '行动计划'),
        (r'市场|行业|竞品|竞争', '市场洞察'),
        (r'产品|功能|研发|技术', '产品进展'),
        (r'团队|组织|人员|角色', '团队协同'),
        (r'数据|指标|仪表盘|KPI', '数据洞察'),
        (r'时间|历程|里程碑|路线图', '发展历程'),
        (r'地图|地区|区域|分布', '区域分布'),
    ]
    for pattern, label in keyword_map:
        if re.search(pattern, text) and label not in keyword_sections:
            keyword_sections.append(label)
    return keyword_sections or ['背景与目标', '核心进展', '问题分析', '行动计划']


def infer_industry(text: str) -> str:
    for pattern, industry in INDUSTRY_PATTERNS:
        if re.search(pattern, text, re.IGNORECASE):
            return industry
    return 'general'


def infer_scenario(text: str) -> str:
    for pattern, scenario in SCENARIO_PATTERNS:
        if re.search(pattern, text, re.IGNORECASE):
            return scenario
    return 'presentation'


def infer_style(prompt: str, style_name: str) -> tuple[str, str, str]:
    text = _normalize(prompt)
    for pattern, color in COLOR_PATTERNS:
        if re.search(pattern, text, re.IGNORECASE):
            return style_name or '简约商务', color, 'light' if color != 'dark' else 'dark'
    if style_name in STYLE_THEME_MAP:
        mapped = STYLE_THEME_MAP[style_name]
        return style_name, mapped[1], mapped[2]
    return style_name or '简约商务', 'blue', 'light'


def build_design_system(prompt: str, style_name: str, data_profile: ExcelProfile | None = None) -> DesignSystem:
    normalized_style, color_token, tone = infer_style(prompt, style_name)
    palette_map = {
        'blue': ('#1B3756', '#245A8D', '#0F766E', '#F7F9FC'),
        'green': ('#0F5132', '#10B981', '#0EA5E9', '#F6FBF8'),
        'red': ('#7F1D1D', '#B91C1C', '#14532D', '#FFF8F8'),
        'purple': ('#4A427C', '#7C3AED', '#9E6F2C', '#FAFAF6'),
        'orange': ('#7C2D12', '#F97316', '#FB923C', '#FFF7ED'),
        'teal': ('#0F4C5C', '#0F766E', '#14B8A6', '#F0FDFA'),
        'gray': ('#1F2937', '#64748B', '#0F766E', '#F8FAFC'),
        'dark': ('#E5E7EB', '#F59E0B', '#22D3EE', '#0F172A'),
        'black_gold': ('#111827', '#F59E0B', '#D4AF37', '#0B0F19'),
    }
    primary, secondary, accent, background = palette_map.get(color_token, palette_map['blue'])
    if data_profile and data_profile.has_map:
        accent = '#2563EB'
    chart_palette = (primary, secondary, accent, '#14B8A6', '#F97316', '#8B5CF6')
    return DesignSystem(
        primary_color=primary,
        secondary_color=secondary,
        accent_color=accent,
        background_color=background,
        title_font='Microsoft YaHei',
        body_font='Microsoft YaHei',
        number_font='Aptos',
        title_size=32,
        subtitle_size=16,
        body_size=18,
        caption_size=11,
        border_radius=14,
        spacing_scale='8pt',
        chart_palette=chart_palette,
        icon_style='line',
        image_style='clean',
        dark_or_light=tone,
    )


def analyze_dataframe_profile(df: pd.DataFrame | None) -> ExcelProfile | None:
    if df is None or df.empty:
        return None
    columns = list(df.columns)
    numeric_columns = tuple(str(col) for col in df.select_dtypes(include='number').columns)
    date_columns = tuple(str(col) for col in columns if any(keyword in str(col).lower() for keyword in ['date', 'time', 'month', 'year', '日期', '时间', '月份', '年度']))
    text_columns = tuple(str(col) for col in columns if str(col) not in numeric_columns and str(col) not in date_columns)
    category_columns = tuple(str(col) for col in columns if str(col) not in numeric_columns and str(col) not in date_columns)
    region_columns = tuple(str(col) for col in columns if re.search(r'地区|区域|省|市|国家|region|area|country', str(col), re.IGNORECASE))
    month_columns = tuple(str(col) for col in columns if re.search(r'月|month|月份', str(col), re.IGNORECASE))
    metric_columns = numeric_columns
    has_percent = any(re.search(r'率|增速|占比|比例|%|percent', str(col), re.IGNORECASE) for col in columns)
    has_map = bool(region_columns)
    has_trend = bool(date_columns or month_columns)
    has_comparison = any(re.search(r'实际|目标|plan|actual|同比|环比|对比', str(col), re.IGNORECASE) for col in columns)
    has_table = len(columns) >= 4
    suggested: list[str] = []
    if metric_columns:
        suggested.extend(['kpi_slide', 'dashboard_slide'])
    if has_trend:
        suggested.extend(['chart_slide', 'timeline_slide'])
    if has_map:
        suggested.append('map_slide')
    if has_table:
        suggested.append('table_slide')
    if has_comparison:
        suggested.append('comparison_slide')
    summary = f"{len(df)} rows x {len(columns)} columns; numeric={len(numeric_columns)}"
    field_roles: dict[str, str] = {}
    for col in columns:
        col_str = str(col)
        if col_str in numeric_columns:
            field_roles[col_str] = 'metric'
        elif col_str in region_columns:
            field_roles[col_str] = 'region'
        elif col_str in date_columns or col_str in month_columns:
            field_roles[col_str] = 'time'
        elif re.search(r'目标|plan|actual|同比|环比|率|占比|增长', col_str, re.IGNORECASE):
            field_roles[col_str] = 'indicator'
        else:
            field_roles[col_str] = 'dimension'
    return ExcelProfile(
        rows=int(len(df)),
        columns=int(len(columns)),
        numeric_columns=numeric_columns,
        text_columns=text_columns,
        date_columns=date_columns,
        category_columns=tuple(col for col in columns if col not in numeric_columns),
        region_columns=region_columns,
        month_columns=month_columns,
        metric_columns=metric_columns,
        has_comparison=has_comparison,
        has_trend=has_trend,
        has_percent=has_percent,
        has_map=has_map,
        has_table=has_table,
        suggested_visuals=tuple(dict.fromkeys(suggested)),
        summary=summary,
        field_roles=field_roles,
    )




def _style_family(style_name: str) -> str:
    text = _normalize(style_name).lower()
    style_rules = [
        (('academic', '\u5b66\u672f', '\u8bba\u6587', '\u7b54\u8fa9'), 'academic'),
        (('creative', '\u521b\u610f', '\u63d2\u753b', '\u5b5f\u83f2\u65af', '\u6d77\u62a5', '\u827a\u672f'), 'creative'),
        (('technology', '\u79d1\u6280', 'tech', 'ai', '\u4eba\u5de5\u667a\u80fd', '\u4e92\u8054\u7f51', '\u4e91', '\u6570\u5b57'), 'technology'),
        (('finance', '\u91d1\u878d', 'bank', 'banking', 'investment', '\u8bc1\u5238', '\u94f6\u884c', '\u4fdd\u9669'), 'finance'),
        (('government', '\u653f\u52a1', '\u653f\u5e9c', '\u515a\u5efa', '\u7ea2\u5934', 'official'), 'government'),
        (('dark', '\u6df1\u8272', '\u9ed1\u91d1', '\u591c\u95f4', '\u6697\u8272'), 'dark'),
        (('minimal', '\u6781\u7b80', '\u7b80\u6d01', '\u767d\u5e95', '\u7eaf\u767d', '\u7559\u767d'), 'minimal'),
    ]
    for tokens, family in style_rules:
        if any(token in text for token in tokens):
            return family
    return 'business'


STYLE_LAYOUT_BIASES: dict[str, list[str]] = {
    'business': ['content_slide', 'comparison_slide', 'timeline_slide', 'data_analysis_slide', 'table_slide'],
    'academic': ['section_slide', 'table_slide', 'chart_slide', 'timeline_slide', 'data_analysis_slide'],
    'creative': ['gallery_slide', 'product_slide', 'people_slide', 'comparison_slide', 'content_slide'],
    'technology': ['dashboard_slide', 'chart_slide', 'data_analysis_slide', 'map_slide', 'process_slide'],
    'finance': ['dashboard_slide', 'table_slide', 'comparison_slide', 'chart_slide', 'kpi_slide'],
    'government': ['section_slide', 'timeline_slide', 'process_slide', 'table_slide', 'strategy_slide'],
    'minimal': ['content_slide', 'section_slide', 'table_slide', 'ending_slide'],
    'dark': ['dashboard_slide', 'comparison_slide', 'gallery_slide', 'ending_slide'],
}

INDUSTRY_LABELS = {
    'general': '\u901a\u7528',
    'finance': '\u91d1\u878d',
    'new_energy': '\u65b0\u80fd\u6e90',
    'technology': '\u79d1\u6280',
    'medical': '\u533b\u7597',
    'education': '\u6559\u80b2',
    'government': '\u653f\u52a1',
    'construction': '\u5efa\u7b51',
    'manufacturing': '\u5236\u9020',
    'logistics': '\u7269\u6d41',
    'commerce': '\u5546\u4e1a',
    'service': '\u670d\u52a1',
}

SCENARIO_LABELS = {
    'presentation': '\u6f14\u793a',
    'annual_report': '\u5e74\u5ea6\u603b\u7ed3',
    'quarterly_report': '\u5b63\u5ea6\u590d\u76d8',
    'monthly_report': '\u6708\u5ea6\u590d\u76d8',
    'weekly_report': '\u5468\u62a5',
    'daily_report': '\u65e5\u62a5',
    'project_report': '\u9879\u76ee\u6c47\u62a5',
    'business_plan': '\u5546\u4e1a\u8ba1\u5212',
    'pitch_deck': '\u878d\u8d44\u8ba1\u5212',
    'product_launch': '\u4ea7\u54c1\u53d1\u5e03',
    'market_analysis': '\u5e02\u573a\u5206\u6790',
    'strategy_analysis': '\u6218\u7565\u5206\u6790',
    'training': '\u57f9\u8bad\u8bfe\u4ef6',
    'academic_defense': '\u5b66\u672f\u7b54\u8fa9',
    'personal_showcase': '\u4e2a\u4eba\u5c55\u793a',
}


def _human_label(token: str, mapping: dict[str, str]) -> str:
    return mapping.get(token, token.replace('_', ' '))


def _pick_layout_for_section(section: str, data_profile: ExcelProfile | None) -> str:
    text = section or ''
    for pattern, layout in LAYOUT_KEYWORDS:
        if re.search(pattern, text, re.IGNORECASE):
            return layout
    if data_profile:
        if data_profile.has_map and not re.search(r'地图|区域|地区|分布', text):
            return 'map_slide'
        if data_profile.has_trend and not re.search(r'时间|历程|里程碑|路线图', text):
            return 'chart_slide'
        if data_profile.metric_columns:
            return 'data_analysis_slide'
    return 'content_slide'


def _page_tags(section: str, layout_type: str, industry: str, scenario: str, style: str) -> tuple[str, ...]:
    base = [industry, scenario, style, layout_type]
    if section:
        base.append(section)
    return tuple(dict.fromkeys(token for token in base if token))


def _sample_bullets(section: str, title: str, variant: int = 0) -> list[str]:
    text = _normalize(section)
    if re.search('\u4e1a\u7ee9|\u6536\u5165|\u9500\u552e|\u8d22\u52a1|\u6210\u957f|\u8003\u6838', text):
        variants = [
            [
                '\u0051\u0033\u6838\u5fc3\u4e1a\u7ee9\u7ee7\u7eed\u4fdd\u6301\u589e\u957f\uff0c\u76ee\u6807\u8fbe\u6210\u7387\u4fdd\u6301\u5728\u8f83\u9ad8\u6c34\u5e73\u3002',
                '\u6536\u5165\u589e\u91cf\u4e3b\u8981\u6765\u81ea\u6838\u5fc3\u533a\u57df\u548c\u65b0\u54c1\u653e\u91cf\uff0c\u8f6c\u5316\u6548\u7387\u540c\u6b65\u6539\u5584\u3002',
                '\u4e0b\u4e00\u9636\u6bb5\u9700\u8981\u7ee7\u7eed\u653e\u5927\u9ad8\u8d21\u732e\u6e20\u9053\uff0c\u5e76\u538b\u7f29\u4f4e\u6548\u6295\u5165\u3002',
            ],
            [
                '\u672c\u671f\u6210\u679c\u80fd\u591f\u652f\u6491\u4e3b\u7ebf\u76ee\u6807\uff0c\u4f46\u8fd8\u9700\u7ee7\u7eed\u63d0\u5347\u4e24\u4e2a\u6838\u5fc3\u6307\u6807\u3002',
                '\u7ee7\u7eed\u5173\u6ce8\u9ad8\u8d28\u91cf\u6210\u957f\uff0c\u800c\u4e0d\u662f\u5355\u7eaf\u7684\u6d41\u91cf\u6269\u5f35\u3002',
                '\u5c06\u9891\u6b21\u8f83\u9ad8\u7684\u52a8\u4f5c\u653e\u5728\u524d\u9762\uff0c\u7528\u5468\u5ea6\u590d\u76d8\u63a7\u5236\u8282\u594f\u3002',
            ],
        ]
        return variants[variant % len(variants)]
    if re.search('\u7528\u6237|\u6d41\u91cf|\u7559\u5b58|\u8f6c\u5316|\u62c9\u65b0|\u4f7f\u7528', text):
        variants = [
            [
                '\u65b0\u589e\u7528\u6237\u4fdd\u6301\u4e0a\u5347\uff0c\u4f46\u7559\u5b58\u73af\u8282\u8fd8\u6709\u63d0\u5347\u7a7a\u95f4\u3002',
                '\u7ecf\u8fc7\u7684\u6d3b\u52a8\u62c9\u65b0\u6548\u679c\u660e\u663e\uff0c\u4f46\u9700\u8981\u628a\u7528\u6237\u8d28\u91cf\u505a\u4e0a\u53bb\u3002',
                '\u4e0b\u4e00\u6b65\u7684\u5173\u952e\u662f\u7a33\u4f4f\u9ad8\u6d3b\u8dc3\u7528\u6237\uff0c\u5e76\u63d0\u9ad8\u8f6c\u5316\u7387\u3002',
            ],
            [
                '\u6d41\u91cf\u589e\u957f\u4e0e\u8f6c\u5316\u6548\u7387\u53ef\u4ee5\u540c\u6b65\u63d0\u5347\uff0c\u4e0d\u53ea\u662f\u62c9\u65b0\u4e0a\u91cf\u3002',
                '\u9700\u8981\u5f3a\u5316\u4f1a\u5458\u89e6\u8fbe\u4e0e\u590d\u8bbf\u673a\u5236\uff0c\u63d0\u9ad8\u7ecf\u8425\u6210\u6548\u3002',
                '\u5c06\u6838\u5fc3\u76ee\u6807\u62c6\u4e3a\u4e09\u4e2a\u9636\u6bb5\uff0c\u6bcf\u5468\u8ffd\u8e2a\u4e00\u6b21\u3002',
            ],
        ]
        return variants[variant % len(variants)]
    if re.search('\u95ee\u9898|\u98ce\u9669|\u75db\u70b9|\u74f6\u9888|\u6311\u6218|\u9669', text):
        variants = [
            [
                '\u73b0\u6709\u95ee\u9898\u4e0d\u662f\u5355\u70b9\u5931\u6548\uff0c\u800c\u662f\u8282\u594f\u3001\u8d23\u4efb\u548c\u6807\u51c6\u4e0d\u591f\u6e05\u6670\u3002',
                '\u9996\u5148\u5904\u7406\u5bf9\u7ed3\u679c\u5f71\u54cd\u6700\u5927\u7684\u963b\u585e\u9879\u3002',
                '\u95ee\u9898\u8981\u5f80\u4e0b\u62c6\uff0c\u4f46\u7ed3\u8bba\u8981\u5f80\u4e0a\u6536\uff0c\u76f4\u63a5\u5bf9\u5e94\u884c\u52a8\u3002',
            ],
            [
                '\u6311\u6218\u4e0d\u662f\u8d44\u6e90\u4e0d\u591f\uff0c\u800c\u662f\u4f18\u5148\u7ea7\u6ca1\u6709\u6392\u597d\u3002',
                '\u5f53\u524d\u95ee\u9898\u7684\u6838\u5fc3\u5728\u4e8e\u8fc7\u7a0b\u4e0d\u8868\u9645\u5f52\u4e0e\u8d1f\u8d23\u4e0d\u6e05\u3002',
                '\u4e0b\u4e00\u6b65\u9700\u8981\u5212\u5206\u6210\u53ef\u6267\u884c\u7684\u6574\u6539\u4efb\u52a1\u3002',
            ],
        ]
        return variants[variant % len(variants)]
    if re.search('\u89c4\u5212|\u8ba1\u5212|\u884c\u52a8|\u4e0b\u6708|\u4e0b\u4e00\u6b65|\u76ee\u6807', text):
        variants = [
            [
                '\u4f18\u5148\u63a8\u8fdb\u4e09\u4e2a\u6700\u80fd\u76f4\u63a5\u5f71\u54cd\u7ed3\u679c\u7684\u52a8\u4f5c\u3002',
                '\u6bcf\u4e00\u9879\u4efb\u52a1\u90fd\u9700\u8981\u660e\u786e\u8d1f\u8d23\u4eba\uff0c\u8282\u70b9\u548c\u9a8c\u6536\u6807\u51c6\u3002',
                '\u7528\u5468\u5ea6\u8ffd\u8e2a\u4fdd\u8bc1\u8ba1\u5212\u4e0d\u53ea\u662f\u5217\u8868\uff0c\u800c\u662f\u771f\u6b63\u843d\u5730\u7684\u8fd0\u8425\u52a8\u4f5c\u3002',
            ],
            [
                '\u672c\u9875\u7684\u76ee\u6807\u662f\u628a\u8ba1\u5212\u53d8\u6210\u53ef\u6267\u884c\u7684\u884c\u52a8\u8868\u3002',
                '\u5148\u5b9a\u4f18\u5148\u7ea7\uff0c\u518d\u5b9a\u8d23\u4efb\u4eba\uff0c\u6700\u540e\u5b9a\u65f6\u95f4\u3002',
                '\u6240\u6709\u52a8\u4f5c\u90fd\u8981\u53ef\u88ab\u8ffd\u8e2a\uff0c\u53ef\u88ab\u590d\u76d8\uff0c\u53ef\u88ab\u9a8c\u6536\u3002',
            ],
        ]
        return variants[variant % len(variants)]

    fallback = [
        f'\u672c\u9875\u56f4\u7ed5{section}\u5c55\u5f00\uff0c\u5148\u7ed9\u7ed3\u8bba\u3002',
        '\u7528\u6570\u636e\u3001\u5bf9\u6bd4\u6216\u6848\u4f8b\u8bf4\u660e\u8fd9\u4ef6\u4e8b\u4e3a\u4ec0\u4e48\u91cd\u8981\u3002',
        '\u6700\u540e\u6c47\u603b\u4e3a\u4e00\u4e2a\u53ef\u6267\u884c\u7684\u4e0b\u4e00\u6b65\u3002',
    ]
    if variant % 2 == 1:
        return [
            '\u5148\u6307\u51fa\u6838\u5fc3\u53d1\u73b0\uff0c\u518d\u7ed9\u51fa\u6570\u636e\u652f\u6491\u3002',
            '\u9700\u8981\u8ba9\u5bf9\u65b9\u4e00\u773c\u770b\u51fa\u6b63\u5728\u53d1\u751f\u4ec0\u4e48\u3002',
            '\u4e0b\u4e00\u6b65\u8981\u76f4\u63a5\u6307\u5411\u884c\u52a8\u3002',
        ]
    return fallback

def _sample_metrics(section: str, data_profile: ExcelProfile | None, index: int) -> list[dict[str, str]]:
    if data_profile and data_profile.metric_columns:
        metrics = []
        for idx, column in enumerate(data_profile.metric_columns[:3]):
            metrics.append({
                'label': column,
                'value': f'{80 + (index + idx) % 17}',
                'trend': '来自Excel数据映射',
            })
        return metrics
    if re.search(r'业绩|营收|销售|经营|财务', section):
        return [
            {'label': '目标达成', 'value': f'{92 + index % 6}%', 'trend': '较上期 +6.2%'},
            {'label': '收入贡献', 'value': f'{180 + index * 8}万', 'trend': '核心区域拉动'},
            {'label': '转化效率', 'value': f'{18 + index % 4}.8%', 'trend': '渠道结构优化'},
        ]
    if re.search(r'用户|增长|流量|留存|转化', section):
        return [
            {'label': '新增用户', 'value': f'{12 + index % 5}.6万', 'trend': '活动引流增强'},
            {'label': '留存率', 'value': f'{34 + index % 4}.5%', 'trend': '会员触达改善'},
            {'label': '获客成本', 'value': f'{28 - index % 3}元', 'trend': '环比下降'},
        ]
    return [
        {'label': '关键项目', 'value': f'{4 + index % 4}项', 'trend': '聚焦落地'},
        {'label': '预计收益', 'value': f'{15 + index % 5}%', 'trend': '持续验证'},
        {'label': '优先级', 'value': 'P1', 'trend': '资源前置'},
    ]


def _choose_recommendation_candidates(page_catalog: pd.DataFrame | None, layout_type: str, industry: str, scenario: str, style: str) -> list[str]:
    if page_catalog is None or page_catalog.empty:
        return []

    layout_key = layout_type.replace('_slide', '')

    def _query(ind: tuple[str, ...], scen: tuple[str, ...], sty: tuple[str, ...]) -> pd.DataFrame:
        return filter_slide_dataframe(
            page_catalog,
            filters=type('F', (), {
                'query': '',
                'layout_types': (layout_key,),
                'slide_types': (),
                'industries': ind,
                'scenarios': scen,
                'styles': sty,
                'colors': (),
                'tags': (),
                'min_score': 0,
                'has_chart': None,
                'has_table': None,
                'has_timeline': None,
                'has_process': None,
                'has_map': None,
                'has_people': None,
                'has_infographic': None,
                'dark_or_light': '',
                'min_quality': 0,
            })(),
        )

    ordered_frames = [
        _query((industry,) if industry else (), (scenario,) if scenario else (), (style,) if style else ()),
        _query((), (scenario,) if scenario else (), (style,) if style else ()),
        _query((), (), (style,) if style else ()),
        _query((), (), ()),
    ]
    collected: list[str] = []

    def _extend(frame: pd.DataFrame) -> None:
        if frame.empty:
            return
        ordered = frame.sort_values(['overall_quality_score', 'design_score', 'layout_score', 'slide_id'], ascending=[False, False, False, True])
        for slide_id in ordered['slide_id'].tolist():
            if slide_id not in collected:
                collected.append(slide_id)
            if len(collected) >= 5:
                return

    for frame in ordered_frames:
        _extend(frame)
        if len(collected) >= 5:
            return collected[:5]

    fallback = page_catalog[page_catalog['layout_type'].astype(str).str.lower().eq(layout_key.lower())]
    if fallback.empty:
        fallback = page_catalog
    _extend(fallback)
    return collected[:5]


def _infer_data_layouts(profile: ExcelProfile | None) -> list[str]:
    layouts = []
    if not profile:
        return layouts
    if profile.has_trend:
        layouts.append('chart_slide')
    if profile.metric_columns:
        layouts.append('kpi_slide')
        layouts.append('dashboard_slide')
    if profile.has_table:
        layouts.append('table_slide')
    if profile.has_map:
        layouts.append('map_slide')
    if profile.has_comparison:
        layouts.append('comparison_slide')
    return list(dict.fromkeys(layouts))


def build_deck_plan(
    prompt: str,
    style_name: str,
    page_range: tuple[int, int] = (8, 14),
    excel_df: pd.DataFrame | None = None,
    page_catalog: pd.DataFrame | None = None,
) -> dict[str, Any]:
    title = infer_title(prompt)
    sections = infer_sections(prompt)
    data_profile = analyze_dataframe_profile(excel_df)
    industry = infer_industry(prompt)
    scenario = infer_scenario(prompt)
    industry_label = _human_label(industry, INDUSTRY_LABELS)
    scenario_label = _human_label(scenario, SCENARIO_LABELS)
    normalized_style, color_token, tone = infer_style(prompt, style_name)
    style_family = _style_family(normalized_style)
    design_system = build_design_system(prompt, normalized_style, data_profile)
    min_pages, max_pages = page_range
    base_total = len(sections) + 3
    target_total = max(base_total, min(max_pages, max(min_pages, (min_pages + max_pages) // 2)))
    data_layouts = _infer_data_layouts(data_profile)
    style_layouts = STYLE_LAYOUT_BIASES.get(style_family, [])

    pages: list[dict[str, Any]] = []
    # cover and agenda
    pages.append({
        'page_no': 1,
        'layout_type': 'cover_slide',
        'section': '封面',
        'title': title,
        'subtitle': f'{normalized_style} - {scenario_label} - {date.today().isoformat()}',
        'bullets': [],
        'metrics': [],
        'tags': _page_tags('封面', 'cover_slide', industry, scenario, normalized_style),
        'recommended_slide_ids': _choose_recommendation_candidates(page_catalog, 'cover_slide', industry, scenario, normalized_style),
    })
    pages.append({
        'page_no': 2,
        'layout_type': 'agenda_slide',
        'section': '目录',
        'title': '目录',
        'subtitle': 'Content Map',
        'bullets': sections,
        'metrics': [],
        'tags': _page_tags('目录', 'agenda_slide', industry, scenario, normalized_style),
        'recommended_slide_ids': _choose_recommendation_candidates(page_catalog, 'agenda_slide', industry, scenario, normalized_style),
    })

    content_slots = max(0, target_total - 3)
    combined_sections = sections[:]
    if data_profile and data_profile.suggested_visuals:
        combined_sections = combined_sections + [f'数据页：{visual}' for visual in data_profile.suggested_visuals[:2]]
    if not combined_sections:
        combined_sections = ['背景与目标', '核心进展', '问题分析', '行动计划']

    for index in range(content_slots):
        section = combined_sections[index % len(combined_sections)]
        layout_type = _pick_layout_for_section(section, data_profile)
        if layout_type == 'content_slide' and style_layouts:
            layout_type = style_layouts[index % len(style_layouts)]
        if index < len(data_layouts) and layout_type == 'content_slide':
            layout_type = data_layouts[index]
        page_number = index + 3
        page = {
            'page_no': page_number,
            'layout_type': layout_type,
            'section': section,
            'title': section if len(section) <= 16 else section[:16],
            'subtitle': f'{industry_label} - {scenario_label}',
            'bullets': _sample_bullets(section, title, index),
            'metrics': _sample_metrics(section, data_profile, index),
            'tags': _page_tags(section, layout_type, industry, scenario, normalized_style),
            'recommended_slide_ids': _choose_recommendation_candidates(page_catalog, layout_type, industry, scenario, normalized_style),
        }
        pages.append(page)

    pages.append({
        'page_no': len(pages) + 1,
        'layout_type': 'ending_slide',
        'section': '结尾',
        'title': '感谢聆听',
        'subtitle': 'Thank You',
        'bullets': ['总结核心结论', '确认下一步行动', '开放问题交流'],
        'metrics': [],
        'tags': _page_tags('结尾', 'ending_slide', industry, scenario, normalized_style),
        'recommended_slide_ids': _choose_recommendation_candidates(page_catalog, 'ending_slide', industry, scenario, normalized_style),
    })

    page_count = len(pages)
    pages = pages[:max_pages] if page_count > max_pages else pages
    summary = {
        'total_pages': len(pages),
        'layout_types': len({page['layout_type'] for page in pages}),
        'industry': industry_label,
        'scenario': scenario_label,
        'style': normalized_style,
        'data_profile': data_profile.to_row() if data_profile else None,
        'page_range': list(page_range),
    }
    return {
        'title': title,
        'subtitle': f'{normalized_style} - {scenario_label} - {date.today().isoformat()}',
        'prompt': prompt,
        'style': normalized_style,
        'design_system': design_system.to_row(),
        'industry': industry,
        'scenario': scenario,
        'page_range': list(page_range),
        'sections': sections,
        'pages': pages,
        'total_pages': len(pages),
        'summary': summary,
        'data_profile': data_profile.to_row() if data_profile else None,
    }


def build_excel_aware_plan(
    prompt: str,
    style_name: str,
    excel_df: pd.DataFrame,
    page_catalog: pd.DataFrame | None = None,
    page_range: tuple[int, int] = (8, 14),
) -> dict[str, Any]:
    return build_deck_plan(prompt, style_name, page_range=page_range, excel_df=excel_df, page_catalog=page_catalog)


def plan_from_request(request: AssemblyRequest, page_catalog: pd.DataFrame | None = None) -> dict[str, Any]:
    excel_df = None
    if request.data_profile is not None:
        excel_df = pd.DataFrame([request.data_profile.to_row()])
    return build_deck_plan(
        prompt=request.prompt,
        style_name=request.style,
        page_range=request.page_range,
        excel_df=excel_df,
        page_catalog=page_catalog,
    )


# Safe overrides appended after the initial draft.
# These avoid regex pitfalls and support both Chinese and English prompts.

SAFE_STYLE_THEME_MAP = {
    '\u7b80\u7ea6\u5546\u52a1': ('business', 'blue', 'light'),
    '\u5b66\u672f\u6c47\u62a5': ('academic', 'purple', 'light'),
    '\u521b\u610f\u6f14\u793a': ('creative', 'orange', 'light'),
    '\u79d1\u6280\u84dd': ('technology', 'teal', 'light'),
    '\u91d1\u878d\u7eff': ('finance', 'green', 'light'),
    '\u653f\u52a1\u7ea2': ('government', 'red', 'light'),
    '\u6df1\u8272\u9ad8\u7ea7': ('dark', 'black_gold', 'dark'),
    '\u6781\u7b80\u767d': ('minimal', 'gray', 'light'),
    '\u5546\u52a1\u84dd': ('business', 'blue', 'light'),
    '\u79d1\u6280\u84dd\u7eff': ('technology', 'teal', 'light'),
    '\u653f\u52a1\u84d7': ('government', 'red', 'light'),
    '\u81ea\u7136\u6e05\u65b0': ('minimal', 'gray', 'light'),
    'business': ('business', 'blue', 'light'),
    'academic': ('academic', 'purple', 'light'),
    'creative': ('creative', 'orange', 'light'),
    'technology': ('technology', 'teal', 'light'),
    'finance': ('finance', 'green', 'light'),
    'government': ('government', 'red', 'light'),
    'minimal': ('minimal', 'gray', 'light'),
    'dark': ('dark', 'black_gold', 'dark'),
}

def _match_terms(text: str, terms: tuple[str, ...]) -> bool:
    lowered = _normalize(text).lower()
    return any(term.lower() in lowered for term in terms)


def infer_industry(text: str) -> str:  # type: ignore[override]
    rules = [
        (("finance", "bank", "insurance", "investment", "security", "\u91d1\u878d", "\u94f6\u884c", "\u4fdd\u9669", "\u6295\u8d44"), 'finance'),
        (("automotive", "auto", "vehicle", "car", "energy", "renewable", "\u65b0\u80fd\u6e90", "\u6c7d\u8f66", "\u5149\u4f0f", "\u7535\u529b", "\u80fd\u6e90"), 'new_energy'),
        (("internet", "software", "ai", "cloud", "security", "technology", "\u4e92\u8054\u7f51", "\u8f6f\u4ef6", "\u4eba\u5de5\u667a\u80fd", "\u5927\u6570\u636e", "\u4e91\u8ba1\u7b97", "\u7f51\u7edc\u5b89\u5168"), 'technology'),
        (("medical", "healthcare", "pharma", "biotech", "\u533b\u7597", "\u533b\u836f", "\u751f\u7269", "\u5065\u5eb7"), 'medical'),
        (("education", "school", "university", "research", "\u6559\u80b2", "\u5b66\u6821", "\u5927\u5b66", "\u79d1\u7814", "\u5b66\u672f"), 'education'),
        (("government", "public", "official", "\u653f\u5e9c", "\u515a\u5efa", "\u516c\u5171", "\u4e8b\u4e1a\u5355\u4f4d"), 'government'),
        (("real estate", "construction", "engineering", "\u623f\u5730\u4ea7", "\u5efa\u7b51", "\u5de5\u7a0b"), 'construction'),
        (("manufacturing", "industrial", "machinery", "\u5236\u9020", "\u5de5\u4e1a", "\u673a\u68b0"), 'manufacturing'),
        (("logistics", "supply chain", "transport", "aviation", "\u7269\u6d41", "\u4f9b\u5e94\u94fe", "\u4ea4\u901a", "\u822a\u7a7a"), 'logistics'),
        (("retail", "e-commerce", "commerce", "marketing", "advertising", "media", "\u96f6\u552e", "\u7535\u5546", "\u8425\u9500", "\u5e7f\u544a", "\u5a92\u4f53"), 'commerce'),
        (("travel", "hotel", "restaurant", "food", "tourism", "\u65c5\u6e38", "\u9152\u5e97", "\u9910\u996e", "\u98df\u54c1"), 'service'),
    ]
    for terms, industry in rules:
        if _match_terms(text, terms):
            return industry
    return 'general'


def infer_scenario(text: str) -> str:  # type: ignore[override]
    rules = [
        (("annual report", "annual summary", "yearly report", "\u5e74\u62a5", "\u5e74\u5ea6\u603b\u7ed3", "\u5e74\u5ea6"), 'annual_report'),
        (("quarterly", "Q1", "Q2", "Q3", "Q4", "\u5b63\u5ea6"), 'quarterly_report'),
        (("monthly", "\u6708\u62a5", "\u6708\u5ea6"), 'monthly_report'),
        (("weekly", "\u5468\u62a5"), 'weekly_report'),
        (("daily", "\u65e5\u62a5"), 'daily_report'),
        (("project report", "project review", "\u9879\u76ee\u6c47\u62a5", "\u9879\u76ee\u603b\u7ed3"), 'project_report'),
        (("business plan", "\u5546\u4e1a\u8ba1\u5212", "\u5546\u4e1a\u8ba1\u5212\u4e66"), 'business_plan'),
        (("pitch deck", "fundraising", "BP", "\u878d\u8d44", "\u8def\u6f14"), 'pitch_deck'),
        (("product launch", "product intro", "\u4ea7\u54c1\u53d1\u5e03", "\u65b0\u54c1"), 'product_launch'),
        (("market analysis", "industry analysis", "competitive analysis", "\u5e02\u573a\u5206\u6790", "\u884c\u4e1a\u5206\u6790", "\u7ade\u54c1", "\u7ade\u4e89\u5206\u6790"), 'market_analysis'),
        (("SWOT", "PEST", "\u6ce2\u7279", "strategy analysis", "strategic analysis", "\u6218\u7565"), 'strategy_analysis'),
        (("training", "courseware", "teaching", "\u57f9\u8bad", "\u8bfe\u4ef6", "\u6559\u5b66"), 'training'),
        (("defense", "thesis", "dissertation", "\u7b54\u8fa9", "\u8bba\u6587"), 'academic_defense'),
        (("resume", "portfolio", "photography", "travel", "wedding", "birthday", "\u7b80\u5386", "\u4f5c\u54c1\u96c6", "\u6444\u5f71", "\u65c5\u884c", "\u5a5a\u793c", "\u751f\u65e5"), 'personal_showcase'),
    ]
    for terms, scenario in rules:
        if _match_terms(text, terms):
            return scenario
    return 'presentation'


def infer_sections(prompt: str) -> list[str]:  # type: ignore[override]
    text = _normalize(prompt)
    match = re.search(r'(?:包含|包括|分为|围绕|cover|include|covering)(.+?)(?:\d+\s*个?章节|章节|部分|模块|section|sections|parts|$)', text, re.IGNORECASE)
    if match:
        raw = match.group(1)
        sections = [item.strip(' 。,，;；:：') for item in re.split(r'[、,，;；/]+', raw) if item.strip(' 。,，;；:：')]
        sections = [item for item in sections if item and len(item) <= 20]
        if sections:
            return sections
    keyword_sections = []
    keyword_map = [
        (("业绩", "营收", "销售", "经营", "财务", "sales", "revenue", "performance"), '业绩概况'),
        (("用户", "增长", "流量", "留存", "转化", "user", "growth", "traffic", "retention", "conversion"), '用户增长'),
        (("问题", "风险", "挑战", "瓶颈", "issue", "risk", "challenge", "problem"), '问题分析'),
        (("规划", "计划", "策略", "下一步", "下月", "下季度", "plan", "strategy", "next step", "roadmap"), '行动计划'),
        (("市场", "行业", "竞品", "竞争", "market", "industry", "competitive"), '市场洞察'),
        (("产品", "功能", "研发", "技术", "product", "feature", "development", "technology"), '产品进展'),
        (("团队", "组织", "人员", "角色", "team", "organization", "people", "role"), '团队协同'),
        (("数据", "指标", "仪表盘", "KPI", "data", "metric", "dashboard"), '数据洞察'),
        (("时间", "历程", "里程碑", "路线图", "timeline", "milestone", "roadmap"), '发展历程'),
        (("地图", "地区", "区域", "分布", "map", "region", "distribution"), '区域分布'),
    ]
    for terms, label in keyword_map:
        if _match_terms(text, terms) and label not in keyword_sections:
            keyword_sections.append(label)
    return keyword_sections or ['背景与目标', '核心进展', '问题分析', '行动计划']


def infer_style(prompt: str, style_name: str) -> tuple[str, str, str]:  # type: ignore[override]
    text = _normalize(prompt)
    for pattern, color in COLOR_PATTERNS:
        if re.search(pattern, text, re.IGNORECASE):
            return style_name or '\u7b80\u7ea6\u5546\u52a1', color, 'light' if color != 'dark' else 'dark'
    if style_name in SAFE_STYLE_THEME_MAP:
        mapped = SAFE_STYLE_THEME_MAP[style_name]
        return style_name, mapped[1], mapped[2]
    if style_name.lower() in SAFE_STYLE_THEME_MAP:
        mapped = SAFE_STYLE_THEME_MAP[style_name.lower()]
        return style_name, mapped[1], mapped[2]
    return style_name or '\u7b80\u7ea6\u5546\u52a1', 'blue', 'light'


def _pick_layout_for_section(section: str, data_profile: ExcelProfile | None) -> str:  # type: ignore[override]
    text = (section or '').lower()
    keyword_layouts = [
        (("封面", "cover", "title", "hero"), 'cover_slide'),
        (("目录", "agenda", "议程"), 'agenda_slide'),
        (("章节", "section", "过渡", "divider"), 'section_slide'),
        (("kpi", "指标", "dashboard", "metric", "看板"), 'kpi_slide'),
        (("数据", "业绩", "财务", "销售", "营收", "运营", "增长", "data", "analysis", "report"), 'data_analysis_slide'),
        (("图表", "chart", "柱状", "折线", "饼图", "雷达", "散点"), 'chart_slide'),
        (("表格", "table", "列表", "清单"), 'table_slide'),
        (("时间轴", "timeline", "历程", "里程碑", "路线图", "roadmap", "milestone"), 'timeline_slide'),
        (("流程", "process", "步骤", "审批", "workflow"), 'process_slide'),
        (("对比", "comparison", "竞品", "优劣", "before", "after"), 'comparison_slide'),
        (("组织", "org", "团队", "team", "人员", "people", "关系"), 'relation_slide'),
        (("swot", "pest", "波特", "战略", "strategy", "定位", "商业模式"), 'strategy_slide'),
        (("计划", "plan", "甘特", "gantt", "任务", "进度", "排期"), 'planning_slide'),
        (("产品", "product", "功能", "特性", "参数", "发布"), 'product_slide'),
        (("地图", "map", "区域", "region", "国家", "分布"), 'map_slide'),
        (("图片", "gallery", "案例", "portfolio", "摄影", "画册", "展示"), 'gallery_slide'),
        (("结束", "ending", "thanks", "q&a", "联系方式", "总结"), 'ending_slide'),
    ]
    for terms, layout in keyword_layouts:
        if _match_terms(text, terms):
            return layout
    if data_profile:
        if data_profile.has_map and 'map' not in text and '地图' not in text:
            return 'map_slide'
        if data_profile.has_trend and 'timeline' not in text and '时间' not in text:
            return 'chart_slide'
        if data_profile.metric_columns:
            return 'data_analysis_slide'
    return 'content_slide'

def infer_title(prompt: str) -> str:  # type: ignore[override]
    text = _normalize(prompt)
    head = re.split(r'包含|包括|分为|围绕|需要|要求|covering|include|includes|with|for|about|on', text, maxsplit=1, flags=re.IGNORECASE)[0]
    head = re.sub(r'^(?:请帮我|帮我|帮忙|麻烦|生成一份|生成一个|制作一份|制作一个|做一份|做一个|生成|制作|创建|输出)?(?:一份|一个|1份)?', '', head).strip()
    head = re.sub(r'^(please\s+)?(create|make|generate|prepare|build|design|produce|draft)\s+', '', head, flags=re.IGNORECASE).strip()
    head = re.sub(r'\b(pptx?|slides?|presentation|deck|report)\b$', '', head, flags=re.IGNORECASE).strip(' -—,，.。')
    head = re.sub(r'\s+', ' ', head).strip()
    if not head:
        head = '\u667a\u80fd\u6c47\u62a5'
    if ('\u6c47\u62a5' not in head and '\u62a5\u544a' not in head and 'report' not in head.lower() and 'deck' not in head.lower() and '\u590d\u76d8' not in head and '\u65b9\u6848' not in head and '\u5206\u6790' not in head and '\u8ba1\u5212' not in head):
        head += '\u6c47\u62a5'
    return head
def infer_scenario(text: str) -> str:  # type: ignore[override]
    rules = [
        (("annual sales analysis report", "sales analysis report", "analysis report", "sales report", "annual report", "annual summary", "yearly report", "\u5e74\u62a5", "\u5e74\u5ea6\u603b\u7ed3", "\u5e74\u5ea6"), 'annual_report'),
        (("quarterly", "Q1", "Q2", "Q3", "Q4", "\u5b63\u5ea6"), 'quarterly_report'),
        (("monthly", "\u6708\u62a5", "\u6708\u5ea6"), 'monthly_report'),
        (("weekly", "\u5468\u62a5"), 'weekly_report'),
        (("daily", "\u65e5\u62a5"), 'daily_report'),
        (("project report", "project review", "\u9879\u76ee\u6c47\u62a5", "\u9879\u76ee\u603b\u7ed3"), 'project_report'),
        (("business plan", "\u5546\u4e1a\u8ba1\u5212", "\u5546\u4e1a\u8ba1\u5212\u4e66"), 'business_plan'),
        (("pitch deck", "fundraising", "BP", "\u878d\u8d44", "\u8def\u6f14"), 'pitch_deck'),
        (("product launch", "product intro", "\u4ea7\u54c1\u53d1\u5e03", "\u65b0\u54c1"), 'product_launch'),
        (("market analysis", "industry analysis", "competitive analysis", "\u5e02\u573a\u5206\u6790", "\u884c\u4e1a\u5206\u6790", "\u7ade\u54c1", "\u7ade\u4e89\u5206\u6790"), 'market_analysis'),
        (("SWOT", "PEST", "\u6ce2\u7279", "strategy analysis", "strategic analysis", "\u6218\u7565"), 'strategy_analysis'),
        (("training", "courseware", "teaching", "\u57f9\u8bad", "\u8bfe\u4ef6", "\u6559\u5b66"), 'training'),
        (("defense", "thesis", "dissertation", "\u7b54\u8fa9", "\u8bba\u6587"), 'academic_defense'),
        (("resume", "portfolio", "photography", "travel", "wedding", "birthday", "\u7b80\u5386", "\u4f5c\u54c1\u96c6", "\u6444\u5f71", "\u65c5\u884c", "\u5a5a\u793c", "\u751f\u65e5"), 'personal_showcase'),
    ]
    return _infer_scenario_base(text, rules)


def _infer_scenario_base(text: str, rules: list[tuple[tuple[str, ...], str]]) -> str:
    for terms, scenario in rules:
        if _match_terms(text, terms):
            return scenario
    return 'presentation'
