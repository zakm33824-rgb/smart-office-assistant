from __future__ import annotations

import pandas as pd

from .page_models import PageSearchFilters


def _contains_any(series: pd.Series, tokens: tuple[str, ...]) -> pd.Series:
    if not tokens:
        return pd.Series(True, index=series.index)
    mask = pd.Series(False, index=series.index)
    lowered = series.astype(str).str.lower()
    for token in tokens:
        mask |= lowered.str.contains(token.lower().replace(' ', '_'), na=False)
    return mask


def filter_slide_dataframe(df: pd.DataFrame, filters: PageSearchFilters) -> pd.DataFrame:
    result = df.copy()
    if filters.query:
        terms = [term for term in filters.query.lower().split() if term]
        if terms:
            text_columns = [
                'slide_id', 'source_template_id', 'source_file', 'source_url',
                'slide_type', 'slide_subtype', 'industry', 'scenario', 'style', 'layout_type',
                'metadata_json', 'tags_json',
            ]
            mask = pd.Series(True, index=result.index)
            for term in terms:
                term_mask = pd.Series(False, index=result.index)
                for column in text_columns:
                    term_mask |= result[column].astype(str).str.lower().str.contains(term, na=False)
                mask &= term_mask
            result = result[mask]
    if filters.layout_types:
        result = result[result['layout_type'].astype(str).str.lower().isin([item.lower() for item in filters.layout_types])]
    if filters.slide_types:
        result = result[result['slide_type'].astype(str).str.lower().isin([item.lower() for item in filters.slide_types])]
    if filters.industries:
        result = result[result['industry'].astype(str).str.lower().isin([item.lower() for item in filters.industries])]
    if filters.scenarios:
        result = result[result['scenario'].astype(str).str.lower().isin([item.lower() for item in filters.scenarios])]
    if filters.styles:
        result = result[result['style'].astype(str).str.lower().isin([item.lower() for item in filters.styles])]
    if filters.colors:
        color_columns = ['primary_color', 'secondary_color', 'background_color']
        mask = pd.Series(False, index=result.index)
        for column in color_columns:
            mask |= _contains_any(result[column], filters.colors)
        result = result[mask]
    if filters.tags:
        result = result[result['tags_json'].astype(str).str.lower().apply(lambda value: all(tag.lower() in value for tag in filters.tags))]
    if filters.has_chart is not None:
        result = result[result['has_chart'].astype(bool).eq(filters.has_chart)]
    if filters.has_table is not None:
        result = result[result['has_table'].astype(bool).eq(filters.has_table)]
    if filters.has_timeline is not None:
        result = result[result['has_timeline'].astype(bool).eq(filters.has_timeline)]
    if filters.has_process is not None:
        result = result[result['has_process'].astype(bool).eq(filters.has_process)]
    if filters.has_map is not None:
        result = result[result['has_map'].astype(bool).eq(filters.has_map)]
    if filters.has_people is not None:
        result = result[result['has_people'].astype(bool).eq(filters.has_people)]
    if filters.has_infographic is not None:
        result = result[result['has_infographic'].astype(bool).eq(filters.has_infographic)]
    if filters.dark_or_light:
        result = result[result['dark_or_light'].astype(str).str.lower().eq(filters.dark_or_light.lower())]
    result = result[result['overall_quality_score'] >= filters.min_quality]
    result = result[result['overall_quality_score'] >= filters.min_score]
    return result.sort_values(['overall_quality_score', 'design_score', 'layout_score', 'slide_id'], ascending=[False, False, False, True])


def search_slide_dataframe(df: pd.DataFrame, filters: PageSearchFilters, limit: int = 200) -> pd.DataFrame:
    result = filter_slide_dataframe(df, filters)
    return result.head(limit)
