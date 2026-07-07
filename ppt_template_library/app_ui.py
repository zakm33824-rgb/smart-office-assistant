from __future__ import annotations

import json
import tempfile
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

from .catalog import build_library_summary, build_source_catalog, filter_source_dataframe, save_source_catalog_json, source_dataframe
from .models import SearchFilters
from .page_engine import build_library_overview, build_page_level_outline, build_page_preview_sheet, extract_pptx_to_slide_catalog, search_components, search_pages, seed_demo_page_library
from .paths import ensure_library_structure
from .slide_storage import initialize_slide_storage, load_component_catalog, load_slide_catalog, load_slide_summary
from .storage import load_source_catalog, save_source_catalog


def _metric_row(summary: dict[str, int | float]) -> None:
    cols = st.columns(4)
    cols[0].metric('来源端点', summary.get('source_count', 0))
    cols[1].metric('仓库来源', summary.get('repository_count', 0))
    cols[2].metric('站点来源', summary.get('site_count', 0))
    cols[3].metric('主题索引', summary.get('topic_count', 0))


def _table_preview(df: pd.DataFrame, columns: list[str], max_rows: int = 100) -> None:
    if df.empty:
        st.info('没有可显示的数据。')
        return
    view_cols = [col for col in columns if col in df.columns]
    st.dataframe(df[view_cols].head(max_rows), use_container_width=True, hide_index=True)


def _page_table_preview(df: pd.DataFrame) -> None:
    if df.empty:
        st.info('当前没有页面记录。')
        return
    cols = [c for c in ['overall_quality_score', 'slide_id', 'slide_type', 'slide_subtype', 'industry', 'scenario', 'style', 'layout_type', 'preview_image'] if c in df.columns]
    st.dataframe(df[cols].head(80), use_container_width=True, hide_index=True)


def _component_table_preview(df: pd.DataFrame) -> None:
    if df.empty:
        st.info('当前没有组件记录。')
        return
    cols = [c for c in ['component_id', 'component_type', 'component_subtype', 'source_slide_id', 'layout_type', 'style_token', 'preview_image'] if c in df.columns]
    st.dataframe(df[cols].head(80), use_container_width=True, hide_index=True)


def _page_filters(df: pd.DataFrame) -> SearchFilters:
    filters = SearchFilters()
    c1, c2, c3 = st.columns(3)
    with c1:
        filters.query = st.text_input('关键词', placeholder='例如：年度报告 / 科技蓝 / 时间轴')
        filters.min_score = int(st.slider('最低分', 0, 100, 70))
    with c2:
        filters.slide_types = tuple(st.multiselect('页面类型', sorted(df['slide_type'].dropna().unique().tolist())))
        filters.layout_types = tuple(st.multiselect('布局类型', sorted(df['layout_type'].dropna().unique().tolist())))
    with c3:
        filters.industries = tuple(st.multiselect('行业', sorted(df['industry'].dropna().unique().tolist())))
        filters.scenarios = tuple(st.multiselect('场景', sorted(df['scenario'].dropna().unique().tolist())))
    filters.styles = tuple(st.multiselect('风格', sorted(df['style'].dropna().unique().tolist())))
    return filters


def _component_filters(df: pd.DataFrame) -> tuple[str, tuple[str, ...]]:
    c1, c2 = st.columns([2, 1])
    with c1:
        query = st.text_input('组件关键词', placeholder='例如：标题 / KPI / 图表 / 时间轴')
    with c2:
        types = tuple(st.multiselect('组件类型', sorted(df['component_type'].dropna().unique().tolist())))
    return query, types


def _save_uploaded_pptx(uploaded_file: Any) -> Path:
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pptx')
    tmp.write(uploaded_file.getvalue())
    tmp.close()
    return Path(tmp.name)


def _render_source_tab() -> None:
    summary = build_library_summary()
    _metric_row(summary)
    st.divider()
    source_df = load_source_catalog()
    if source_df.empty:
        source_df = source_dataframe()
    with st.expander('筛选条件', expanded=True):
        filters = SearchFilters()
        c1, c2, c3 = st.columns(3)
        with c1:
            filters.query = st.text_input('关键词', placeholder='例如：财务 / 演示文稿 / 大学 / 商业')
            filters.min_score = int(st.slider('最低评分', 0, 100, 70, key='source_min_score'))
        with c2:
            filters.source_types = tuple(st.multiselect('来源类型', sorted(source_df['source_type'].dropna().unique().tolist())))
            filters.licenses = tuple(st.multiselect('许可证', sorted(source_df['license_hint'].dropna().unique().tolist())))
        with c3:
            filters.languages = tuple(st.multiselect('语言', sorted(source_df['language'].dropna().unique().tolist())))
            filters.commercial_only = st.checkbox('仅看可商用', value=False)
    filtered = filter_source_dataframe(source_df, filters)
    st.write(f'匹配来源：{len(filtered)}')
    _table_preview(filtered, ['score', 'name', 'source_type', 'license_hint', 'commercial_use_hint', 'programmatic_access', 'source_website', 'source_url'])
    st.code('模板库/\n  资源/\n  组件/\n  预览/\n  元数据/\n  数据库/\n  可商用/\n  个人使用/\n  许可待确认/\n  高质量/\n  行业/\n  场景/\n  风格/\n  色彩/\n  布局/\n  日志/\n  来源记录/', language='text')


def _render_page_tab() -> None:
    seed_demo_page_library()
    summary = load_slide_summary()
    cols = st.columns(5)
    cols[0].metric('页面数', summary.get('slide_count', 0))
    cols[1].metric('组件数', summary.get('component_count', 0))
    cols[2].metric('页面类型', summary.get('layout_types', 0))
    cols[3].metric('高质量页', summary.get('premium_count', 0))
    cols[4].metric('平均分', summary.get('average_score', 0.0))
    st.divider()

    left, right = st.columns([1.1, 1])
    with left:
        refresh = st.button('刷新示例页面库', use_container_width=True)
        uploaded = st.file_uploader('上传演示文稿解析为页面库', type=['pptx'])
        if uploaded is not None and st.button('解析并入库', type='primary', use_container_width=True):
            try:
                with st.spinner('正在解析演示文稿页面...'):
                    tmp_path = _save_uploaded_pptx(uploaded)
                    result = extract_pptx_to_slide_catalog(tmp_path, source_template_id=Path(uploaded.name).stem, source_url=f'file://{uploaded.name}')
                st.success(f"已解析 {result.get('ingested', 0)} 页")
                st.rerun()
            except Exception as exc:  # noqa: BLE001
                st.error(f'解析失败：{exc}')
        if refresh and st.button('重新生成示例库', type='secondary', use_container_width=True):
            seed_demo_page_library(refresh=True)
            st.rerun()
    with right:
        st.info('页面库支持页级检索、组件检索、演示文稿解析和智能组装计划。')

    page_df = load_slide_catalog()
    if page_df.empty:
        page_df = pd.DataFrame()
    st.subheader('页面检索')
    page_filters = _page_filters(page_df if not page_df.empty else pd.DataFrame({'slide_type': [], 'layout_type': [], 'industry': [], 'scenario': [], 'style': []})) if not page_df.empty else SearchFilters()
    if page_df.empty:
        st.info('页面库暂无数据。')
    else:
        filtered = search_pages(page_filters)
        st.write(f'匹配页面：{len(filtered)}')
        _page_table_preview(filtered)
        if not filtered.empty:
            try:
                sheet = build_page_preview_sheet(filtered, Path('ppt_template_library') / 'preview' / 'page_contact_sheet.png', columns=3, limit=9)
                st.image(str(sheet), caption='页面预览缩略图', use_container_width=True)
            except Exception:
                pass

    st.subheader('页面结构摘要')
    st.json(build_library_overview())


def _render_component_tab() -> None:
    seed_demo_page_library()
    comp_df = load_component_catalog()
    cols = st.columns(4)
    cols[0].metric('组件数', len(comp_df))
    cols[1].metric('组件类型', int(comp_df['component_type'].nunique()) if not comp_df.empty else 0)
    cols[2].metric('来源页', int(comp_df['source_slide_id'].nunique()) if not comp_df.empty else 0)
    cols[3].metric('预览图', int(comp_df['preview_image'].astype(str).str.len().gt(0).sum()) if not comp_df.empty else 0)
    st.divider()
    if comp_df.empty:
        st.info('当前没有组件数据。')
        return
    query, types = _component_filters(comp_df)
    filtered = search_components(query=query, component_types=types)
    st.write(f'匹配组件：{len(filtered)}')
    _component_table_preview(filtered)


def _render_assembly_tab() -> None:
    seed_demo_page_library()
    st.subheader('智能组装方案')
    prompt = st.text_area('需求描述', value='生成一份新能源汽车企业年度经营分析报告，科技蓝风格，需要数据分析、时间轴和总结页。', height=120)
    c1, c2 = st.columns([2, 1])
    with c1:
        style = st.selectbox('风格', ['简约商务', '学术汇报', '创意演示', '科技蓝', '金融绿', '政务红', '极简白', '黑金高端'])
        page_range = st.slider('页数范围', 5, 20, (8, 12))
    with c2:
        st.write('')
        st.write('')
        build = st.button('生成页面级方案', type='primary', use_container_width=True)
    if build:
        try:
            with st.spinner('正在生成页面级组装方案...'):
                outline = build_page_level_outline(prompt, style, page_range)
            st.session_state['page_outline'] = outline
            st.success('方案已生成')
        except Exception as exc:  # noqa: BLE001
            st.error(f'生成失败：{exc}')
    outline = st.session_state.get('page_outline')
    if outline:
        st.markdown(f"**{outline['title']}**  ·  {outline['subtitle']}")
        st.json(outline['design_system'])
        st.dataframe(pd.DataFrame(outline['page_plan']), use_container_width=True, hide_index=True)
        st.dataframe(pd.DataFrame(outline['pages']), use_container_width=True, hide_index=True)
        st.download_button('下载方案文件', data=json.dumps(outline, ensure_ascii=False, indent=2).encode('utf-8'), file_name='page_assembly_outline.json', mime='application/json', use_container_width=True)


def render_template_library_page() -> None:
    st.header('演示文稿模板资源库')
    st.caption('页面级模板库、组件级素材库、智能页面检索和跨模板组装计划。')
    ensure_library_structure()
    initialize_slide_storage()
    tabs = st.tabs(['来源总览', '页面级模板库', '组件级素材库', '智能组装'])
    with tabs[0]:
        _render_source_tab()
    with tabs[1]:
        _render_page_tab()
    with tabs[2]:
        _render_component_tab()
    with tabs[3]:
        _render_assembly_tab()

