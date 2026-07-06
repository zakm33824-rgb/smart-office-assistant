"""轻量化智能办公 Web 小程序。

运行方式：
    streamlit run app.py
"""

from __future__ import annotations

import io
import re
import tempfile
from datetime import date
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


APP_TITLE = "轻量化智能办公助手"
DEFAULT_PPT_PROMPT = (
    "生成一份Q3产品运营复盘PPT，包含业绩概况、用户增长、问题分析、下月规划4个章节"
)
DEFAULT_TABLE_PROMPT = (
    "生成2024年各门店营收表，包含门店名、月度营收、同比增速、排名4列，填充合理示例数据"
)


PPT_STYLES: dict[str, dict[str, Any]] = {
    "简约商务": {
        "bg": RGBColor(247, 249, 252),
        "panel": RGBColor(255, 255, 255),
        "primary": RGBColor(32, 83, 117),
        "accent": RGBColor(0, 152, 138),
        "text": RGBColor(38, 45, 52),
        "muted": RGBColor(99, 110, 123),
        "light": RGBColor(228, 238, 244),
    },
    "学术汇报": {
        "bg": RGBColor(248, 248, 245),
        "panel": RGBColor(255, 255, 255),
        "primary": RGBColor(74, 66, 124),
        "accent": RGBColor(153, 111, 44),
        "text": RGBColor(42, 42, 46),
        "muted": RGBColor(105, 103, 98),
        "light": RGBColor(235, 232, 244),
    },
    "创意演示": {
        "bg": RGBColor(250, 250, 255),
        "panel": RGBColor(255, 255, 255),
        "primary": RGBColor(31, 76, 153),
        "accent": RGBColor(228, 98, 74),
        "text": RGBColor(32, 39, 52),
        "muted": RGBColor(94, 103, 120),
        "light": RGBColor(235, 241, 255),
    },
}


def setup_page() -> None:
    """配置 Streamlit 页面基础样式。"""
    st.set_page_config(page_title=APP_TITLE, page_icon="📊", layout="wide")
    st.markdown(
        """
        <style>
        .block-container {padding-top: 1.6rem; padding-bottom: 2.4rem;}
        [data-testid="stSidebar"] {background: #f6f8fb;}
        [data-testid="stToolbar"], .stDeployButton, #MainMenu, footer {display: none !important;}
        h1, h2, h3 {letter-spacing: 0 !important;}
        .office-card {
            border: 1px solid #e6ebf2;
            border-radius: 8px;
            padding: 16px 18px;
            background: #ffffff;
            margin-bottom: 12px;
        }
        .small-muted {color: #667085; font-size: 0.92rem;}
        .metric-pill {
            display: inline-block;
            padding: 4px 10px;
            border-radius: 999px;
            background: #eef5ff;
            color: #245a8d;
            font-size: 0.85rem;
            margin-right: 6px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip())


def safe_filename(name: str, suffix: str) -> str:
    """生成适合下载的文件名。"""
    cleaned = re.sub(r"[\\/:*?\"<>|\s]+", "_", name.strip())[:36].strip("_")
    return f"{cleaned or '智能办公结果'}{suffix}"


def split_items(raw_text: str) -> list[str]:
    """按中文/英文常见分隔符拆分项目。"""
    cleaned = normalize_text(raw_text)
    cleaned = re.sub(r"\d+\s*个?章节", "", cleaned)
    cleaned = re.sub(r"\d+\s*列", "", cleaned)
    cleaned = cleaned.replace("以及", "、").replace("和", "、").replace("与", "、")
    parts = re.split(r"[、,，;；/]+", cleaned)
    return [item.strip(" .。:：") for item in parts if item.strip(" .。:：")]


def infer_ppt_title(prompt: str) -> str:
    """从提示词中粗略提取 PPT 标题。"""
    text = normalize_text(prompt)
    title_part = re.split(r"包含|包括|分为|围绕|需要|要求", text, maxsplit=1)[0]
    title_part = re.sub(r"^(请|帮我|帮忙|麻烦)?(生成|制作|创建|输出)?(一份|一个|1份)?", "", title_part)
    title_part = re.sub(r"(PPT|ppt|幻灯片|演示文稿|报告)$", "", title_part).strip(" ，,。")
    if not title_part:
        title_part = "智能办公汇报"
    if "PPT" not in title_part.upper() and "汇报" not in title_part and "复盘" not in title_part:
        title_part = f"{title_part}汇报"
    return title_part


def infer_ppt_sections(prompt: str) -> list[str]:
    """从提示词中提取章节，无法识别时使用内置办公汇报结构。"""
    text = normalize_text(prompt)
    match = re.search(r"(?:包含|包括|分为|围绕)(.+?)(?:\d+\s*个?章节|章节|部分|模块|$)", text)
    if match:
        sections = split_items(match.group(1))
        sections = [re.sub(r"(四个|三个|两个|五个)$", "", section).strip() for section in sections]
        sections = [section for section in sections if section and len(section) <= 18]
        if sections:
            return sections

    keyword_sections = []
    keyword_map = [
        ("业绩|营收|销售|经营", "业绩概况"),
        ("用户|增长|流量|留存", "用户增长"),
        ("问题|风险|挑战|不足", "问题分析"),
        ("规划|计划|策略|下月|下一步", "下月规划"),
        ("市场|行业|竞品", "市场洞察"),
        ("产品|功能|研发", "产品进展"),
    ]
    for pattern, section in keyword_map:
        if re.search(pattern, text) and section not in keyword_sections:
            keyword_sections.append(section)
    return keyword_sections or ["背景与目标", "核心进展", "问题分析", "行动计划"]


def section_bullets(section: str, deck_title: str, page_variant: int = 0) -> list[str]:
    """根据章节关键词生成简洁分点，保证页面层级清晰。"""
    templates = [
        (
            r"业绩|营收|销售|经营|概况|结果",
            [
                "梳理核心指标完成情况，突出目标达成率与关键变化",
                "拆解主要贡献来源，识别高增长业务与重点区域",
                "对比上一周期表现，说明趋势、波动原因与结构变化",
                "沉淀可复用经验，为后续资源配置提供依据",
            ],
        ),
        (
            r"用户|增长|流量|留存|转化",
            [
                "展示新增、活跃、留存等关键用户指标的阶段表现",
                "分析主要增长渠道，区分自然增长与活动拉动贡献",
                "关注用户分层变化，识别高价值用户群体特征",
                "提出提升转化与复购的运营抓手",
            ],
        ),
        (
            r"问题|风险|挑战|不足|瓶颈",
            [
                "归纳当前最影响结果的关键问题，避免泛化罗列",
                "从流程、资源、产品和市场四个角度定位原因",
                "评估问题影响范围，明确优先级和处理节奏",
                "设置预警指标，降低后续执行中的不确定性",
            ],
        ),
        (
            r"规划|计划|策略|下月|下一步|行动",
            [
                "明确下一阶段目标，拆分为可执行的重点项目",
                "安排时间节点、责任人和交付标准，提升落地确定性",
                "匹配预算、人员和渠道资源，保证重点事项优先推进",
                "建立复盘机制，按周跟踪进展并动态调整策略",
            ],
        ),
        (
            r"市场|行业|竞品|洞察",
            [
                "概述行业环境变化，提炼对业务有影响的外部信号",
                "对比竞品动作，识别产品、渠道和定价层面的差异",
                "判断机会窗口与潜在威胁，形成可验证假设",
                "将洞察转化为后续策略输入，服务业务决策",
            ],
        ),
        (
            r"产品|功能|研发|技术",
            [
                "说明核心功能迭代进度，突出用户价值与业务目标",
                "梳理需求池优先级，平衡体验、效率与技术成本",
                "总结上线效果与反馈，定位需要继续优化的环节",
                "明确下一轮版本节奏和跨团队协作事项",
            ],
        ),
    ]
    for pattern, bullets in templates:
        if re.search(pattern, section):
            return bullets if page_variant == 0 else detail_bullets(section, page_variant)

    if page_variant:
        return detail_bullets(section, page_variant)
    return [
        f"围绕“{section}”明确本页核心结论，服务《{deck_title}》整体叙事",
        "拆分现状、原因、影响和建议，形成便于汇报的逻辑链",
        "使用数据或案例支撑判断，避免仅停留在主观描述",
        "输出下一步动作，确保内容能够转化为办公决策",
    ]


def detail_bullets(section: str, variant: int) -> list[str]:
    detail_sets = [
        [
            "补充关键数据口径，明确统计范围、周期和对比基准",
            "提炼两到三个最重要发现，用结论先行方式呈现",
            "标注需要管理层关注的事项，便于现场快速决策",
        ],
        [
            "拆解执行路径，将目标分为短期动作和中期建设",
            "识别跨部门协作节点，提前明确输入与输出要求",
            "设置衡量指标，确保后续复盘能够闭环",
        ],
        [
            "总结可复制经验，沉淀为流程、模板或标准动作",
            "保留待验证假设，下一周期通过数据继续校准",
            "形成负责人、截止日期和优先级清单",
        ],
    ]
    return detail_sets[(variant - 1) % len(detail_sets)]


def build_ppt_outline(prompt: str, style_name: str, page_range: tuple[int, int]) -> dict[str, Any]:
    """构建 PPT 大纲与每页内容。"""
    title = infer_ppt_title(prompt)
    sections = infer_ppt_sections(prompt)
    min_pages, max_pages = page_range
    base_total = len(sections) + 3
    target_total = max(base_total, min(max_pages, max(min_pages, (min_pages + max_pages) // 2)))
    content_count = max(len(sections), target_total - 3)

    pages = []
    for index in range(content_count):
        section = sections[index % len(sections)]
        variant = index // len(sections)
        page_title = section if variant == 0 else f"{section}：重点展开"
        pages.append(
            {
                "section": section,
                "title": page_title,
                "bullets": section_bullets(section, title, variant),
                "tag": "核心章节" if variant == 0 else "补充页",
            }
        )

    return {
        "title": title,
        "subtitle": f"{style_name} · 精美版式 · {date.today().strftime('%Y-%m-%d')}",
        "sections": sections,
        "pages": pages,
        "total_pages": len(pages) + 3,
        "style": style_name,
    }


def set_shape_fill(shape: Any, color: RGBColor) -> None:
    shape.fill.solid()
    shape.fill.fore_color.rgb = color


def set_shape_line(shape: Any, color: RGBColor | None = None, transparency: int = 100000) -> None:
    shape.line.fill.background()
    if color:
        shape.line.fill.solid()
        shape.line.fill.fore_color.rgb = color
    shape.line.transparency = transparency


def add_text(
    slide: Any,
    text: str,
    left: float,
    top: float,
    width: float,
    height: float,
    font_size: int = 20,
    color: RGBColor | None = None,
    bold: bool = False,
    align: PP_ALIGN = PP_ALIGN.LEFT,
    font_name: str = "Microsoft YaHei",
) -> Any:
    """添加文本框并统一字体样式。"""
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    paragraph = frame.paragraphs[0]
    paragraph.text = text
    paragraph.alignment = align
    run = paragraph.runs[0]
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    return box


def add_bullets(
    slide: Any,
    bullets: list[str],
    left: float,
    top: float,
    width: float,
    height: float,
    color: RGBColor,
    font_size: int = 18,
) -> Any:
    """添加项目符号列表。"""
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    for index, bullet in enumerate(bullets):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.text = bullet
        paragraph.level = 0
        paragraph.space_after = Pt(8)
        paragraph.line_spacing = 1.12
        run = paragraph.runs[0]
        run.font.name = "Microsoft YaHei"
        run.font.size = Pt(font_size)
        run.font.color.rgb = color
    return box


def add_background(slide: Any, palette: dict[str, RGBColor]) -> None:
    background = slide.background
    background.fill.solid()
    background.fill.fore_color.rgb = palette["bg"]


def add_footer(slide: Any, palette: dict[str, RGBColor], page_no: int, total_pages: int) -> None:
    add_text(
        slide,
        f"{page_no:02d} / {total_pages:02d}",
        11.75,
        6.95,
        1.0,
        0.25,
        font_size=9,
        color=palette["muted"],
        align=PP_ALIGN.RIGHT,
    )


def create_pptx(outline: dict[str, Any]) -> bytes:
    """将大纲渲染为标准 pptx 文件。"""
    palette = PPT_STYLES[outline["style"]]
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    total_pages = outline["total_pages"]

    # 封面页
    slide = prs.slides.add_slide(blank)
    add_background(slide, palette)
    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(0.28), Inches(7.5))
    set_shape_fill(band, palette["accent"])
    set_shape_line(band)
    accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.95), Inches(1.08), Inches(1.4), Inches(0.08))
    set_shape_fill(accent, palette["accent"])
    set_shape_line(accent)
    add_text(slide, outline["title"], 0.95, 1.35, 9.4, 1.35, 34, palette["primary"], True)
    add_text(slide, outline["subtitle"], 0.98, 2.78, 7.8, 0.45, 15, palette["muted"])
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.85), Inches(4.55), Inches(3.55), Inches(1.15))
    set_shape_fill(card, palette["panel"])
    set_shape_line(card, palette["light"], transparency=0)
    add_text(slide, "精美版式大纲 / 可导出 Office 与 WPS", 9.12, 4.84, 3.05, 0.35, 12, palette["text"])
    add_footer(slide, palette, 1, total_pages)

    # 目录页
    slide = prs.slides.add_slide(blank)
    add_background(slide, palette)
    add_text(slide, "目录", 0.78, 0.52, 2.4, 0.62, 26, palette["primary"], True)
    add_text(slide, "Agenda", 0.82, 1.14, 2.2, 0.28, 10, palette["muted"])
    for idx, section in enumerate(outline["sections"], start=1):
        row_top = 1.65 + (idx - 1) * 0.78
        num = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.0), Inches(row_top), Inches(0.38), Inches(0.38))
        set_shape_fill(num, palette["primary"])
        set_shape_line(num)
        add_text(slide, f"{idx}", 1.0, row_top + 0.055, 0.38, 0.18, 10, RGBColor(255, 255, 255), True, PP_ALIGN.CENTER)
        add_text(slide, section, 1.58, row_top - 0.02, 8.7, 0.45, 18, palette["text"], True)
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.58), Inches(row_top + 0.5), Inches(8.8), Inches(0.02))
        set_shape_fill(line, palette["light"])
        set_shape_line(line)
    add_footer(slide, palette, 2, total_pages)

    # 内容页
    for index, page in enumerate(outline["pages"], start=3):
        slide = prs.slides.add_slide(blank)
        add_background(slide, palette)
        title_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.16))
        set_shape_fill(title_bar, palette["accent"])
        set_shape_line(title_bar)
        add_text(slide, page["title"], 0.72, 0.48, 9.4, 0.55, 24, palette["primary"], True)
        add_text(slide, page["tag"], 10.65, 0.58, 1.6, 0.24, 10, palette["muted"], align=PP_ALIGN.RIGHT)

        panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.72), Inches(1.45), Inches(8.15), Inches(4.85))
        set_shape_fill(panel, palette["panel"])
        set_shape_line(panel, palette["light"], transparency=0)
        add_bullets(slide, page["bullets"], 1.05, 1.82, 7.45, 3.9, palette["text"], 16)

        insight = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.28), Inches(1.45), Inches(3.3), Inches(4.85))
        set_shape_fill(insight, palette["light"])
        set_shape_line(insight)
        add_text(slide, "汇报要点", 9.58, 1.82, 1.7, 0.35, 16, palette["primary"], True)
        insight_text = (
            "结论先行\n"
            "数据佐证\n"
            "动作闭环"
        )
        add_text(slide, insight_text, 9.6, 2.48, 2.4, 1.5, 20, palette["text"], True)
        add_text(slide, "建议替换为真实业务数据后直接汇报。", 9.62, 4.86, 2.25, 0.58, 11, palette["muted"])
        add_footer(slide, palette, index, total_pages)

    # 致谢页
    slide = prs.slides.add_slide(blank)
    add_background(slide, palette)
    add_text(slide, "感谢聆听", 0.0, 2.55, 13.333, 0.72, 34, palette["primary"], True, PP_ALIGN.CENTER)
    add_text(slide, "Thank You", 0.0, 3.35, 13.333, 0.35, 16, palette["muted"], align=PP_ALIGN.CENTER)
    end_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5.3), Inches(4.08), Inches(2.75), Inches(0.06))
    set_shape_fill(end_line, palette["accent"])
    set_shape_line(end_line)
    add_footer(slide, palette, total_pages, total_pages)

    output = io.BytesIO()
    prs.save(output)
    return output.getvalue()


def infer_table_columns(prompt: str) -> list[str]:
    """从自然语言中提取表格列名。"""
    text = normalize_text(prompt)
    match = re.search(r"(?:包含|包括|字段为|列为)(.+?)(?:\d+\s*列|列|字段|，填充|,填充|$)", text)
    if match:
        columns = split_items(match.group(1))
        columns = [re.sub(r"(分别为|为)$", "", col).strip() for col in columns]
        columns = [col for col in columns if col and len(col) <= 18]
        if columns:
            return columns

    if re.search(r"门店|营收|销售", text):
        return ["门店名", "月度营收", "同比增速", "排名"]
    if re.search(r"项目|任务|进度", text):
        return ["项目名称", "负责人", "当前状态", "完成进度", "风险等级"]
    if re.search(r"员工|人事|绩效", text):
        return ["员工姓名", "部门", "岗位", "绩效评分", "备注"]
    return ["项目", "类别", "数值", "状态", "备注"]


def infer_row_count(prompt: str, fallback: int = 8) -> int:
    match = re.search(r"(\d{1,2})\s*(?:行|条|个|家|门店)", prompt)
    if match:
        return max(3, min(50, int(match.group(1))))
    return fallback


def generated_values_for_column(column: str, rows: int, prompt: str) -> list[Any]:
    """按列名生成办公场景示例数据。"""
    shop_names = ["上海旗舰店", "北京朝阳店", "广州天河店", "深圳南山店", "杭州西湖店", "成都高新店", "南京新街口店", "武汉光谷店"]
    project_names = ["客户增长项目", "会员运营项目", "流程优化项目", "数据看板项目", "新品推广项目", "服务升级项目"]
    owners = ["张晨", "李安", "王悦", "赵宁", "陈一", "刘佳", "周楠", "孙禾"]
    statuses = ["进行中", "已完成", "待启动", "有风险"]
    risks = ["低", "中", "高"]

    values: list[Any] = []
    for index in range(rows):
        base = index + 1
        if re.search(r"门店|店名", column):
            values.append(shop_names[index % len(shop_names)])
        elif re.search(r"项目|任务", column):
            values.append(project_names[index % len(project_names)])
        elif re.search(r"负责人|姓名|员工", column):
            values.append(owners[index % len(owners)])
        elif re.search(r"部门", column):
            values.append(["运营部", "市场部", "销售部", "产品部", "客服部"][index % 5])
        elif re.search(r"日期|时间", column):
            values.append(f"2024-{(index % 12) + 1:02d}-01")
        elif re.search(r"月|月份", column) and not re.search(r"营收|金额|收入", column):
            values.append(f"2024-{(index % 12) + 1:02d}")
        elif re.search(r"营收|收入|销售额|金额|GMV|费用|成本", column, re.IGNORECASE):
            values.append(round(86.5 + index * 13.7 + (index % 3) * 5.2, 2))
        elif re.search(r"同比|环比|增速|增长率|完成率|进度|占比", column):
            values.append(f"{round(8.5 + index * 1.7 + (index % 4) * 0.8, 1)}%")
        elif re.search(r"排名|序号", column):
            values.append(base)
        elif re.search(r"数量|人数|订单|销量", column):
            values.append(120 + index * 17)
        elif re.search(r"评分|得分|分数", column):
            values.append(round(82 + (index % 6) * 2.5, 1))
        elif re.search(r"状态", column):
            values.append(statuses[index % len(statuses)])
        elif re.search(r"风险", column):
            values.append(risks[index % len(risks)])
        elif re.search(r"备注|说明", column):
            values.append("示例数据，可按实际情况替换")
        else:
            values.append(f"{column}{base}")
    return values


def generate_table_from_prompt(prompt: str, rows: int | None = None) -> pd.DataFrame:
    """根据自然语言提示词生成结构化表格。"""
    columns = infer_table_columns(prompt)
    row_count = rows or infer_row_count(prompt)
    data = {column: generated_values_for_column(column, row_count, prompt) for column in columns}
    df = pd.DataFrame(data)

    rank_columns = [col for col in df.columns if re.search(r"排名", col)]
    amount_columns = [col for col in df.columns if re.search(r"营收|收入|销售额|金额|GMV", col, re.IGNORECASE)]
    if rank_columns and amount_columns:
        rank = df[amount_columns[0]].rank(method="first", ascending=False).astype(int)
        df[rank_columns[0]] = rank
        df = df.sort_values(rank_columns[0]).reset_index(drop=True)
    return df


def demo_dataframe() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "门店名": ["上海旗舰店", "北京朝阳店", "广州天河店", "深圳南山店", "杭州西湖店", "成都高新店"],
            "月度营收": [186.5, 172.3, None, 153.8, 149.2, 141.6],
            "同比增速": [0.184, 0.163, 0.128, None, 0.112, 0.097],
            "区域": ["华东", "华北", "华南", "华南", "华东", "西南"],
            "排名": [1, 2, 3, 4, 5, 6],
        }
    )


def read_table_file(uploaded_file: Any) -> pd.DataFrame:
    """读取 CSV/XLSX 文件并返回 DataFrame。"""
    suffix = Path(uploaded_file.name).suffix.lower()
    if suffix == ".xlsx":
        return pd.read_excel(io.BytesIO(uploaded_file.getvalue()), engine="openpyxl")
    if suffix == ".csv":
        raw = uploaded_file.getvalue()
        last_error: Exception | None = None
        for encoding in ("utf-8-sig", "utf-8", "gbk"):
            try:
                return pd.read_csv(io.BytesIO(raw), encoding=encoding)
            except Exception as exc:  # noqa: BLE001 - 需要尝试多种常见编码
                last_error = exc
        raise ValueError(f"CSV 文件读取失败：{last_error}") from last_error
    raise ValueError("仅支持 CSV 或 XLSX 格式文件。")


def flatten_columns(columns: pd.Index) -> list[str]:
    flattened = []
    for col in columns:
        if isinstance(col, tuple):
            flattened.append("_".join(str(part) for part in col if str(part) != ""))
        else:
            flattened.append(str(col))
    return flattened


def fill_missing_values(df: pd.DataFrame, method: str, fill_value: str) -> pd.DataFrame:
    """按用户选择填充空值。"""
    result = df.copy()
    if method == "固定值":
        return result.fillna(fill_value)
    if method == "前向填充":
        return result.ffill()
    if method == "后向填充":
        return result.bfill()

    for column in result.columns:
        series = result[column]
        if method == "均值" and pd.api.types.is_numeric_dtype(series):
            result[column] = series.fillna(series.mean())
        elif method == "中位数" and pd.api.types.is_numeric_dtype(series):
            result[column] = series.fillna(series.median())
        else:
            mode = series.mode(dropna=True)
            if not mode.empty:
                result[column] = series.fillna(mode.iloc[0])
    return result


def add_summary_row(df: pd.DataFrame) -> pd.DataFrame:
    """为数值列添加合计行。"""
    if df.empty:
        return df
    summary: dict[str, Any] = {}
    numeric_cols = df.select_dtypes(include="number").columns
    for column in df.columns:
        if column in numeric_cols:
            summary[column] = df[column].sum()
        elif not summary:
            summary[column] = "合计"
        else:
            summary[column] = ""
    return pd.concat([df, pd.DataFrame([summary])], ignore_index=True)


def process_dataframe(
    df: pd.DataFrame,
    drop_duplicates: bool,
    duplicate_subset: list[str],
    fill_missing: bool,
    fill_method: str,
    fill_value: str,
    sort_column: str | None,
    sort_ascending: bool,
    group_columns: list[str],
    agg_columns: list[str],
    agg_methods: list[str],
    append_summary: bool,
) -> tuple[pd.DataFrame, list[str]]:
    """执行去重、填充、排序、分组统计和汇总行等操作。"""
    result = df.copy()
    logs: list[str] = []

    if drop_duplicates:
        before = len(result)
        subset = duplicate_subset or None
        result = result.drop_duplicates(subset=subset)
        logs.append(f"去重：{before} 行 -> {len(result)} 行")

    if fill_missing:
        result = fill_missing_values(result, fill_method, fill_value)
        logs.append(f"空值填充：使用 {fill_method}")

    if sort_column:
        result = result.sort_values(by=sort_column, ascending=sort_ascending, na_position="last").reset_index(drop=True)
        logs.append(f"排序：按“{sort_column}”{'升序' if sort_ascending else '降序'}")

    if group_columns:
        if not agg_methods:
            agg_methods = ["sum"]
        numeric_cols = list(result.select_dtypes(include="number").columns)
        selected_agg_cols = agg_columns or [col for col in numeric_cols if col not in group_columns]
        if selected_agg_cols:
            grouped = result.groupby(group_columns, dropna=False)[selected_agg_cols].agg(agg_methods).reset_index()
            grouped.columns = flatten_columns(grouped.columns)
        else:
            grouped = result.groupby(group_columns, dropna=False).size().reset_index(name="记录数")
        result = grouped
        logs.append(f"分组统计：按 {', '.join(group_columns)} 汇总")

    if append_summary:
        result = add_summary_row(result)
        logs.append("汇总行：已添加数值列合计")

    return result, logs


def dataframe_to_xlsx(df: pd.DataFrame, sheet_name: str = "结果") -> bytes:
    """导出为带基础样式的 XLSX。"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        worksheet = writer.sheets[sheet_name]
        header_fill = PatternFill(fill_type="solid", fgColor="EAF2F8")
        header_font = Font(bold=True, color="1F2937")
        border = Border(bottom=Side(style="thin", color="D7DEE8"))
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border
        for column_cells in worksheet.columns:
            max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
            worksheet.column_dimensions[column_cells[0].column_letter].width = min(max(max_length + 4, 12), 28)
        worksheet.freeze_panes = "A2"
    return output.getvalue()


def dataframe_to_csv(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")


def render_outline_preview(outline: dict[str, Any]) -> None:
    """在页面中预览 PPT 大纲。"""
    st.markdown(
        f"""
        <div class="office-card">
            <span class="metric-pill">共 {outline['total_pages']} 页</span>
            <span class="metric-pill">{outline['style']}</span>
            <span class="metric-pill">{len(outline['sections'])} 个章节</span>
            <h3>{outline['title']}</h3>
            <div class="small-muted">{outline['subtitle']}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    preview_rows = [
        {"页码": 1, "页面": "封面页", "标题": outline["title"], "要点": "标题 / 副标题 / 风格信息"},
        {"页码": 2, "页面": "目录页", "标题": "目录", "要点": "、".join(outline["sections"])},
    ]
    for index, page in enumerate(outline["pages"], start=3):
        preview_rows.append(
            {
                "页码": index,
                "页面": page["tag"],
                "标题": page["title"],
                "要点": "；".join(page["bullets"][:2]),
            }
        )
    preview_rows.append({"页码": outline["total_pages"], "页面": "结束页", "标题": "感谢聆听", "要点": "致谢与收尾"})
    st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)


def ppt_page() -> None:
    st.header("智能 PPT 生成")
    st.caption("输入一句办公需求，自动拆解主题与章节，生成可下载的标准 .pptx 文件。")

    input_col, option_col = st.columns([2, 1])
    with input_col:
        prompt = st.text_area("PPT 提示词", value=DEFAULT_PPT_PROMPT, height=132)
    with option_col:
        style_name = st.selectbox("PPT 风格", list(PPT_STYLES.keys()))
        page_range = st.slider("大致页数范围", min_value=5, max_value=20, value=(6, 10), step=1)

    generate = st.button("一键生成精美 PPT", type="primary", use_container_width=True)
    if generate:
        if not normalize_text(prompt):
            st.warning("请输入 PPT 主题或需求，例如：生成一份 Q3 产品运营复盘 PPT。")
        else:
            try:
                with st.spinner("正在解析提示词并生成 PPT，请稍候..."):
                    outline = build_ppt_outline(prompt, style_name, page_range)
                    pptx_bytes = create_pptx(outline)
                    st.session_state["ppt_outline"] = outline
                    st.session_state["pptx_bytes"] = pptx_bytes
                st.success("PPT 已生成，可在下方预览大纲并下载。")
            except Exception as exc:  # noqa: BLE001 - Streamlit 需要友好兜底提示
                st.error(f"PPT 生成失败：{exc}")

    outline = st.session_state.get("ppt_outline")
    pptx_bytes = st.session_state.get("pptx_bytes")
    if outline and pptx_bytes:
        st.subheader("结果预览")
        render_outline_preview(outline)
        st.download_button(
            "下载精美 PPTX 文件",
            data=pptx_bytes,
            file_name=safe_filename(outline["title"], ".pptx"),
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )
    else:
        st.info("首次使用可直接点击“一键生成精美 PPT”，系统会用默认提示词生成一份演示文件。")


def download_table_buttons(df: pd.DataFrame, base_name: str) -> None:
    xlsx_bytes = dataframe_to_xlsx(df)
    csv_bytes = dataframe_to_csv(df)
    left, right = st.columns(2)
    with left:
        st.download_button(
            "下载精美 XLSX",
            data=xlsx_bytes,
            file_name=safe_filename(base_name, ".xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with right:
        st.download_button(
            "下载 CSV",
            data=csv_bytes,
            file_name=safe_filename(base_name, ".csv"),
            mime="text/csv",
            use_container_width=True,
        )


def text_to_table_tab() -> None:
    st.subheader("文本生成精美表格")
    prompt = st.text_area("表格描述", value=DEFAULT_TABLE_PROMPT, height=112, key="table_prompt")
    row_count = st.slider("示例数据行数", min_value=3, max_value=30, value=8, step=1)

    if st.button("生成精美表格", type="primary", use_container_width=True):
        if not normalize_text(prompt):
            st.warning("请输入表格描述，例如：生成 2024 年各门店营收表。")
        else:
            try:
                with st.spinner("正在生成结构化表格..."):
                    df = generate_table_from_prompt(prompt, row_count)
                    st.session_state["generated_df"] = df
                st.success("表格已生成。")
            except Exception as exc:  # noqa: BLE001
                st.error(f"表格生成失败：{exc}")

    df = st.session_state.get("generated_df")
    if df is not None:
        st.markdown("#### 实时预览")
        st.dataframe(style_dataframe_for_preview(df), use_container_width=True, hide_index=True)
        download_table_buttons(df, "文本生成精美表格")
    else:
        st.info("首次使用可直接点击“生成精美表格”，系统会使用默认描述生成示例数据。")


def uploaded_or_demo_dataframe(uploaded_file: Any, use_demo: bool) -> pd.DataFrame | None:
    if uploaded_file is not None:
        return read_table_file(uploaded_file)
    if use_demo:
        return demo_dataframe()
    return None


def process_table_tab() -> None:
    st.subheader("表格数据处理")
    uploaded_file = st.file_uploader("上传 CSV / XLSX 文件", type=["csv", "xlsx"])
    use_demo = st.checkbox("未上传时使用示例销售数据", value=True)

    try:
        source_df = uploaded_or_demo_dataframe(uploaded_file, use_demo)
    except Exception as exc:  # noqa: BLE001
        st.error(f"文件读取失败：{exc}")
        source_df = None

    if source_df is None:
        st.info("请上传 CSV / XLSX 文件，或勾选使用示例数据。")
        return

    st.markdown("#### 原始数据预览")
    st.dataframe(style_dataframe_for_preview(source_df.head(100)), use_container_width=True, hide_index=True)

    st.markdown("#### 处理操作")
    col_left, col_right = st.columns(2)
    with col_left:
        drop_duplicates = st.checkbox("去重")
        duplicate_subset = st.multiselect("去重依据列（留空表示整行去重）", list(source_df.columns))
        fill_missing = st.checkbox("空值填充")
        fill_method = st.selectbox("填充方式", ["固定值", "均值", "中位数", "众数", "前向填充", "后向填充"])
        fill_value = st.text_input("固定填充值", value="未填写")
    with col_right:
        enable_sort = st.checkbox("按列排序")
        sort_column = st.selectbox("排序列", [""] + list(source_df.columns))
        sort_ascending = st.radio("排序方式", ["升序", "降序"], horizontal=True) == "升序"
        enable_group = st.checkbox("分组统计")
        group_columns = st.multiselect("分组列", list(source_df.columns))
        numeric_cols = list(source_df.select_dtypes(include="number").columns)
        agg_columns = st.multiselect("统计数值列", numeric_cols, default=numeric_cols[:1])
        agg_methods = st.multiselect("统计方式", ["sum", "mean", "count", "max", "min"], default=["sum"])
        append_summary = st.checkbox("添加汇总行")

    if st.button("执行处理", type="primary", use_container_width=True):
        try:
            with st.spinner("正在处理表格数据..."):
                processed_df, logs = process_dataframe(
                    source_df,
                    drop_duplicates=drop_duplicates,
                    duplicate_subset=duplicate_subset,
                    fill_missing=fill_missing,
                    fill_method=fill_method,
                    fill_value=fill_value,
                    sort_column=sort_column if enable_sort and sort_column else None,
                    sort_ascending=sort_ascending,
                    group_columns=group_columns if enable_group else [],
                    agg_columns=agg_columns if enable_group else [],
                    agg_methods=agg_methods if enable_group else [],
                    append_summary=append_summary,
                )
                st.session_state["processed_df"] = processed_df
                st.session_state["process_logs"] = logs
            st.success("表格处理完成。")
        except Exception as exc:  # noqa: BLE001
            st.error(f"表格处理失败：{exc}")

    processed_df = st.session_state.get("processed_df")
    if processed_df is not None:
        st.markdown("#### 处理结果预览")
        logs = st.session_state.get("process_logs", [])
        if logs:
            st.write("；".join(logs))
        st.dataframe(style_dataframe_for_preview(processed_df), use_container_width=True, hide_index=True)
        download_table_buttons(processed_df, "处理后表格")


def table_page() -> None:
    st.header("智能表格生成与处理")
    st.caption("支持文本生成精美表格、CSV/XLSX 上传处理、实时预览与导出。")
    tab_generate, tab_process = st.tabs(["文本生成精美表格", "表格数据处理"])
    with tab_generate:
        text_to_table_tab()
    with tab_process:
        process_table_tab()


def main() -> None:
    setup_page()
    st.sidebar.title(APP_TITLE)
    st.sidebar.caption("面向普通办公用户的单文件 Web 小程序")
    module = st.sidebar.radio("功能模块", ["PPT生成", "表格工具"])
    st.sidebar.divider()
    st.sidebar.info("所有示例内容均在本地规则生成，可直接导出后继续编辑。")

    if module == "PPT生成":
        ppt_page()
    else:
        table_page()



# --- enhanced-output-overrides ---
# 这组函数在 main() 执行前覆盖基础生成逻辑，让 PPT 与 XLSX 输出更像成品模板。

def _office_palette(palette: dict[str, RGBColor]) -> dict[str, RGBColor]:
    return {
        "bg": palette.get("bg", RGBColor(246, 248, 252)),
        "paper": palette.get("panel", RGBColor(255, 255, 255)),
        "primary": palette.get("primary", RGBColor(27, 55, 86)),
        "primary_2": RGBColor(49, 106, 150),
        "accent": palette.get("accent", RGBColor(0, 159, 149)),
        "accent_2": RGBColor(244, 169, 69),
        "text": palette.get("text", RGBColor(36, 44, 54)),
        "muted": palette.get("muted", RGBColor(104, 116, 132)),
        "line": palette.get("light", RGBColor(218, 226, 236)),
        "soft": RGBColor(230, 242, 247),
    }


def _add_card(slide: Any, left: float, top: float, width: float, height: float, fill: RGBColor, line: RGBColor) -> None:
    shadow = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left + 0.04), Inches(top + 0.05), Inches(width), Inches(height))
    set_shape_fill(shadow, RGBColor(222, 229, 238))
    set_shape_line(shadow)
    panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    set_shape_fill(panel, fill)
    set_shape_line(panel, line, transparency=0)


def _metrics_for(section: str, index: int) -> list[dict[str, str]]:
    if re.search(r"业绩|营收|销售|经营", section):
        return [
            {"label": "目标达成", "value": f"{92 + index % 6}%", "trend": "较上期 +6.2%"},
            {"label": "收入贡献", "value": f"{180 + index * 8}万", "trend": "核心区域拉动"},
            {"label": "转化效率", "value": f"{18 + index % 4}.8%", "trend": "渠道结构优化"},
        ]
    if re.search(r"用户|增长|流量|留存", section):
        return [
            {"label": "新增用户", "value": f"{12 + index % 5}.6万", "trend": "活动引流增强"},
            {"label": "留存率", "value": f"{34 + index % 4}.5%", "trend": "会员触达改善"},
            {"label": "获客成本", "value": f"{28 - index % 3}元", "trend": "环比下降"},
        ]
    if re.search(r"问题|风险|挑战|不足", section):
        return [
            {"label": "高优先级", "value": f"{3 + index % 3}项", "trend": "需协同解决"},
            {"label": "影响范围", "value": "中", "trend": "聚焦关键流程"},
            {"label": "解决周期", "value": f"{2 + index % 2}周", "trend": "按周跟踪"},
        ]
    return [
        {"label": "关键项目", "value": f"{4 + index % 4}项", "trend": "聚焦落地"},
        {"label": "预计收益", "value": f"{15 + index % 5}%", "trend": "持续验证"},
        {"label": "优先级", "value": "P1", "trend": "资源前置"},
    ]


def _metric_card(slide: Any, p: dict[str, RGBColor], metric: dict[str, str], left: float, top: float, width: float) -> None:
    _add_card(slide, left, top, width, 1.05, p["paper"], p["line"])
    add_text(slide, metric["label"], left + 0.18, top + 0.14, width - 0.35, 0.22, 9, p["muted"])
    add_text(slide, metric["value"], left + 0.18, top + 0.39, width - 0.35, 0.34, 19, p["primary"], True)
    add_text(slide, metric["trend"], left + 0.18, top + 0.76, width - 0.35, 0.2, 8, p["accent"])


def _tiny_bar_chart(slide: Any, p: dict[str, RGBColor], left: float, top: float) -> None:
    values = [0.48, 0.7, 0.58, 0.86, 0.76]
    labels = ["M1", "M2", "M3", "M4", "M5"]
    for idx, value in enumerate(values):
        x = left + idx * 0.46
        height = 1.35 * value
        bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(top + 1.5 - height), Inches(0.24), Inches(height))
        set_shape_fill(bar, p["accent"] if idx == 3 else p["primary_2"])
        set_shape_line(bar)
        add_text(slide, labels[idx], x - 0.03, top + 1.56, 0.33, 0.16, 7, p["muted"], align=PP_ALIGN.CENTER)


def _bullet_cards(slide: Any, p: dict[str, RGBColor], bullets: list[str]) -> None:
    for idx, bullet in enumerate(bullets[:4], start=1):
        top = 2.55 + (idx - 1) * 0.78
        badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.98), Inches(top + 0.02), Inches(0.35), Inches(0.35))
        set_shape_fill(badge, p["accent"] if idx % 2 else p["primary_2"])
        set_shape_line(badge)
        add_text(slide, str(idx), 1.0, top + 0.08, 0.31, 0.12, 8, RGBColor(255, 255, 255), True, PP_ALIGN.CENTER)
        add_text(slide, bullet, 1.52, top - 0.03, 7.0, 0.48, 14, p["text"])


def create_pptx(outline: dict[str, Any]) -> bytes:
    """增强版 PPT：封面视觉区、目录时间线、指标卡片和行动建议区。"""
    p = _office_palette(PPT_STYLES[outline["style"]])
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    total_pages = outline["total_pages"]

    slide = prs.slides.add_slide(blank)
    add_background(slide, p)
    hero = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(2.05))
    set_shape_fill(hero, p["primary"])
    set_shape_line(hero)
    ribbon = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(2.05), Inches(13.333), Inches(0.18))
    set_shape_fill(ribbon, p["accent"])
    set_shape_line(ribbon)
    block = slide.shapes.add_shape(MSO_SHAPE.PARALLELOGRAM, Inches(9.7), Inches(0.28), Inches(2.25), Inches(1.45))
    set_shape_fill(block, p["accent_2"])
    set_shape_line(block)
    add_text(slide, "BUSINESS REVIEW", 0.82, 0.46, 3.0, 0.22, 9, RGBColor(220, 232, 245), True)
    add_text(slide, outline["title"], 0.78, 0.86, 9.0, 0.75, 33, RGBColor(255, 255, 255), True)
    add_text(slide, f"{outline['style']} · 精美版式 · {date.today().strftime('%Y-%m-%d')}", 0.82, 1.62, 7.0, 0.3, 12, RGBColor(218, 229, 239))
    _add_card(slide, 0.78, 2.92, 7.55, 2.85, p["paper"], p["line"])
    add_text(slide, "汇报结构", 1.1, 3.18, 1.8, 0.3, 15, p["primary"], True)
    for idx, section in enumerate(outline["sections"][:5], start=1):
        top = 3.72 + (idx - 1) * 0.38
        add_text(slide, f"0{idx}", 1.12, top, 0.45, 0.2, 9, p["accent"], True)
        add_text(slide, section, 1.72, top - 0.02, 5.8, 0.25, 12, p["text"])
    _add_card(slide, 8.82, 2.92, 3.65, 2.85, p["soft"], p["line"])
    add_text(slide, "一页一结论", 9.15, 3.24, 2.5, 0.3, 18, p["primary"], True)
    add_text(slide, "精美版式版式已加入指标卡片、强调区、章节层次和商务配色。", 9.15, 3.82, 2.55, 0.78, 12, p["text"])
    _tiny_bar_chart(slide, p, 9.22, 4.33)
    add_footer(slide, p, 1, total_pages)

    slide = prs.slides.add_slide(blank)
    add_background(slide, p)
    add_text(slide, "目录", 0.78, 0.52, 2.4, 0.54, 27, p["primary"], True)
    add_text(slide, "Agenda / Content Map", 0.82, 1.09, 3.0, 0.24, 9, p["muted"], True)
    timeline = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.15), Inches(2.05), Inches(10.6), Inches(0.03))
    set_shape_fill(timeline, p["line"])
    set_shape_line(timeline)
    card_width = 10.5 / max(1, len(outline["sections"]))
    for idx, section in enumerate(outline["sections"], start=1):
        left = 0.92 + (idx - 1) * card_width
        node = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(left + 0.22), Inches(1.82), Inches(0.45), Inches(0.45))
        set_shape_fill(node, p["accent"] if idx % 2 else p["primary_2"])
        set_shape_line(node)
        add_text(slide, str(idx), left + 0.22, 1.9, 0.45, 0.12, 9, RGBColor(255, 255, 255), True, PP_ALIGN.CENTER)
        _add_card(slide, left, 2.48, max(1.65, card_width - 0.18), 1.7, p["paper"], p["line"])
        add_text(slide, section, left + 0.18, 2.72, max(1.1, card_width - 0.54), 0.42, 15, p["primary"], True)
        add_text(slide, "关键结论 · 数据支撑 · 行动建议", left + 0.18, 3.48, max(1.1, card_width - 0.5), 0.35, 9, p["muted"])
    add_footer(slide, p, 2, total_pages)

    for index, page in enumerate(outline["pages"], start=3):
        slide = prs.slides.add_slide(blank)
        add_background(slide, p)
        side = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(0.28), Inches(7.5))
        set_shape_fill(side, p["primary"])
        set_shape_line(side)
        add_text(slide, f"{index - 2:02d}", 0.62, 0.45, 0.65, 0.28, 10, p["accent"], True)
        add_text(slide, page["title"], 0.78, 0.75, 8.9, 0.48, 24, p["primary"], True)
        add_text(slide, page["tag"], 10.35, 0.83, 1.9, 0.22, 9, p["muted"], True, PP_ALIGN.RIGHT)
        title_rule = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.82), Inches(1.36), Inches(1.1), Inches(0.06))
        set_shape_fill(title_rule, p["accent"])
        set_shape_line(title_rule)
        for metric_index, metric in enumerate(_metrics_for(page["section"], index)):
            _metric_card(slide, p, metric, 0.82 + metric_index * 2.82, 1.62, 2.55)
        _add_card(slide, 0.82, 2.42, 8.18, 3.85, p["paper"], p["line"])
        add_text(slide, "核心结论", 1.05, 2.68, 1.6, 0.26, 14, p["primary"], True)
        _bullet_cards(slide, p, page["bullets"])
        _add_card(slide, 9.35, 1.62, 3.0, 4.65, p["soft"], p["line"])
        add_text(slide, "建议动作", 9.65, 1.92, 1.6, 0.28, 15, p["primary"], True)
        for item_idx, (color_key, action) in enumerate([("primary_2", "明确负责人和截止日期"), ("accent", "按周复盘指标变化"), ("accent_2", "把结论沉淀为模板")], start=1):
            top = 2.47 + (item_idx - 1) * 0.76
            badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(9.6), Inches(top), Inches(0.36), Inches(0.36))
            set_shape_fill(badge, p[color_key])
            set_shape_line(badge)
            add_text(slide, str(item_idx), 9.68, top + 0.08, 0.25, 0.18, 9, RGBColor(255, 255, 255), True, PP_ALIGN.CENTER)
            add_text(slide, action, 10.1, top, 1.8, 0.42, 11, p["text"])
        _tiny_bar_chart(slide, p, 9.72, 4.72)
        add_footer(slide, p, index, total_pages)

    slide = prs.slides.add_slide(blank)
    add_background(slide, p)
    block = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(2.1), Inches(13.333), Inches(1.9))
    set_shape_fill(block, p["primary"])
    set_shape_line(block)
    add_text(slide, "感谢聆听", 0.0, 2.58, 13.333, 0.56, 32, RGBColor(255, 255, 255), True, PP_ALIGN.CENTER)
    add_text(slide, "Thank You / Questions", 0.0, 3.22, 13.333, 0.28, 12, RGBColor(217, 228, 240), True, PP_ALIGN.CENTER)
    add_text(slide, "建议将本页替换为汇报人、部门或联系方式。", 0.0, 4.36, 13.333, 0.3, 11, p["muted"], align=PP_ALIGN.CENTER)
    add_footer(slide, p, total_pages, total_pages)

    output = io.BytesIO()
    prs.save(output)
    return output.getvalue()


def _percent_like(column: str) -> bool:
    return bool(re.search(r"率|增速|占比|进度", column))


def _amount_like(column: str) -> bool:
    return bool(re.search(r"营收|收入|销售额|金额|GMV|费用|成本", column, re.IGNORECASE))


def style_dataframe_for_preview(df: pd.DataFrame) -> Any:
    formats = {}
    for column in df.columns:
        if pd.api.types.is_numeric_dtype(df[column]):
            formats[column] = "{:.1%}" if _percent_like(str(column)) else "{:,.2f}"
    return (
        df.style.format(formats, na_rep="-")
        .set_properties(**{"text-align": "left"})
        .set_table_styles([
            {"selector": "th", "props": [("background-color", "#1b3756"), ("color", "#ffffff"), ("font-weight", "700")]},
            {"selector": "td", "props": [("border-bottom", "1px solid #e6ebf2")]},
        ])
    )


def dataframe_to_xlsx(df: pd.DataFrame, sheet_name: str = "结果") -> bytes:
    """增强版 XLSX：数据概览页、主题表格、冻结、筛选和条件格式。"""
    from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        overview = workbook.create_sheet("数据概览", 0)
        overview.sheet_view.showGridLines = False
        overview["A1"] = "智能表格结果概览"
        overview["A1"].font = Font(size=18, bold=True, color="1B3756")
        overview["A2"] = f"生成日期：{date.today().strftime('%Y-%m-%d')}"
        overview["A2"].font = Font(color="667085")
        for idx, (label, value) in enumerate([("总行数", len(df)), ("总列数", len(df.columns)), ("数值列", len(df.select_dtypes(include="number").columns)), ("空值数", int(df.isna().sum().sum()))]):
            col = 1 + idx * 2
            overview.cell(row=4, column=col, value=label)
            overview.cell(row=5, column=col, value=value)
            overview.cell(row=4, column=col).fill = PatternFill(fill_type="solid", fgColor="EAF2F8")
            overview.cell(row=4, column=col).font = Font(bold=True, color="1B3756")
            overview.cell(row=5, column=col).font = Font(size=16, bold=True, color="0F766E")
            overview.column_dimensions[get_column_letter(col)].width = 16
        overview["A8"] = "说明"
        overview["A8"].font = Font(bold=True, color="1B3756")
        overview["A9"] = "结果页已添加筛选、冻结表头、主题样式与数值条件格式，可直接用于办公汇报。"
        overview["A9"].alignment = Alignment(wrap_text=True)
        overview.column_dimensions["A"].width = 32

        worksheet = writer.sheets[sheet_name]
        max_row, max_col = worksheet.max_row, worksheet.max_column
        header_fill = PatternFill(fill_type="solid", fgColor="1B3756")
        header_font = Font(bold=True, color="FFFFFF")
        border = Border(bottom=Side(style="thin", color="D9E2EC"))
        even_fill = PatternFill(fill_type="solid", fgColor="F8FBFF")
        odd_fill = PatternFill(fill_type="solid", fgColor="FFFFFF")
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border
        for row in worksheet.iter_rows(min_row=2, max_row=max_row):
            fill = even_fill if row[0].row % 2 == 0 else odd_fill
            for cell in row:
                cell.fill = fill
                cell.border = border
                cell.alignment = Alignment(vertical="center")
        for idx, column_name in enumerate(df.columns, start=1):
            letter = get_column_letter(idx)
            values = [worksheet.cell(row=row, column=idx).value for row in range(1, max_row + 1)]
            worksheet.column_dimensions[letter].width = min(max(max(len(str(v)) if v is not None else 0 for v in values) + 4, 12), 30)
            if pd.api.types.is_numeric_dtype(df[column_name]):
                data_range = f"{letter}2:{letter}{max_row}"
                if _percent_like(str(column_name)):
                    for cell in worksheet[letter][1:]:
                        cell.number_format = "0.0%"
                    worksheet.conditional_formatting.add(data_range, DataBarRule(start_type="min", end_type="max", color="2AA7A1", showValue=True))
                else:
                    for cell in worksheet[letter][1:]:
                        cell.number_format = "#,##0.00"
                    worksheet.conditional_formatting.add(data_range, ColorScaleRule(start_type="min", start_color="FDE68A", mid_type="percentile", mid_value=50, mid_color="BFDBFE", end_type="max", end_color="2563EB"))
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        worksheet.sheet_view.showGridLines = False
        table_ref = f"A1:{get_column_letter(max_col)}{max_row}"
        table = Table(displayName="OfficeResultTable", ref=table_ref)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        worksheet.add_table(table)
    return output.getvalue()

if __name__ == "__main__":
    main()


# =========================
# 依赖安装命令（requirements.txt 内容）
# =========================
# streamlit>=1.36,<2.0
# pandas>=2.2,<3.0
# openpyxl>=3.1,<4.0
# python-pptx>=0.6.23,<1.1
#
# 一键安装：
# pip install "streamlit>=1.36,<2.0" "pandas>=2.2,<3.0" "openpyxl>=3.1,<4.0" "python-pptx>=0.6.23,<1.1"
#
# 启动步骤：
# 1. 将本文件保存为 app.py
# 2. 在当前目录执行依赖安装命令
# 3. 运行：streamlit run app.py
# 4. 浏览器打开命令行输出的本地地址即可使用
