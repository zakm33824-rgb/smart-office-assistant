from __future__ import annotations

from io import BytesIO
from pathlib import Path
from typing import Iterable

from PIL import Image, ImageColor, ImageDraw, ImageFont


def _font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    try:
        if bold:
            candidates = [
                "C:/Windows/Fonts/arialbd.ttf",
                "C:/Windows/Fonts/segoeuib.ttf",
                "C:/Windows/Fonts/msyhbd.ttc",
            ]
        else:
            candidates = [
                "C:/Windows/Fonts/arial.ttf",
                "C:/Windows/Fonts/segoeui.ttf",
                "C:/Windows/Fonts/msyh.ttc",
            ]
        for candidate in candidates:
            if Path(candidate).exists():
                return ImageFont.truetype(candidate, size=size)
    except Exception:
        pass
    return ImageFont.load_default()


def create_source_preview(
    title: str,
    subtitle: str,
    score: int,
    tags: Iterable[str],
    out_path: str | Path,
    width: int = 1280,
    height: int = 720,
) -> Path:
    background = Image.new("RGB", (width, height), ImageColor.getrgb("#F7F9FC"))
    draw = ImageDraw.Draw(background)
    draw.rounded_rectangle((40, 40, width - 40, height - 40), radius=32, fill="#FFFFFF", outline="#D8E0EA", width=3)
    draw.rounded_rectangle((70, 70, width - 70, 190), radius=28, fill="#1B3756")
    draw.text((100, 105), "PPT Template Library", fill="#FFFFFF", font=_font(24, bold=True))
    draw.text((100, 145), "Source card preview", fill="#CFE3FA", font=_font(15))

    draw.text((100, 240), title, fill="#102A43", font=_font(42, bold=True))
    draw.text((100, 305), subtitle, fill="#5B6B7A", font=_font(22))

    draw.rounded_rectangle((100, 380, 340, 455), radius=22, fill="#EAF2F8")
    draw.text((135, 405), f"Score {score}", fill="#1B3756", font=_font(28, bold=True))

    x = 100
    y = 510
    for tag in list(tags)[:8]:
        tag_width = max(90, 22 * len(tag) + 28)
        draw.rounded_rectangle((x, y, x + tag_width, y + 50), radius=22, fill="#EEF5FF")
        draw.text((x + 18, y + 14), tag, fill="#245A8D", font=_font(18, bold=True))
        x += tag_width + 16
        if x > width - 220:
            x = 100
            y += 64

    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    background.save(out)
    return out


def build_contact_sheet(image_paths: list[str | Path], out_path: str | Path, columns: int = 3) -> Path:
    if not image_paths:
        raise ValueError("No images provided")
    imgs = [Image.open(path).convert("RGB") for path in image_paths]
    thumb_w, thumb_h = 420, 236
    margin = 24
    rows = (len(imgs) + columns - 1) // columns
    canvas = Image.new("RGB", (columns * thumb_w + (columns + 1) * margin, rows * thumb_h + (rows + 1) * margin), "#F7F9FC")
    for index, img in enumerate(imgs):
        row, col = divmod(index, columns)
        thumb = img.resize((thumb_w, thumb_h))
        x = margin + col * (thumb_w + margin)
        y = margin + row * (thumb_h + margin)
        canvas.paste(thumb, (x, y))
    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(out)
    return out



def _draw_tag(draw: ImageDraw.ImageDraw, x: int, y: int, text: str, fill: str = '#EEF5FF', fg: str = '#245A8D') -> int:
    font = _font(18, bold=True)
    box_w = max(92, int(draw.textbbox((0, 0), text, font=font)[2] - draw.textbbox((0, 0), text, font=font)[0] + 34))
    draw.rounded_rectangle((x, y, x + box_w, y + 42), radius=18, fill=fill)
    draw.text((x + 16, y + 10), text, fill=fg, font=font)
    return box_w


def create_page_preview(
    title: str,
    subtitle: str,
    layout_type: str,
    score: int,
    tags: Iterable[str],
    out_path: str | Path,
    width: int = 1280,
    height: int = 720,
) -> Path:
    background = Image.new('RGB', (width, height), ImageColor.getrgb('#F7F9FC'))
    draw = ImageDraw.Draw(background)
    draw.rounded_rectangle((36, 36, width - 36, height - 36), radius=34, fill='#FFFFFF', outline='#D8E0EA', width=3)
    draw.rectangle((36, 36, width - 36, 120), fill='#1B3756')
    draw.text((84, 58), 'PPT Page Library', fill='#FFFFFF', font=_font(24, bold=True))
    draw.text((84, 92), layout_type, fill='#CFE3FA', font=_font(16, bold=True))
    draw.rounded_rectangle((width - 240, 58, width - 86, 104), radius=18, fill='#245A8D')
    draw.text((width - 212, 68), f'Score {score}', fill='#FFFFFF', font=_font(20, bold=True))
    draw.text((84, 168), title, fill='#102A43', font=_font(40, bold=True))
    draw.text((84, 232), subtitle, fill='#5B6B7A', font=_font(22))

    # Layout sketch area
    panel_left, panel_top, panel_right, panel_bottom = 84, 308, width - 84, height - 112
    draw.rounded_rectangle((panel_left, panel_top, panel_right, panel_bottom), radius=28, fill='#F9FBFD', outline='#E1E8F0', width=2)

    if layout_type in {'cover', 'cover_slide'}:
        draw.rounded_rectangle((120, 350, 620, 560), radius=26, fill='#1B3756')
        draw.rounded_rectangle((680, 350, 1140, 454), radius=22, fill='#EEF5FF')
        draw.rounded_rectangle((680, 480, 980, 560), radius=22, fill='#EAFBF9')
    elif layout_type in {'agenda', 'agenda_slide'}:
        for idx in range(4):
            x = 140 + idx * 230
            draw.ellipse((x, 386, x + 58, 444), fill='#245A8D')
            draw.rectangle((x + 28, 414, x + 164, 422), fill='#D8E0EA')
    elif layout_type in {'section', 'section_slide'}:
        draw.rounded_rectangle((120, 360, 1120, 476), radius=22, fill='#1B3756')
        draw.rounded_rectangle((120, 510, 900, 566), radius=18, fill='#EEF5FF')
    elif layout_type in {'data', 'dashboard', 'data_analysis_slide', 'kpi'}:
        for idx, h in enumerate((110, 160, 130, 190, 150)):
            x = 130 + idx * 180
            draw.rounded_rectangle((x, 520 - h, x + 90, 520), radius=16, fill='#245A8D' if idx % 2 == 0 else '#0F766E')
        draw.rounded_rectangle((760, 360, 1120, 564), radius=24, fill='#EEF5FF')
        draw.line((790, 520, 1080, 400), fill='#1B3756', width=6)
    elif layout_type in {'chart', 'chart_slide'}:
        draw.rounded_rectangle((120, 360, 560, 560), radius=24, fill='#EEF5FF')
        draw.rounded_rectangle((620, 360, 1120, 560), radius=24, fill='#FFF7ED')
        for idx, h in enumerate((96, 146, 124, 174)):
            x = 160 + idx * 86
            draw.rectangle((x, 540 - h, x + 44, 540), fill='#1B3756')
    elif layout_type in {'table', 'table_slide'}:
        for row in range(5):
            y = 350 + row * 42
            draw.line((120, y, 1120, y), fill='#D8E0EA', width=2)
        for col in range(5):
            x = 120 + col * 200
            draw.line((x, 350, x, 550), fill='#D8E0EA', width=2)
    elif layout_type in {'timeline', 'timeline_slide', 'roadmap'}:
        draw.line((150, 450, 1090, 450), fill='#245A8D', width=6)
        for idx in range(5):
            x = 170 + idx * 220
            draw.ellipse((x - 18, 432, x + 18, 468), fill='#0F766E')
            draw.rounded_rectangle((x - 70, 360 if idx % 2 == 0 else 480, x + 70, 430 if idx % 2 == 0 else 550), radius=16, fill='#EEF5FF')
    elif layout_type in {'process', 'process_slide'}:
        for idx in range(4):
            y = 360 + idx * 90
            draw.rounded_rectangle((130, y, 980, y + 56), radius=16, fill='#EEF5FF')
            draw.polygon([(1030, y + 16), (1072, y + 28), (1030, y + 40)], fill='#1B3756')
    elif layout_type in {'comparison', 'comparison_slide'}:
        draw.rounded_rectangle((120, 360, 540, 560), radius=24, fill='#EEF5FF')
        draw.rounded_rectangle((660, 360, 1120, 560), radius=24, fill='#FFF7ED')
        draw.line((620, 360, 620, 560), fill='#D8E0EA', width=4)
    elif layout_type in {'strategy', 'strategy_slide', 'swot'}:
        coords = [(150, 370), (650, 370), (150, 500), (650, 500)]
        fills = ['#EEF5FF', '#EAFBF9', '#FFF7ED', '#FCE7F3']
        for (x, y), fill in zip(coords, fills):
            draw.rounded_rectangle((x, y, x + 380, y + 92), radius=18, fill=fill)
    elif layout_type in {'people', 'team', 'people_slide'}:
        for idx in range(4):
            x = 150 + idx * 250
            draw.ellipse((x, 360, x + 120, 480), fill='#245A8D')
            draw.rounded_rectangle((x - 10, 490, x + 130, 560), radius=18, fill='#EEF5FF')
    elif layout_type in {'product', 'product_slide'}:
        draw.rounded_rectangle((120, 360, 620, 560), radius=24, fill='#FCE7F3')
        draw.rounded_rectangle((700, 360, 1120, 424), radius=18, fill='#EEF5FF')
        draw.rounded_rectangle((700, 444, 1120, 506), radius=18, fill='#EAFBF9')
        draw.rounded_rectangle((700, 526, 980, 560), radius=18, fill='#FFF7ED')
    elif layout_type in {'map', 'map_slide'}:
        draw.rounded_rectangle((150, 350, 1090, 560), radius=28, fill='#EEF5FF')
        for idx in range(8):
            x = 190 + idx * 110
            draw.ellipse((x, 400 + (idx % 3) * 30, x + 20, 420 + (idx % 3) * 30), fill='#1B3756')
    else:
        for idx in range(3):
            x = 140 + idx * 320
            draw.rounded_rectangle((x, 380, x + 240, 540), radius=22, fill='#EEF5FF')

    x = 84
    y = height - 84
    for tag in list(tags)[:7]:
        x += _draw_tag(draw, x, y, str(tag)) + 12
        if x > width - 180:
            break

    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    background.save(out)
    return out


def create_component_preview(
    title: str,
    subtitle: str,
    component_type: str,
    score: int,
    tags: Iterable[str],
    out_path: str | Path,
    width: int = 960,
    height: int = 540,
) -> Path:
    background = Image.new('RGB', (width, height), ImageColor.getrgb('#F7F9FC'))
    draw = ImageDraw.Draw(background)
    draw.rounded_rectangle((24, 24, width - 24, height - 24), radius=28, fill='#FFFFFF', outline='#D8E0EA', width=2)
    draw.rectangle((24, 24, width - 24, 92), fill='#1B3756')
    draw.text((52, 42), 'Component Library', fill='#FFFFFF', font=_font(22, bold=True))
    draw.text((52, 122), title, fill='#102A43', font=_font(34, bold=True))
    draw.text((52, 174), subtitle, fill='#5B6B7A', font=_font(18))
    draw.rounded_rectangle((width - 198, 120, width - 52, 168), radius=16, fill='#245A8D')
    draw.text((width - 164, 130), f'Score {score}', fill='#FFFFFF', font=_font(18, bold=True))
    draw.rounded_rectangle((74, 244, width - 74, height - 90), radius=22, fill='#F9FBFD', outline='#E1E8F0', width=2)

    if component_type in {'title', 'subtitle', 'text', 'caption'}:
        draw.line((120, 294, width - 140, 294), fill='#1B3756', width=8)
        draw.line((120, 338, width - 260, 338), fill='#D8E0EA', width=6)
        draw.line((120, 378, width - 180, 378), fill='#D8E0EA', width=6)
    elif component_type in {'kpi', 'number', 'metric'}:
        for idx in range(3):
            x = 120 + idx * 220
            draw.rounded_rectangle((x, 286, x + 160, 406), radius=18, fill='#EEF5FF' if idx != 1 else '#EAFBF9')
    elif component_type in {'chart', 'dashboard'}:
        draw.rounded_rectangle((130, 286, 420, 444), radius=18, fill='#EEF5FF')
        draw.rounded_rectangle((460, 286, 820, 444), radius=18, fill='#FFF7ED')
        for idx, h in enumerate((54, 98, 72, 122)):
            x = 520 + idx * 56
            draw.rectangle((x, 424 - h, x + 26, 424), fill='#1B3756')
    elif component_type in {'timeline', 'process'}:
        draw.line((140, 338, width - 150, 338), fill='#245A8D', width=5)
        for idx in range(4):
            x = 180 + idx * 180
            draw.ellipse((x - 14, 324, x + 14, 352), fill='#0F766E')
    elif component_type in {'table', 'matrix'}:
        for row in range(4):
            y = 286 + row * 34
            draw.line((120, y, width - 120, y), fill='#D8E0EA', width=2)
        for col in range(4):
            x = 120 + col * 200
            draw.line((x, 286, x, 390), fill='#D8E0EA', width=2)
    elif component_type in {'people', 'person'}:
        for idx in range(3):
            x = 160 + idx * 220
            draw.ellipse((x, 286, x + 88, 374), fill='#245A8D')
            draw.rounded_rectangle((x - 10, 388, x + 98, 432), radius=12, fill='#EEF5FF')
    else:
        draw.rounded_rectangle((120, 286, width - 120, 420), radius=18, fill='#EEF5FF')

    x = 52
    y = height - 54
    for tag in list(tags)[:6]:
        x += _draw_tag(draw, x, y, str(tag), fill='#EEF5FF', fg='#245A8D') + 10
        if x > width - 150:
            break

    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    background.save(out)
    return out
