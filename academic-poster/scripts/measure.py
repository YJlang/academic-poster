"""Text-height measurement via font metrics (Pillow).

The poster builder uses this to lay out sections in a *flow* (each section is
pushed down by the measured height of the one above it) instead of fixed
coordinates, which is what prevents text from overflowing into the next
section. See reference/layout-engine.md.

All sizes are in millimeters (mm) and points (pt). 1pt = 0.352778mm.
"""
from PIL import ImageFont

MM_PER_PT = 0.352778
_FONT_CACHE = {}


def _font(path, size_pt):
    """Load a font at 4x the point size (px) for finer getlength() resolution.
    `path` may be None -> Pillow's default bitmap font (rough but keeps working)."""
    key = (path, round(size_pt * 4))
    if key not in _FONT_CACHE:
        px = max(4, int(round(size_pt * 4)))
        try:
            _FONT_CACHE[key] = ImageFont.truetype(path, px) if path else ImageFont.load_default(px)
        except Exception:
            _FONT_CACHE[key] = ImageFont.load_default()
    return _FONT_CACHE[key]


def line_count(text, width_mm, size_pt, font_path):
    """How many wrapped lines `text` takes in a box `width_mm` wide.
    Greedy character wrap (works for CJK + Latin); honors explicit '\\n'."""
    f = _font(font_path, size_pt)
    max_px = (width_mm / MM_PER_PT) * 4
    n, cur = 1, 0.0
    for ch in text:
        if ch == "\n":
            n += 1
            cur = 0.0
            continue
        w = f.getlength(ch)
        if cur + w > max_px and cur > 0:
            n += 1
            cur = w
        else:
            cur += w
    return n


def text_height_mm(paragraphs, width_mm, size_pt, font_path,
                   line_spacing=1.3, para_gap_pt=4.0):
    """Rendered height (mm) of a list of paragraph strings in a column.

    line_spacing is intentionally a touch larger than the value you draw with
    (~1.27): real renderers (PowerPoint, LibreOffice) lay lines out slightly
    taller than the raw font metric, so over-estimating here keeps a safety
    margin and stops the next section from being overlapped.
    """
    h = 0.0
    for p in paragraphs:
        lines = line_count(p, width_mm, size_pt, font_path)
        h += lines * size_pt * line_spacing * MM_PER_PT
        h += para_gap_pt * MM_PER_PT
    return h


def fit_one_line_size(parts, avail_mm, font_path, max_size=24.0, sub_ratio=0.66):
    """Largest main font size (pt) that fits `parts` on ONE line within avail_mm.

    parts: list of (text, is_main). is_main=False renders at size*sub_ratio
    (e.g. subscripts in an equation). Used to make an equation fill its box.
    """
    def width(s):
        total = sum(_font(font_path, s if m else s * sub_ratio).getlength(t)
                    for t, m in parts)
        return total / 4.0 * MM_PER_PT
    base = 18.0
    w = width(base)
    if w <= 0:
        return max_size
    return max(8.0, min(max_size, base * avail_mm / w))
