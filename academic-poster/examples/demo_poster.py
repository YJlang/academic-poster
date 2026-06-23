"""Build one A1 poster from config.example.yaml (all dummy content).

    python demo_poster.py                  # -> demo_poster.pptx
    python demo_poster.py my.yaml out.pptx

Then render a preview to check it:
    python ../scripts/render_preview.py demo_poster.pptx
"""
import os
import sys

import yaml

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "scripts"))
from poster import Poster   # noqa: E402


def find_font():
    """Pick a common sans-serif present on the machine for measuring + display.
    Returns (font_name, ttf_path). Extend the lists for your platform/fonts."""
    candidates = [
        ("Arial", "/Library/Fonts/Arial.ttf"),
        ("Arial", "/System/Library/Fonts/Supplemental/Arial.ttf"),
        ("Helvetica Neue", "/System/Library/Fonts/HelveticaNeue.ttc"),
        ("Helvetica", "/System/Library/Fonts/Helvetica.ttc"),
        ("DejaVu Sans", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        ("Liberation Sans", "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"),
        ("Arial", "C:/Windows/Fonts/arial.ttf"),
    ]
    for name, path in candidates:
        if os.path.exists(path):
            return name, path
    return "Helvetica", None   # None -> measure.py falls back to a rough default


def main():
    cfg_path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "config.example.yaml")
    out = sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, "demo_poster.pptx")
    with open(cfg_path, encoding="utf-8") as f:
        cfg = yaml.safe_load(f)

    theme = cfg.setdefault("theme", {})
    if theme.get("font") in (None, "auto") or theme.get("font_path") in (None, "auto"):
        name, path = find_font()
        if theme.get("font") in (None, "auto"):
            theme["font"] = name
        if theme.get("font_path") in (None, "auto"):
            theme["font_path"] = path

    Poster(cfg).save(out)


if __name__ == "__main__":
    main()
