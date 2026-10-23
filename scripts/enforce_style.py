#!/usr/bin/env python3
"""Enforce the User Style Preset on a generated PPTX and run layout lint.

Usage:
    python3 scripts/enforce_style.py deck.pptx [--report qa_style_report.md]

What it enforces (per the SKILL.md User Style Preset):
  1. Typography  : latin font = Arial, East Asian font = 微软雅黑 on EVERY run.
  2. Sizes       : slide-title runs -> 20 pt bold; body runs -> 18 pt.
                   Runs already >= 21 pt (cover title / hero quote) are kept.
  3. Line spacing: 1.5 on every paragraph.
  4. Source footers: text shapes starting with "Source:" are removed.
  5. Layout lint : picture-vs-text overlaps are fixed by shrinking the
                   picture; out-of-slide-bounds shapes are reported.

Idempotent: safe to run multiple times. Only python-pptx is required.
Exit code 0 when all checks pass after enforcement, 1 otherwise.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.oxml.ns import qn

FONT_LATIN = "Arial"
FONT_EA = "微软雅黑"
TITLE_SIZE = Pt(20)
BODY_SIZE = Pt(18)
LINE_SPACING = 1.5
EMU_TOL = 2000  # ~0.0002 inch tolerance for bounds checks


def _set_run_fonts(run) -> None:
    """Set latin=Arial and ea=微软雅黑 on one run (idempotent)."""
    rPr = run._r.get_or_add_rPr()
    for tag, face in (("a:latin", FONT_LATIN), ("a:ea", FONT_EA)):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", face)


def _is_title_shape(shape, slide_w) -> bool:
    """Heuristic: the slide title is a text shape near the top of the slide."""
    if not shape.has_text_frame or not shape.text_frame.text.strip():
        return False
    return shape.top is not None and shape.top < Emu(int(slide_w * 0)) + 1371600  # < 1.5 inch


def _box(shape):
    return (shape.left, shape.top, shape.left + shape.width, shape.top + shape.height)


def _overlap(a, b, tol=2000) -> bool:
    return not (a[2] <= b[0] + tol or b[2] <= a[0] + tol or
                a[3] <= b[1] + tol or b[3] <= a[1] + tol)


def enforce(path: Path, report_path: Path | None) -> int:
    prs = Presentation(str(path))
    slide_w, slide_h = prs.slide_width, prs.slide_height
    stats = {"runs": 0, "titles": 0, "bodies": 0, "kept_big": 0,
             "spacing": 0, "sources_removed": 0, "shrunk": 0}
    issues: list[str] = []

    for idx, slide in enumerate(prs.slides, 1):
        title_elems = {sh._element for sh in slide.shapes if _is_title_shape(sh, slide_w)}

        # 4. remove "Source:" footers
        for sh in list(slide.shapes):
            if sh.has_text_frame and sh.text_frame.text.strip().lower().startswith("source:"):
                sh._element.getparent().remove(sh._element)
                stats["sources_removed"] += 1

        for shape in slide.shapes:
            # --- text handling ---
            if shape.has_text_frame:
                is_title = shape._element in title_elems
                for para in shape.text_frame.paragraphs:
                    if para.line_spacing != LINE_SPACING:
                        para.line_spacing = LINE_SPACING
                        stats["spacing"] += 1
                    for run in para.runs:
                        if not run.text.strip():
                            continue
                        _set_run_fonts(run)
                        stats["runs"] += 1
                        size = run.font.size
                        if size is None:
                            run.font.size = TITLE_SIZE if is_title else BODY_SIZE
                            if is_title:
                                run.font.bold = True
                            stats["titles" if is_title else "bodies"] += 1
                        elif size.pt >= 21:
                            stats["kept_big"] += 1  # cover / hero quote: keep
                        elif 14 <= size.pt <= 16:
                            stats["kept_big"] += 1  # micro-label per preset: keep
                        elif is_title:
                            run.font.size = TITLE_SIZE
                            run.font.bold = True
                            stats["titles"] += 1
                        else:
                            run.font.size = BODY_SIZE
                            stats["bodies"] += 1

            # --- layout lint: out of bounds ---
            if shape.left is None:
                continue
            b = _box(shape)
            if b[0] < -EMU_TOL or b[1] < -EMU_TOL or \
               b[2] > slide_w + EMU_TOL or b[3] > slide_h + EMU_TOL:
                issues.append(f"slide {idx}: '{shape.name}' out of slide bounds")

        # --- layout lint: picture vs text overlap -> shrink picture ---
        pics = [sh for sh in slide.shapes if sh.shape_type == 13]
        texts = [sh for sh in slide.shapes
                 if sh.shape_type != 13 and sh.has_text_frame
                 and sh.text_frame.text.strip()]
        for pic in pics:
            orig = _box(pic)
            moved = False
            for _ in range(20):
                if not any(_overlap(_box(pic), _box(t)) for t in texts):
                    break
                pic.width = int(pic.width * 0.95)
                pic.height = int(pic.height * 0.95)
                stats["shrunk"] += 1
            if any(_overlap(_box(pic), _box(t)) for t in texts):
                # 候选位置：右侧留白带 / 底部内容区，逐个尝试
                for left, top in (
                    (slide_w - pic.width - Emu(457200), Emu(1828800)),   # 右缘留 0.5in
                    (Emu(457200), slide_h - pic.height - Emu(228600)),  # 底部留 0.25in
                ):
                    old_l, old_t = pic.left, pic.top
                    pic.left, pic.top = int(left), int(top)
                    if not any(_overlap(_box(pic), _box(t)) for t in texts):
                        moved = True
                        break
                    pic.left, pic.top = old_l, old_t
            if not moved and any(_overlap(_box(pic), _box(t)) for t in texts):
                pic.left, pic.top, pic.width, pic.height = orig  # 恢复原位，如实报告
                issues.append(f"slide {idx}: picture '{pic.name}' overlaps text and no auto-fix position found — manual layout change required")

    prs.save(str(path))

    # ----- report -----
    lines = [
        "# Style Enforcement Report",
        "",
        f"File: `{path}`",
        f"Runs re-fonted (Arial + 微软雅黑): {stats['runs']}",
        f"Title runs set to 20 pt bold: {stats['titles']}",
        f"Body runs set to 18 pt: {stats['bodies']}",
        f"Hero/cover runs kept (>=21 pt): {stats['kept_big']}",
        f"Paragraphs set to 1.5 line spacing: {stats['spacing']}",
        f"'Source:' footers removed: {stats['sources_removed']}",
        f"Picture shrink iterations (overlap fix): {stats['shrunk']}",
        "",
        "## Issues",
        "\n".join(f"- {i}" for i in issues) if issues else "- none",
    ]
    text = "\n".join(lines)
    print(text)
    if report_path:
        report_path.write_text(text + "\n", encoding="utf-8")
    return 1 if issues else 0


def main() -> int:
    ap = argparse.ArgumentParser(description="Enforce User Style Preset on a PPTX")
    ap.add_argument("pptx", type=Path)
    ap.add_argument("--report", type=Path, default=None)
    args = ap.parse_args()
    if not args.pptx.exists():
        print(f"ERROR: {args.pptx} not found", file=sys.stderr)
        return 2
    return enforce(args.pptx, args.report)


if __name__ == "__main__":
    sys.exit(main())
