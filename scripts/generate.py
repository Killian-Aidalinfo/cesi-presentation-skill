#!/usr/bin/env python3
"""Generate an editable PPTX and/or a separately rendered PDF from a JSON deck."""
import argparse
import json
import sys
from pathlib import Path

# Indicative, not official CESI identity.
NAVY = '#000000'
BLUE = '#FBE216'
INK = '#000000'
WHITE = '#FFFFFF'
W, H = 960, 540


def validate(data):
    if not isinstance(data, dict) or not isinstance(data.get('slides'), list) or not data['slides']:
        raise ValueError('slides must be a non-empty array')
    if not isinstance(data.get('footer', ''), str) or len(data.get('footer', '')) > 70:
        raise ValueError('footer must be a string of at most 70 characters')
    for i, s in enumerate(data['slides'], 1):
        if not isinstance(s, dict):
            raise ValueError(f'slide {i}: must be an object')
        if s.get('type', 'content') not in ('title', 'content', 'section'):
            raise ValueError(f'slide {i}: type must be title, section or content')
        for key, limit in [('title', 80), ('subtitle', 110)]:
            value = s.get(key, '')
            if not isinstance(value, str) or len(value) > limit or '\n' in value:
                raise ValueError(f'slide {i}: {key} must be a single-line string of at most {limit} characters')
        if not s.get('title', '').strip():
            raise ValueError(f'slide {i}: title is required')
        bullets = s.get('bullets', [])
        if not isinstance(bullets, list) or len(bullets) > 5 or any(not isinstance(b, str) or not b.strip() or len(b) > 120 or '\n' in b for b in bullets):
            raise ValueError(f'slide {i}: bullets must be an array of at most 5 non-empty single-line strings, each at most 120 characters')
        if s.get('type', 'content') != 'content' and bullets:
            raise ValueError(f'slide {i}: bullets are supported only on content slides')
    return data


def layout(slide, number, footer):
    """Return simple vector primitives in 960x540 top-left coordinates."""
    dark = slide.get('type', 'content') in ('title', 'section')
    p = [('rect', 0, 0, W, H, NAVY if dark else WHITE),
         ('rect', 0, 0, 12, H, BLUE)]
    if dark:
        p += [('rect', 64, 123, 75, 6, BLUE),
              ('text', 64, 162, 820, 84, slide['title'], 38, WHITE, True),
              ('text', 64, 275, 820, 55, slide.get('subtitle', ''), 21, '#D8E9ED', False)]
    else:
        p += [('rect', 0, 0, W, 14, BLUE),
              ('text', 64, 49, 830, 73, slide['title'], 30, NAVY, True),
              ('rect', 64, 127, 74, 4, BLUE)]
        y = 157
        if slide.get('subtitle'):
            p.append(('text', 64, y, 815, 47, slide['subtitle'], 17, INK, False))
            y += 55
        for bullet in slide.get('bullets', []):
            p += [('rect', 66, y + 9, 7, 7, BLUE), ('text', 88, y, 792, 48, bullet, 20, INK, False)]
            y += 53
    p.append(('text', 64, 505, 750, 22, footer, 10, '#D8E9ED' if dark else INK, False))
    p.append(('text', 895, 505, 35, 22, str(number), 10, '#D8E9ED' if dark else INK, False))
    return p


def pptx_out(data, path):
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR
    r = Presentation(); r.slide_width = Inches(13.333333); r.slide_height = Inches(7.5)
    unit = 72  # points per inch, and coordinates use points
    for n, slide in enumerate(data['slides'], 1):
        page = r.slides.add_slide(r.slide_layouts[6])
        for item in layout(slide, n, data.get('footer', '')):
            kind, x, y, w, h, *rest = item
            if kind == 'rect':
                shp = page.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(x), Pt(y), Pt(w), Pt(h))
                shp.fill.solid(); shp.fill.fore_color.rgb = RGBColor.from_string(rest[0][1:]); shp.line.fill.background()
            else:
                text, size, color, bold = rest
                if not text: continue
                box = page.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
                frame = box.text_frame; frame.clear(); frame.word_wrap = True
                frame.margin_left = frame.margin_right = frame.margin_top = frame.margin_bottom = 0
                frame.vertical_anchor = MSO_ANCHOR.TOP
                para = frame.paragraphs[0]; para.text = text
                para.font.name = 'Arial'; para.font.size = Pt(size); para.font.bold = bold
                para.font.color.rgb = RGBColor.from_string(color[1:])
    r.save(str(path))


def pdf_out(data, path):
    from reportlab.pdfgen import canvas
    from reportlab.lib.colors import HexColor
    from reportlab.pdfbase.pdfmetrics import stringWidth
    c = canvas.Canvas(str(path), pagesize=(W, H))
    c.setTitle(data['slides'][0]['title'])
    for n, slide in enumerate(data['slides'], 1):
        for item in layout(slide, n, data.get('footer', '')):
            kind, x, y, w, h, *rest = item
            if kind == 'rect':
                c.setFillColor(HexColor(rest[0])); c.rect(x, H-y-h, w, h, fill=1, stroke=0)
            else:
                text, size, color, bold = rest
                font = 'Helvetica-Bold' if bold else 'Helvetica'
                c.setFillColor(HexColor(color)); c.setFont(font, size)
                # Word-wrap against the same text box. No artificial truncation.
                lines = []
                for word in text.split():
                    candidate = (lines[-1] + ' ' + word) if lines else word
                    if lines and stringWidth(candidate, font, size) <= w: lines[-1] = candidate
                    else: lines.append(word)
                if len(lines) * size * 1.15 > h + size * .15:
                    raise ValueError(f'slide {n}: text exceeds its box; shorten: {text[:45]}')
                for j, line in enumerate(lines):
                    if stringWidth(line, font, size) > w:
                        raise ValueError(f'slide {n}: word exceeds its box; shorten: {line[:45]}')
                    c.drawString(x, H-y-size-j*size*1.15, line)
        c.showPage()
    c.save()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('spec', type=Path, help='UTF-8 JSON deck')
    parser.add_argument('--pptx', type=Path, help='output PPTX path')
    parser.add_argument('--pdf', type=Path, help='output PDF path (independently rendered, not a PPTX conversion)')
    args = parser.parse_args()
    if not (args.pptx or args.pdf): parser.error('provide --pptx and/or --pdf')
    outputs = [p for p in (args.pptx, args.pdf) if p]
    if any(p.resolve() == args.spec.resolve() for p in outputs) or len({p.resolve() for p in outputs}) != len(outputs):
        parser.error('input and output paths must be distinct')
    if args.pptx and args.pptx.suffix.lower() != '.pptx': parser.error('--pptx must end in .pptx')
    if args.pdf and args.pdf.suffix.lower() != '.pdf': parser.error('--pdf must end in .pdf')
    try:
        data = validate(json.loads(args.spec.read_text(encoding='utf-8')))
        for p in outputs: p.parent.mkdir(parents=True, exist_ok=True)
        # Render PDF first so layout errors do not leave an apparently successful PPTX.
        if args.pdf: pdf_out(data, args.pdf)
        if args.pptx: pptx_out(data, args.pptx)
    except (OSError, ValueError, ImportError) as e:
        print(f'error: {e}', file=sys.stderr); return 1
    print(json.dumps({'slides': len(data['slides']), 'pptx': str(args.pptx) if args.pptx else None, 'pdf': str(args.pdf) if args.pdf else None}, ensure_ascii=False))
    return 0

if __name__ == '__main__': sys.exit(main())
