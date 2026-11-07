"""
Render assets/cover.png (1600x1200, 4:3) - the Upwork portfolio thumbnail.

Numbers come from reports/results.json. Once you have timed the two files in
Excel yourself, regenerate the cover with your Excel numbers:

    python demo/make_cover.py --recalc-before 95 --recalc-after 0.4 --recalc-label "Excel 365"

Requires Pillow. Fonts: Inter + Caladea if installed, otherwise Pillow's default.
"""

import argparse
import json
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[1]
W, H = 1600, 1200
PAPER, INK, MUTED, RULE, ACCENT, TILE = "#F4EEE3", "#1E1A15", "#6B5F52", "#BDB2A1", "#9A3F1B", "#EDE4D4"
FONTS = {
    "serif": "/usr/share/fonts/truetype/crosextra/Caladea-Bold.ttf",
    "bold": "/usr/share/fonts/opentype/inter/InterDisplay-Bold.otf",
    "semi": "/usr/share/fonts/opentype/inter/Inter-SemiBold.otf",
    "reg": "/usr/share/fonts/opentype/inter/Inter-Regular.otf",
}


def font(kind, size):
    try:
        return ImageFont.truetype(FONTS[kind], size)
    except OSError:
        return ImageFont.load_default(size)


def fmt_secs(s):
    if s >= 60:
        return f"{s / 60:.0f} min" if s >= 600 else f"{s / 60:.1f} min"
    return f"{s:.1f} s" if s < 10 else f"{s:.0f} s"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--recalc-before", type=float)
    ap.add_argument("--recalc-after", type=float)
    ap.add_argument("--recalc-label", default="LibreOffice")
    ap.add_argument("--name", default="Said Ali")
    a = ap.parse_args()
    r = json.loads((ROOT / "reports" / "results.json").read_text())
    b, f = r["before"], r["after"]
    rb = a.recalc_before if a.recalc_before is not None else b["libreoffice_recalc_s"]
    ra = a.recalc_after if a.recalc_after is not None else f["libreoffice_recalc_s"]

    img = Image.new("RGB", (W, H), PAPER)
    d = ImageDraw.Draw(img)
    m = 96

    # header rule + kicker
    d.rectangle([m, 80, W - m, 86], fill=INK)
    d.text((m, 116), "CASE STUDY  ·  EXCEL PERFORMANCE", font=font("semi", 26), fill=ACCENT)
    d.text((m, 168), "Bloated work order tracker,", font=font("serif", 88), fill=INK)
    d.text((m, 268), "repaired", font=font("serif", 88), fill=INK)
    d.text((m, 384), f"{b['data_rows']:,} work orders  ·  identical results before and after  ·  rebuilt properly",
           font=font("reg", 30), fill=MUTED)

    tiles = [
        ("File size", f"{b['size_mb']:.1f} MB", f"{f['size_mb']:.2f} MB",
         f"{(1 - f['size_mb'] / b['size_mb']) * 100:.0f}% smaller"),
        (f"Full recalculation ({a.recalc_label})", fmt_secs(rb), fmt_secs(ra), f"{rb / ra:,.0f}× faster"),
        ("Distinct cell formats", f"{b['cell_formats']:,}", f"{f['cell_formats']:,}", "one consistent style set"),
        ("Volatile formulas", f"{b['volatile_formulas']:,}", f"{f['volatile_formulas']:,}",
         "TODAY() lives in one cell"),
    ]
    gap, top, th = 32, 456, 272
    tw = (W - 2 * m - gap) // 2
    for i, (label, before, after, note) in enumerate(tiles):
        x = m + (i % 2) * (tw + gap)
        y = top + (i // 2) * (th + gap)
        d.rectangle([x, y, x + tw, y + th], fill=TILE)
        d.rectangle([x, y, x + tw, y + 6], fill=ACCENT)
        d.text((x + 36, y + 36), label.upper(), font=font("semi", 22), fill=MUTED)
        # before value, struck through
        bf = font("semi", 40)
        d.text((x + 36, y + 84), before, font=bf, fill=MUTED)
        bw = d.textlength(before, font=bf)
        d.line([x + 32, y + 110, x + 40 + bw, y + 110], fill=MUTED, width=3)
        d.text((x + 52 + bw, y + 84), "→", font=bf, fill=MUTED)
        d.text((x + 36, y + 140), after, font=font("bold", 76), fill=INK)
        d.text((x + 36, y + 234), note, font=font("reg", 24), fill=INK)

    # footer
    fy = H - 112
    d.line([m, fy, W - m, fy], fill=RULE, width=2)
    d.text((m, fy + 26), f"{a.name}  ·  Excel & Outlook automation for maintenance and ops teams",
           font=font("semi", 24), fill=INK)
    note = "Synthetic demo data"
    nf = font("reg", 22)
    d.text((W - m - d.textlength(note, font=nf), fy + 28), note, font=nf, fill=MUTED)

    out = ROOT / "assets" / "cover.png"
    out.parent.mkdir(exist_ok=True)
    img.save(out, optimize=True)
    print(out)


if __name__ == "__main__":
    main()
