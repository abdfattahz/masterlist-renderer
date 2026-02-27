import os
import math
import platform
import argparse
import pandas as pd
from PIL import Image, ImageDraw, ImageFont, ImageStat


def get_sheet_names(path: str):
    xl = pd.ExcelFile(path)
    return xl.sheet_names


def load_rows_from_sheet(
    path: str, sheet_name: str, company_col="COMPANY NAME", brn_col="COMPANY NO."
):
    df = pd.read_excel(path, sheet_name=sheet_name)
    df.columns = [c.strip() for c in df.columns]

    if company_col not in df.columns or brn_col not in df.columns:
        raise ValueError(
            f"[Sheet: {sheet_name}] Expected columns '{company_col}' and '{brn_col}'. Found: {list(df.columns)}"
        )

    df = df[[company_col, brn_col]].dropna(how="all")
    df[company_col] = df[company_col].astype(str).str.strip()
    df[brn_col] = df[brn_col].astype(str).str.strip()

    rows = list(df.itertuples(index=False, name=None))
    # return rows as list[(company, brn)]
    return rows


def pick_default_font() -> str:
    system = platform.system()

    if system == "Windows":
        win_dir = os.environ.get("WINDIR", r"C:\\Windows")
        candidates = [
            os.path.join(win_dir, "Fonts", "arial.ttf"),
            os.path.join(win_dir, "Fonts", "segoeui.ttf"),
            os.path.join(win_dir, "Fonts", "calibri.ttf"),
        ]
    elif system == "Darwin":
        candidates = [
            "/System/Library/Fonts/Supplemental/Arial.ttf",
            "/System/Library/Fonts/Supplemental/Helvetica.ttf",
            "/Library/Fonts/Arial.ttf",
        ]
    else:
        candidates = [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSansCondensed.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
            "/usr/share/fonts/truetype/freefont/FreeSans.ttf",
        ]

    for c in candidates:
        if os.path.exists(c):
            return c

    raise RuntimeError("No usable TTF font found. Pass --font /path/to/font.ttf")


def pick_default_font_linux() -> str:
    return pick_default_font()


def wrap_lines(
    draw: ImageDraw.ImageDraw,
    text: str,
    font: ImageFont.FreeTypeFont,
    max_width: int,
    max_lines: int,
):
    """
    Wrap text into up to max_lines without changing font size.
    If too long, truncate last line with ellipsis.
    """
    words = text.split()
    if not words:
        return [""]

    lines = []
    line = ""

    for w in words:
        test = (line + " " + w).strip()
        if draw.textlength(test, font=font) <= max_width:
            line = test
        else:
            if line:
                lines.append(line)
            line = w

    if line:
        lines.append(line)

    # Clamp and add ellipsis if needed
    if len(lines) > max_lines:
        lines = lines[:max_lines]
        last = lines[-1]

        while last and draw.textlength(last + "…", font=font) > max_width:
            last = last[:-1].rstrip()

        lines[-1] = (last + "…") if last else "…"

    return lines


def _clamp_color(value: float) -> int:
    return max(0, min(255, int(round(value))))


def _blend(color, target, ratio: float):
    return tuple(_clamp_color(c + (t - c) * ratio) for c, t in zip(color, target))


def _relative_luminance(color) -> float:
    r, g, b = color
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def derive_table_palette_from_background(
    bg_image: Image.Image, cell_alpha: int, header_alpha: int = 235
):
    rgb_img = bg_image.convert("RGB")
    stat = ImageStat.Stat(rgb_img)
    avg = tuple(_clamp_color(c) for c in stat.mean)

    row_light = _blend(avg, (255, 255, 255), 0.25)
    row_dark = _blend(avg, (0, 0, 0), 0.15)
    header_color = _blend(avg, (0, 0, 0), 0.35)
    border_color = _blend(header_color, (0, 0, 0), 0.4)

    base_lum = _relative_luminance(avg)
    header_lum = _relative_luminance(header_color)
    light_text = (245, 245, 245)
    dark_text = (25, 25, 25)

    body_text = light_text if base_lum < 120 else dark_text
    header_text = light_text if header_lum < 120 else dark_text

    return {
        "row_a": row_light + (cell_alpha,),
        "row_b": row_dark + (cell_alpha,),
        "header_bg": header_color + (header_alpha,),
        "border": border_color + (255,),
        "body_text": body_text,
        "header_text": header_text,
    }


def render_streamed_pages(
    rows_stream,
    total_rows: int,
    out_dir: str,
    width: int = 1920,
    height: int = 1080,
    font_path: str | None = None,
    background_path: str | None = None,
    pairs_per_row: int = 3,
    rows_per_page: int = 18,
    margin: int = 0,
    gutter: int = 0,
    header_h: int = 70,
    name_ratio: float = 0.72,
    cell_alpha: int = 175,
    border_width: int = 1,
    body_font_size: int = 14,
    header_font_size: int = 18,
    body_text_color=(20, 0, 0),
    header_text_color=(255, 255, 255),
    row_a_color: tuple[int, int, int] | None = None,
    row_b_color: tuple[int, int, int] | None = None,
    header_bg_color: tuple[int, int, int] | None = None,
    border_color: tuple[int, int, int] | None = None,
    match_table_to_bg: bool = False,
    allow_auto_body_text: bool = True,
    allow_auto_header_text: bool = True,
    progress_callback=None,
):
    """
    Renders pages from a *single continuous stream* of rows.
    It fills page slots fully before creating a new page.
    """
    os.makedirs(out_dir, exist_ok=True)

    if font_path is None:
        font_path = pick_default_font()

    header_font = ImageFont.truetype(font_path, header_font_size)
    company_font = ImageFont.truetype(font_path, body_font_size)
    brn_font = ImageFont.truetype(font_path, body_font_size)

    # Background
    if background_path and os.path.exists(background_path):
        base_bg = Image.open(background_path).convert("RGB").resize((width, height))
    else:
        base_bg = Image.new("RGB", (width, height), (255, 255, 255))
    base_bg = base_bg.convert("RGBA")

    usable_w = width - 2 * margin
    usable_h = height - 2 * margin
    body_h = usable_h - header_h

    pair_w = (usable_w - (pairs_per_row - 1) * gutter) // pairs_per_row
    name_w = int(pair_w * name_ratio)
    brn_w = pair_w - name_w

    base_h = body_h // rows_per_page
    extra = body_h % rows_per_page  # leftover pixels

    # Heights list sums EXACTLY to body_h
    row_heights = [base_h + (1 if i < extra else 0) for i in range(rows_per_page)]

    # y positions for each row start (top of each row)
    row_tops = []
    y = margin + header_h
    for h in row_heights:
        row_tops.append(y)
        y += h

    per_page = pairs_per_row * rows_per_page
    total_pages = math.ceil(total_rows / per_page)

    # Colors
    base_row_a = row_a_color or (243, 166, 166)
    base_row_b = row_b_color or (232, 126, 126)
    base_header_bg = header_bg_color or (180, 40, 40)
    base_border = border_color or (20, 0, 0)

    palette = {
        "row_a": (*base_row_a, cell_alpha),
        "row_b": (*base_row_b, cell_alpha),
        "header_bg": (*base_header_bg, 235),
        "border": (*base_border, 255),
        "body_text": body_text_color,
        "header_text": header_text_color,
    }

    if match_table_to_bg and background_path and os.path.exists(background_path):
        auto_palette = derive_table_palette_from_background(base_bg, cell_alpha)
        palette["row_a"] = auto_palette["row_a"]
        palette["row_b"] = auto_palette["row_b"]
        palette["header_bg"] = auto_palette["header_bg"]
        palette["border"] = auto_palette["border"]
        if allow_auto_body_text:
            palette["body_text"] = auto_palette["body_text"]
        if allow_auto_header_text:
            palette["header_text"] = auto_palette["header_text"]

    if row_a_color is not None:
        palette["row_a"] = (*row_a_color, cell_alpha)
    if row_b_color is not None:
        palette["row_b"] = (*row_b_color, cell_alpha)
    if header_bg_color is not None:
        palette["header_bg"] = (*header_bg_color, 235)
    if border_color is not None:
        palette["border"] = (*border_color, 255)

    row_a = palette["row_a"]
    row_b = palette["row_b"]
    header_bg = palette["header_bg"]
    border = palette["border"]
    text = (*palette["body_text"], 255)
    header_text = (*palette["header_text"], 255)

    def header_rect(pair_i, is_brn: bool):
        x0 = margin + pair_i * (pair_w + gutter) + (name_w if is_brn else 0)
        x1 = x0 + (brn_w if is_brn else name_w)
        y0 = margin
        y1 = margin + header_h
        return x0, y0, x1, y1

    def cell_rect(pair_i, row_i, is_brn: bool):
        x0 = margin + pair_i * (pair_w + gutter) + (name_w if is_brn else 0)
        x1 = x0 + (brn_w if is_brn else name_w)

        y0 = row_tops[row_i]
        y1 = y0 + row_heights[row_i]

        return x0, y0, x1, y1

    # Consume the stream page-by-page
    for page in range(1, total_pages + 1):
        img = base_bg.copy()  # RGBA

        overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
        odraw = ImageDraw.Draw(overlay)

        # Header blocks
        for p in range(pairs_per_row):
            for is_brn, label in [(False, "COMPANY NAME"), (True, "BRN NO")]:
                x0, y0, x1, y1 = header_rect(p, is_brn)
                odraw.rectangle(
                    [x0, y0, x1, y1], fill=header_bg, outline=border, width=border_width
                )

        # Pull exactly per_page rows from the stream (or less on last page)
        page_rows = []
        for _ in range(per_page):
            try:
                page_rows.append(next(rows_stream))
            except StopIteration:
                break

        # Draw body rectangles first (translucent)
        for i, (company, brn) in enumerate(page_rows):
            row_i = i // pairs_per_row
            pair_i = i % pairs_per_row
            bg = row_a if (row_i % 2 == 0) else row_b

            x0, y0, x1, y1 = cell_rect(pair_i, row_i, is_brn=False)
            xb0, yb0, xb1, yb1 = cell_rect(pair_i, row_i, is_brn=True)

            odraw.rectangle(
                [x0, y0, x1, y1], fill=bg, outline=border, width=border_width
            )
            odraw.rectangle(
                [xb0, yb0, xb1, yb1], fill=bg, outline=border, width=border_width
            )

        img = Image.alpha_composite(img, overlay)
        draw = ImageDraw.Draw(img)

        # Header text
        for p in range(pairs_per_row):
            for is_brn, label in [(False, "COMPANY NAME"), (True, "BRN NO")]:
                x0, y0, x1, y1 = header_rect(p, is_brn)
                tw = draw.textlength(label, font=header_font)
                th = header_font.size
                draw.text(
                    (x0 + (x1 - x0 - tw) / 2, y0 + (y1 - y0 - th) / 2),
                    label,
                    font=header_font,
                    fill=header_text,
                )

        # Body text (solid)
        pad = 10
        for i, (company, brn) in enumerate(page_rows):
            row_i = i // pairs_per_row
            pair_i = i % pairs_per_row

            x0, y0, x1, y1 = cell_rect(pair_i, row_i, is_brn=False)
            xb0, yb0, xb1, yb1 = cell_rect(pair_i, row_i, is_brn=True)

            company_lines = wrap_lines(
                draw, company, company_font, (x1 - x0 - 2 * pad), 2
            )
            brn_lines = wrap_lines(draw, brn, brn_font, (xb1 - xb0 - 2 * pad), 1)
            brn_line = brn_lines[0] if brn_lines else brn

            # Center company
            line_gap = 2
            total_h = len(company_lines) * (company_font.size + line_gap) - line_gap
            cy = y0 + ((y1 - y0) - total_h) / 2
            for line in company_lines:
                tw = draw.textlength(line, font=company_font)
                draw.text(
                    (x0 + (x1 - x0 - tw) / 2, cy), line, font=company_font, fill=text
                )
                cy += company_font.size + line_gap

            # Center BRN
            tw = draw.textlength(brn_line, font=brn_font)
            draw.text(
                (xb0 + (xb1 - xb0 - tw) / 2, yb0 + ((yb1 - yb0) - brn_font.size) / 2),
                brn_line,
                font=brn_font,
                fill=text,
            )

        out_path = os.path.join(out_dir, f"masterlist_{page:02d}.png")
        # rotate image 90° clockwise
        rotated = img.transpose(Image.Transpose.ROTATE_270)

        rotated.convert("RGB").save(out_path, "PNG", optimize=True)

        if callable(progress_callback):
            progress_callback(page, total_pages)

    return total_pages


def parse_rgb(value: str):
    try:
        parts = [p.strip() for p in value.split(",")]
        if len(parts) != 3:
            raise ValueError

        r, g, b = map(int, parts)
        if any(c < 0 or c > 255 for c in (r, g, b)):
            raise ValueError

        return (r, g, b)
    except Exception:
        raise argparse.ArgumentTypeError(
            "Color must be in format R,G,B (e.g. 255,255,255)"
        )


def load_all_rows(path: str):
    sheet_names = get_sheet_names(path)
    all_rows = []

    for sheet in sheet_names:
        all_rows.extend(load_rows_from_sheet(path, sheet_name=sheet))

    return all_rows


def run_render_process(
    excel_path: str,
    out_dir: str = "output",
    bg_path: str | None = None,
    font_path: str | None = None,
    pairs: int = 3,
    rows: int = 18,
    alpha: int = 175,
    font_size: int = 14,
    header_font_size: int = 18,
    text_color: str = "20,0,0",
    header_text_color: str = "255,255,255",
    row_a_color: str | None = None,
    row_b_color: str | None = None,
    header_bg_color: str | None = None,
    border_color: str | None = None,
    match_table_to_bg: bool = False,
    progress_callback=None,
):
    default_text_color = "20,0,0"
    default_header_text_color = "255,255,255"

    all_rows = load_all_rows(excel_path)
    total_rows = len(all_rows)

    body_text_color = parse_rgb(text_color)
    header_text_color_value = parse_rgb(header_text_color)
    row_a_color_value = parse_rgb(row_a_color) if row_a_color else None
    row_b_color_value = parse_rgb(row_b_color) if row_b_color else None
    header_bg_color_value = parse_rgb(header_bg_color) if header_bg_color else None
    border_color_value = parse_rgb(border_color) if border_color else None
    text_color_overridden = text_color != default_text_color
    header_text_color_overridden = header_text_color != default_header_text_color

    def row_generator():
        for row in all_rows:
            yield row

    pages = render_streamed_pages(
        rows_stream=row_generator(),
        total_rows=total_rows,
        out_dir=out_dir,
        width=1080,
        height=1920,
        font_path=font_path,
        background_path=bg_path,
        pairs_per_row=pairs,
        rows_per_page=rows,
        margin=0,
        gutter=0,
        header_h=70,
        name_ratio=0.72,
        cell_alpha=alpha,
        border_width=1,
        body_font_size=font_size,
        header_font_size=header_font_size,
        body_text_color=body_text_color,
        header_text_color=header_text_color_value,
        row_a_color=row_a_color_value,
        row_b_color=row_b_color_value,
        header_bg_color=header_bg_color_value,
        border_color=border_color_value,
        match_table_to_bg=match_table_to_bg,
        allow_auto_body_text=not text_color_overridden,
        allow_auto_header_text=not header_text_color_overridden,
        progress_callback=progress_callback,
    )

    return pages, total_rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", required=True, help="Path to masterlist Excel file")
    ap.add_argument("--out", default="output", help="Output folder for PNGs")
    ap.add_argument("--bg", default=None, help="Optional background image")
    ap.add_argument("--font", default=None, help="Optional TTF font path")
    ap.add_argument(
        "--pairs", type=int, default=3, help="How many (Name+BRN) pairs across"
    )
    ap.add_argument("--rows", type=int, default=18, help="Rows per page")
    ap.add_argument("--alpha", type=int, default=175, help="Cell opacity 0-255")
    ap.add_argument("--font_size", type=int, default=14, help="Body font size (fixed)")
    ap.add_argument("--header_font_size", type=int, default=18, help="Header font size")
    ap.add_argument(
        "--text_color",
        default="20,0,0",
        help="Body text RGB, e.g. 0,0,0 or 255,255,255",
    )
    ap.add_argument(
        "--header_text_color",
        default="255,255,255",
        help="Header text RGB, e.g. 255,255,255",
    )
    ap.add_argument(
        "--match_table_to_bg",
        action="store_true",
        help="Automatically derive table and text colors from the background image",
    )
    args = ap.parse_args()

    pages, _ = run_render_process(
        excel_path=args.excel,
        out_dir=args.out,
        bg_path=args.bg,
        font_path=args.font,
        pairs=args.pairs,
        rows=args.rows,
        alpha=args.alpha,
        font_size=args.font_size,
        header_font_size=args.header_font_size,
        text_color=args.text_color,
        header_text_color=args.header_text_color,
        match_table_to_bg=args.match_table_to_bg,
    )

    print(f"Done. Generated {pages} page(s) into: {args.out}")


if __name__ == "__main__":
    main()
