from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path
from tempfile import TemporaryDirectory

import xlrd
import xlwt
from PIL import Image, ImageDraw


MAIN_TABLE_LAST_COL = 8
CURRENT_FORMULA_COL = 9
CURRENT_FORMULA_WIDTH = 2816
TEMPLATE_LAST_COL = 8

CLASSIC_PALETTE = {
    "navy": (0x21, (31, 53, 79)),
    "slate": (0x22, (69, 104, 142)),
    "mist": (0x23, (223, 235, 247)),
    "ivory": (0x24, (248, 250, 252)),
    "gold": (0x25, (214, 178, 61)),
    "coral": (0x26, (194, 87, 87)),
    "sage": (0x27, (129, 161, 135)),
    "stone": (0x28, (161, 171, 181)),
}

CLANCY_PALETTE = {
    "navy": (0x21, (36, 46, 55)),
    "slate": (0x22, (41, 52, 62)),
    "mist": (0x23, (242, 245, 249)),
    "ivory": (0x24, (238, 238, 238)),
    "gold": (0x25, (255, 96, 0)),
    "coral": (0x26, (255, 96, 0)),
    "sage": (0x27, (62, 62, 62)),
    "stone": (0x28, (190, 190, 190)),
}

CLANCY_LIGHT_PALETTE = {
    "navy": (0x21, (36, 46, 55)),
    "slate": (0x22, (41, 52, 62)),
    "mist": (0x23, (242, 245, 249)),
    "ivory": (0x24, (255, 255, 255)),
    "gold": (0x25, (255, 96, 0)),
    "coral": (0x26, (255, 212, 141)),
    "sage": (0x27, (120, 120, 120)),
    "stone": (0x28, (190, 190, 190)),
}

CORPORATE_OLD_SCHOOL_PALETTE = {
    "navy": (0x21, (31, 78, 121)),
    "slate": (0x22, (79, 129, 189)),
    "mist": (0x23, (220, 230, 241)),
    "ivory": (0x24, (242, 242, 242)),
    "gold": (0x25, (184, 204, 228)),
    "coral": (0x26, (217, 225, 242)),
    "sage": (0x27, (89, 89, 89)),
    "stone": (0x28, (166, 166, 166)),
}

RETRO_90S_PALETTE = {
    "navy": (0x21, (0, 0, 128)),
    "slate": (0x22, (0, 128, 128)),
    "mist": (0x23, (192, 192, 192)),
    "ivory": (0x24, (236, 233, 216)),
    "gold": (0x25, (255, 255, 0)),
    "coral": (0x26, (0, 255, 255)),
    "sage": (0x27, (0, 0, 0)),
    "stone": (0x28, (128, 128, 128)),
}

VEGAS_CASINO_PALETTE = {
    "navy": (0x21, (46, 0, 74)),
    "slate": (0x22, (124, 0, 145)),
    "mist": (0x23, (255, 230, 250)),
    "ivory": (0x24, (36, 0, 46)),
    "gold": (0x25, (255, 215, 0)),
    "coral": (0x26, (255, 20, 147)),
    "sage": (0x27, (255, 255, 255)),
    "stone": (0x28, (255, 105, 180)),
}

CYBERPUNK_PALETTE = {
    "navy": (0x21, (10, 10, 28)),
    "slate": (0x22, (36, 14, 76)),
    "mist": (0x23, (17, 245, 255)),
    "ivory": (0x24, (24, 12, 38)),
    "gold": (0x25, (255, 0, 153)),
    "coral": (0x26, (255, 255, 0)),
    "sage": (0x27, (240, 240, 240)),
    "stone": (0x28, (120, 75, 220)),
}

HIGH_CONTRAST_PALETTE = {
    "navy": (0x21, (0, 0, 0)),
    "slate": (0x22, (255, 255, 255)),
    "mist": (0x23, (255, 255, 0)),
    "ivory": (0x24, (255, 255, 255)),
    "gold": (0x25, (255, 0, 0)),
    "coral": (0x26, (0, 255, 255)),
    "sage": (0x27, (0, 0, 0)),
    "stone": (0x28, (0, 0, 0)),
}

MEDIEVAL_MANUSCRIPT_PALETTE = {
    "navy": (0x21, (73, 46, 26)),
    "slate": (0x22, (120, 86, 52)),
    "mist": (0x23, (232, 214, 176)),
    "ivory": (0x24, (245, 235, 207)),
    "gold": (0x25, (168, 123, 29)),
    "coral": (0x26, (129, 32, 32)),
    "sage": (0x27, (58, 41, 23)),
    "stone": (0x28, (112, 93, 74)),
}

TREASURE_MAP_PALETTE = {
    "navy": (0x21, (44, 78, 55)),
    "slate": (0x22, (74, 108, 78)),
    "mist": (0x23, (214, 201, 141)),
    "ivory": (0x24, (240, 228, 181)),
    "gold": (0x25, (163, 109, 26)),
    "coral": (0x26, (143, 76, 35)),
    "sage": (0x27, (60, 68, 38)),
    "stone": (0x28, (111, 99, 66)),
}

SYNTHWAVE_OUTRUN_PALETTE = {
    "navy": (0x21, (30, 12, 58)),
    "slate": (0x22, (78, 25, 119)),
    "mist": (0x23, (50, 236, 255)),
    "ivory": (0x24, (25, 13, 45)),
    "gold": (0x25, (255, 94, 194)),
    "coral": (0x26, (255, 196, 0)),
    "sage": (0x27, (241, 238, 255)),
    "stone": (0x28, (108, 58, 191)),
}

BRUTALIST_MONOCHROME_PALETTE = {
    "navy": (0x21, (18, 18, 18)),
    "slate": (0x22, (48, 48, 48)),
    "mist": (0x23, (235, 235, 235)),
    "ivory": (0x24, (255, 255, 255)),
    "gold": (0x25, (0, 0, 0)),
    "coral": (0x26, (210, 210, 210)),
    "sage": (0x27, (20, 20, 20)),
    "stone": (0x28, (120, 120, 120)),
}

STAR_WARS_PALETTE = {
    "navy": (0x21, (12, 12, 32)),
    "slate": (0x22, (33, 33, 74)),
    "mist": (0x23, (255, 232, 110)),
    "ivory": (0x24, (20, 20, 20)),
    "gold": (0x25, (255, 214, 10)),
    "coral": (0x26, (83, 210, 255)),
    "sage": (0x27, (214, 214, 214)),
    "stone": (0x28, (90, 90, 120)),
}


@dataclass(frozen=True)
class ThemeProfile:
    name: str
    palette: dict[str, tuple[int, tuple[int, int, int]]]
    title_font_name: str = "Arial"
    heading_font_name: str = "Arial"
    body_font_name: str = "Arial"


THEMES = {
    "classic": ThemeProfile(name="classic", palette=CLASSIC_PALETTE),
    "clancy": ThemeProfile(
        name="clancy",
        palette=CLANCY_PALETTE,
        title_font_name="Exo",
        heading_font_name="Exo",
        body_font_name="Arial",
    ),
    "clancy-light": ThemeProfile(
        name="clancy-light",
        palette=CLANCY_LIGHT_PALETTE,
        title_font_name="Exo",
        heading_font_name="Exo",
        body_font_name="Arial",
    ),
    "corporate-oldschool": ThemeProfile(
        name="corporate-oldschool",
        palette=CORPORATE_OLD_SCHOOL_PALETTE,
        title_font_name="Arial",
        heading_font_name="Arial",
        body_font_name="Arial",
    ),
    "retro-90s": ThemeProfile(
        name="retro-90s",
        palette=RETRO_90S_PALETTE,
        title_font_name="Arial",
        heading_font_name="Arial",
        body_font_name="Arial",
    ),
    "vegas-casino": ThemeProfile(
        name="vegas-casino",
        palette=VEGAS_CASINO_PALETTE,
        title_font_name="Arial Black",
        heading_font_name="Arial Black",
        body_font_name="Arial",
    ),
    "cyberpunk": ThemeProfile(
        name="cyberpunk",
        palette=CYBERPUNK_PALETTE,
        title_font_name="Arial Black",
        heading_font_name="Arial Black",
        body_font_name="Arial",
    ),
    "high-contrast": ThemeProfile(
        name="high-contrast",
        palette=HIGH_CONTRAST_PALETTE,
        title_font_name="Arial Black",
        heading_font_name="Arial Black",
        body_font_name="Arial",
    ),
    "medieval-manuscript": ThemeProfile(
        name="medieval-manuscript",
        palette=MEDIEVAL_MANUSCRIPT_PALETTE,
        title_font_name="Book Antiqua",
        heading_font_name="Book Antiqua",
        body_font_name="Times New Roman",
    ),
    "fantasy-forest-map": ThemeProfile(
        name="fantasy-forest-map",
        palette=TREASURE_MAP_PALETTE,
        title_font_name="Book Antiqua",
        heading_font_name="Book Antiqua",
        body_font_name="Georgia",
    ),
    "synthwave-outrun": ThemeProfile(
        name="synthwave-outrun",
        palette=SYNTHWAVE_OUTRUN_PALETTE,
        title_font_name="Arial Black",
        heading_font_name="Arial Black",
        body_font_name="Arial",
    ),
    "brutalist-monochrome": ThemeProfile(
        name="brutalist-monochrome",
        palette=BRUTALIST_MONOCHROME_PALETTE,
        title_font_name="Arial Black",
        heading_font_name="Arial Black",
        body_font_name="Arial",
    ),
    "star-wars": ThemeProfile(
        name="star-wars",
        palette=STAR_WARS_PALETTE,
        title_font_name="Arial Black",
        heading_font_name="Arial Black",
        body_font_name="Arial",
    ),
}


@dataclass(frozen=True)
class ThemeStyle:
    font_name: str = "Arial"
    font_height: int = 200
    bold: bool = False
    italic: bool = False
    font_colour: str = "black"
    bg_colour: str | None = None
    border_colour: str = "stone"
    left: int = 1
    right: int = 1
    top: int = 1
    bottom: int = 1
    horz: int = xlwt.Alignment.HORZ_LEFT
    vert: int = xlwt.Alignment.VERT_CENTER
    wrap: bool = True


class StyleFactory:
    def __init__(self, workbook: xlwt.Workbook, palette: dict[str, tuple[int, tuple[int, int, int]]]) -> None:
        self.workbook = workbook
        self.palette = palette
        self.cache: dict[tuple[str, str], xlwt.XFStyle] = {}

    def color_index(self, color_name: str) -> int:
        if color_name == "black":
            return 0x08
        if color_name == "white":
            return 0x09
        return self.palette[color_name][0]

    def make(self, spec: ThemeStyle, number_format: str = "General") -> xlwt.XFStyle:
        key = (repr(spec), number_format)
        if key in self.cache:
            return self.cache[key]

        font = xlwt.Font()
        font.name = spec.font_name
        font.height = spec.font_height
        font.bold = spec.bold
        font.italic = spec.italic
        font.colour_index = self.color_index(spec.font_colour)

        alignment = xlwt.Alignment()
        alignment.horz = spec.horz
        alignment.vert = spec.vert
        alignment.wrap = 1 if spec.wrap else 0

        pattern = xlwt.Pattern()
        if spec.bg_colour:
            pattern.pattern = xlwt.Pattern.SOLID_PATTERN
            pattern.pattern_fore_colour = self.color_index(spec.bg_colour)
        else:
            pattern.pattern = xlwt.Pattern.NO_PATTERN

        borders = xlwt.Borders()
        borders.left = spec.left
        borders.right = spec.right
        borders.top = spec.top
        borders.bottom = spec.bottom
        colour_index = self.color_index(spec.border_colour)
        borders.left_colour = colour_index
        borders.right_colour = colour_index
        borders.top_colour = colour_index
        borders.bottom_colour = colour_index

        style = xlwt.XFStyle()
        style.font = font
        style.alignment = alignment
        style.pattern = pattern
        style.borders = borders
        style.num_format_str = number_format

        self.cache[key] = style
        return style


def configure_palette(workbook: xlwt.Workbook, palette: dict[str, tuple[int, tuple[int, int, int]]]) -> None:
    for _, (index, rgb) in palette.items():
        workbook.set_colour_RGB(index, *rgb)


def content_bounds(sheet: xlrd.sheet.Sheet) -> tuple[int, int]:
    max_row = 0
    max_col = 0
    for rowx in range(sheet.nrows):
        for colx in range(sheet.ncols):
            if sheet.cell(rowx, colx).ctype not in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                max_row = max(max_row, rowx)
                max_col = max(max_col, colx)
    for row_lo, row_hi, col_lo, col_hi in sheet.merged_cells:
        max_row = max(max_row, row_hi - 1)
        max_col = max(max_col, col_hi - 1)
    return max_row, max_col


def format_for_cell(book: xlrd.book.Book, sheet: xlrd.sheet.Sheet, rowx: int, colx: int) -> str:
    xf = book.xf_list[sheet.cell_xf_index(rowx, colx)]
    fmt = book.format_map.get(xf.format_key)
    return fmt.format_str if fmt else "General"


def normalize_status(value: object) -> str:
    if isinstance(value, str):
        return value.strip().lower()
    return ""


def row_role(sheet_name: str, sheet: xlrd.sheet.Sheet, rowx: int) -> str:
    col0 = str(sheet.cell_value(rowx, 0)).strip() if sheet.ncols > 0 else ""
    col1 = str(sheet.cell_value(rowx, 1)).strip() if sheet.ncols > 1 else ""
    row_values = [sheet.cell_value(rowx, c) for c in range(sheet.ncols)]

    if sheet_name == "Compatibility Report":
        if rowx in {0, 1}:
            return "compat_title"
        if rowx == 5:
            return "compat_header"
        if rowx in {3, 7}:
            return "compat_text"
        return "compat_body"

    if sheet_name == "Current" and rowx == 1:
        return "title"
    if sheet_name == "Template" and rowx == 6:
        return "title"

    if col0 in {"Total Sent:", "Total Rec:", "Not Found:"}:
        return col0.lower().replace(":", "").replace(" ", "_")

    if col0 and not col1 and (
        "TOTAL" in [str(v) for v in row_values]
        or all(str(v).strip() == "" for v in row_values[1:])
    ):
        return "section_header"

    if col0.lower().startswith("no need"):
        return "note"

    status = normalize_status(col1)
    if status == "sent":
        return "sent"
    if status == "rec'd":
        return "received"
    return "body"


def should_skip_row(sheet_name: str, sheet: xlrd.sheet.Sheet, rowx: int) -> bool:
    if sheet_name not in {"Current", "Template"}:
        return False

    col0 = str(sheet.cell_value(rowx, 0)).strip().lower() if sheet.ncols > 0 else ""
    col1 = normalize_status(sheet.cell_value(rowx, 1)) if sheet.ncols > 1 else ""
    prev_col0 = str(sheet.cell_value(rowx - 1, 0)).strip().lower() if rowx > 0 and sheet.ncols > 0 else ""

    if "republic parking" in col0 or col0.startswith("no need to do republic"):
        return True
    return col1 == "rec'd" and (
        "republic parking" in prev_col0 or prev_col0.startswith("no need to do republic")
    )


def retained_rows(sheet_name: str, sheet: xlrd.sheet.Sheet) -> list[int]:
    return [rowx for rowx in range(sheet.nrows) if not should_skip_row(sheet_name, sheet, rowx)]


def style_spec_for_cell(
    theme: ThemeProfile,
    sheet_name: str,
    role: str,
    rowx: int,
    colx: int,
    value: object,
    fmt: str,
) -> ThemeStyle:
    is_clancy_dark = theme.name == "clancy"
    is_clancy_light = theme.name == "clancy-light"
    is_oldschool = theme.name == "corporate-oldschool"
    is_star_wars = theme.name == "star-wars"
    is_windows_90s = theme.name == "retro-90s"
    is_vegas = theme.name == "vegas-casino"
    is_cyberpunk = theme.name == "cyberpunk"
    is_high_contrast = theme.name == "high-contrast"
    is_medieval = theme.name == "medieval-manuscript"
    is_treasure = theme.name == "fantasy-forest-map"
    is_outrun = theme.name == "synthwave-outrun"
    is_brutalist = theme.name == "brutalist-monochrome"
    is_clancy = is_clancy_dark or is_clancy_light
    title_font = theme.title_font_name
    heading_font = theme.heading_font_name
    body_font = theme.body_font_name
    centered = ThemeStyle(horz=xlwt.Alignment.HORZ_CENTER, vert=xlwt.Alignment.VERT_CENTER)
    centered_right = ThemeStyle(horz=xlwt.Alignment.HORZ_RIGHT, vert=xlwt.Alignment.VERT_CENTER)

    if sheet_name == "Current" and colx == CURRENT_FORMULA_COL:
        return ThemeStyle(
            font_name=body_font,
            font_height=200,
            font_colour="mist" if is_star_wars else "mist" if (is_cyberpunk or is_outrun) else "gold" if is_vegas else "sage" if is_medieval else "navy" if is_treasure else "white" if is_brutalist else "black" if is_high_contrast else "sage" if (is_clancy or is_oldschool) else "black",
            border_colour="stone",
            left=0,
            right=0,
            top=0,
            bottom=0,
            horz=xlwt.Alignment.HORZ_CENTER,
            wrap=False,
        )

    if role == "title":
        if is_medieval:
            return ThemeStyle(
                font_name=title_font,
                font_height=320,
                bold=True,
                font_colour="gold",
                bg_colour="navy",
                border_colour="coral",
                left=2,
                right=2,
                top=2,
                bottom=2,
                horz=xlwt.Alignment.HORZ_CENTER,
                wrap=False,
            )
        if is_treasure:
            return ThemeStyle(
                font_name=title_font,
                font_height=320,
                bold=True,
                italic=True,
                font_colour="sage",
                bg_colour="ivory",
                border_colour="gold",
                left=2,
                right=2,
                top=2,
                bottom=2,
                horz=xlwt.Alignment.HORZ_CENTER,
                wrap=False,
            )
        if is_outrun:
            return ThemeStyle(
                font_name=title_font,
                font_height=340,
                bold=True,
                font_colour="mist",
                bg_colour="navy",
                border_colour="gold",
                left=2,
                right=2,
                top=2,
                bottom=2,
                horz=xlwt.Alignment.HORZ_CENTER,
                wrap=False,
            )
        if is_brutalist:
            return ThemeStyle(
                font_name=title_font,
                font_height=320,
                bold=True,
                font_colour="white",
                bg_colour="black",
                border_colour="stone",
                left=2,
                right=2,
                top=2,
                bottom=2,
                horz=xlwt.Alignment.HORZ_CENTER,
                wrap=False,
            )
        if is_windows_90s:
            return ThemeStyle(
                font_name=title_font,
                font_height=300,
                bold=True,
                font_colour="white",
                bg_colour="navy",
                border_colour="stone",
                left=2,
                right=2,
                top=2,
                bottom=2,
                horz=xlwt.Alignment.HORZ_CENTER,
                wrap=False,
            )
        if is_vegas:
            return ThemeStyle(
                font_name=title_font,
                font_height=340,
                bold=True,
                font_colour="gold",
                bg_colour="navy",
                border_colour="coral",
                left=2,
                right=2,
                top=2,
                bottom=2,
                horz=xlwt.Alignment.HORZ_CENTER,
                wrap=False,
            )
        if is_cyberpunk:
            return ThemeStyle(
                font_name=title_font,
                font_height=340,
                bold=True,
                font_colour="coral",
                bg_colour="navy",
                border_colour="gold",
                left=2,
                right=2,
                top=2,
                bottom=2,
                horz=xlwt.Alignment.HORZ_CENTER,
                wrap=False,
            )
        if is_brutalist:
            return ThemeStyle(
                font_name=title_font,
                font_height=240,
                bold=True,
                font_colour="white",
                bg_colour="black",
                border_colour="stone",
                horz=xlwt.Alignment.HORZ_LEFT,
            )
        if is_outrun:
            return ThemeStyle(
                font_name=title_font,
                font_height=240,
                bold=True,
                font_colour="mist",
                bg_colour="navy",
                border_colour="gold",
                horz=xlwt.Alignment.HORZ_LEFT,
            )
        if is_treasure:
            return ThemeStyle(
                font_name=title_font,
                font_height=240,
                bold=True,
                italic=True,
                font_colour="sage",
                bg_colour="ivory",
                border_colour="gold",
                horz=xlwt.Alignment.HORZ_LEFT,
            )
        if is_medieval:
            return ThemeStyle(
                font_name=title_font,
                font_height=240,
                bold=True,
                font_colour="gold",
                bg_colour="navy",
                border_colour="coral",
                horz=xlwt.Alignment.HORZ_LEFT,
            )
        if is_high_contrast:
            return ThemeStyle(
                font_name=title_font,
                font_height=320,
                bold=True,
                font_colour="white",
                bg_colour="black",
                border_colour="gold",
                left=2,
                right=2,
                top=2,
                bottom=2,
                horz=xlwt.Alignment.HORZ_CENTER,
                wrap=False,
            )
        if is_star_wars:
            return ThemeStyle(
                font_name=title_font,
                font_height=360,
                bold=True,
                font_colour="gold",
                bg_colour="navy",
                border_colour="coral",
                left=0,
                right=0,
                top=0,
                bottom=2,
                horz=xlwt.Alignment.HORZ_CENTER,
                wrap=False,
            )
        if is_oldschool:
            return ThemeStyle(
                font_name=title_font,
                font_height=300,
                bold=True,
                font_colour="white",
                bg_colour="navy",
                border_colour="stone",
                left=0,
                right=0,
                top=0,
                bottom=2,
                horz=xlwt.Alignment.HORZ_CENTER,
                wrap=False,
            )
        if is_clancy_light:
            return ThemeStyle(
                font_name=title_font,
                font_height=320,
                bold=True,
                font_colour="navy",
                bg_colour="ivory",
                border_colour="gold",
                left=0,
                right=0,
                top=0,
                bottom=2,
                horz=xlwt.Alignment.HORZ_CENTER,
                wrap=False,
            )
        return ThemeStyle(
            font_name=title_font,
            font_height=320 if is_clancy else 300,
            bold=True,
            font_colour="white" if is_clancy else "navy",
            bg_colour="navy" if is_clancy else None,
            border_colour="gold" if is_clancy else "slate",
            left=0,
            right=0,
            top=0,
            bottom=2,
            horz=xlwt.Alignment.HORZ_CENTER,
            wrap=False,
        )

    if role == "section_header":
        if is_medieval and str(value).strip().upper() == "TOTAL":
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="gold",
                bg_colour="coral",
                border_colour="gold",
                left=2,
                right=2,
                top=2,
                bottom=2,
                horz=xlwt.Alignment.HORZ_CENTER,
            )
        if is_medieval:
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="sage",
                bg_colour="mist",
                border_colour="gold",
                left=2,
                right=2,
                top=2,
                bottom=2,
                horz=xlwt.Alignment.HORZ_LEFT if colx == 0 else xlwt.Alignment.HORZ_CENTER,
            )
        if is_treasure and str(value).strip().upper() == "TOTAL":
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="ivory",
                bg_colour="navy",
                border_colour="gold",
                left=2,
                right=2,
                top=2,
                bottom=2,
                horz=xlwt.Alignment.HORZ_CENTER,
            )
        if is_treasure:
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="sage",
                bg_colour="mist",
                border_colour="coral",
                left=2,
                right=2,
                top=2,
                bottom=2,
                horz=xlwt.Alignment.HORZ_LEFT if colx == 0 else xlwt.Alignment.HORZ_CENTER,
            )
        if is_outrun and str(value).strip().upper() == "TOTAL":
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="navy",
                bg_colour="coral",
                border_colour="mist",
                left=2,
                right=2,
                top=2,
                bottom=2,
                horz=xlwt.Alignment.HORZ_CENTER,
            )
        if is_outrun:
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="mist",
                bg_colour="slate",
                border_colour="gold",
                left=2,
                right=2,
                top=2,
                bottom=2,
                horz=xlwt.Alignment.HORZ_LEFT if colx == 0 else xlwt.Alignment.HORZ_CENTER,
            )
        if is_brutalist and str(value).strip().upper() == "TOTAL":
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="white",
                bg_colour="black",
                border_colour="stone",
                left=2,
                right=2,
                top=2,
                bottom=2,
                horz=xlwt.Alignment.HORZ_CENTER,
            )
        if is_brutalist:
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="black",
                bg_colour="mist",
                border_colour="stone",
                left=2,
                right=2,
                top=2,
                bottom=2,
                horz=xlwt.Alignment.HORZ_LEFT if colx == 0 else xlwt.Alignment.HORZ_CENTER,
            )
        if is_windows_90s and str(value).strip().upper() == "TOTAL":
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="white",
                bg_colour="slate",
                border_colour="stone",
                left=2,
                right=2,
                top=2,
                bottom=2,
                horz=xlwt.Alignment.HORZ_CENTER,
            )
        if is_windows_90s:
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="navy",
                bg_colour="mist",
                border_colour="stone",
                left=2,
                right=2,
                top=2,
                bottom=2,
                horz=xlwt.Alignment.HORZ_LEFT if colx == 0 else xlwt.Alignment.HORZ_CENTER,
            )
        if is_vegas and str(value).strip().upper() == "TOTAL":
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="navy",
                bg_colour="gold",
                border_colour="coral",
                left=2,
                right=2,
                top=2,
                bottom=2,
                horz=xlwt.Alignment.HORZ_CENTER,
            )
        if is_vegas:
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="gold",
                bg_colour="slate",
                border_colour="coral",
                left=2,
                right=2,
                top=2,
                bottom=2,
                horz=xlwt.Alignment.HORZ_LEFT if colx == 0 else xlwt.Alignment.HORZ_CENTER,
            )
        if is_cyberpunk and str(value).strip().upper() == "TOTAL":
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="navy",
                bg_colour="gold",
                border_colour="coral",
                left=2,
                right=2,
                top=2,
                bottom=2,
                horz=xlwt.Alignment.HORZ_CENTER,
            )
        if is_cyberpunk:
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="coral",
                bg_colour="slate",
                border_colour="gold",
                left=2,
                right=2,
                top=2,
                bottom=2,
                horz=xlwt.Alignment.HORZ_LEFT if colx == 0 else xlwt.Alignment.HORZ_CENTER,
            )
        if is_brutalist:
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="black",
                bg_colour="mist",
                border_colour="stone",
                horz=xlwt.Alignment.HORZ_CENTER if colx >= 4 else xlwt.Alignment.HORZ_LEFT,
            )
        if is_outrun:
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="mist",
                bg_colour="slate",
                border_colour="gold",
                horz=xlwt.Alignment.HORZ_CENTER if colx >= 4 else xlwt.Alignment.HORZ_LEFT,
            )
        if is_treasure:
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="sage",
                bg_colour="mist",
                border_colour="coral",
                horz=xlwt.Alignment.HORZ_CENTER if colx >= 4 else xlwt.Alignment.HORZ_LEFT,
            )
        if is_medieval:
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="sage",
                bg_colour="mist",
                border_colour="gold",
                horz=xlwt.Alignment.HORZ_CENTER if colx >= 4 else xlwt.Alignment.HORZ_LEFT,
            )
        if is_high_contrast:
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="white" if str(value).strip().upper() != "TOTAL" else "black",
                bg_colour="black" if str(value).strip().upper() != "TOTAL" else "gold",
                border_colour="gold",
                left=2,
                right=2,
                top=2,
                bottom=2,
                horz=xlwt.Alignment.HORZ_LEFT if colx == 0 else xlwt.Alignment.HORZ_CENTER,
            )
        if is_star_wars and str(value).strip().upper() == "TOTAL":
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="navy",
                bg_colour="gold",
                border_colour="coral",
                left=1,
                right=1,
                top=1,
                bottom=2,
                horz=xlwt.Alignment.HORZ_CENTER,
            )
        if is_star_wars:
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="gold",
                bg_colour="slate",
                border_colour="coral",
                left=1,
                right=1,
                top=1,
                bottom=2,
                horz=xlwt.Alignment.HORZ_LEFT if colx == 0 else xlwt.Alignment.HORZ_CENTER,
            )
        if is_oldschool and str(value).strip().upper() == "TOTAL":
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="navy",
                bg_colour="gold",
                border_colour="stone",
                left=1,
                right=1,
                top=1,
                bottom=2,
                horz=xlwt.Alignment.HORZ_CENTER,
            )
        if is_oldschool:
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="navy",
                bg_colour="mist",
                border_colour="stone",
                left=1,
                right=1,
                top=1,
                bottom=2,
                horz=xlwt.Alignment.HORZ_LEFT if colx == 0 else xlwt.Alignment.HORZ_CENTER,
            )
        if is_clancy and str(value).strip().upper() == "TOTAL":
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="white",
                bg_colour="gold",
                border_colour="gold",
                left=1,
                right=1,
                top=1,
                bottom=2,
                horz=xlwt.Alignment.HORZ_CENTER,
            )
        if is_clancy_light:
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="navy",
                bg_colour="mist",
                border_colour="gold",
                left=1,
                right=1,
                top=1,
                bottom=2,
                horz=xlwt.Alignment.HORZ_LEFT if colx == 0 else xlwt.Alignment.HORZ_CENTER,
            )
        return ThemeStyle(
            font_name=heading_font,
            font_height=220,
            bold=True,
            font_colour="navy",
            bg_colour="ivory" if is_clancy else "mist",
            border_colour="gold" if is_clancy else "slate",
            left=1,
            right=1,
            top=1,
            bottom=2,
            horz=xlwt.Alignment.HORZ_LEFT if colx == 0 else xlwt.Alignment.HORZ_CENTER,
        )

    if role == "total_sent":
        if is_medieval:
            bg_colour = "coral"
            font_colour = "gold"
        elif is_treasure:
            bg_colour = "navy"
            font_colour = "ivory"
        elif is_outrun:
            bg_colour = "gold"
            font_colour = "navy"
        elif is_brutalist:
            bg_colour = "black"
            font_colour = "white"
        elif is_windows_90s:
            bg_colour = "slate"
            font_colour = "white"
        elif is_vegas:
            bg_colour = "gold"
            font_colour = "navy"
        elif is_cyberpunk:
            bg_colour = "coral"
            font_colour = "navy"
        elif is_high_contrast:
            bg_colour = "gold"
            font_colour = "black"
        elif is_star_wars:
            bg_colour = "slate"
            font_colour = "gold"
        elif is_oldschool:
            bg_colour = "slate"
            font_colour = "white"
        elif is_clancy_dark:
            bg_colour = "navy"
            font_colour = "white"
        elif is_clancy_light:
            bg_colour = "mist"
            font_colour = "navy"
        else:
            bg_colour = "mist"
            font_colour = "navy"
    elif role == "total_rec":
        if is_medieval:
            bg_colour = "mist"
            font_colour = "sage"
        elif is_treasure:
            bg_colour = "mist"
            font_colour = "sage"
        elif is_outrun:
            bg_colour = "mist"
            font_colour = "navy"
        elif is_brutalist:
            bg_colour = "mist"
            font_colour = "black"
        elif is_windows_90s:
            bg_colour = "mist"
            font_colour = "navy"
        elif is_vegas:
            bg_colour = "coral"
            font_colour = "navy"
        elif is_cyberpunk:
            bg_colour = "gold"
            font_colour = "navy"
        elif is_high_contrast:
            bg_colour = "black"
            font_colour = "white"
        elif is_star_wars:
            bg_colour = "coral"
            font_colour = "navy"
        elif is_oldschool:
            bg_colour = "mist"
            font_colour = "navy"
        elif is_clancy_dark:
            bg_colour = "slate"
            font_colour = "white"
        elif is_clancy_light:
            bg_colour = "ivory"
            font_colour = "navy"
        else:
            bg_colour = "ivory"
            font_colour = "navy"
    elif role == "not_found":
        if is_medieval:
            bg_colour = "gold"
            font_colour = "sage"
        elif is_treasure:
            bg_colour = "coral"
            font_colour = "ivory"
        elif is_outrun:
            bg_colour = "slate"
            font_colour = "mist"
        elif is_brutalist:
            bg_colour = "white"
            font_colour = "black"
        elif is_windows_90s:
            bg_colour = "gold"
            font_colour = "navy"
        elif is_vegas:
            bg_colour = "mist"
            font_colour = "navy"
        elif is_cyberpunk:
            bg_colour = "slate"
            font_colour = "coral"
        elif is_high_contrast:
            bg_colour = "white"
            font_colour = "black"
        elif is_star_wars:
            bg_colour = "gold"
            font_colour = "navy"
        elif is_oldschool:
            bg_colour = "coral"
            font_colour = "navy"
        elif is_clancy_dark:
            bg_colour = "gold"
            font_colour = "white"
        elif is_clancy_light:
            bg_colour = "coral"
            font_colour = "navy"
        else:
            bg_colour = None
            font_colour = "navy"
    else:
        bg_colour = None
        font_colour = "black"

    if role in {"total_sent", "total_rec", "not_found"}:
        if colx == 0:
            return ThemeStyle(
                font_name=heading_font if (is_clancy or is_oldschool or is_star_wars or is_windows_90s or is_vegas or is_cyberpunk or is_high_contrast or is_medieval or is_treasure or is_outrun or is_brutalist) else body_font,
                font_height=220,
                bold=True,
                font_colour=font_colour,
                bg_colour=bg_colour,
                border_colour="stone" if is_brutalist else "gold" if (is_high_contrast or is_cyberpunk or is_outrun or is_medieval) else "coral" if (is_star_wars or is_vegas or is_treasure) else "stone",
                left=1,
                right=1,
                top=1,
                bottom=2,
                horz=xlwt.Alignment.HORZ_RIGHT,
            )
        return ThemeStyle(
            font_name=heading_font if (is_clancy or is_oldschool or is_star_wars or is_windows_90s or is_vegas or is_cyberpunk or is_high_contrast or is_medieval or is_treasure or is_outrun or is_brutalist) else body_font,
            font_height=220,
            bold=True,
            font_colour=font_colour,
            bg_colour=bg_colour,
            border_colour="stone" if is_brutalist else "gold" if (is_high_contrast or is_cyberpunk or is_outrun or is_medieval) else "coral" if (is_star_wars or is_vegas or is_treasure) else "stone",
            left=1,
            right=1,
            top=1,
            bottom=2,
            horz=xlwt.Alignment.HORZ_CENTER,
        )

    if role == "note":
        return ThemeStyle(
            font_name=body_font,
            bold=colx == 7,
            font_colour="white" if is_high_contrast else "mist" if (is_star_wars or is_outrun) else "gold" if is_vegas else "sage" if (is_clancy or is_oldschool or is_medieval or is_treasure) else "black" if is_brutalist else "navy",
            bg_colour="black" if (is_high_contrast or is_brutalist) else "slate" if (is_star_wars or is_outrun) else "navy" if is_vegas else "ivory" if (is_oldschool or is_treasure) else "mist" if (is_medieval or is_clancy_dark) else "ivory" if is_clancy_light else "ivory",
            border_colour="stone" if is_brutalist else "gold" if (is_high_contrast or is_cyberpunk or is_outrun or is_medieval) else "coral" if (is_star_wars or is_vegas or is_treasure) else "stone",
            left=1,
            right=1,
            top=1,
            bottom=2,
            horz=xlwt.Alignment.HORZ_CENTER if colx >= 1 else xlwt.Alignment.HORZ_LEFT,
        )

    if role in {"sent", "received", "body"}:
        if is_brutalist:
            bg_colour = "coral" if role == "received" else "ivory" if colx <= MAIN_TABLE_LAST_COL else None
        elif is_outrun:
            bg_colour = "slate" if role == "received" else "ivory" if colx <= MAIN_TABLE_LAST_COL else None
        elif is_treasure:
            bg_colour = "mist" if role == "received" else "ivory" if colx <= MAIN_TABLE_LAST_COL else None
        elif is_medieval:
            bg_colour = "mist" if role == "received" else "ivory" if colx <= MAIN_TABLE_LAST_COL else None
        elif is_high_contrast:
            bg_colour = "black" if role == "received" else "white" if colx <= MAIN_TABLE_LAST_COL else None
        elif is_vegas:
            bg_colour = "coral" if role == "received" else "ivory" if colx <= MAIN_TABLE_LAST_COL else None
        elif is_cyberpunk:
            bg_colour = "slate" if role == "received" else "ivory" if colx <= MAIN_TABLE_LAST_COL else None
        elif is_windows_90s:
            bg_colour = "mist" if role == "received" else "ivory" if colx <= MAIN_TABLE_LAST_COL else None
        elif is_star_wars:
            bg_colour = "slate" if role == "received" else "ivory" if colx <= MAIN_TABLE_LAST_COL else None
        elif is_oldschool:
            bg_colour = "coral" if role == "received" else "ivory" if colx <= MAIN_TABLE_LAST_COL else None
        elif is_clancy_dark:
            bg_colour = "mist" if role == "received" else "ivory" if colx <= MAIN_TABLE_LAST_COL else None
        elif is_clancy_light:
            bg_colour = "mist" if role == "received" else "ivory" if colx <= MAIN_TABLE_LAST_COL else None
        else:
            bg_colour = "mist" if role == "received" else None
        spec = ThemeStyle(
            font_name=body_font,
            font_height=200,
            font_colour="white" if is_high_contrast and role != "body" else "black" if is_high_contrast else "white" if is_brutalist and role != "body" else "black" if is_brutalist else "mist" if (is_star_wars or is_outrun) else "gold" if is_vegas else "mist" if is_cyberpunk else "sage" if (is_clancy or is_oldschool or is_medieval or is_treasure) else "navy" if role != "body" else "black",
            bg_colour=bg_colour,
            border_colour="stone" if is_brutalist and colx <= MAIN_TABLE_LAST_COL else "gold" if (is_high_contrast or is_cyberpunk or is_outrun or is_medieval) and colx <= MAIN_TABLE_LAST_COL else "coral" if (is_star_wars or is_vegas or is_treasure) and colx <= MAIN_TABLE_LAST_COL else "stone",
            horz=xlwt.Alignment.HORZ_LEFT if colx == 0 else xlwt.Alignment.HORZ_CENTER,
            bold=colx == 7 and sheet_name != "Compatibility Report",
        )
        if colx == 1:
            return ThemeStyle(
                font_name=body_font,
                font_height=200,
                bold=True,
                font_colour="black" if is_high_contrast else "white" if is_brutalist else "gold" if (is_star_wars or is_outrun) else "navy" if (is_oldschool or is_medieval or is_treasure) else "gold" if (is_clancy or is_vegas or is_cyberpunk) else "navy",
                bg_colour=bg_colour,
                border_colour="stone" if is_brutalist else "gold" if (is_high_contrast or is_cyberpunk or is_outrun or is_medieval) else "coral" if (is_star_wars or is_vegas or is_treasure) else "stone",
                horz=xlwt.Alignment.HORZ_CENTER,
            )
        if colx == 7 and sheet_name != "Compatibility Report":
            return ThemeStyle(
                font_name=body_font,
                font_height=200,
                bold=True,
                font_colour="black" if is_high_contrast else "white" if is_brutalist else "navy" if (is_oldschool or is_star_wars or is_vegas or is_cyberpunk or is_medieval or is_treasure or is_outrun) else "white" if is_clancy else "navy",
                bg_colour="mist" if is_brutalist else "gold" if (is_clancy or is_oldschool or is_star_wars or is_vegas or is_cyberpunk or is_high_contrast or is_medieval or is_outrun) else "navy" if is_treasure else None,
                border_colour="stone" if is_brutalist else "gold" if (is_high_contrast or is_cyberpunk or is_outrun or is_medieval) else "coral" if (is_star_wars or is_vegas or is_treasure) else "stone" if is_oldschool else "gold" if is_clancy else "stone",
                horz=xlwt.Alignment.HORZ_CENTER,
            )
        if "%" in fmt or colx >= 8:
            return ThemeStyle(
                font_name=body_font,
                font_height=200,
                font_colour="black" if is_high_contrast else "black" if is_brutalist else "coral" if is_star_wars else "gold" if is_vegas else "coral" if (is_cyberpunk or is_outrun) else "sage" if (is_oldschool or is_medieval or is_treasure) else "slate" if is_clancy else "navy",
                bg_colour="white" if (is_high_contrast or is_brutalist) and colx <= MAIN_TABLE_LAST_COL else "ivory" if (is_clancy or is_oldschool or is_star_wars or is_vegas or is_cyberpunk or is_medieval or is_treasure or is_outrun) and colx <= MAIN_TABLE_LAST_COL else None,
                border_colour="stone" if is_brutalist and colx <= MAIN_TABLE_LAST_COL else "gold" if (is_high_contrast or is_cyberpunk or is_outrun or is_medieval) and colx <= MAIN_TABLE_LAST_COL else "coral" if (is_star_wars or is_vegas or is_treasure) and colx <= MAIN_TABLE_LAST_COL else "stone",
                horz=xlwt.Alignment.HORZ_CENTER,
            )
        if colx >= 2:
            return ThemeStyle(
                font_name=body_font,
                font_height=200,
                font_colour="black" if is_high_contrast else "black" if is_brutalist else "mist" if (is_star_wars or is_outrun) else "gold" if is_vegas else "mist" if is_cyberpunk else "sage" if (is_clancy or is_oldschool or is_medieval or is_treasure) else "navy" if role != "body" else "black",
                bg_colour=bg_colour,
                border_colour="stone" if is_brutalist else "gold" if (is_high_contrast or is_cyberpunk or is_outrun or is_medieval) else "coral" if (is_star_wars or is_vegas or is_treasure) else "stone",
                horz=xlwt.Alignment.HORZ_CENTER,
            )
        return spec

    if role == "compat_title":
        if is_high_contrast:
            return ThemeStyle(
                font_name=title_font,
                font_height=240,
                bold=True,
                font_colour="white",
                bg_colour="black",
                border_colour="gold",
                horz=xlwt.Alignment.HORZ_LEFT,
            )
        if is_cyberpunk:
            return ThemeStyle(
                font_name=title_font,
                font_height=240,
                bold=True,
                font_colour="coral",
                bg_colour="navy",
                border_colour="gold",
                horz=xlwt.Alignment.HORZ_LEFT,
            )
        if is_vegas:
            return ThemeStyle(
                font_name=title_font,
                font_height=240,
                bold=True,
                font_colour="gold",
                bg_colour="navy",
                border_colour="coral",
                horz=xlwt.Alignment.HORZ_LEFT,
            )
        if is_windows_90s:
            return ThemeStyle(
                font_name=title_font,
                font_height=240,
                bold=True,
                font_colour="white",
                bg_colour="navy",
                border_colour="stone",
                horz=xlwt.Alignment.HORZ_LEFT,
            )
        if is_star_wars:
            return ThemeStyle(
                font_name=title_font,
                font_height=240,
                bold=True,
                font_colour="gold",
                bg_colour="navy",
                border_colour="coral",
                horz=xlwt.Alignment.HORZ_LEFT,
            )
        if is_oldschool:
            return ThemeStyle(
                font_name=title_font,
                font_height=240,
                bold=True,
                font_colour="white",
                bg_colour="navy",
                border_colour="stone",
                horz=xlwt.Alignment.HORZ_LEFT,
            )
        if is_clancy_light:
            return ThemeStyle(
                font_name=title_font,
                font_height=240,
                bold=True,
                font_colour="navy",
                bg_colour="ivory",
                border_colour="gold",
                horz=xlwt.Alignment.HORZ_LEFT,
            )
        return ThemeStyle(
            font_name=title_font,
            font_height=240,
            bold=True,
            font_colour="white",
            bg_colour="navy",
            border_colour="navy",
            horz=xlwt.Alignment.HORZ_LEFT,
        )

    if role == "compat_header":
        if is_high_contrast:
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="black",
                bg_colour="gold",
                border_colour="gold",
                horz=xlwt.Alignment.HORZ_CENTER if colx >= 4 else xlwt.Alignment.HORZ_LEFT,
            )
        if is_cyberpunk:
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="navy",
                bg_colour="gold",
                border_colour="coral",
                horz=xlwt.Alignment.HORZ_CENTER if colx >= 4 else xlwt.Alignment.HORZ_LEFT,
            )
        if is_vegas:
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="gold",
                bg_colour="slate",
                border_colour="coral",
                horz=xlwt.Alignment.HORZ_CENTER if colx >= 4 else xlwt.Alignment.HORZ_LEFT,
            )
        if is_windows_90s:
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="navy",
                bg_colour="mist",
                border_colour="stone",
                horz=xlwt.Alignment.HORZ_CENTER if colx >= 4 else xlwt.Alignment.HORZ_LEFT,
            )
        if is_star_wars:
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="navy",
                bg_colour="gold",
                border_colour="coral",
                horz=xlwt.Alignment.HORZ_CENTER if colx >= 4 else xlwt.Alignment.HORZ_LEFT,
            )
        if is_oldschool:
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="navy",
                bg_colour="mist",
                border_colour="stone",
                horz=xlwt.Alignment.HORZ_CENTER if colx >= 4 else xlwt.Alignment.HORZ_LEFT,
            )
        if is_clancy_light:
            return ThemeStyle(
                font_name=heading_font,
                font_height=220,
                bold=True,
                font_colour="navy",
                bg_colour="coral",
                border_colour="gold",
                horz=xlwt.Alignment.HORZ_CENTER if colx >= 4 else xlwt.Alignment.HORZ_LEFT,
            )
        return ThemeStyle(
            font_name=heading_font,
            font_height=220,
            bold=True,
            font_colour="white",
            bg_colour="gold" if is_clancy else "slate",
            border_colour="gold" if is_clancy else "navy",
            horz=xlwt.Alignment.HORZ_CENTER if colx >= 4 else xlwt.Alignment.HORZ_LEFT,
        )

    if role == "compat_text":
        return ThemeStyle(
            font_name=body_font,
            font_height=200,
            font_colour="white" if is_high_contrast else "black" if is_brutalist else "mist" if is_star_wars else "gold" if is_vegas else "mist" if (is_cyberpunk or is_outrun) else "sage" if (is_medieval or is_treasure) else "black",
            bg_colour="black" if (is_high_contrast or is_brutalist) else "slate" if is_star_wars else "navy" if is_vegas else "slate" if (is_cyberpunk or is_outrun) else "ivory" if (is_oldschool or is_treasure) else "mist",
            border_colour="stone" if is_brutalist else "gold" if (is_high_contrast or is_cyberpunk or is_outrun or is_medieval) else "coral" if (is_star_wars or is_vegas or is_treasure) else "stone",
            horz=xlwt.Alignment.HORZ_LEFT,
        )

    if role == "compat_body":
        return centered_right if colx == 4 else centered if colx == 5 else ThemeStyle(
            font_name=body_font,
            font_height=200,
            font_colour="black" if is_high_contrast else "black" if is_brutalist else "mist" if is_star_wars else "gold" if is_vegas else "coral" if (is_cyberpunk or is_outrun) else "sage" if (is_medieval or is_treasure) else "black",
            bg_colour="white" if (is_high_contrast or is_brutalist) else "ivory" if (is_star_wars or is_vegas or is_cyberpunk or is_treasure or is_outrun) else "mist" if (is_oldschool or is_medieval) else "ivory",
            border_colour="stone" if is_brutalist else "gold" if (is_high_contrast or is_cyberpunk or is_outrun or is_medieval) else "coral" if (is_star_wars or is_vegas or is_treasure) else "stone",
        )

    return ThemeStyle()


def merged_ranges(
    sheet_name: str,
    sheet: xlrd.sheet.Sheet,
    row_map: dict[int, int],
    max_col: int,
) -> dict[tuple[int, int], tuple[int, int, int, int]]:
    merged: dict[tuple[int, int], tuple[int, int, int, int]] = {
        (row_map[row_lo], col_lo): (row_map[row_lo], row_map[row_hi - 1], col_lo, min(col_hi - 1, max_col))
        for row_lo, row_hi, col_lo, col_hi in sheet.merged_cells
        if all(rowx in row_map for rowx in range(row_lo, row_hi)) and col_lo <= max_col
    }
    if sheet_name == "Current" and 1 in row_map:
        merged[(row_map[1], 0)] = (row_map[1], row_map[1], 0, MAIN_TABLE_LAST_COL)
    return merged


def covered_cells(merged: dict[tuple[int, int], tuple[int, int, int, int]]) -> set[tuple[int, int]]:
    cells: set[tuple[int, int]] = set()
    for row_lo, row_hi, col_lo, col_hi in merged.values():
        for rowx in range(row_lo, row_hi + 1):
            for colx in range(col_lo, col_hi + 1):
                cells.add((rowx, colx))
    return cells


def write_value(
    worksheet: xlwt.Worksheet,
    rowx: int,
    colx: int,
    cell: xlrd.sheet.Cell,
    value: object,
    style: xlwt.XFStyle,
) -> None:
    if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
        worksheet.write(rowx, colx, "", style)
    elif cell.ctype == xlrd.XL_CELL_TEXT:
        worksheet.write(rowx, colx, value, style)
    elif cell.ctype == xlrd.XL_CELL_NUMBER:
        worksheet.write(rowx, colx, value, style)
    elif cell.ctype == xlrd.XL_CELL_DATE:
        worksheet.write(rowx, colx, value, style)
    elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
        worksheet.write(rowx, colx, bool(value), style)
    elif cell.ctype == xlrd.XL_CELL_ERROR:
        worksheet.write(rowx, colx, value, style)
    else:
        worksheet.write(rowx, colx, str(value), style)


def output_cell_value(sheet_name: str, output_filename: str, rowx: int, colx: int, cell: xlrd.sheet.Cell) -> object:
    if sheet_name == "Compatibility Report" and rowx == 0 and colx == 1:
        return f"Compatibility Report for {output_filename}"
    if cell.ctype == xlrd.XL_CELL_ERROR:
        return xlrd.error_text_from_code.get(cell.value, "")
    return cell.value


def themed_cell_value(
    theme: ThemeProfile,
    sheet_name: str,
    output_filename: str,
    rowx: int,
    colx: int,
    cell: xlrd.sheet.Cell,
) -> object:
    value = output_cell_value(sheet_name, output_filename, rowx, colx, cell)
    if not isinstance(value, str):
        return value

    if theme.name == "star-wars":
        if sheet_name == "Current" and rowx == 1 and colx == 0:
            return "GALACTIC DMV HOLOCRON - APRIL 2025"
        if sheet_name == "Template" and rowx == 6 and colx == 0:
            return "GALACTIC DMV HOLOCRON TEMPLATE"
        if sheet_name == "Compatibility Report" and rowx == 0 and colx == 1:
            return f"JEDI ARCHIVE COMPATIBILITY REPORT - {output_filename}"
        replacements = {
            "Other States": "Outer Rim Systems",
            "Other States-POLK": "Outer Rim Systems - POLK SECTOR",
            "Colorado DMV ": "Coruscant DMV",
            "TOTAL": "GALACTIC TOTAL",
            "Total Sent:": "Transmissions Sent:",
            "Total Rec:": "Signals Received:",
            "Not Found:": "Lost in Hyperspace:",
            "Sent": "Dispatched",
            "Rec'd": "Received",
            "Compatibility Report": "Jedi Archive Report",
            "Minor loss of fidelity": "Minor disturbance in the Force",
        }
        return replacements.get(value, value)
    if theme.name == "retro-90s":
        if sheet_name == "Current" and rowx == 1 and colx == 0:
            return "Monthly Report - Windows 95 Edition"
        if sheet_name == "Template" and rowx == 6 and colx == 0:
            return "Template - Windows 95 Edition"
        if sheet_name == "Compatibility Report" and rowx == 0 and colx == 1:
            return f"System Compatibility Check - {output_filename}"
        replacements = {
            "TOTAL": "TOTAL.EXE",
            "Total Sent:": "Packets Sent:",
            "Total Rec:": "Packets Received:",
            "Not Found:": "Missing Files:",
            "Minor loss of fidelity": "Display driver warning",
        }
        return replacements.get(value, value)
    if theme.name == "vegas-casino":
        if sheet_name == "Current" and rowx == 1 and colx == 0:
            return "THE GOLDMAN GRAND - HIGH ROLLER REPORT"
        if sheet_name == "Template" and rowx == 6 and colx == 0:
            return "THE GOLDMAN GRAND - TABLE TEMPLATE"
        if sheet_name == "Compatibility Report" and rowx == 0 and colx == 1:
            return f"Casino Floor Compatibility Report - {output_filename}"
        replacements = {
            "Other States": "High Roller Tables",
            "Other States-POLK": "High Roller Tables - Polk Pit",
            "Colorado DMV ": "Main Casino Floor",
            "TOTAL": "JACKPOT TOTAL",
            "Total Sent:": "Chips In Play:",
            "Total Rec:": "Chips Returned:",
            "Not Found:": "House Losses:",
            "Sent": "Wagered",
            "Rec'd": "Paid",
            "Minor loss of fidelity": "Minor lighting loss",
        }
        return replacements.get(value, value)
    if theme.name == "cyberpunk":
        if sheet_name == "Current" and rowx == 1 and colx == 0:
            return "NEON GRID OPERATIONS // APRIL 2025"
        if sheet_name == "Template" and rowx == 6 and colx == 0:
            return "NEON GRID TEMPLATE // CONTROL PANEL"
        if sheet_name == "Compatibility Report" and rowx == 0 and colx == 1:
            return f"Neon Grid Diagnostics - {output_filename}"
        replacements = {
            "Other States": "Night City Sectors",
            "Other States-POLK": "Night City Sectors - Polk Node",
            "Colorado DMV ": "Central Neon District",
            "TOTAL": "NET TOTAL",
            "Total Sent:": "Signals Pushed:",
            "Total Rec:": "Signals Synced:",
            "Not Found:": "Ghost Packets:",
            "Sent": "Uploaded",
            "Rec'd": "Synced",
            "Minor loss of fidelity": "Signal degradation detected",
        }
        return replacements.get(value, value)
    if theme.name == "high-contrast":
        if sheet_name == "Current" and rowx == 1 and colx == 0:
            return "HIGH CONTRAST REPORT - APRIL 2025"
        if sheet_name == "Template" and rowx == 6 and colx == 0:
            return "HIGH CONTRAST TEMPLATE"
        if sheet_name == "Compatibility Report" and rowx == 0 and colx == 1:
            return f"High Contrast Accessibility Report - {output_filename}"
        replacements = {
            "TOTAL": "TOTAL",
            "Total Sent:": "TOTAL SENT:",
            "Total Rec:": "TOTAL RECEIVED:",
            "Not Found:": "NOT FOUND:",
            "Minor loss of fidelity": "Accessibility contrast warning",
        }
        return replacements.get(value, value)
    if theme.name == "medieval-manuscript":
        if sheet_name == "Current" and rowx == 1 and colx == 0:
            return "Royal Ledger of the Golden Month"
        if sheet_name == "Template" and rowx == 6 and colx == 0:
            return "Royal Ledger Template"
        if sheet_name == "Compatibility Report" and rowx == 0 and colx == 1:
            return f"Scribe's Compatibility Chronicle - {output_filename}"
        replacements = {
            "Other States": "Western Kingdoms",
            "Other States-POLK": "Western Kingdoms - Polk Province",
            "Colorado DMV ": "Capital Registry",
            "TOTAL": "GRAND LEDGER",
            "Total Sent:": "Missives Sent:",
            "Total Rec:": "Missives Returned:",
            "Not Found:": "Lost to the Realm:",
            "Sent": "Dispatched",
            "Rec'd": "Received",
            "Minor loss of fidelity": "Minor ink loss detected",
        }
        return replacements.get(value, value)
    if theme.name == "fantasy-forest-map":
        if sheet_name == "Current" and rowx == 1 and colx == 0:
            return "Forest Treasure Map - Monthly Survey"
        if sheet_name == "Template" and rowx == 6 and colx == 0:
            return "Forest Treasure Map Template"
        if sheet_name == "Compatibility Report" and rowx == 0 and colx == 1:
            return f"Treasure Map Survey Report - {output_filename}"
        replacements = {
            "Other States": "Whispering Woods",
            "Other States-POLK": "Whispering Woods - Polk Trail",
            "Colorado DMV ": "Emerald Keep",
            "TOTAL": "TREASURE TOTAL",
            "Total Sent:": "Clues Marked:",
            "Total Rec:": "Caches Found:",
            "Not Found:": "Lost in the Thicket:",
            "Sent": "Mapped",
            "Rec'd": "Found",
            "Minor loss of fidelity": "Minor parchment wear",
        }
        return replacements.get(value, value)
    if theme.name == "synthwave-outrun":
        if sheet_name == "Current" and rowx == 1 and colx == 0:
            return "SUNSET CIRCUIT REPORT // APRIL 2025"
        if sheet_name == "Template" and rowx == 6 and colx == 0:
            return "SUNSET CIRCUIT TEMPLATE"
        if sheet_name == "Compatibility Report" and rowx == 0 and colx == 1:
            return f"Outrun System Check - {output_filename}"
        replacements = {
            "Other States": "Sunset Sectors",
            "Other States-POLK": "Sunset Sectors - Polk Drive",
            "Colorado DMV ": "Chrome Coast Hub",
            "TOTAL": "SUNSET TOTAL",
            "Total Sent:": "Pulses Sent:",
            "Total Rec:": "Echoes Returned:",
            "Not Found:": "Signal Fade:",
            "Sent": "Broadcast",
            "Rec'd": "Echoed",
            "Minor loss of fidelity": "Minor VHS drift",
        }
        return replacements.get(value, value)
    if theme.name == "brutalist-monochrome":
        if sheet_name == "Current" and rowx == 1 and colx == 0:
            return "MONTHLY REPORT / MONO EDITION"
        if sheet_name == "Template" and rowx == 6 and colx == 0:
            return "TEMPLATE / MONO EDITION"
        if sheet_name == "Compatibility Report" and rowx == 0 and colx == 1:
            return f"MONO COMPATIBILITY REPORT / {output_filename}"
        replacements = {
            "TOTAL": "TOTAL BLOCK",
            "Total Sent:": "OUTPUT TOTAL:",
            "Total Rec:": "INPUT TOTAL:",
            "Not Found:": "NULL SET:",
            "Minor loss of fidelity": "Monochrome degradation warning",
        }
        return replacements.get(value, value)

    return value


def create_star_wars_assets(asset_dir: Path) -> dict[str, Path]:
    banner_path = asset_dir / "star_wars_banner.bmp"
    side_art_path = asset_dir / "star_wars_side.bmp"
    compat_art_path = asset_dir / "star_wars_compat.bmp"

    banner = Image.new("RGB", (900, 180), (8, 8, 20))
    draw = ImageDraw.Draw(banner)
    for x in range(0, 900, 37):
        for y in range((x * 7) % 23, 180, 29):
            draw.ellipse((x, y, x + 2, y + 2), fill=(255, 255, 255))
    draw.rectangle((0, 0, 899, 25), fill=(12, 12, 32))
    draw.text((180, 25), "STAR WARS", fill=(255, 214, 10))
    draw.text((85, 75), "GALACTIC DMV INTELLIGENCE REPORT", fill=(255, 214, 10))
    draw.text((170, 120), "MAY THE FORMS BE WITH YOU", fill=(83, 210, 255))
    draw.polygon([(15, 160), (60, 150), (35, 110)], fill=(83, 210, 255))
    draw.polygon([(845, 15), (890, 25), (865, 65)], fill=(255, 214, 10))
    banner.save(banner_path)

    side = Image.new("RGB", (180, 260), (12, 12, 32))
    draw = ImageDraw.Draw(side)
    for x in range(0, 180, 20):
        for y in range((x * 5) % 13, 260, 21):
            draw.point((x, y), fill=(255, 255, 255))
    draw.rectangle((25, 20, 155, 50), outline=(255, 214, 10), width=3)
    draw.text((40, 26), "REBELS", fill=(255, 214, 10))
    draw.rectangle((30, 85, 150, 125), outline=(83, 210, 255), width=3)
    draw.text((44, 95), "EMPIRE", fill=(83, 210, 255))
    draw.line((30, 180, 150, 180), fill=(255, 0, 0), width=4)
    draw.line((90, 150, 90, 210), fill=(83, 210, 255), width=4)
    draw.ellipse((60, 215, 120, 255), outline=(255, 214, 10), width=3)
    draw.text((69, 228), "R2", fill=(255, 214, 10))
    side.save(side_art_path)

    compat = Image.new("RGB", (420, 120), (12, 12, 32))
    draw = ImageDraw.Draw(compat)
    for x in range(0, 420, 25):
        for y in range((x * 3) % 17, 120, 23):
            draw.point((x, y), fill=(255, 255, 255))
    draw.text((25, 15), "JEDI ARCHIVE", fill=(255, 214, 10))
    draw.text((25, 50), "COMPATIBILITY CHECK", fill=(83, 210, 255))
    draw.text((25, 82), "NO DISINTEGRATIONS", fill=(255, 214, 10))
    compat.save(compat_art_path)

    return {
        "banner": banner_path,
        "side": side_art_path,
        "compat": compat_art_path,
    }


def create_windows_90s_assets(asset_dir: Path) -> dict[str, Path]:
    banner_path = asset_dir / "windows_90s_banner.bmp"
    side_art_path = asset_dir / "windows_90s_side.bmp"
    compat_art_path = asset_dir / "windows_90s_compat.bmp"

    banner = Image.new("RGB", (900, 180), (192, 192, 192))
    draw = ImageDraw.Draw(banner)
    draw.rectangle((0, 0, 899, 30), fill=(0, 0, 128))
    draw.rectangle((8, 42, 892, 168), fill=(236, 233, 216), outline=(128, 128, 128), width=3)
    draw.rectangle((16, 50, 884, 70), fill=(0, 128, 128))
    draw.text((30, 8), "Galactic Report Manager", fill=(255, 255, 255))
    draw.text((40, 88), "RETRO 90s WINDOWS COMMAND CENTER", fill=(0, 0, 0))
    draw.rectangle((40, 115, 160, 145), fill=(192, 192, 192), outline=(0, 0, 0), width=1)
    draw.text((70, 123), "OK", fill=(0, 0, 0))
    draw.rectangle((185, 115, 305, 145), fill=(192, 192, 192), outline=(0, 0, 0), width=1)
    draw.text((202, 123), "CANCEL", fill=(0, 0, 0))
    banner.save(banner_path)

    side = Image.new("RGB", (180, 260), (236, 233, 216))
    draw = ImageDraw.Draw(side)
    draw.rectangle((12, 12, 168, 248), outline=(128, 128, 128), width=3)
    draw.rectangle((18, 18, 162, 44), fill=(0, 0, 128))
    draw.text((30, 24), "Control Panel", fill=(255, 255, 255))
    draw.rectangle((28, 70, 150, 95), fill=(192, 192, 192), outline=(0, 0, 0))
    draw.text((40, 76), "Spreadsheet", fill=(0, 0, 0))
    draw.rectangle((28, 110, 150, 135), fill=(192, 192, 192), outline=(0, 0, 0))
    draw.text((55, 116), "Inbox", fill=(0, 0, 0))
    draw.rectangle((28, 150, 150, 175), fill=(192, 192, 192), outline=(0, 0, 0))
    draw.text((57, 156), "Start", fill=(0, 0, 0))
    draw.rectangle((20, 222, 160, 240), fill=(0, 128, 128))
    draw.text((35, 226), "Ready", fill=(255, 255, 255))
    side.save(side_art_path)

    compat = Image.new("RGB", (420, 120), (236, 233, 216))
    draw = ImageDraw.Draw(compat)
    draw.rectangle((0, 0, 419, 26), fill=(0, 0, 128))
    draw.text((15, 7), "Compatibility Wizard", fill=(255, 255, 255))
    draw.text((18, 45), "Your workbook completed successfully.", fill=(0, 0, 0))
    draw.text((18, 72), "Press Finish to continue.", fill=(0, 0, 0))
    draw.rectangle((300, 86, 385, 108), fill=(192, 192, 192), outline=(0, 0, 0))
    draw.text((320, 91), "Finish", fill=(0, 0, 0))
    compat.save(compat_art_path)

    return {
        "banner": banner_path,
        "side": side_art_path,
        "compat": compat_art_path,
    }


def create_vegas_assets(asset_dir: Path) -> dict[str, Path]:
    banner_path = asset_dir / "vegas_banner.bmp"
    side_art_path = asset_dir / "vegas_side.bmp"
    compat_art_path = asset_dir / "vegas_compat.bmp"

    banner = Image.new("RGB", (900, 180), (35, 0, 45))
    draw = ImageDraw.Draw(banner)
    for x in range(0, 900, 25):
        draw.ellipse((x + 5, 10, x + 12, 17), fill=(255, 215, 0))
    draw.rectangle((30, 35, 870, 150), outline=(255, 215, 0), width=5)
    draw.text((260, 55), "VEGAS CASINO REPORT", fill=(255, 215, 0))
    draw.text((175, 100), "JACKPOT METRICS - ALL TABLES HOT", fill=(255, 80, 160))
    banner.save(banner_path)

    side = Image.new("RGB", (180, 260), (50, 0, 50))
    draw = ImageDraw.Draw(side)
    for y in range(20, 240, 28):
        draw.rectangle((20, y, 160, y + 18), fill=(255, 215, 0))
    draw.text((50, 26), "777", fill=(50, 0, 50))
    draw.text((46, 82), "BAR", fill=(50, 0, 50))
    draw.text((38, 138), "SPIN", fill=(50, 0, 50))
    draw.text((37, 194), "WIN!", fill=(50, 0, 50))
    side.save(side_art_path)

    compat = Image.new("RGB", (420, 120), (50, 0, 50))
    draw = ImageDraw.Draw(compat)
    draw.rectangle((10, 10, 410, 110), outline=(255, 215, 0), width=4)
    draw.text((35, 24), "CASINO FLOOR COMPATIBILITY", fill=(255, 215, 0))
    draw.text((55, 64), "HOUSE RULES APPLIED", fill=(255, 80, 160))
    compat.save(compat_art_path)

    return {
        "banner": banner_path,
        "side": side_art_path,
        "compat": compat_art_path,
    }


def create_cyberpunk_assets(asset_dir: Path) -> dict[str, Path]:
    banner_path = asset_dir / "cyberpunk_banner.bmp"
    side_art_path = asset_dir / "cyberpunk_side.bmp"
    compat_art_path = asset_dir / "cyberpunk_compat.bmp"

    banner = Image.new("RGB", (900, 180), (18, 10, 36))
    draw = ImageDraw.Draw(banner)
    for x in range(0, 900, 40):
        draw.line((x, 0, x + 80, 180), fill=(255, 0, 180), width=2)
    draw.rectangle((20, 25, 880, 155), outline=(0, 255, 255), width=4)
    draw.text((250, 55), "CYBERPUNK GRID REPORT", fill=(0, 255, 255))
    draw.text((210, 100), "NEON SIGNALS // NIGHT CITY METRICS", fill=(255, 0, 180))
    banner.save(banner_path)

    side = Image.new("RGB", (180, 260), (10, 10, 20))
    draw = ImageDraw.Draw(side)
    for y in range(15, 250, 24):
        draw.line((15, y, 165, y), fill=(0, 255, 255), width=1)
    for x in range(15, 165, 24):
        draw.line((x, 15, x, 245), fill=(255, 0, 180), width=1)
    draw.rectangle((35, 40, 145, 90), outline=(255, 255, 0), width=3)
    draw.text((58, 58), "NEXUS", fill=(255, 255, 0))
    draw.rectangle((35, 130, 145, 180), outline=(0, 255, 255), width=3)
    draw.text((52, 148), "GRID", fill=(0, 255, 255))
    side.save(side_art_path)

    compat = Image.new("RGB", (420, 120), (10, 10, 20))
    draw = ImageDraw.Draw(compat)
    draw.rectangle((0, 0, 419, 119), outline=(0, 255, 255), width=3)
    draw.text((30, 24), "CYBERPUNK SYSTEM CHECK", fill=(0, 255, 255))
    draw.text((55, 64), "NEON STACK ONLINE", fill=(255, 0, 180))
    compat.save(compat_art_path)

    return {
        "banner": banner_path,
        "side": side_art_path,
        "compat": compat_art_path,
    }


def create_high_contrast_assets(asset_dir: Path) -> dict[str, Path]:
    banner_path = asset_dir / "high_contrast_banner.bmp"
    compat_art_path = asset_dir / "high_contrast_compat.bmp"

    banner = Image.new("RGB", (900, 180), (0, 0, 0))
    draw = ImageDraw.Draw(banner)
    draw.rectangle((20, 20, 880, 160), outline=(255, 255, 0), width=5)
    draw.text((180, 55), "HIGH CONTRAST COMMAND REPORT", fill=(255, 255, 255))
    draw.text((260, 100), "MAXIMUM LEGIBILITY MODE", fill=(255, 255, 0))
    banner.save(banner_path)

    compat = Image.new("RGB", (420, 120), (0, 0, 0))
    draw = ImageDraw.Draw(compat)
    draw.rectangle((8, 8, 412, 112), outline=(255, 255, 255), width=3)
    draw.text((18, 26), "ACCESSIBILITY COMPATIBILITY", fill=(255, 255, 0))
    draw.text((38, 66), "HIGH VISIBILITY ENABLED", fill=(255, 255, 255))
    compat.save(compat_art_path)

    return {
        "banner": banner_path,
        "compat": compat_art_path,
    }


def create_medieval_assets(asset_dir: Path) -> dict[str, Path]:
    banner_path = asset_dir / "medieval_banner.bmp"
    side_art_path = asset_dir / "medieval_side.bmp"
    compat_art_path = asset_dir / "medieval_compat.bmp"

    banner = Image.new("RGB", (900, 180), (245, 235, 207))
    draw = ImageDraw.Draw(banner)
    draw.rectangle((10, 10, 890, 170), outline=(168, 123, 29), width=6)
    draw.rectangle((28, 28, 872, 152), outline=(129, 32, 32), width=2)
    draw.text((210, 38), "ROYAL LEDGER OF THE REALM", fill=(73, 46, 26))
    draw.text((255, 88), "ILLUMINATED MONTHLY RECORD", fill=(129, 32, 32))
    draw.ellipse((40, 45, 90, 95), outline=(168, 123, 29), width=3)
    draw.ellipse((810, 45, 860, 95), outline=(168, 123, 29), width=3)
    banner.save(banner_path)

    side = Image.new("RGB", (180, 260), (245, 235, 207))
    draw = ImageDraw.Draw(side)
    draw.rectangle((12, 12, 168, 248), outline=(168, 123, 29), width=4)
    draw.text((36, 22), "S C R I B E", fill=(73, 46, 26))
    draw.rectangle((30, 60, 150, 100), outline=(129, 32, 32), width=2)
    draw.text((56, 74), "SEAL", fill=(129, 32, 32))
    draw.rectangle((30, 130, 150, 170), outline=(168, 123, 29), width=2)
    draw.text((54, 144), "CROWN", fill=(73, 46, 26))
    draw.line((38, 208, 142, 208), fill=(129, 32, 32), width=3)
    side.save(side_art_path)

    compat = Image.new("RGB", (420, 120), (245, 235, 207))
    draw = ImageDraw.Draw(compat)
    draw.rectangle((8, 8, 412, 112), outline=(168, 123, 29), width=4)
    draw.text((35, 22), "SCRIBE'S COMPATIBILITY CHRONICLE", fill=(73, 46, 26))
    draw.text((72, 66), "SEALED BY THE ARCHIVIST", fill=(129, 32, 32))
    compat.save(compat_art_path)

    return {"banner": banner_path, "side": side_art_path, "compat": compat_art_path}


def create_treasure_map_assets(asset_dir: Path) -> dict[str, Path]:
    banner_path = asset_dir / "treasure_banner.bmp"
    side_art_path = asset_dir / "treasure_side.bmp"
    compat_art_path = asset_dir / "treasure_compat.bmp"

    banner = Image.new("RGB", (900, 180), (240, 228, 181))
    draw = ImageDraw.Draw(banner)
    draw.rectangle((18, 18, 882, 162), outline=(163, 109, 26), width=4)
    draw.text((215, 34), "FANTASY FOREST TREASURE MAP", fill=(60, 68, 38))
    draw.text((260, 78), "SURVEY OF HIDDEN BOOTY", fill=(143, 76, 35))
    draw.line((80, 130, 180, 90), fill=(143, 76, 35), width=4)
    draw.line((180, 90, 280, 120), fill=(143, 76, 35), width=4)
    draw.text((290, 106), "X", fill=(163, 109, 26))
    draw.ellipse((720, 40, 830, 150), outline=(44, 78, 55), width=3)
    draw.line((775, 48, 775, 142), fill=(44, 78, 55), width=2)
    draw.line((728, 95, 822, 95), fill=(44, 78, 55), width=2)
    banner.save(banner_path)

    side = Image.new("RGB", (180, 260), (240, 228, 181))
    draw = ImageDraw.Draw(side)
    draw.rectangle((12, 12, 168, 248), outline=(163, 109, 26), width=3)
    draw.ellipse((45, 28, 135, 118), outline=(44, 78, 55), width=3)
    draw.line((90, 36, 90, 110), fill=(44, 78, 55), width=2)
    draw.line((52, 72, 128, 72), fill=(44, 78, 55), width=2)
    draw.text((79, 42), "N", fill=(44, 78, 55))
    draw.text((80, 91), "S", fill=(44, 78, 55))
    draw.text((57, 68), "W", fill=(44, 78, 55))
    draw.text((116, 68), "E", fill=(44, 78, 55))
    draw.text((42, 160), "X MARKS", fill=(143, 76, 35))
    draw.text((62, 188), "THE SPOT", fill=(143, 76, 35))
    side.save(side_art_path)

    compat = Image.new("RGB", (420, 120), (240, 228, 181))
    draw = ImageDraw.Draw(compat)
    draw.rectangle((10, 10, 410, 110), outline=(163, 109, 26), width=3)
    draw.text((35, 22), "TREASURE MAP SURVEY", fill=(60, 68, 38))
    draw.text((66, 64), "ALL TRAILS ACCOUNTED FOR", fill=(143, 76, 35))
    compat.save(compat_art_path)

    return {"banner": banner_path, "side": side_art_path, "compat": compat_art_path}


def create_outrun_assets(asset_dir: Path) -> dict[str, Path]:
    banner_path = asset_dir / "outrun_banner.bmp"
    side_art_path = asset_dir / "outrun_side.bmp"
    compat_art_path = asset_dir / "outrun_compat.bmp"

    banner = Image.new("RGB", (900, 180), (20, 8, 40))
    draw = ImageDraw.Draw(banner)
    draw.rectangle((0, 0, 899, 179), fill=(30, 12, 58))
    draw.ellipse((330, 24, 570, 170), fill=(255, 94, 194), outline=(255, 196, 0))
    for y in range(95, 170, 12):
        draw.line((120, y, 780, y), fill=(50, 236, 255), width=2)
    for x in range(180, 721, 45):
        draw.line((450, 170, x, 95), fill=(50, 236, 255), width=1)
    draw.text((248, 32), "SYNTHWAVE OUTRUN REPORT", fill=(241, 238, 255))
    banner.save(banner_path)

    side = Image.new("RGB", (180, 260), (20, 8, 40))
    draw = ImageDraw.Draw(side)
    for y in range(20, 240, 20):
        draw.line((15, y, 165, y), fill=(50, 236, 255), width=1)
    for x in range(20, 160, 20):
        draw.line((x, 20, x, 240), fill=(255, 94, 194), width=1)
    draw.rectangle((30, 42, 150, 90), outline=(255, 196, 0), width=3)
    draw.text((51, 57), "SUNSET", fill=(255, 196, 0))
    draw.rectangle((30, 146, 150, 194), outline=(50, 236, 255), width=3)
    draw.text((58, 161), "GRID", fill=(50, 236, 255))
    side.save(side_art_path)

    compat = Image.new("RGB", (420, 120), (20, 8, 40))
    draw = ImageDraw.Draw(compat)
    draw.rectangle((8, 8, 412, 112), outline=(255, 196, 0), width=3)
    draw.text((35, 24), "OUTRUN SYSTEM CHECK", fill=(241, 238, 255))
    draw.text((55, 68), "HORIZON LOCK ENGAGED", fill=(50, 236, 255))
    compat.save(compat_art_path)

    return {"banner": banner_path, "side": side_art_path, "compat": compat_art_path}


def create_brutalist_assets(asset_dir: Path) -> dict[str, Path]:
    banner_path = asset_dir / "brutalist_banner.bmp"
    side_art_path = asset_dir / "brutalist_side.bmp"
    compat_art_path = asset_dir / "brutalist_compat.bmp"

    banner = Image.new("RGB", (900, 180), (255, 255, 255))
    draw = ImageDraw.Draw(banner)
    draw.rectangle((0, 0, 899, 45), fill=(0, 0, 0))
    draw.rectangle((0, 135, 899, 179), fill=(0, 0, 0))
    draw.rectangle((40, 55, 220, 125), fill=(0, 0, 0))
    draw.rectangle((260, 55, 860, 125), outline=(0, 0, 0), width=4)
    draw.text((280, 76), "BRUTALIST MONOCHROME REPORT", fill=(0, 0, 0))
    banner.save(banner_path)

    side = Image.new("RGB", (180, 260), (255, 255, 255))
    draw = ImageDraw.Draw(side)
    draw.rectangle((15, 15, 165, 245), outline=(0, 0, 0), width=4)
    draw.rectangle((30, 30, 150, 70), fill=(0, 0, 0))
    draw.rectangle((30, 90, 90, 150), fill=(0, 0, 0))
    draw.rectangle((105, 90, 150, 150), outline=(0, 0, 0), width=3)
    draw.rectangle((30, 170, 150, 215), outline=(0, 0, 0), width=3)
    side.save(side_art_path)

    compat = Image.new("RGB", (420, 120), (255, 255, 255))
    draw = ImageDraw.Draw(compat)
    draw.rectangle((8, 8, 412, 112), outline=(0, 0, 0), width=4)
    draw.rectangle((20, 20, 140, 96), fill=(0, 0, 0))
    draw.text((165, 38), "MONO COMPATIBILITY", fill=(0, 0, 0))
    draw.text((165, 70), "STRUCTURE VERIFIED", fill=(0, 0, 0))
    compat.save(compat_art_path)

    return {"banner": banner_path, "side": side_art_path, "compat": compat_art_path}


def create_theme_assets(asset_dir: Path, theme: ThemeProfile) -> dict[str, Path] | None:
    if theme.name == "star-wars":
        return create_star_wars_assets(asset_dir)
    if theme.name == "retro-90s":
        return create_windows_90s_assets(asset_dir)
    if theme.name == "vegas-casino":
        return create_vegas_assets(asset_dir)
    if theme.name == "cyberpunk":
        return create_cyberpunk_assets(asset_dir)
    if theme.name == "high-contrast":
        return create_high_contrast_assets(asset_dir)
    if theme.name == "medieval-manuscript":
        return create_medieval_assets(asset_dir)
    if theme.name == "fantasy-forest-map":
        return create_treasure_map_assets(asset_dir)
    if theme.name == "synthwave-outrun":
        return create_outrun_assets(asset_dir)
    if theme.name == "brutalist-monochrome":
        return create_brutalist_assets(asset_dir)
    return None


def apply_theme_art(
    worksheet: xlwt.Worksheet,
    sheet_name: str,
    theme: ThemeProfile,
    assets: dict[str, Path] | None,
) -> None:
    if not assets:
        return
    if sheet_name == "Current":
        worksheet.insert_bitmap(str(assets["banner"]), 0, 0, scale_x=0.72, scale_y=0.72)
        if "side" in assets:
            worksheet.insert_bitmap(str(assets["side"]), 71, 9, scale_x=0.55, scale_y=0.55)
    elif sheet_name == "Template":
        worksheet.insert_bitmap(str(assets["banner"]), 0, 0, scale_x=0.72, scale_y=0.72)
    elif sheet_name == "Compatibility Report":
        worksheet.insert_bitmap(str(assets["compat"]), 0, 1, scale_x=0.68, scale_y=0.68)


def restyle_sheet(
    book: xlrd.book.Book,
    workbook: xlwt.Workbook,
    sheet_name: str,
    theme: ThemeProfile,
    output_filename: str,
    assets: dict[str, Path] | None,
) -> None:
    source = book.sheet_by_name(sheet_name)
    target = workbook.add_sheet(sheet_name, cell_overwrite_ok=True)
    style_factory = workbook._style_factory  # type: ignore[attr-defined]
    rows = retained_rows(sheet_name, source)
    row_map = {source_rowx: target_rowx for target_rowx, source_rowx in enumerate(rows)}

    _, max_col = content_bounds(source)
    max_row = len(rows) - 1
    if sheet_name == "Template":
        max_col = TEMPLATE_LAST_COL
    if sheet_name == "Current":
        max_col = max(max_col, CURRENT_FORMULA_COL)

    merged = merged_ranges(sheet_name, source, row_map, max_col)
    covered = covered_cells(merged)

    if sheet_name == "Compatibility Report":
        max_col = max(max_col, source.ncols - 1)
        for colx in range(max_col + 1):
            source_col = source.colinfo_map.get(colx)
            target.col(colx).width = source_col.width if source_col else 3584

        for rowx in range(max_row + 1):
            source_row = source.rowinfo_map.get(rowx)
            target.row(rowx).height_mismatch = True
            target.row(rowx).height = source_row.height if source_row else 360
    else:
        for colx in range(max_col + 1):
            if sheet_name == "Current" and colx == CURRENT_FORMULA_COL:
                target.col(colx).width = CURRENT_FORMULA_WIDTH
            else:
                target.col(colx).width = 4096

        for rowx in range(max_row + 1):
            target.row(rowx).height_mismatch = True
            target.row(rowx).height = 420

    for target_rowx, source_rowx in enumerate(rows):
        role = row_role(sheet_name, source, source_rowx)
        for colx in range(max_col + 1):
            if (target_rowx, colx) in covered and (target_rowx, colx) not in merged:
                continue

            cell = source.cell(source_rowx, colx) if source_rowx < source.nrows and colx < source.ncols else xlrd.sheet.Cell(
                xlrd.XL_CELL_EMPTY, ""
            )
            value = themed_cell_value(theme, sheet_name, output_filename, source_rowx, colx, cell)
            fmt = format_for_cell(book, source, source_rowx, colx) if source_rowx < source.nrows and colx < source.ncols else "General"
            spec = style_spec_for_cell(theme, sheet_name, role, target_rowx, colx, value, fmt)
            style = style_factory.make(spec, fmt if fmt else "General")

            if (target_rowx, colx) in merged:
                row_lo, row_hi, col_lo, col_hi = merged[(target_rowx, colx)]
                target.write_merge(row_lo, row_hi, col_lo, col_hi, value, style)
            else:
                write_value(target, target_rowx, colx, cell, value, style)

    apply_theme_art(target, sheet_name, theme, assets)


def build_workbook(input_path: str, output_path: str, theme_name: str = "classic") -> None:
    source_book = xlrd.open_workbook(input_path, formatting_info=True)
    target_book = xlwt.Workbook(style_compression=2)
    theme = THEMES[theme_name]
    output_filename = Path(output_path).name
    configure_palette(target_book, theme.palette)
    target_book._style_factory = StyleFactory(target_book, theme.palette)  # type: ignore[attr-defined]
    with TemporaryDirectory() as tmpdir:
        assets = create_theme_assets(Path(tmpdir), theme)
        for sheet_name in source_book.sheet_names():
            restyle_sheet(source_book, target_book, sheet_name, theme, output_filename, assets)
        target_book.save(output_path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Restyle the Goldman monthly workbook.")
    parser.add_argument("input_path", nargs="?", default="Goldman_Monthly_Report.xls")
    parser.add_argument("output_path", nargs="?", default="Goldman_Monthly_Report.xls")
    parser.add_argument("--theme", choices=sorted(THEMES), default="classic")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    build_workbook(args.input_path, args.output_path, args.theme)


if __name__ == "__main__":
    main()
