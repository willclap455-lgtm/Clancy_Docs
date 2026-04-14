from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path

import xlrd
import xlwt


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
            font_colour="sage" if (is_clancy or is_oldschool) else "black",
            border_colour="stone",
            left=0,
            right=0,
            top=0,
            bottom=0,
            horz=xlwt.Alignment.HORZ_CENTER,
            wrap=False,
        )

    if role == "title":
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
        if is_oldschool:
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
        if is_oldschool:
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
        if is_oldschool:
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
                font_name=heading_font if (is_clancy or is_oldschool) else body_font,
                font_height=220,
                bold=True,
                font_colour=font_colour,
                bg_colour=bg_colour,
                border_colour="stone" if is_oldschool else "gold" if is_clancy else "slate",
                left=1,
                right=1,
                top=1,
                bottom=2,
                horz=xlwt.Alignment.HORZ_RIGHT,
            )
        return ThemeStyle(
            font_name=heading_font if (is_clancy or is_oldschool) else body_font,
            font_height=220,
            bold=True,
            font_colour=font_colour,
            bg_colour=bg_colour,
            border_colour="stone" if is_oldschool else "gold" if is_clancy else "slate",
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
            font_colour="sage" if (is_clancy or is_oldschool) else "navy",
            bg_colour="ivory" if is_oldschool else "mist" if is_clancy_dark else "ivory" if is_clancy_light else "ivory",
            border_colour="stone" if is_oldschool else "gold" if is_clancy else "slate",
            left=1,
            right=1,
            top=1,
            bottom=2,
            horz=xlwt.Alignment.HORZ_CENTER if colx >= 1 else xlwt.Alignment.HORZ_LEFT,
        )

    if role in {"sent", "received", "body"}:
        if is_oldschool:
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
            font_colour="sage" if (is_clancy or is_oldschool) else "navy" if role != "body" else "black",
            bg_colour=bg_colour,
            border_colour="stone",
            horz=xlwt.Alignment.HORZ_LEFT if colx == 0 else xlwt.Alignment.HORZ_CENTER,
            bold=colx == 7 and sheet_name != "Compatibility Report",
        )
        if colx == 1:
            return ThemeStyle(
                font_name=body_font,
                font_height=200,
                bold=True,
                font_colour="navy" if is_oldschool else "gold" if is_clancy else "navy",
                bg_colour=bg_colour,
                border_colour="stone",
                horz=xlwt.Alignment.HORZ_CENTER,
            )
        if colx == 7 and sheet_name != "Compatibility Report":
            return ThemeStyle(
                font_name=body_font,
                font_height=200,
                bold=True,
                font_colour="navy" if is_oldschool else "white" if is_clancy else "navy",
                bg_colour="gold" if (is_clancy or is_oldschool) else None,
                border_colour="stone" if is_oldschool else "gold" if is_clancy else "stone",
                horz=xlwt.Alignment.HORZ_CENTER,
            )
        if "%" in fmt or colx >= 8:
            return ThemeStyle(
                font_name=body_font,
                font_height=200,
                font_colour="sage" if is_oldschool else "slate" if is_clancy else "navy",
                bg_colour="ivory" if (is_clancy or is_oldschool) and colx <= MAIN_TABLE_LAST_COL else None,
                border_colour="stone",
                horz=xlwt.Alignment.HORZ_CENTER,
            )
        if colx >= 2:
            return ThemeStyle(
                font_name=body_font,
                font_height=200,
                font_colour="sage" if (is_clancy or is_oldschool) else "navy" if role != "body" else "black",
                bg_colour=bg_colour,
                border_colour="stone",
                horz=xlwt.Alignment.HORZ_CENTER,
            )
        return spec

    if role == "compat_title":
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
            bg_colour="ivory" if is_oldschool else "mist",
            border_colour="stone",
            horz=xlwt.Alignment.HORZ_LEFT,
        )

    if role == "compat_body":
        return centered_right if colx == 4 else centered if colx == 5 else ThemeStyle(
            font_name=body_font,
            font_height=200,
            bg_colour="mist" if is_oldschool else "ivory",
            border_colour="stone",
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


def restyle_sheet(
    book: xlrd.book.Book,
    workbook: xlwt.Workbook,
    sheet_name: str,
    theme: ThemeProfile,
    output_filename: str,
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
            value = output_cell_value(sheet_name, output_filename, source_rowx, colx, cell)
            fmt = format_for_cell(book, source, source_rowx, colx) if source_rowx < source.nrows and colx < source.ncols else "General"
            spec = style_spec_for_cell(theme, sheet_name, role, target_rowx, colx, value, fmt)
            style = style_factory.make(spec, fmt if fmt else "General")

            if (target_rowx, colx) in merged:
                row_lo, row_hi, col_lo, col_hi = merged[(target_rowx, colx)]
                target.write_merge(row_lo, row_hi, col_lo, col_hi, value, style)
            else:
                write_value(target, target_rowx, colx, cell, value, style)


def build_workbook(input_path: str, output_path: str, theme_name: str = "classic") -> None:
    source_book = xlrd.open_workbook(input_path, formatting_info=True)
    target_book = xlwt.Workbook(style_compression=2)
    theme = THEMES[theme_name]
    output_filename = Path(output_path).name
    configure_palette(target_book, theme.palette)
    target_book._style_factory = StyleFactory(target_book, theme.palette)  # type: ignore[attr-defined]

    for sheet_name in source_book.sheet_names():
        restyle_sheet(source_book, target_book, sheet_name, theme, output_filename)

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
