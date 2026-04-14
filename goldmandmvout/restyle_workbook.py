from __future__ import annotations

import argparse
from dataclasses import dataclass

import xlrd
import xlwt


TITLE_COLS_CURRENT = 9

PALETTE = {
    "navy": (0x21, (31, 53, 79)),
    "slate": (0x22, (69, 104, 142)),
    "mist": (0x23, (223, 235, 247)),
    "ivory": (0x24, (248, 250, 252)),
    "gold": (0x25, (214, 178, 61)),
    "coral": (0x26, (194, 87, 87)),
    "sage": (0x27, (129, 161, 135)),
    "stone": (0x28, (161, 171, 181)),
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
    def __init__(self, workbook: xlwt.Workbook) -> None:
        self.workbook = workbook
        self.cache: dict[tuple[str, str], xlwt.XFStyle] = {}

    def color_index(self, color_name: str) -> int:
        if color_name == "black":
            return 0x08
        if color_name == "white":
            return 0x09
        return PALETTE[color_name][0]

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


def configure_palette(workbook: xlwt.Workbook) -> None:
    for _, (index, rgb) in PALETTE.items():
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
    if sheet_name == "Template" and rowx in {0, 6}:
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


def style_spec_for_cell(
    sheet_name: str,
    role: str,
    rowx: int,
    colx: int,
    value: object,
    fmt: str,
) -> ThemeStyle:
    centered = ThemeStyle(horz=xlwt.Alignment.HORZ_CENTER, vert=xlwt.Alignment.VERT_CENTER)
    centered_right = ThemeStyle(horz=xlwt.Alignment.HORZ_RIGHT, vert=xlwt.Alignment.VERT_CENTER)

    if role == "title":
        return ThemeStyle(
            font_height=320,
            bold=True,
            font_colour="white",
            bg_colour="navy",
            border_colour="navy",
            horz=xlwt.Alignment.HORZ_CENTER,
            wrap=True,
        )

    if role == "section_header":
        if colx == 0:
            return ThemeStyle(
                font_height=220,
                bold=True,
                font_colour="white",
                bg_colour="slate",
                border_colour="navy",
                left=2,
                right=2,
                top=2,
                bottom=2,
            )
        if str(value).strip().upper() == "TOTAL":
            return ThemeStyle(
                font_height=220,
                bold=True,
                font_colour="white",
                bg_colour="gold",
                border_colour="navy",
                left=2,
                right=2,
                top=2,
                bottom=2,
                horz=xlwt.Alignment.HORZ_CENTER,
            )
        return ThemeStyle(
            font_height=220,
            bold=True,
            font_colour="white",
            bg_colour="slate",
            border_colour="navy",
            left=2,
            right=2,
            top=2,
            bottom=2,
            horz=xlwt.Alignment.HORZ_CENTER,
        )

    if role == "total_sent":
        bg_colour = "gold"
        font_colour = "white"
    elif role == "total_rec":
        bg_colour = "sage"
        font_colour = "white"
    elif role == "not_found":
        bg_colour = "coral"
        font_colour = "white"
    else:
        bg_colour = None
        font_colour = "black"

    if role in {"total_sent", "total_rec", "not_found"}:
        if colx == 0:
            return ThemeStyle(
                font_height=220,
                bold=True,
                font_colour=font_colour,
                bg_colour=bg_colour,
                border_colour="navy",
                left=2,
                right=2,
                top=2,
                bottom=2,
                horz=xlwt.Alignment.HORZ_RIGHT,
            )
        return ThemeStyle(
            font_height=220,
            bold=True,
            font_colour=font_colour,
            bg_colour=bg_colour,
            border_colour="navy",
            left=2,
            right=2,
            top=2,
            bottom=2,
            horz=xlwt.Alignment.HORZ_CENTER,
        )

    if role == "note":
        return ThemeStyle(
            bold=colx == 7,
            italic=colx in {0, 1},
            font_colour="navy",
            bg_colour="mist",
            border_colour="slate",
            left=1,
            right=1,
            top=1,
            bottom=2,
            horz=xlwt.Alignment.HORZ_CENTER if colx >= 1 else xlwt.Alignment.HORZ_LEFT,
        )

    if role in {"sent", "received", "body"}:
        bg_colour = "ivory" if role == "sent" else "mist" if role == "received" else None
        spec = ThemeStyle(
            font_height=200,
            font_colour="navy" if role != "body" else "black",
            bg_colour=bg_colour,
            border_colour="stone",
            horz=xlwt.Alignment.HORZ_LEFT if colx == 0 else xlwt.Alignment.HORZ_CENTER,
            bold=colx == 7 and sheet_name != "Compatibility Report",
        )
        if colx == 1:
            return ThemeStyle(
                font_height=200,
                bold=True,
                font_colour="navy",
                bg_colour=bg_colour,
                border_colour="stone",
                horz=xlwt.Alignment.HORZ_CENTER,
            )
        if colx == 7 and sheet_name != "Compatibility Report":
            return ThemeStyle(
                font_height=200,
                bold=True,
                font_colour="navy",
                bg_colour="gold",
                border_colour="slate",
                horz=xlwt.Alignment.HORZ_CENTER,
            )
        if "%" in fmt or colx >= 8:
            return ThemeStyle(
                font_height=200,
                font_colour="navy",
                bg_colour="mist" if role == "sent" else "ivory",
                border_colour="stone",
                horz=xlwt.Alignment.HORZ_CENTER,
            )
        if colx >= 2:
            return ThemeStyle(
                font_height=200,
                font_colour="navy" if role != "body" else "black",
                bg_colour=bg_colour,
                border_colour="stone",
                horz=xlwt.Alignment.HORZ_CENTER,
            )
        return spec

    if role == "compat_title":
        return ThemeStyle(
            font_height=240,
            bold=True,
            font_colour="white",
            bg_colour="navy",
            border_colour="navy",
            horz=xlwt.Alignment.HORZ_LEFT,
        )

    if role == "compat_header":
        return ThemeStyle(
            font_height=220,
            bold=True,
            font_colour="white",
            bg_colour="slate",
            border_colour="navy",
            horz=xlwt.Alignment.HORZ_CENTER if colx >= 4 else xlwt.Alignment.HORZ_LEFT,
        )

    if role == "compat_text":
        return ThemeStyle(
            font_height=200,
            bg_colour="mist",
            border_colour="stone",
            horz=xlwt.Alignment.HORZ_LEFT,
        )

    if role == "compat_body":
        return centered_right if colx == 4 else centered if colx == 5 else ThemeStyle(
            font_height=200,
            bg_colour="ivory",
            border_colour="stone",
        )

    return ThemeStyle()


def merged_ranges(sheet_name: str, sheet: xlrd.sheet.Sheet, max_col: int) -> dict[tuple[int, int], tuple[int, int, int, int]]:
    merged: dict[tuple[int, int], tuple[int, int, int, int]] = {
        (row_lo, col_lo): (row_lo, row_hi - 1, col_lo, col_hi - 1)
        for row_lo, row_hi, col_lo, col_hi in sheet.merged_cells
    }
    if sheet_name == "Current":
        merged[(1, 0)] = (1, 1, 0, max_col)
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
    style: xlwt.XFStyle,
) -> None:
    if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
        worksheet.write(rowx, colx, "", style)
    elif cell.ctype == xlrd.XL_CELL_TEXT:
        worksheet.write(rowx, colx, cell.value, style)
    elif cell.ctype == xlrd.XL_CELL_NUMBER:
        worksheet.write(rowx, colx, cell.value, style)
    elif cell.ctype == xlrd.XL_CELL_DATE:
        worksheet.write(rowx, colx, cell.value, style)
    elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
        worksheet.write(rowx, colx, bool(cell.value), style)
    elif cell.ctype == xlrd.XL_CELL_ERROR:
        worksheet.write(rowx, colx, xlrd.error_text_from_code.get(cell.value, ""), style)
    else:
        worksheet.write(rowx, colx, str(cell.value), style)


def restyle_sheet(book: xlrd.book.Book, workbook: xlwt.Workbook, sheet_name: str) -> None:
    source = book.sheet_by_name(sheet_name)
    target = workbook.add_sheet(sheet_name, cell_overwrite_ok=True)
    style_factory = workbook._style_factory  # type: ignore[attr-defined]

    max_row, max_col = content_bounds(source)
    if sheet_name == "Template":
        max_col = max(max_col, 14)
    if sheet_name == "Current":
        max_col = max(max_col, TITLE_COLS_CURRENT)

    merged = merged_ranges(sheet_name, source, max_col)
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
            target.col(colx).width = 4096

        for rowx in range(max_row + 1):
            target.row(rowx).height_mismatch = True
            target.row(rowx).height = 420

    for rowx in range(max_row + 1):
        role = row_role(sheet_name, source, rowx)
        for colx in range(max_col + 1):
            if (rowx, colx) in covered and (rowx, colx) not in merged:
                continue

            cell = source.cell(rowx, colx) if rowx < source.nrows and colx < source.ncols else xlrd.sheet.Cell(
                xlrd.XL_CELL_EMPTY, ""
            )
            fmt = format_for_cell(book, source, rowx, colx) if rowx < source.nrows and colx < source.ncols else "General"
            spec = style_spec_for_cell(sheet_name, role, rowx, colx, cell.value, fmt)
            style = style_factory.make(spec, fmt if fmt else "General")

            if (rowx, colx) in merged:
                row_lo, row_hi, col_lo, col_hi = merged[(rowx, colx)]
                target.write_merge(row_lo, row_hi, col_lo, col_hi, cell.value, style)
            else:
                write_value(target, rowx, colx, cell, style)


def build_workbook(input_path: str, output_path: str) -> None:
    source_book = xlrd.open_workbook(input_path, formatting_info=True)
    target_book = xlwt.Workbook(style_compression=2)
    configure_palette(target_book)
    target_book._style_factory = StyleFactory(target_book)  # type: ignore[attr-defined]

    for sheet_name in source_book.sheet_names():
        restyle_sheet(source_book, target_book, sheet_name)

    target_book.save(output_path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Restyle the Goldman monthly workbook.")
    parser.add_argument("input_path", nargs="?", default="Goldman_Monthly_Report.xls")
    parser.add_argument("output_path", nargs="?", default="Goldman_Monthly_Report.xls")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    build_workbook(args.input_path, args.output_path)


if __name__ == "__main__":
    main()
