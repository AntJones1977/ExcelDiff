"""Load a .xlsx file and parse it into a WorkbookData snapshot."""

from __future__ import annotations

import openpyxl

from exceldiff.parsing.models import DefinedNameData, WorkbookData
from exceldiff.parsing.sheet_parser import parse_sheet


def parse_workbook(file_path: str) -> WorkbookData:
    """Load an Excel workbook and return a fully parsed WorkbookData snapshot.

    Uses data_only=False to read formulas instead of cached values.
    Uses read_only=False to access merged cells, styles, and data validations.
    """
    wb = openpyxl.load_workbook(file_path, data_only=False, read_only=False)

    sheet_names = tuple(wb.sheetnames)
    sheets = {}
    for name in sheet_names:
        ws = wb[name]
        sheets[name] = parse_sheet(ws)

    # Extract defined names
    defined_names: list[DefinedNameData] = []
    for name in wb.defined_names:
        dn = wb.defined_names[name]
        scope_name = None
        if dn.localSheetId is not None and dn.localSheetId < len(sheet_names):
            scope_name = sheet_names[dn.localSheetId]
        defined_names.append(
            DefinedNameData(
                name=dn.name,
                value=str(dn.attr_text),
                scope=scope_name,
            )
        )

    wb.close()

    return WorkbookData(
        file_path=file_path,
        sheet_names=sheet_names,
        sheets=sheets,
        defined_names=tuple(defined_names),
    )
