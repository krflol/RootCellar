#!/usr/bin/env python3
"""Generate deterministic XLSX fixtures for part-graph corpus validation."""

from __future__ import annotations

import argparse
import pathlib
import zipfile


def build_content_types(include_shared_strings: bool) -> str:
    shared_override = (
        '<Override PartName="/xl/sharedStrings.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        if include_shared_strings
        else ""
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        f"{shared_override}"
        "</Types>"
    )


def build_root_rels() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        "</Relationships>"
    )


def build_workbook_xml(sheet_name: str = "Sheet1") -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<sheets>"
        f'<sheet name="{sheet_name}" sheetId="1" r:id="rId1"/>'
        "</sheets>"
        "</workbook>"
    )


def build_workbook_rels(extra_dangling: bool = False) -> str:
    dangling = (
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/missing-sheet.xml"/>'
        if extra_dangling
        else ""
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet1.xml"/>'
        f"{dangling}"
        "</Relationships>"
    )


def write_fixture(
    out_path: pathlib.Path,
    sheet_xml: str,
    *,
    include_shared_strings: bool = False,
    shared_strings_xml: str | None = None,
    extra_dangling: bool = False,
    extra_parts: dict[str, bytes] | None = None,
) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out_path, mode="w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", build_content_types(include_shared_strings))
        archive.writestr("_rels/.rels", build_root_rels())
        archive.writestr("xl/workbook.xml", build_workbook_xml())
        archive.writestr("xl/_rels/workbook.xml.rels", build_workbook_rels(extra_dangling))
        archive.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        if include_shared_strings and shared_strings_xml:
            archive.writestr("xl/sharedStrings.xml", shared_strings_xml)
        if extra_parts:
            for part_name, payload in extra_parts.items():
                archive.writestr(part_name, payload)


def generate(output_dir: pathlib.Path) -> list[pathlib.Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    created: list[pathlib.Path] = []

    minimal_sheet = (
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>"
        "</worksheet>"
    )
    minimal_path = output_dir / "minimal.xlsx"
    write_fixture(minimal_path, minimal_sheet)
    created.append(minimal_path)

    formula_sheet = (
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        "<sheetData>"
        "<row r=\"1\"><c r=\"A1\"><v>10</v></c></row>"
        "<row r=\"2\"><c r=\"A2\"><v>2</v></c></row>"
        "<row r=\"3\"><c r=\"A3\"><f>A1+A2</f><v>12</v></c></row>"
        "</sheetData>"
        "</worksheet>"
    )
    formula_path = output_dir / "formula.xlsx"
    write_fixture(formula_path, formula_sheet)
    created.append(formula_path)

    dangling_sheet = (
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        "<sheetData><row r=\"1\"><c r=\"A1\"><v>7</v></c></row></sheetData>"
        "</worksheet>"
    )
    dangling_path = output_dir / "dangling-edge.xlsx"
    write_fixture(dangling_path, dangling_sheet, extra_dangling=True)
    created.append(dangling_path)

    unknown_part_sheet = (
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        "<sheetData><row r=\"1\"><c r=\"A1\"><v>99</v></c></row></sheetData>"
        "</worksheet>"
    )
    unknown_path = output_dir / "unknown-part.xlsx"
    write_fixture(
        unknown_path,
        unknown_part_sheet,
        extra_parts={"xCustom/blob.bin": b"opaque"},
    )
    created.append(unknown_path)

    return created


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate deterministic XLSX fixtures for corpus validation."
    )
    parser.add_argument(
        "--output-dir",
        default=".ci/corpus-fixtures",
        help="Directory to write generated .xlsx fixtures.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    output_dir = pathlib.Path(args.output_dir)
    generated = generate(output_dir)
    print(f"Generated {len(generated)} fixture(s) in {output_dir}")
    for path in generated:
        print(f" - {path}")


if __name__ == "__main__":
    main()
