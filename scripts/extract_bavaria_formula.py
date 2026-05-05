#!/usr/bin/env python3
from __future__ import annotations
import argparse
import csv
import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from pathlib import Path

NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def excel_serial_to_date(value: float) -> str:
    base = datetime(1899, 12, 30)
    return (base + timedelta(days=float(value))).strftime("%Y-%m-%d")


def col_to_num(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def num_to_col(n: int) -> str:
    s = ""
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def parse_xlsx_cells(xlsx_path: Path, sheet_regex: str) -> tuple[str, dict[str, str]]:
    with zipfile.ZipFile(xlsx_path) as zf:
        wb = ET.fromstring(zf.read("xl/workbook.xml"))
        sheets = []
        for sh in wb.find("a:sheets", NS):
            sheets.append((sh.attrib["name"], sh.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]))

        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rid_to_target = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}

        shared = []
        if "xl/sharedStrings.xml" in zf.namelist():
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in root.findall("a:si", NS):
                texts = [t.text or "" for t in si.findall(".//a:t", NS)]
                shared.append("".join(texts))

        match_name, match_target = None, None
        pattern = re.compile(sheet_regex)
        for name, rid in sheets:
            if pattern.search(name):
                match_name = name
                match_target = rid_to_target[rid]
                break

        if not match_name:
            raise ValueError(f"No worksheet matched regex: {sheet_regex}")

        xml_path = "xl/" + match_target.lstrip("/")
        ws_xml = ET.fromstring(zf.read(xml_path))

        data: dict[str, str] = {}
        for c in ws_xml.findall(".//a:sheetData/a:row/a:c", NS):
            ref = c.attrib.get("r")
            if not ref:
                continue
            t = c.attrib.get("t")
            v = c.find("a:v", NS)
            if v is None or v.text is None:
                data[ref] = ""
            elif t == "s":
                data[ref] = shared[int(v.text)]
            else:
                data[ref] = v.text

        return match_name, data


def extract_to_long(data: dict[str, str], sheet_name: str, row_start: int, row_end: int, col_start: str, col_end: str):
    c_start, c_end = col_to_num(col_start), col_to_num(col_end)
    month_labels = []
    for c in range(c_start, c_end + 1):
        col_ref = f"{num_to_col(c)}{row_start}"
        raw = data.get(col_ref, "")
        label = raw
        if raw:
            try:
                f = float(raw)
                if f > 30000:
                    label = excel_serial_to_date(f)
            except ValueError:
                pass
        month_labels.append((num_to_col(c), raw, label))

    rows = []
    for r in range(row_start + 1, row_end + 1):
        metric = data.get(f"B{r}", "") or data.get(f"C{r}", "")
        for col_letter, raw_month, norm_month in month_labels:
            val = data.get(f"{col_letter}{r}", "")
            rows.append({
                "source_file": "1 Bavaria - Precio Marzo.xlsx",
                "sheet_name": sheet_name,
                "metric_row": r,
                "metric_name": metric,
                "column_letter": col_letter,
                "period_raw": raw_month,
                "period_normalized": norm_month,
                "value_raw": val,
            })
    return rows


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract Bavaria formula range and normalize to long format.")
    parser.add_argument("--input", default="data/1 Bavaria - Precio Marzo.xlsx")
    parser.add_argument("--output", default="output/bavaria_formula_mar26_long.csv")
    parser.add_argument("--sheet-regex", default=r"^Fórmula.*26")
    args = parser.parse_args()

    sheet_name, data = parse_xlsx_cells(Path(args.input), args.sheet_regex)
    rows = extract_to_long(data, sheet_name, 15, 26, "C", "CV")

    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    with out.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)

    print(f"Sheet selected: {sheet_name}")
    print(f"Rows written: {len(rows)}")
    print(f"Output: {out}")


if __name__ == "__main__":
    main()
