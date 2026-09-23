import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def col_number(cell_ref):
    letters = re.match(r"[A-Z]+", cell_ref).group(0)
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def read_xml(zf, name):
    return ET.fromstring(zf.read(name))


def relationship_map(zf, rel_path):
    root = read_xml(zf, rel_path)
    return {r.attrib["Id"]: r.attrib["Target"] for r in root.findall("p:Relationship", NS)}


def normalize_target(base, target):
    parts = (Path(base) / target).parts
    stack = []
    for part in parts:
        if part == "..":
            stack.pop()
        elif part != ".":
            stack.append(part)
    return "/".join(stack)


def inspect(path):
    result = {"file": str(path), "sheets": []}
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        shared = []
        if "xl/sharedStrings.xml" in names:
            root = read_xml(zf, "xl/sharedStrings.xml")
            for si in root.findall("m:si", NS):
                shared.append("".join(t.text or "" for t in si.iterfind(".//m:t", NS)))

        wb = read_xml(zf, "xl/workbook.xml")
        rels = relationship_map(zf, "xl/_rels/workbook.xml.rels")
        defined_names = []
        dns = wb.find("m:definedNames", NS)
        if dns is not None:
            for dn in dns:
                defined_names.append({"name": dn.attrib.get("name"), "value": dn.text})
        result["defined_names"] = defined_names

        for sheet in wb.find("m:sheets", NS):
            rid = sheet.attrib[f"{{{NS['r']}}}id"]
            sheet_path = normalize_target("xl", rels[rid])
            root = read_xml(zf, sheet_path)
            info = {
                "name": sheet.attrib["name"],
                "state": sheet.attrib.get("state", "visible"),
                "dimension": None,
                "merged": [],
                "rows": [],
                "validations": [],
                "comments": [],
            }
            dim = root.find("m:dimension", NS)
            if dim is not None:
                info["dimension"] = dim.attrib.get("ref")
            merges = root.find("m:mergeCells", NS)
            if merges is not None:
                info["merged"] = [x.attrib["ref"] for x in merges]

            data = root.find("m:sheetData", NS)
            if data is not None:
                for row in data.findall("m:row", NS):
                    values = {}
                    for c in row.findall("m:c", NS):
                        ref = c.attrib.get("r")
                        typ = c.attrib.get("t")
                        formula = c.find("m:f", NS)
                        value = c.find("m:v", NS)
                        inline = c.find("m:is", NS)
                        rendered = None
                        if typ == "s" and value is not None:
                            rendered = shared[int(value.text)]
                        elif typ == "inlineStr" and inline is not None:
                            rendered = "".join(t.text or "" for t in inline.iterfind(".//m:t", NS))
                        elif value is not None:
                            rendered = value.text
                        if formula is not None:
                            rendered = {"formula": formula.text, "cached": rendered}
                        if rendered not in (None, ""):
                            values[ref] = rendered
                    if values:
                        info["rows"].append({"row": int(row.attrib["r"]), "cells": values})

            dvs = root.find("m:dataValidations", NS)
            if dvs is not None:
                for dv in dvs:
                    f1 = dv.find("m:formula1", NS)
                    f2 = dv.find("m:formula2", NS)
                    info["validations"].append({
                        "sqref": dv.attrib.get("sqref"),
                        "type": dv.attrib.get("type"),
                        "formula1": f1.text if f1 is not None else None,
                        "formula2": f2.text if f2 is not None else None,
                        "error": dv.attrib.get("error"),
                        "prompt": dv.attrib.get("prompt"),
                    })

            rel_name = str(Path(sheet_path).parent / "_rels" / (Path(sheet_path).name + ".rels")).replace("\\", "/")
            if rel_name in names:
                srels = relationship_map(zf, rel_name)
                for target in srels.values():
                    if "comments" not in target:
                        continue
                    cp = normalize_target(str(Path(sheet_path).parent), target)
                    cr = read_xml(zf, cp)
                    authors = [a.text or "" for a in cr.findall("m:authors/m:author", NS)]
                    for comment in cr.findall("m:commentList/m:comment", NS):
                        text = "".join(t.text or "" for t in comment.iterfind(".//m:t", NS))
                        aid = int(comment.attrib.get("authorId", 0))
                        info["comments"].append({"ref": comment.attrib["ref"], "author": authors[aid], "text": text})
            result["sheets"].append(info)
    return result


if __name__ == "__main__":
    for arg in sys.argv[1:]:
        print(json.dumps(inspect(Path(arg)), ensure_ascii=False, indent=2))
