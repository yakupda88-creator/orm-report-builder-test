from __future__ import annotations

import html
import math
import re
import zipfile
from collections import Counter, defaultdict
from datetime import datetime, timedelta
from pathlib import Path
from xml.etree import ElementTree as ET

from inspect_xlsx import inspect


HEADERS = ["Дата", "Время", "Текст", "Источник", "Ссылка", "Автор", "Дубли", "Тональность", "Продукт", "Процесс", "Классификатор"]
COLS = "ABCDEFGHIJK"
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
SOURCE_MAP = {"2ГИС": "2gis.ru", "Google": "google.com", "Яндекс": "maps.yandex.ru"}
CLASSIFIER_ALIASES = {
    "стоа (выбор, обслуживание)": "Не устраивает СТОА (выбор, обслуживание)",
    "не устраивает стоа (выбор, обслуживание)": "Не устраивает СТОА (выбор, обслуживание)",
}
MONTHS_RU = ["Янв", "Фев", "Мар", "Апр", "Май", "Июн", "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек"]
ALIASES = {
    "Дата": ["Дата", "Дата отзыва"], "Время": ["Время", "Время отзыва"],
    "Текст": ["Текст", "Отзыв"], "Источник": ["Источник"],
    "Ссылка": ["Ссылка", "URL"], "Автор": ["Автор"], "Дубли": ["Дубли"],
    "Тональность": ["Тональность"], "Продукт": ["Продукт"],
    "Процесс": ["Процесс"], "Классификатор": ["Классификатор"],
    "Удален": ["Удален", "Удалён"], "Геосервис": ["Геосервис"],
    "Оценка": ["Оценка"], "Тег": ["Тег"],
}


def norm(value):
    return re.sub(r"\s+", " ", str(value or "").strip()).casefold()


def canonical_classifier(value):
    """Merge legacy spellings that represent the same classifier."""
    return CLASSIFIER_ALIASES.get(norm(value), value)


def col_of(ref):
    return re.match(r"[A-Z]+", ref).group(0)


def row_values(sheet, number):
    row = next((r for r in sheet["rows"] if r["row"] == number), None)
    return row["cells"] if row else {}


def header_map(sheet):
    for row in sheet["rows"][:20]:
        values = {norm(v): col_of(k) for k, v in row["cells"].items() if not isinstance(v, dict)}
        if len(values) >= 4:
            yield row["row"], values


def find_sheet(book, required):
    requested = list(required)
    best = None
    for sheet in book["sheets"]:
        for row_no, mapping in header_map(sheet):
            resolved = dict(mapping)
            score = 0
            for field in requested:
                col = next((mapping.get(norm(alias)) for alias in ALIASES.get(field, [field]) if mapping.get(norm(alias))), None)
                if col:
                    score += 1
                    resolved[norm(field)] = col
            if best is None or score > best[0]:
                best = (score, sheet, row_no, resolved)
    if not best or best[0] < len(requested):
        raise ValueError("Не найден лист с полями: " + ", ".join(requested))
    return best[1], best[2], best[3]


def cell(cells, col, row):
    value = cells.get(f"{col}{row}", "")
    if isinstance(value, dict):
        return value.get("cached", "")
    return value


def excel_date(value):
    try:
        return datetime(1899, 12, 30) + timedelta(days=float(value))
    except Exception:
        for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
            try:
                return datetime.strptime(str(value).strip(), fmt)
            except ValueError:
                pass
    return None


def in_month(value, month):
    if not month:
        return True
    dt = excel_date(value)
    return bool(dt and dt.strftime("%Y-%m") == month)


def esc(value):
    return html.escape(str(value), quote=False)


def cell_xml(ref, value, style=0):
    if value in (None, ""):
        return f'<c r="{ref}" s="{style}"/>'
    text = str(value)
    if re.fullmatch(r"-?\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?", text):
        return f'<c r="{ref}" s="{style}"><v>{text}</v></c>'
    preserve = ' xml:space="preserve"' if text != text.strip() else ""
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t{preserve}>{esc(text)}</t></is></c>'


def worksheet_xml(rows, widths, freeze=True, autofilter=True):
    max_col = max((len(r[0]) for r in rows), default=1)
    max_row = len(rows)
    row_parts = []
    for number, (values, styles) in enumerate(rows, 1):
        cells = "".join(cell_xml(f"{chr(64+i)}{number}", value, styles[i-1] if i-1 < len(styles) else 0) for i, value in enumerate(values, 1))
        row_parts.append(f'<row r="{number}">{cells}</row>')
    cols_xml = "".join(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>' for i, w in enumerate(widths, 1))
    pane = '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' if freeze else ""
    filt = f'<autoFilter ref="A1:{chr(64+max_col)}{max_row}"/>' if autofilter and max_row else ""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<dimension ref="A1:{chr(64+max_col)}{max_row}"/><sheetViews><sheetView workbookViewId="0">{pane}</sheetView></sheetViews>'
        f'<sheetFormatPr defaultRowHeight="18"/><cols>{cols_xml}</cols><sheetData>{"".join(row_parts)}</sheetData>{filt}</worksheet>'
    )


def styles_xml():
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="4"><numFmt numFmtId="164" formatCode="dd.mm.yyyy"/><numFmt numFmtId="165" formatCode="hh:mm"/><numFmt numFmtId="166" formatCode="0.0%"/><numFmt numFmtId="167" formatCode="mm.yyyy"/></numFmts>
<fonts count="3"><font><sz val="9"/><name val="Arial"/></font><font><b/><color rgb="FF000000"/><sz val="9"/><name val="Arial"/></font><font><b/><color rgb="FF000000"/><sz val="9"/><name val="Arial"/></font></fonts>
<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9EAF7"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFEB9C"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border/><border><left style="thin"><color rgb="FFD9E1EC"/></left><right style="thin"><color rgb="FFD9E1EC"/></right><top style="thin"><color rgb="FFD9E1EC"/></top><bottom style="thin"><color rgb="FFD9E1EC"/></bottom></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="11"><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf><xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="166" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1"/><xf numFmtId="167" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1"/><xf numFmtId="166" fontId="1" fillId="2" borderId="1" xfId="0" applyNumberFormat="1"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'''


def write_workbook(path, sheets, drawings=None):
    path = Path(path)
    drawings = drawings or [None] * len(sheets)
    drawing_map = {}
    for sheet_index, chart in enumerate(drawings, 1):
        if chart:
            drawing_map[sheet_index] = len(drawing_map) + 1
    sheet_overrides = "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1, len(sheets)+1))
    chart_overrides = "".join(f'<Override PartName="/xl/charts/chart{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' for i in drawing_map.values())
    drawing_overrides = "".join(f'<Override PartName="/xl/drawings/drawing{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' for i in drawing_map.values())
    content = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>{sheet_overrides}{drawing_overrides}{chart_overrides}</Types>'''
    wb_sheets = "".join(f'<sheet name="{esc(name)}" sheetId="{i}" r:id="rId{i}"/>' for i, (name, _) in enumerate(sheets, 1))
    workbook = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>{wb_sheets}</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>'''
    wb_rels = "".join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1, len(sheets)+1))
    wb_rels += f'<Relationship Id="rId{len(sheets)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'''
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        z.writestr("[Content_Types].xml", content)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{wb_rels}</Relationships>')
        z.writestr("xl/styles.xml", styles_xml())
        for i, (_, xml) in enumerate(sheets, 1):
            if i in drawing_map:
                drawing_index = drawing_map[i]
                xml = xml.replace("</worksheet>", '<drawing r:id="rId1"/></worksheet>')
                z.writestr(f"xl/worksheets/_rels/sheet{i}.xml.rels", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing{drawing_index}.xml"/></Relationships>')
            z.writestr(f"xl/worksheets/sheet{i}.xml", xml)
        for sheet_index, drawing_index in drawing_map.items():
            drawing_item = drawings[sheet_index-1]
            chart_data = drawing_item["chart"] if isinstance(drawing_item, dict) else drawing_item
            start_col = drawing_item.get("col", 0) if isinstance(drawing_item, dict) else 8
            start_row = drawing_item.get("row", 8) if isinstance(drawing_item, dict) else 8
            end_col = start_col + (drawing_item.get("width", 14) if isinstance(drawing_item, dict) else 13)
            end_row = start_row + (drawing_item.get("height", 22) if isinstance(drawing_item, dict) else 22)
            drawing = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor><xdr:from><xdr:col>{start_col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{start_row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>{end_col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{end_row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Диаграмма 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>'''
            rels = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart{drawing_index}.xml"/></Relationships>'
            z.writestr(f"xl/drawings/drawing{drawing_index}.xml", drawing)
            z.writestr(f"xl/drawings/_rels/drawing{drawing_index}.xml.rels", rels)
            z.writestr(f"xl/charts/chart{drawing_index}.xml", chart_data)
    with zipfile.ZipFile(path) as z:
        if z.testzip():
            raise RuntimeError("Создан повреждённый Excel-файл")


def _add_review_styles(styles_text, base_styles):
    """Create Calibri 9 normal/red/yellow variants of the template styles."""
    style_root = ET.fromstring(styles_text)
    fonts = style_root.find(f"{{{MAIN_NS}}}fonts")
    calibri9 = None
    for idx, font in enumerate(list(fonts or [])):
        name = font.find(f"{{{MAIN_NS}}}name")
        size = font.find(f"{{{MAIN_NS}}}sz")
        if name is not None and size is not None and name.get("val", "").casefold() == "calibri" and size.get("val") == "9":
            calibri9 = idx
            break
    if calibri9 is None:
        font = ET.Element(f"{{{MAIN_NS}}}font")
        ET.SubElement(font, f"{{{MAIN_NS}}}sz", {"val": "9"})
        ET.SubElement(font, f"{{{MAIN_NS}}}name", {"val": "Calibri"})
        calibri9 = len(list(fonts))
        fonts.append(font)
        fonts.set("count", str(calibri9 + 1))

    fills = style_root.find(f"{{{MAIN_NS}}}fills")
    fill_id = len(list(fills))
    for color in ("FFFFC7CE", "FFFFEB9C"):
        fill = ET.SubElement(fills, f"{{{MAIN_NS}}}fill")
        pattern = ET.SubElement(fill, f"{{{MAIN_NS}}}patternFill", {"patternType": "solid"})
        ET.SubElement(pattern, f"{{{MAIN_NS}}}fgColor", {"rgb": color})
        ET.SubElement(pattern, f"{{{MAIN_NS}}}bgColor", {"indexed": "64"})
    fills.set("count", str(len(list(fills))))

    num_fmts = style_root.find(f"{{{MAIN_NS}}}numFmts")
    if num_fmts is None:
        num_fmts = ET.Element(f"{{{MAIN_NS}}}numFmts", {"count": "0"})
        fonts_index = list(style_root).index(fonts)
        style_root.insert(fonts_index, num_fmts)
    format_ids = {}
    used_ids = set()
    for item in list(num_fmts):
        used_ids.add(int(item.get("numFmtId", "0")))
        format_ids[item.get("formatCode")] = int(item.get("numFmtId", "0"))
    for code in ("dd.mm.yyyy", "h:mm:ss"):
        if code not in format_ids:
            new_id = max(used_ids | {163}) + 1
            used_ids.add(new_id)
            ET.SubElement(num_fmts, f"{{{MAIN_NS}}}numFmt", {"numFmtId": str(new_id), "formatCode": code})
            format_ids[code] = new_id
    num_fmts.set("count", str(len(list(num_fmts))))

    xfs = style_root.find(f"{{{MAIN_NS}}}cellXfs")
    original_xfs = list(xfs)
    normal_map, red_map, yellow_map = {}, {}, {}

    def clone(base_style, column, fill=None):
        source = original_xfs[base_style] if base_style < len(original_xfs) else original_xfs[0]
        copied = ET.fromstring(ET.tostring(source, encoding="utf-8"))
        copied.set("fontId", str(calibri9))
        copied.set("applyFont", "1")
        if column == "A":
            copied.set("numFmtId", str(format_ids["dd.mm.yyyy"]))
            copied.set("applyNumberFormat", "1")
        elif column == "B":
            copied.set("numFmtId", str(format_ids["h:mm:ss"]))
            copied.set("applyNumberFormat", "1")
        if fill is not None:
            copied.set("fillId", str(fill))
            copied.set("applyFill", "1")
        new_style = len(list(xfs))
        xfs.append(copied)
        return new_style

    for column, base_style in base_styles.items():
        normal_map[column] = clone(base_style, column)
        red_map[column] = clone(base_style, column, fill_id)
        yellow_map[column] = clone(base_style, column, fill_id + 1)
    xfs.set("count", str(len(list(xfs))))
    ET.register_namespace("", MAIN_NS)
    return ET.tostring(style_root, encoding="unicode", xml_declaration=True), normal_map, red_map, yellow_map


def write_base_from_template(template_path, output_path, records, last_history_row, history_classifier_updates=None):
    template_path, output_path = Path(template_path), Path(output_path)
    ns = {"m":MAIN_NS,"r":REL_NS,"p":PKG_REL_NS}
    with zipfile.ZipFile(template_path) as zin:
        workbook = ET.fromstring(zin.read("xl/workbook.xml"))
        rels_root = ET.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
        rels = {x.get("Id"):x.get("Target") for x in rels_root.findall("p:Relationship",ns)}
        source_sheet = next(x for x in workbook.find("m:sheets",ns) if x.get("name")=="Исходник")
        target = rels[source_sheet.get(f"{{{REL_NS}}}id")].replace("\\","/").lstrip("/")
        sheet_path = target if target.startswith("xl/") else "xl/"+target
        root = ET.fromstring(zin.read(sheet_path))
        data = root.find("m:sheetData",ns)
        rows = list(data.findall("m:row",ns))
        style_map = {c:0 for c in COLS}
        style_found = set()
        history_classifier_updates = history_classifier_updates or {}
        for row in rows:
            number=int(row.get("r"))
            for c in list(row.findall("m:c",ns)):
                col=col_of(c.get("r",""))
                col_num=sum((ord(ch)-64)*26**i for i,ch in enumerate(reversed(col)))
                if col_num>11: row.remove(c)
                elif number>1 and c.get("s") is not None and col not in style_found:
                    style_map[col]=int(c.get("s")); style_found.add(col)
                if col == "K" and number in history_classifier_updates:
                    for child in list(c):
                        c.remove(child)
                    c.set("t", "inlineStr")
                    is_el=ET.SubElement(c,f"{{{MAIN_NS}}}is")
                    t=ET.SubElement(is_el,f"{{{MAIN_NS}}}t")
                    t.text=str(history_classifier_updates[number])
            if number>last_history_row: data.remove(row)
        cols_node=root.find("m:cols",ns)
        if cols_node is not None:
            for col in list(cols_node):
                minimum=int(col.get("min","1")); maximum=int(col.get("max",str(minimum)))
                if minimum>11: cols_node.remove(col)
                elif maximum>11: col.set("max","11")
        styles_text=zin.read("xl/styles.xml").decode("utf-8")
        styles_text,normal_styles,red_styles,yellow_styles=_add_review_styles(styles_text,style_map)
        def add_cell(row_el,ref,value,style):
            c=ET.SubElement(row_el,f"{{{MAIN_NS}}}c",{"r":ref,"s":str(style)})
            text="" if value in (None, "") else str(value)
            if not text: return
            if re.fullmatch(r"-?\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?",text):
                ET.SubElement(c,f"{{{MAIN_NS}}}v").text=text
            else:
                c.set("t","inlineStr");is_el=ET.SubElement(c,f"{{{MAIN_NS}}}is");t=ET.SubElement(is_el,f"{{{MAIN_NS}}}t");t.text=text
        for offset,item in enumerate(records,1):
            number=last_history_row+offset
            row_el=ET.SubElement(data,f"{{{MAIN_NS}}}row",{"r":str(number)})
            for idx,(col,value) in enumerate(zip(COLS,item["values"])):
                style=red_styles[col] if item["styles"][idx]==4 else yellow_styles[col] if item["styles"][idx]==5 else normal_styles[col]
                add_cell(row_el,f"{col}{number}",value,style)
        final_row=last_history_row+len(records)
        dim=root.find("m:dimension",ns)
        if dim is not None: dim.set("ref",f"A1:K{final_row}")
        af=root.find("m:autoFilter",ns)
        if af is not None: af.set("ref",f"A1:K{final_row}")
        sheet_bytes=ET.tostring(root,encoding="utf-8",xml_declaration=True)
        output_path.parent.mkdir(parents=True,exist_ok=True)
        with zipfile.ZipFile(output_path,"w",zipfile.ZIP_DEFLATED,compresslevel=6) as zout:
            for item in zin.infolist():
                if item.filename==sheet_path: zout.writestr(item,sheet_bytes)
                elif item.filename=="xl/styles.xml": zout.writestr(item,styles_text.encode("utf-8"))
                else: zout.writestr(item,zin.read(item.filename))
    with zipfile.ZipFile(output_path) as z:
        if z.testzip(): raise RuntimeError("Создан повреждённый Excel-файл")


def tokens(text):
    return re.findall(r"[а-яёa-z]{3,}", str(text).casefold())


def has_review_text(text):
    """A dash or punctuation-only value is a placeholder, not review text."""
    return bool(re.search(r"[а-яёa-z0-9]", str(text or ""), flags=re.I))


class TextClassifier:
    def __init__(self):
        self.docs = defaultdict(Counter)
        self.totals = Counter()
        self.labels = Counter()

    def add(self, text, label):
        if not label:
            return
        self.labels[label] += 1
        counts = Counter(tokens(text))
        self.docs[label].update(counts)
        self.totals[label] += sum(counts.values())

    def predict(self, text, exclude=None):
        words = tokens(text)
        if not words or not self.labels:
            return ""
        excluded = {norm(x) for x in (exclude or [])}
        vocab = max(1000, len({w for c in self.docs.values() for w in c}))
        candidates = [(label, count) for label, count in self.labels.items() if norm(label) not in excluded]
        if not candidates:
            return ""
        total_docs = sum(count for _, count in candidates)
        best = None
        for label, count in candidates:
            score = math.log(count / total_docs)
            denom = self.totals[label] + vocab
            for word in words:
                score += math.log((self.docs[label][word] + 1) / denom)
            if best is None or score > best[0]:
                best = (score, label)
        return best[1]


def parse_lists(staff_book):
    sheet, header_row, mapping = find_sheet(staff_book, ["Продукт", "Процесс", "Классификатор"])
    cols = {key: mapping.get(norm(key)) for key in ("Продукт", "Процесс", "Классификатор")}
    result = {k: [] for k in cols}
    for row in sheet["rows"]:
        n, c = row["row"], row["cells"]
        if n <= header_row:
            continue
        for key, col in cols.items():
            value = cell(c, col, n) if col else ""
            if value and value not in result[key]:
                result[key].append(value)
    return result


def split_tag(tag, lists):
    raw = str(tag or "").strip()
    found = {k: [] for k in lists}
    for key, values in lists.items():
        for value in sorted(values, key=len, reverse=True):
            if norm(value) == "благодарность":
                continue
            pattern = rf"(?:^|,\s*){re.escape(str(value))}(?=,\s*|,?$)"
            if re.search(pattern, raw, flags=re.I):
                found[key].append(value)
    for ambiguous in ("Другое", "HR"):
        if ambiguous in found["Процесс"] and ambiguous in found["Классификатор"]:
            if len(found["Классификатор"]) > 1:
                found["Процесс"].remove(ambiguous)
            elif len(found["Процесс"]) > 1:
                found["Классификатор"].remove(ambiguous)
    found["Классификатор"] = list(dict.fromkeys(canonical_classifier(value) for value in found["Классификатор"]))
    return found


def canonical_value(value, allowed):
    key = norm(value)
    return next((item for item in allowed if norm(item) == key), "")


def dimension_values(value, allowed):
    """Return recognized values from a standalone Product/Process/Classifier cell."""
    raw = str(value or "").strip()
    if not raw:
        return []
    exact = canonical_value(raw, allowed)
    if exact:
        return [exact]
    found = []
    for item in sorted(allowed, key=lambda x: len(str(x)), reverse=True):
        pattern = rf"(?:^|\s*[|;]\s*|,\s*){re.escape(str(item))}(?=\s*[|;]\s*|,\s*|$)"
        if re.search(pattern, raw, flags=re.I):
            found.append(item)
    return found or [raw]


def mentioned_products(text, products):
    source = str(text or "").casefold()
    found = []
    for product in products:
        candidate = str(product).strip()
        if not candidate or norm(candidate) == norm("Нет продукта"):
            continue
        if re.search(rf"(?<![а-яёa-z0-9]){re.escape(candidate.casefold())}(?![а-яёa-z0-9])", source):
            found.append(product)
    return found


def infer_process(text, processes, model):
    """Apply explicit business rules first, then use the trained suggestion model."""
    source = str(text or "").casefold()

    def choose(name):
        return canonical_value(name, processes)

    # A claim event always has priority over interface/service wording.
    if re.search(r"\bдтп\b|авари|страхов(?:ой|ого|ому) случа|урегулирован|ремонт|выплат", source):
        value = choose("Урегулирование")
        if value:
            return value

    # When a review lists several different policy/app capabilities, its subject
    # is general support/functionality rather than one transaction.
    action_groups = (
        r"просмотр\w*\s+полис|посмотреть\s+полис",
        r"внест\w*\s+изменен|изменить\s+полис",
        r"продл\w*|пролонгац",
        r"оформ\w*\s+(?:нов\w*|полис)",
    )
    action_count = sum(bool(re.search(pattern, source)) for pattern in action_groups)
    if action_count >= 2:
        value = choose("Сопровождение/консультация")
        if value:
            return value

    if re.search(r"оплат\w*\s+полис|оплат[аы]\s+полис|\bкупить\b|приобрест|оформ\w*\s+(?:нов\w*\s+)?полис", source):
        value = choose("Покупка")
        if value:
            return value

    if re.search(r"продл\w*|пролонгац", source):
        value = choose("Пролонгация")
        if value:
            return value

    if re.search(r"приложен|личн\w*\s+кабинет|\bлк\b|интерфейс|функционал|сервис|менеджер|поддержк", source):
        value = choose("Сопровождение/консультация")
        if value:
            return value

    return model.predict(text, {"Другое"})


def build_base(partner_path, staff_path, master_path, output_path, month):
    partner, staff, master = inspect(partner_path), inspect(staff_path), inspect(master_path)
    staff_sheet, staff_header, sm = find_sheet(staff, HEADERS)
    partner_sheet, partner_header, pm = find_sheet(partner, ["Геосервис", "Дата", "Оценка", "Автор", "Отзыв", "Тег", "Удален"])
    master_sheet, master_header, mm = find_sheet(master, HEADERS)
    lists = parse_lists(staff)
    records = []
    history_classifier_updates = {}
    last_history_row = master_header
    training = [TextClassifier(), TextClassifier(), TextClassifier()]
    dimensions = ("Продукт", "Процесс", "Классификатор")
    defaults = ("Нет продукта", "Другое", "Другое")
    for row in master_sheet["rows"]:
        n, c = row["row"], row["cells"]
        if n <= master_header:
            continue
        vals = [cell(c, mm[norm(h)], n) for h in HEADERS]
        if any(v not in (None, "") for v in vals):
            original_classifier = vals[10]
            vals[10] = canonical_classifier(vals[10])
            if vals[10] != original_classifier:
                history_classifier_updates[n] = vals[10]
            last_history_row = max(last_history_row, n)
            records.append({"values": vals, "styles": [2,3,0,0,0,0,0,0,0,0,0], "source":"История", "row":n})
            for model, dimension, label in zip(training, dimensions, vals[8:11]):
                canonical = canonical_value(label, lists[dimension])
                if canonical:
                    model.add(vals[2], canonical)
    history_count = len(records)
    staff_added = 0
    yellow_count = 0
    for row in staff_sheet["rows"]:
        n, c = row["row"], row["cells"]
        if n <= staff_header:
            continue
        vals = [cell(c, sm[norm(h)], n) for h in HEADERS]
        if not any(v not in (None, "") for v in vals) or not in_month(vals[0], month):
            continue
        styles = [2,3,0,0,0,0,0,0,0,0,0]
        if vals[1] in (None, ""):
            vals[1] = 0
        text_present = has_review_text(vals[2])
        for model_idx, (idx, dimension, default) in enumerate(zip((8, 9, 10), dimensions, defaults)):
            original = vals[idx]
            found = dimension_values(original, lists[dimension])
            if not found:
                if dimension == "Продукт":
                    mentions = mentioned_products(vals[2], lists[dimension]) if text_present else []
                    vals[idx] = mentions[0] if mentions else "Нет продукта"
                elif dimension == "Процесс":
                    vals[idx] = infer_process(vals[2], lists[dimension], training[model_idx]) if text_present else default
                else:
                    vals[idx] = training[model_idx].predict(vals[2], {default}) if text_present else default
                styles[idx] = 5
                yellow_count += 1
            elif dimension == "Продукт" and len(found) > 1:
                mentions = mentioned_products(vals[2], lists[dimension]) if text_present else []
                vals[idx] = mentions[0] if mentions else default
                styles[idx] = 5
                yellow_count += 1
            elif dimension != "Продукт" and len(found) > 1:
                vals[idx] = " | ".join(found)
                styles[idx] = 4
            else:
                vals[idx] = found[0]
        vals[10] = canonical_classifier(vals[10])
        records.append({"values": vals, "styles":styles, "source":"Сотрудники", "row":n})
        staff_added += 1
        for model, dimension, label in zip(training, dimensions, vals[8:11]):
            canonical = canonical_value(label, lists[dimension])
            if canonical:
                model.add(vals[2], canonical)
    checks = []
    red_count = 0
    partner_added = deleted = 0
    pcol = {h: pm[norm(h)] for h in ("Геосервис","Дата","Оценка","Автор","Отзыв","Тег","Удален")}
    link_col = pm.get(norm("URL"), pm.get(norm("Ссылка"), "W"))
    for row in partner_sheet["rows"]:
        n, c = row["row"], row["cells"]
        if n <= partner_header:
            continue
        if str(cell(c,pcol["Удален"],n)).strip() in {"1","1.0"}:
            deleted += 1; continue
        date = cell(c,pcol["Дата"],n)
        if not in_month(date, month):
            continue
        score = str(cell(c,pcol["Оценка"],n)).strip().replace(".0","")
        tone = "Негативная" if score in {"1","2"} else "Нейтральная" if score == "3" else "Позитивная" if score in {"4","5"} else ""
        text = cell(c,pcol["Отзыв"],n)
        tag = cell(c,pcol["Тег"],n)
        styles = [2,3,0,0,0,0,0,0,0,0,0]
        parsed = split_tag(tag, lists) if str(tag).strip() else {key: [] for key in dimensions}
        text_present = has_review_text(text)
        resolved = []
        problems=[]
        for model_idx, (idx, dimension, default) in enumerate(zip((8, 9, 10), dimensions, defaults)):
            found = parsed[dimension]
            if not found:
                if dimension == "Продукт":
                    mentions = mentioned_products(text, lists[dimension]) if text_present else []
                    value = mentions[0] if mentions else "Нет продукта"
                elif dimension == "Процесс":
                    value = infer_process(text, lists[dimension], training[model_idx]) if text_present else default
                else:
                    value = training[model_idx].predict(text, {default}) if text_present else default
                styles[idx] = 5
                yellow_count += 1
            elif dimension == "Продукт" and len(found) > 1:
                mentions = mentioned_products(text, lists[dimension]) if text_present else []
                value = mentions[0] if mentions else default
                styles[idx] = 5
                yellow_count += 1
            elif dimension != "Продукт" and len(found) > 1:
                value = " | ".join(found)
                styles[idx] = 4
                red_count += 1
                problems.append(f"{dimension.casefold()}: найдено {len(found)}")
            else:
                value = found[0]
            resolved.append(value)
        product, process, classifier = resolved
        classifier = canonical_classifier(classifier)
        reason="; ".join(problems); check_type="Противоречие" if problems else ""
        source_raw=cell(c,pcol["Геосервис"],n)
        vals=[date,0,text,SOURCE_MAP.get(source_raw,source_raw),cell(c,link_col,n),cell(c,pcol["Автор"],n),"Первичный",tone,product,process,classifier]
        item={"values":vals,"styles":styles,"source":"Партнёр","row":n,"tag":tag,"reason":reason,"type":check_type}
        records.append(item); partner_added += 1
        if reason: checks.append(item)
    write_base_from_template(master_path, output_path, records[history_count:], last_history_row, history_classifier_updates)
    red_count += sum(1 for item in records[history_count:history_count+staff_added] for style in item["styles"] if style == 4)
    return {"history":history_count,"staff":staff_added,"partner":partner_added,"deleted":deleted,"checks":len(checks),"red":red_count,"yellow":yellow_count}


def chart_xml(title, sheet, cat_range, val_ranges, series_names, chart_type="bar", category_values=None):
    series=[]
    colors=["38056C","9A92FF","79FF3A"]
    category_values = category_values or []
    category_cache = ""
    if category_values:
        points = "".join(f'<c:pt idx="{idx}"><c:v>{esc(value)}</c:v></c:pt>' for idx, value in enumerate(category_values))
        category_cache = f'<c:strCache><c:ptCount val="{len(category_values)}"/>{points}</c:strCache>'
    for i,(vr,name) in enumerate(zip(val_ranges,series_names)):
        series.append(f'''<c:ser><c:idx val="{i}"/><c:order val="{i}"/><c:tx><c:v>{esc(name)}</c:v></c:tx><c:spPr><a:solidFill><a:srgbClr val="{colors[i%len(colors)]}"/></a:solidFill></c:spPr><c:dLbls><c:numFmt formatCode="0.0%" sourceLinked="0"/><c:dLblPos val="outEnd"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showLeaderLines val="0"/></c:dLbls><c:cat><c:strRef><c:f>'{sheet}'!{cat_range}</c:f>{category_cache}</c:strRef></c:cat><c:val><c:numRef><c:f>'{sheet}'!{vr}</c:f></c:numRef></c:val></c:ser>''')
    grouping="clustered"
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ru-RU" typeface="Arial" sz="900"/><a:t>{esc(title)}</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:layout/><c:barChart><c:barDir val="{'bar' if chart_type=='bar' else 'col'}"/><c:grouping val="{grouping}"/>{''.join(series)}<c:axId val="123456"/><c:axId val="654321"/></c:barChart><c:catAx><c:axId val="123456"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="nextTo"/><c:crossAx val="654321"/><c:crosses val="autoZero"/></c:catAx><c:valAx><c:axId val="654321"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="0%" sourceLinked="0"/><c:crossAx val="123456"/><c:crosses val="autoZero"/></c:valAx></c:plotArea><c:legend><c:legendPos val="b"/></c:legend><c:plotVisOnly val="1"/></c:chart><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr typeface="Arial" sz="900"/></a:pPr><a:endParaRPr lang="ru-RU"/></a:p></c:txPr></c:chartSpace>'''


def analytics_drawings(charts):
    anchors=[]
    for i in range(len(charts)):
        col=(i%2)*9; row=(i//2)*18+7
        anchors.append(f'''<xdr:twoCellAnchor><xdr:from><xdr:col>{col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>{col+8}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row+16}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="{i+2}" name="Диаграмма {i+1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId{i+1}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>''')
    drawing=f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">{''.join(anchors)}</xdr:wsDr>'''
    rels=''.join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart{i}.xml"/>' for i in range(1,len(charts)+1))
    return {"drawing":drawing,"rels":f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>',"charts":charts}


def build_analytics(base_path, output_path, month):
    book = inspect(base_path)
    sheet, header, mapping = find_sheet(book, HEADERS)
    rows = []
    for row in sheet["rows"]:
        n, c = row["row"], row["cells"]
        if n <= header:
            continue
        vals = {h: cell(c, mapping[norm(h)], n) for h in HEADERS}
        vals["Классификатор"] = canonical_classifier(vals["Классификатор"])
        dt = excel_date(vals["Дата"])
        if dt:
            vals["_date"] = dt
            rows.append(vals)

    report_date = datetime.strptime(month, "%Y-%m") if month else max((r["_date"] for r in rows), default=datetime.now()).replace(day=1)
    start_date = datetime(2025, 1, 1)
    month_keys = []
    cursor = start_date
    while cursor <= report_date:
        month_keys.append(cursor.strftime("%Y-%m"))
        cursor = datetime(cursor.year + (cursor.month == 12), 1 if cursor.month == 12 else cursor.month + 1, 1)

    rows = [r for r in rows if start_date <= r["_date"] < cursor]
    selected = [r for r in rows if r["_date"].strftime("%Y-%m") == report_date.strftime("%Y-%m")]
    tone = Counter(r["Тональность"] for r in selected if r["Тональность"])
    active_months = {r["_date"].strftime("%Y-%m") for r in rows}
    display_months = [key for key in month_keys if key in active_months]
    last_three = month_keys[-3:]

    def excel_serial(dt):
        return (dt - datetime(1899, 12, 30)).days

    def make_sheet(wanted_tone, sheet_name, title):
        counts = defaultdict(Counter)
        totals = Counter()
        for r in rows:
            if r["Тональность"] != wanted_tone:
                continue
            key = r["_date"].strftime("%Y-%m")
            label = str(r["Классификатор"] or "Другое").strip()
            counts[key][label] += 1
            totals[key] += 1

        categories = sorted({label for counter in counts.values() for label in counter}, key=norm)
        total_all = sum(totals.values())

        report_key = report_date.strftime("%Y-%m")
        top = sorted(categories, key=lambda label: (counts[report_key][label], label), reverse=True)[:5]
        while len(top) < 5:
            top.append("")

        table = []
        merges = []
        total_columns = len(display_months) + 2

        def column_name(number):
            result = ""
            while number:
                number, remainder = divmod(number - 1, 26)
                result = chr(65 + remainder) + result
            return result

        def append_matrix(section_title, percent=False):
            title_row = len(table) + 1
            values = [section_title] + [""] * (total_columns - 1)
            table.append((values, [1] * total_columns))
            merges.append(f"A{title_row}:{column_name(total_columns)}{title_row}")

            year_row = len(table) + 1
            years = [datetime.strptime(key, "%Y-%m").year for key in display_months]
            year_values = ["Классификатор"] + [""] * len(display_months) + ["Общий итог"]
            for index, year in enumerate(years, 1):
                if index == 1 or years[index-2] != year:
                    year_values[index] = str(year)
            table.append((year_values, [1] * total_columns))

            month_row = len(table) + 1
            month_values = [""] + [MONTHS_RU[datetime.strptime(key, "%Y-%m").month - 1].casefold() for key in display_months] + [""]
            table.append((month_values, [1] * total_columns))
            merges.append(f"A{year_row}:A{month_row}")
            merges.append(f"{column_name(total_columns)}{year_row}:{column_name(total_columns)}{month_row}")
            for year in sorted(set(years)):
                positions = [i + 2 for i, value in enumerate(years) if value == year]
                if positions:
                    merges.append(f"{column_name(min(positions))}{year_row}:{column_name(max(positions))}{year_row}")

            for label in categories:
                row_values = [label]
                for key in display_months:
                    amount = counts[key][label]
                    value = (amount / totals[key]) if percent and amount and totals[key] else amount
                    row_values.append(value if value else "")
                overall = (sum(counts[key][label] for key in display_months) / total_all) if percent and total_all else sum(counts[key][label] for key in display_months)
                row_values.append(overall if overall else "")
                table.append((row_values, [0] + ([8] if percent else [0]) * (total_columns - 1)))

            total_values = ["Общий итог"]
            for key in display_months:
                total_values.append(1 if percent and totals[key] else totals[key])
            total_values.append(1 if percent and total_all else total_all)
            table.append((total_values, [1] + ([10] if percent else [1]) * (total_columns - 1)))

        append_matrix("Количество по полю Классификатор", percent=False)
        table.extend([([""], [7]), ([""], [7])])
        append_matrix("Доля классификатора от общего количества за месяц", percent=True)
        table.extend([([""], [7]), ([""], [7])])

        helper_header_row = len(table) + 1
        helper_header = ["Классификатор"] + [datetime.strptime(key, "%Y-%m").strftime("%m.%Y") for key in last_three]
        table.append((helper_header, [1, 1, 1, 1]))
        for label in top:
            values = [label]
            for key in last_three:
                values.append(counts[key][label] / totals[key] if label and totals[key] else 0)
            table.append((values, [0, 8, 8, 8]))

        month_names = " — ".join(datetime.strptime(key, "%Y-%m").strftime("%m.%Y") for key in last_three)
        first_data_row = helper_header_row + 1
        last_data_row = helper_header_row + 5
        chart = chart_xml(f"{title}: {month_names}", sheet_name, f"$A${first_data_row}:$A${last_data_row}", [f"$B${first_data_row}:$B${last_data_row}", f"$C${first_data_row}:$C${last_data_row}", f"$D${first_data_row}:$D${last_data_row}"], [datetime.strptime(key, "%Y-%m").strftime("%m.%Y") for key in last_three], "col", top)
        widths = [42] + [9] * len(display_months) + [12]
        xml = worksheet_xml(table, widths, autofilter=False)
        if merges:
            merge_xml = f'<mergeCells count="{len(merges)}">' + "".join(f'<mergeCell ref="{ref}"/>' for ref in merges) + '</mergeCells>'
            xml = xml.replace("</worksheet>", merge_xml + "</worksheet>")
        chart_anchor_row = helper_header_row + 7
        return xml, {"chart": chart, "col": 0, "row": chart_anchor_row, "width": 15, "height": 24}, len(categories)

    negative_xml, negative_chart, negative_rows = make_sheet("Негативная", "Классификатор Негатив", "ТОП-5 упоминаний по причинам негатива")
    positive_xml, positive_chart, positive_rows = make_sheet("Позитивная", "Классификатор Позитив", "ТОП-5 упоминаний по причинам позитива")
    sheets = [("Классификатор Негатив", negative_xml), ("Классификатор Позитив", positive_xml)]
    write_workbook(output_path, sheets, [negative_chart, positive_chart])
    return {"rows": len(rows), "selected": len(selected), "negative": tone.get("Негативная", 0), "neutral": tone.get("Нейтральная", 0), "positive": tone.get("Позитивная", 0), "negative_table_rows": negative_rows, "positive_table_rows": positive_rows}
