"""Edits the inspected original using python-docx's XML tree, preserving ZIP parts."""
from copy import deepcopy
from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
import re
import sys
import zipfile

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

from .model import ClientData, ValidationError, atomic_write

ROOT = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
TEMPLATE = ROOT / "templates" / "Hindi Ah Boi  4 Old Vrndavan Mandkheda 22.09.2026.docx"
TEMPLATE_SHA256 = "825a595e9a5004f662ea2c1d19944a1780b005756e63ab9b480ae597dd9f4820"
OLD_NAME = "o`Unkou"
OLD_FATHER = "txUukFk"
OLD_VILLAGE = "eaM[ksMk"


class TemplateError(ValueError):
    pass


def text(element):
    return "".join(t.text or "" for t in element.iter(qn("w:t")))


def set_text(node, value):
    node.text = value
    node.set(qn("xml:space"), "preserve")


def set_font(run, font):
    props = run.find(qn("w:rPr"))
    if props is None:
        props = OxmlElement("w:rPr")
        run.insert(0, props)
    fonts = props.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        props.insert(0, fonts)
    for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        fonts.attrib.pop(qn("w:" + attr), None)
    for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(qn("w:" + attr), font)


def replace_span(paragraph, start, end, value, font=None):
    """Replace an offset range even across runs; never flatten a paragraph.

    Legacy/numeric values reuse source text nodes. Unicode values receive an
    isolated run so the source prefix/suffix never inherit a Unicode font.
    """
    nodes = list(paragraph.iter(qn("w:t")))
    positions, offset = [], 0
    for node in nodes:
        positions.append((node, offset, offset + len(node.text or "")))
        offset += len(node.text or "")
    affected = [(n, a, b) for n, a, b in positions if a < end and b > start]
    if not affected or not 0 <= start < end <= offset:
        raise TemplateError("Invalid field span in template.")
    first, a, _ = affected[0]
    last, last_a, _ = affected[-1]
    prefix = (first.text or "")[:start-a]
    suffix = (last.text or "")[end-last_a:]
    if font is None:
        if first is last:
            set_text(first, prefix + value + suffix)
        else:
            set_text(first, prefix + value)
            for node, _, _ in affected[1:-1]:
                set_text(node, "")
            set_text(last, suffix)
        return
    run = first.getparent()
    if run.tag != qn("w:r") or run.getparent() is not paragraph or len(run.findall(qn("w:t"))) != 1:
        raise TemplateError("Unicode replacement encountered an unsupported rich-text field.")
    inserted = OxmlElement("w:r")
    props = run.find(qn("w:rPr"))
    if props is not None:
        inserted.append(deepcopy(props))
    set_font(inserted, font)
    node = OxmlElement("w:t")
    set_text(node, value)
    inserted.append(node)
    set_text(first, prefix)
    index = paragraph.index(run) + 1
    paragraph.insert(index, inserted)
    if first is last:
        if suffix:
            tail = OxmlElement("w:r")
            if props is not None:
                tail.append(deepcopy(props))
            tail_text = OxmlElement("w:t")
            set_text(tail_text, suffix)
            tail.append(tail_text)
            paragraph.insert(index + 1, tail)
    else:
        for node, _, _ in affected[1:-1]:
            set_text(node, "")
        set_text(last, suffix)


def cells(row):
    # Physical XML cells avoid python-docx's aliasing of vertical merged cells.
    return row.findall(qn("w:tc"))


def land_tables(root):
    for table in root.iter(qn("w:tbl")):
        rows = table.findall(qn("w:tr"))
        if not rows:
            continue
        headers = [text(c).strip() for c in cells(rows[0])]
        if len(headers) in (4, 5) and headers[0] == "dz-" and "[kljk" in headers[1] and headers[3] == "yxku":
            yield table


def fill_cell(cell, value, font=None):
    paragraphs = cell.findall(qn("w:p"))
    paragraph = next((p for p in paragraphs if text(p).strip()), paragraphs[0])
    nodes = list(paragraph.iter(qn("w:t")))
    # Select the actual value run, not the leading whitespace run in Bahi cells.
    target = next((n for n in nodes if (n.text or "").strip()), nodes[0] if nodes else None)
    if target is None:
        run = OxmlElement("w:r")
        props = paragraph.find("./" + qn("w:pPr") + "/" + qn("w:rPr"))
        if props is not None:
            run.append(deepcopy(props))
        paragraph.append(run)
        target = OxmlElement("w:t")
        run.append(target)
    for p in paragraphs:
        for n in p.iter(qn("w:t")):
            set_text(n, "")
    set_text(target, value)
    if font:
        set_font(target.getparent(), font)


def update_land_tables(root, client):
    count = 0
    area, lagan = map(lambda n: format(n, "f"), client.totals)
    font = client.unicode_font if client.text_mode == "unicode" else None
    for table in list(land_tables(root)):
        rows = table.findall(qn("w:tr"))
        if len(rows) < 3:
            raise TemplateError("Land table has no body/total prototype.")
        prototype, total = deepcopy(rows[1]), deepcopy(rows[-1])
        width = len(cells(prototype))
        for row in rows[1:]:
            table.remove(row)
        for index, record in enumerate(client.records):
            row = deepcopy(prototype)
            values = [f"{index+1:02d}", record.khasra, record.area, record.lagan]
            if width == 5:
                values.append(client.bahi if index == 0 else "")
            for col, (cell, value) in enumerate(zip(cells(row), values)):
                fill_cell(cell, value, font if col in (1, 4) else None)
            if width == 5:
                merge = cells(row)[4].find("./" + qn("w:tcPr") + "/" + qn("w:vMerge"))
                if merge is not None:
                    merge.set(qn("w:val"), "restart" if index == 0 else "continue")
            table.append(row)
        total_cells = cells(total)
        for col, value in ((0, ""), (2, area), (3, lagan)):
            if col in (2, 3):
                # The five-column source has SUM(ABOVE) fields. Word/LO can
                # refresh them using float arithmetic or round to two places.
                # Materialize only these totals as exact Decimal result text,
                # retaining the original result runs and all their formatting.
                for field_node in list(total_cells[col].iter()):
                    if field_node.tag in (qn("w:fldChar"), qn("w:instrText")):
                        field_node.getparent().remove(field_node)
            fill_cell(total_cells[col], value)
        # The count uses Times New Roman while the Hindi suffix uses DevLys.
        # Replacing the entire cell would render the suffix as Latin "fdrk".
        count_paragraph = next(p for p in total_cells[1].findall(qn("w:p")) if text(p).strip())
        match = re.search(r"[0-9]+", text(count_paragraph))
        if match is None:
            raise TemplateError("Total row has no record count.")
        replace_span(count_paragraph, match.start(), match.end(), f"{len(client.records):02d}")
        if width == 5:
            fill_cell(total_cells[4], "")
        table.append(total)
        # Cloning must not duplicate Word's paragraph identity metadata.
        for row in table.findall(qn("w:tr"))[1:]:
            for element in row.iter():
                for key in list(element.attrib):
                    if key.rsplit("}", 1)[-1] in ("paraId", "textId"):
                        del element.attrib[key]
        count += 1
    return count


def paragraph_edits(paragraph, client):
    """Offsets are computed from original text, so new values are never rescanned."""
    source = text(paragraph)
    edits = []
    font = client.unicode_font if client.text_mode == "unicode" else None

    def token(old, value, chosen_font=font):
        for match in re.finditer(re.escape(old), source):
            edits.append((match.start(), match.end(), value, chosen_font))

    token(OLD_NAME, client.client_name)
    token(OLD_FATHER, client.father_name)
    token("czkEg.k", client.caste)
    # Both inspected affidavit age slots are spaces between these legacy labels.
    # Keep the labels' DevLys runs. ASCII age digits use the source font in
    # legacy mode; Unicode mode isolates the age just like other new values.
    for match in re.finditer(r"(?<=vk;q)( +)(?=tkfr)", source):
        edits.append((match.start(), match.end(), " " + client.age + " ", font))
    # Only the residence occurrence becomes address; other occurrences are village.
    address_spans = {m.span(1) for m in re.finditer(r"fuoklh(?: &)? +(?:xzke +)?(eaM\[ksMk)", source)}
    for match in re.finditer(re.escape(OLD_VILLAGE), source):
        edits.append((*match.span(), client.address if match.span() in address_spans else client.village, font))
    # Client geography is the first tehsil and district following residence/land village.
    # This avoids bank district before the name and registrar/advocate districts later.
    if OLD_VILLAGE in source:
        after = source.index(OLD_VILLAGE) + len(OLD_VILLAGE)
        tehsil = re.search(r"[’']?;keiqj", source[after:])
        if tehsil:
            edits.append((after+tehsil.start(), after+tehsil.end(), client.tehsil, font))
        district = re.search(r"ftyk +(lhgksj)", source[after:])
        if district:
            edits.append((after+district.start(1), after+district.end(1), client.district, font))
    token("LD-445692", client.bahi)
    if "1.229" in source and ("gSDV" in source or "gSDVj" in source):
        token("1.229", format(client.totals[0], "f"), None)
    date_value = client.document_date[8:10] + "." + client.document_date[5:7] + "." + client.document_date[:4]
    date_patterns = [
        r"(?<=vkt fnukad)( +)(?=dks)",
        r"(?<=fnukad%&)( +[-]+)",
        r"(?<=fnukad%&)( +)$",
        r"(?<=fnukad)( +)-09-2026(?= rd)",
        r"(?<=fnukad)( +[-]+)(?= *$)",
        r"(?<=rkjh\[k %&)( +)$",
    ]
    for pattern in date_patterns:
        for m in re.finditer(pattern, source):
            # Date separators must be literal punctuation, not DevLys glyph codes.
            edits.append((m.start(), m.end(), " " + date_value + " ", "Times New Roman"))
    edits.sort(reverse=True)
    previous = len(source) + 1
    for start, end, value, chosen_font in edits:
        if end > previous:
            raise TemplateError("Overlapping edit slots.")
        previous = start
        if source[start:end] != value or chosen_font:
            replace_span(paragraph, start, end, value, chosen_font)
    return len(edits)


@dataclass(frozen=True)
class GenerationResult:
    path: Path
    land_tables_updated: int
    field_spans_updated: int


def generate_docx(client: ClientData, destination, template=TEMPLATE):
    client.require_complete()
    template, destination = Path(template), Path(destination)
    if destination.suffix.lower() != ".docx":
        raise ValidationError("Output must have a .docx extension.")
    if destination.resolve() == template.resolve() or destination.resolve().is_relative_to((ROOT / "templates").resolve()):
        raise ValidationError("Save generated documents outside templates/.")
    if destination.exists():
        raise FileExistsError("Output already exists. Choose a new filename.")
    # Read once; source identity and the package used for output cannot diverge.
    from io import BytesIO
    original = template.read_bytes()
    if sha256(original).hexdigest() != TEMPLATE_SHA256:
        raise TemplateError("Template differs from the inspected original. Inspect it and update the engine profile before generating.")
    document = Document(BytesIO(original))
    root = document._element
    tables = list(land_tables(root))
    if len(tables) != 3:
        raise TemplateError("Expected all three inspected land tables, including the nested table.")
    protected = {p for table in tables for p in table.iter(qn("w:p"))}
    spans = sum(paragraph_edits(p, client) for p in list(root.iter(qn("w:p"))) if p not in protected and not any(a.tag == qn("w:txbxContent") for a in p.iterancestors()))
    # Schedule cells are separate paragraphs with no village context for tehsil/district.
    schedule = list(root.iter(qn("w:tbl")))[5]
    location = cells(schedule.findall(qn("w:tr"))[2])
    font = client.unicode_font if client.text_mode == "unicode" else None
    fill_cell(location[2], client.tehsil, font)
    fill_cell(location[3], client.district, font)
    count = update_land_tables(root, client)
    new_xml = etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)

    def write_package(path):
        with zipfile.ZipFile(BytesIO(original)) as source, zipfile.ZipFile(path, "w") as output:
            output.comment = source.comment
            for entry in source.infolist():
                output.writestr(entry, new_xml if entry.filename == "word/document.xml" else source.read(entry.filename))

    atomic_write(destination, write_package)
    return GenerationResult(destination, count, spans)
