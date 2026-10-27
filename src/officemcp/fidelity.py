"""The fidelity check: nothing the input has may disappear unless an action declared it would.

For Office files (OOXML packages) the inventory is: the number of parts of each content type,
the hashes of binary parts, and counts of features a library can drop. For PDF it is the page
count, form fields, outline entries, attachments, signatures and annotations. A batch may remove
only what its actions declared (`Edit.allow`).
"""

import hashlib
import zipfile
from collections import Counter
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"

# Elements whose loss is a loss of content or behaviour, by a readable name.
FEATURES = {
    f"{{{W}}}drawing": "Word drawings and images",
    f"{{{W}}}object": "Word embedded objects",
    f"{{{W}}}ins": "Word tracked insertions",
    f"{{{W}}}del": "Word tracked deletions",
    f"{{{W}}}sdt": "Word content controls",
    f"{{{W}}}fldSimple": "Word fields",
    f"{{{W}}}fldChar": "Word fields",
    f"{{{W}}}footnoteReference": "Word footnote references",
    f"{{{W}}}endnoteReference": "Word endnote references",
    f"{{{W}}}commentReference": "Word comment references",
    f"{{{W}}}txbxContent": "Word text boxes",
    f"{{{W}}}hyperlink": "Word hyperlinks",
    f"{{{W}}}bookmarkStart": "Word bookmarks",
    f"{{{M}}}oMath": "equations",
    f"{{{C}}}chart": "chart references",
    f"{{{S}}}dataValidation": "Excel data validations",
    f"{{{S}}}conditionalFormatting": "Excel conditional formats",
    f"{{{S}}}hyperlink": "Excel hyperlinks",
    f"{{{S}}}drawing": "Excel sheet drawings",
    f"{{{S}}}legacyDrawing": "Excel comments and form controls (VML)",
    f"{{{S}}}control": "Excel form controls",
    f"{{{S}}}oleObject": "Excel embedded objects",
    f"{{{S}}}tablePart": "Excel tables",
    f"{{{S}}}mergeCell": "Excel merged cells",
    f"{{{S}}}definedName": "Excel defined names",
    f"{{{X14}}}sparklineGroup": "Excel sparklines",
    f"{{{P}}}oleObj": "PowerPoint embedded objects",
    f"{{{P}}}sldId": "PowerPoint slides",
    f"{{{A}}}tbl": "tables in drawings",
    f"{{{A}}}graphicData": "graphic frames (charts, SmartArt, tables)",
    f"{{{MC}}}AlternateContent": "alternate content (newer Office features)",
}

# Parts any producer regenerates or Office rebuilds: their disappearance loses nothing.
REGENERATED = {
    "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml",
    "application/vnd.openxmlformats-officedocument.extended-properties+xml",
    # Text moves between the shared-strings part and the cells (inline strings); it stays in the cells.
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
}
X15AC = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"

# Office extensions (<ext uri="...">) are counted one by one. These record only how a file was
# made (calculation features, a theme's family name) and are ignored.
BOOKKEEPING_EXTENSIONS = {
    "{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}",  # calculation features
    "{05A4C25C-085E-4340-85A3-A5531E510DB2}",  # theme family
}
# Style extensions that only name a default (slicer, timeline): ignored unless they define styles.
DEFAULT_STYLE_EXTENSIONS = {"{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}", "{9260A510-F301-46a8-8635-F512D64BE5F5}"}
EXTENSIONS = {
    "{78C0D931-6437-407d-A8EE-F0AAD7539E65}": "Excel conditional formats (2010)",
    "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}": "Excel data validations (2010)",
    "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}": "Excel sparklines",
    "{A8765BA9-456A-4dab-B4F3-ACF838C121DE}": "Excel slicers",
    "{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}": "Excel table slicers",
    "{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}": "Excel dynamic arrays",
    "{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}": "Excel rich values",
    "{46F421CA-312F-682f-3DD2-61675219B42D}": "Excel 2010 styles (dxfs)",
}
THUMBNAIL = "docProps/thumbnail"


def _content_types(archive: zipfile.ZipFile) -> dict[str, str]:
    root = etree.fromstring(archive.read("[Content_Types].xml"))
    defaults = {e.get("Extension").lower(): e.get("ContentType") for e in root if e.tag.endswith("Default")}
    overrides = {e.get("PartName").lstrip("/"): e.get("ContentType") for e in root if e.tag.endswith("Override")}
    types = {}
    for name in archive.namelist():
        if name.endswith("/") or name == "[Content_Types].xml":
            continue
        types[name] = overrides.get(name) or defaults.get(name.rsplit(".", 1)[-1].lower(), "application/octet-stream")
    return types


def ooxml_inventory(path: Path) -> Counter:
    inventory = Counter()
    with zipfile.ZipFile(path) as archive:
        for name, kind in _content_types(archive).items():
            # Relationship parts come and go with the parts they belong to; those are counted.
            if kind in REGENERATED or name.startswith(THUMBNAIL) or name.endswith(".rels"):
                continue
            inventory[f"part:{kind}"] += 1
            data = archive.read(name)
            if name.endswith(".xml") or name.endswith(".rels") or kind.endswith("+xml") or kind.endswith("/xml"):
                try:
                    root = etree.fromstring(data, etree.XMLParser(huge_tree=True, resolve_entities=False))
                    for element in root.iter():
                        label = FEATURES.get(element.tag) or _extension(element)
                        if label and not _bookkeeping(element):
                            inventory[f"feature:{label}"] += 1
                except etree.XMLSyntaxError:
                    inventory[f"binary:{hashlib.sha256(data).hexdigest()}"] += 1
            else:
                inventory[f"binary:{hashlib.sha256(data).hexdigest()}"] += 1
    return inventory


def _extension(element) -> str | None:
    if not isinstance(element.tag, str) or not element.tag.endswith("}ext") or element.get("uri") is None:
        return None
    uri = element.get("uri")
    if uri in BOOKKEEPING_EXTENSIONS or (uri in DEFAULT_STYLE_EXTENSIONS and not any(len(child) for child in element)):
        return None
    return EXTENSIONS.get(uri, f"Office extension {uri}")


def _bookkeeping(element) -> bool:
    """Alternate content that only records where Excel last saved the file (x15ac:absPath)."""
    if element.tag != f"{{{MC}}}AlternateContent":
        return False
    leaves = [e for e in element.iter() if len(e) == 0]
    return bool(leaves) and all(e.tag == f"{{{X15AC}}}absPath" or e.tag.endswith("Fallback") for e in leaves)


def pdf_inventory(path: Path) -> Counter:
    from pypdf import PdfReader

    reader = PdfReader(path)
    inventory = Counter()
    inventory["pdf:pages"] = len(reader.pages)
    fields = reader.get_fields() or {}
    for name, value in fields.items():
        inventory["pdf:form fields"] += 1
        if value.get("/FT") == "/Sig":
            inventory["pdf:signatures"] += 1

    def walk(items):
        for item in items:
            if isinstance(item, list):
                walk(item)
            else:
                inventory["pdf:outline entries"] += 1

    try:
        walk(reader.outline)
    except Exception:
        pass
    inventory["pdf:attachments"] = len(reader.attachments or {})
    for page in reader.pages:
        inventory["pdf:annotations"] += len(page.get("/Annots") or [])
    return +inventory


def inventory(path: Path, format: str) -> Counter:
    return pdf_inventory(path) if format == "pdf" else ooxml_inventory(path)


def losses(before: Counter, after: Counter, allowed: Counter) -> list[dict]:
    """What the output lacks that the input had and no action declared it removes."""
    lost = []
    for key, count in before.items():
        missing = count - after.get(key, 0) - allowed.get(key, 0)
        if missing > 0:
            kind, _, name = key.partition(":")
            label = {"part": "parts of type", "binary": "binary part", "feature": "", "pdf": ""}.get(kind, kind)
            lost.append({"what": f"{label} {name}".strip() if kind != "binary" else f"binary part sha256:{name[:16]}…", "count": missing})
    return lost
