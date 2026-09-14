from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED, ZIP_STORED
import xml.etree.ElementTree as ET

DOCX = Path("Le Logiciel d'incarnation - se liberer du piege.docx")
TXT = Path("Le Logiciel d'incarnation - se liberer du piege.txt")
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
W = f"{{{W_NS}}}"
ET.register_namespace("w", W_NS)

with ZipFile(DOCX, "r") as archive:
    root = ET.fromstring(archive.read("word/document.xml"))

changed = 0
for paragraph in root.findall(".//w:body/w:p", NS):
    nodes = paragraph.findall(".//w:t", NS)
    text = "".join(node.text or "" for node in nodes)
    old = "Chaque reprise de la scène de vie annulée accomplissait donc une double suppression. Il tuait"
    if old in text:
        nodes[0].text = text.replace(old, "Chaque reprise de la scène de vie annulée accomplissait donc une double suppression. Elle tuait")
        for node in nodes[1:]:
            node.text = ""
        changed += 1

if changed != 1:
    raise RuntimeError(f"Correction attendue une fois, trouvée {changed} fois.")

xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
temp = DOCX.with_suffix(".tmp.docx")
with ZipFile(DOCX, "r") as zin, ZipFile(temp, "w") as zout:
    for item in zin.infolist():
        data = xml if item.filename == "word/document.xml" else zin.read(item.filename)
        compression = ZIP_STORED if item.filename == "mimetype" else ZIP_DEFLATED
        zout.writestr(item, data, compress_type=compression)
temp.replace(DOCX)

lines = []
for paragraph in root.findall(".//w:body/w:p", NS):
    text = "".join(node.text or "" for node in paragraph.findall(".//w:t", NS)).strip()
    if text:
        lines.append(text)
TXT.write_text("\n\n".join(lines) + "\n", encoding="utf-8")
