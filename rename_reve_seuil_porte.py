from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED, ZIP_STORED
import xml.etree.ElementTree as ET


DOCX = Path("Le Logiciel d'incarnation - se liberer du piege.docx")
TXT = Path("Le Logiciel d'incarnation - se liberer du piege.txt")
OLD = "Le rêve du seuil : à la recherche de la version originelle"
NEW = "Le rêve du seuil de la porte : à la recherche de la version originelle"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
W = f"{{{W_NS}}}"
ET.register_namespace("w", W_NS)


def paragraph_text(paragraph):
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS)).strip()


with ZipFile(DOCX, "r") as archive:
    root = ET.fromstring(archive.read("word/document.xml"))

replaced = 0
for paragraph in root.findall(".//w:body/w:p", NS):
    if paragraph_text(paragraph) != OLD:
        continue
    text_nodes = paragraph.findall(".//w:t", NS)
    text_nodes[0].text = NEW
    for node in text_nodes[1:]:
        node.text = ""
    replaced += 1

if replaced != 2:
    raise RuntimeError(f"Deux titres étaient attendus, {replaced} trouvé(s).")

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
    text = paragraph_text(paragraph)
    if text:
        lines.append(text)
TXT.write_text("\n\n".join(lines) + "\n", encoding="utf-8")

with ZipFile(DOCX, "r") as archive:
    if archive.testzip() is not None:
        raise RuntimeError("Le fichier DOCX est invalide.")

print(f"Titres remplacés : {replaced}")
