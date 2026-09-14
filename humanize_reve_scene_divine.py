from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED, ZIP_STORED
import xml.etree.ElementTree as ET


DOCX = Path("Le Logiciel d'incarnation - se liberer du piege.docx")
TXT = Path("Le Logiciel d'incarnation - se liberer du piege.txt")
TITLE = "Le rêve du seuil de la porte : à la recherche de la version originelle"
NEXT_PART = "Partie II - Les boucles collectives"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
W = f"{{{W_NS}}}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
ET.register_namespace("w", W_NS)

GOD_OLD = "Dans le rêve, Dieu lui-même prend la forme de l'abandon. Le mot peut sembler terrible, mais il ne désigne ni une faute divine ni un manque d'amour. Dieu abandonne l'individu au sens où il le remet à sa propre autonomie consciente. Il abandonne également la masse humaine à sa nature de groupe : elle s'organise, se corrige et se reproduit selon ses lois, puis se confie à des égrégores religieux, politiques, économiques ou idéologiques. Ces puissances collectives promettent l'appartenance, mais conduisent souvent l'individu non vers la rencontre avec lui-même, plutôt vers sa dilution et sa division en sous-versions toujours plus compatibles. Dieu ne vient pas interrompre chaque boucle pour sauver personnellement chaque version de nous. Il n'y a ni miracle exigible, ni dérogation particulière qui abolirait les conséquences de l'incarnation. Le retrait divin laisse la création tenir debout par elle-même. Il laisse aussi l'individu assez seul pour devenir réellement conscient."
GOD_NEW = "Dans le rêve, Dieu se révèle à moi sous la forme de l'abandon. Il ne se révèle pas en descendant dans le scénario humain pour le corriger, mais en me mettant à genoux devant le vide de ma propre dilution. La douleur devient révélation : j'ai passé ma vie à me disperser dans les attentes, les rôles et les avenirs du collectif au lieu d'être conscient de moi, de mes désirs et des futurs que mes choix pouvaient ouvrir. Dieu a abandonné l'individu au sein d'un scénario humain collectif qui s'autogère. Cet abandon ne désigne ni une faute divine ni un manque d'amour. Il signifie que Dieu n'a pas besoin d'interférer dans chaque scène, car la masse humaine s'est donné ses propres lois, ses mécanismes de correction et ses égrégores religieux, politiques, économiques ou idéologiques. Le groupe protège sa continuité en étouffant ce qui, dans l'individu, pourrait devenir une loi singulière. Il promet l'appartenance, mais conduit souvent non vers la rencontre avec soi, plutôt vers la dilution et la multiplication de versions toujours plus conformes. Dieu laisse cette humanité autonome suivre sa nature collective ; mais, dans le même retrait, il offre à celui qui prend conscience la possibilité de devenir lui aussi une nature autonome."

NATURE_OLD = "Cette autonomie ressemble à celle des éléments de la nature. La mer appartient au monde, échange avec l'air, reçoit les fleuves, répond à la lune et touche les rivages ; pourtant, elle demeure la mer. Elle n'a pas besoin de quitter la nature pour posséder sa propre puissance. Elle peut se lever en vague, se retirer, devenir calme ou ouragan. De même, la conscience éveillée reste parmi les humains et participe à leur vaste système interconnecté, mais elle ne se confond plus entièrement avec lui. Elle devient un élément singulier de la création, capable de composer sa forme avec ce qui la traverse sans être constamment réécrit par le scénario collectif."
NATURE_NEW = "La révélation me rend alors à une évidence oubliée : je suis moi aussi un élément autonome de la création. La mer appartient au monde, reçoit les fleuves, répond à la lune et touche les rivages ; pourtant, elle demeure la mer. Elle a le droit de se déchaîner en vagues, de se retirer ou de devenir calme. Le vent traverse la terre, rencontre les reliefs et influe sur tous les éléments qu'il touche ; pourtant, il obéit à sa propre nature lorsqu'il se lève. De même, l'individu conscient n'a pas à disparaître dans la masse humaine pour appartenir au réel. Il peut exprimer sa puissance, faire entendre ses désirs, choisir les avenirs qu'il veut ouvrir et influer sur le monde commun. Il reste relié aux humains, mais il n'est plus dissous en eux. Il devient une force singulière, capable de se donner des lois conscientes et de composer avec les autres éléments sans leur abandonner son centre."

PHRASE_REPLACEMENTS = {
    "à chaque redémarrage": "chaque fois que la scène de vie annulée m'était imposée de nouveau",
    "à la continuité du scénario destiné. Celui": "à la continuité du scénario destiné. La vie qu'il avait ouverte était annulée, et la scène devait être rejouée depuis son seuil. Celui",
    "rebootait la scène quelques instants plus tôt": "annulait cette scène de vie et m'obligeait à la rejouer quelques instants plus tôt",
    "Chaque redémarrage accomplissait": "Chaque reprise de la scène de vie annulée accomplissait",
    "les corrections, les redémarrages et les clonages successifs": "les corrections, les scènes annulées puis rejouées et les clonages successifs",
    "les reboots qui l'ont façonné": "les scènes de vie annulées puis rejouées qui l'ont façonné",
    "Le reboot du rêve ressemble ainsi": "Cette obligation de rejouer une scène de vie annulée ressemble ainsi",
    "immunisé contre le reboot intérieur": "devenu conscient du mécanisme qui annulait ses choix et le contraignait à rejouer sa vie",
    "sans provoquer de reboot": "sans être annulée ni imposée une nouvelle fois",
    "Il n'y a plus de reboot intérieur": "Il n'y a plus d'annulation intérieure de la scène choisie",
}


def paragraph_text(paragraph):
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS)).strip()


def paragraph_style(paragraph):
    style = paragraph.find("w:pPr/w:pStyle", NS)
    return style.get(f"{W}val") if style is not None else ""


def set_text(paragraph, text):
    nodes = paragraph.findall(".//w:t", NS)
    nodes[0].text = text
    nodes[0].set(XML_SPACE, "preserve")
    for node in nodes[1:]:
        node.text = ""


with ZipFile(DOCX, "r") as archive:
    root = ET.fromstring(archive.read("word/document.xml"))

in_chapter = False
changed = 0
god_changed = False
nature_changed = False
for paragraph in root.findall(".//w:body/w:p", NS):
    text = paragraph_text(paragraph)
    style = paragraph_style(paragraph)
    if style == "Heading2" and text == TITLE:
        in_chapter = True
        continue
    if in_chapter and style == "Heading1" and text == NEXT_PART:
        break
    if not in_chapter:
        continue

    new_text = text
    if text == GOD_OLD:
        new_text = GOD_NEW
        god_changed = True
    elif text == NATURE_OLD:
        new_text = NATURE_NEW
        nature_changed = True
    else:
        for old, new in PHRASE_REPLACEMENTS.items():
            new_text = new_text.replace(old, new)

    if new_text != text:
        set_text(paragraph, new_text)
        changed += 1

if not in_chapter or not god_changed or not nature_changed:
    raise RuntimeError("Le chapitre ou les paragraphes philosophiques attendus sont introuvables.")

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
        raise RuntimeError("Le DOCX est invalide.")

print(f"Paragraphes adaptés : {changed}")
