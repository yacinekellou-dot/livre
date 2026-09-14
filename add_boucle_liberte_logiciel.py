from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED, ZIP_STORED
import shutil
import xml.etree.ElementTree as ET

DOCX = Path("Le Logiciel d'incarnation - se liberer du piege.docx")
TXT = Path("Le Logiciel d'incarnation - se liberer du piege.txt")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
NS = {"w": W_NS}
ET.register_namespace("w", W_NS)

W = f"{{{W_NS}}}"
XML_SPACE = f"{{{XML_NS}}}space"

TITLE = "La boucle de la recherche de la liberté"
TOC_AFTER = "La lassitude de la boucle"
BODY_AFTER = "La lassitude de la boucle"
BODY_BEFORE = "Partie II - Les boucles collectives"

CHAPTER_PARAGRAPHS = [
    "Il existe une autre boucle, plus subtile que celle du travail, plus noble en apparence, et pourtant tout aussi capable de nous posséder : la boucle de la recherche de la liberté. Elle commence souvent par une sensation d'étouffement. Quelque chose nous serre. Un lieu, une relation, un métier, une règle, une attente, un regard. Nous croyons alors que la liberté se trouve exactement de l'autre côté de cette entrave.",
    "Nous nous disons : quand je serai débarrassé de cela, je serai libre. Quand je quitterai ce travail, je serai libre. Quand je ne dépendrai plus de ce patron, je serai libre. Quand je n'aurai plus cette obligation, cette contrainte, cette peur, cette personne à satisfaire, alors enfin je respirerai.",
    "Et parfois, il faut réellement partir. Certaines chaînes existent. Certaines contraintes abîment. Certaines situations doivent être quittées pour que la conscience cesse de se réduire. Mais le piège commence lorsque nous confondons toute liberté avec la suppression d'une entrave. Car une fois l'entrave supprimée, le désir ne s'arrête pas toujours. Il se déplace.",
    "La conscience croit avoir gagné sa liberté, puis elle découvre un nouveau manque. Elle erre de désir en désir. Elle quitte une cage, puis cherche aussitôt un autre objet à conquérir. Elle libère une zone de sa vie, puis se sent enfermée ailleurs. Le maître change de visage, mais la structure demeure : il faut encore atteindre quelque chose pour se sentir enfin vivant.",
    "La liberté n'est donc peut-être pas le simple fait de n'avoir aucune entrave. Une vie sans entrave totale n'existe pas. Le corps lui-même impose ses lois. Il faut dormir, manger, respirer, vieillir, habiter un temps, porter une histoire. Même celui qui refuse toutes les règles reste soumis à la gravité de ses besoins, à la mémoire de ses blessures, au mouvement de ses désirs.",
    "Passés les besoins de base, le désir devient plus incertain. Quand le corps a mangé, dormi, récupéré, quand la survie immédiate ne commande plus tout, le désir cherche un objet. Et s'il ne trouve pas en nous une direction profonde, il emprunte celle du dehors. Nous désirons alors ce que les autres désirent. Nous appelons cela choix personnel, alors que le désir prend ses ordres dans le regard social.",
    "C'est une servitude étrange : servir un désir dont on n'est pas vraiment le maître. Croire que l'on suit sa volonté, alors que l'on imite une image de réussite. Croire que l'on se libère, alors que l'on obéit à une comparaison. Croire que l'on choisit, alors que le logiciel a simplement installé en nous une nouvelle faim.",
    "Le désir peut devenir notre nouveau maître. Il peut prendre la place du patron, de la famille, de la morale, de la société, de l'ancien système que nous avons combattu. Nous pensions avoir quitté l'autorité extérieure, mais nous avons parfois seulement intériorisé une tyrannie plus mobile, plus séduisante, plus difficile à nommer.",
    "C'est pour cela que le désir de révolte ne naît pas toujours chez ceux qui sont objectivement les plus contraints. L'histoire le montre, et la Révolution française en donne une image forte : la volonté de renversement ne surgit pas uniquement de la misère absolue. Elle apparaît souvent lorsque la conscience commence à voir l'écart entre ce qu'elle vit et ce qu'elle croit pouvoir devenir.",
    "La volonté de libération n'est donc pas toujours le signe qu'il reste trop de chaînes. Elle peut aussi être le signe que le désir s'est éveillé, qu'il ne supporte plus son ancien maître, mais qu'il ne sait pas encore vers quelle nécessité intérieure se réorienter. Alors il attaque le maître de l'étape précédente : le patron, le travail, la famille, le pays, la religion, la morale, le système.",
    "Parfois cette attaque est juste. Mais parfois elle manque sa véritable cible. Nous croyons lutter contre l'enfermement alors que nous luttons seulement contre sa dernière forme visible. Nous supprimons une contrainte, puis la soif demeure. Nous changeons de décor, puis le manque recommence. Aucune suppression de contrainte ne fera taire à elle seule notre désir de liberté si ce désir reste sans orientation profonde.",
    "La liberté n'est donc pas laisser libre cours à tous ses désirs. Cette idée est séduisante, mais elle transforme vite l'être humain en serviteur de ce qui le traverse. Si tout désir devient un ordre, alors je ne suis pas libre. Je suis disponible à toutes les impulsions, à toutes les images, à toutes les promesses. Je ne suis plus maître de mon incarnation ; je deviens le lieu où passent des commandes venues d'ailleurs.",
    "Le véritable point de départ de la liberté commence peut-être lorsque l'on accepte cette phrase difficile : mon désir doit être éduqué, non pas écrasé. Un désir ne se supprime pas par décret. On ne devient pas libre en déclarant simplement : je ne veux plus. La volonté pure échoue souvent parce qu'elle croit pouvoir commander seule à l'affect. La raison explique, mais elle ne pèse pas toujours assez lourd face à une force vivante.",
    "Spinoza avait vu cela avec une lucidité redoutable : une passion ne se dépasse pas par une idée froide, mais par un affect plus puissant. Pour sortir d'un désir qui nous possède, il ne suffit pas de le condamner. Il faut lui opposer un désir plus fort, plus vrai, mieux orienté. Arrêter de fumer, par exemple, ne consiste pas seulement à haïr la cigarette. Il faut désirer davantage respirer, durer, retrouver son corps, respecter la vie qui nous porte.",
    "La vraie liberté n'est donc pas l'absence de loi. Sans loi intérieure, le désir devient errance. Sans forme, l'énergie se disperse. Sans direction, la révolte recommence sans cesse. Une liberté sans loi morale finit souvent par se retourner contre celui qui la revendique, parce qu'elle ne sait plus quoi faire de sa propre ouverture.",
    "Être libre, ce n'est pas obéir à n'importe quelle loi imposée de l'extérieur. Mais ce n'est pas non plus vivre sans loi. C'est suivre une loi que l'on s'est donnée à soi-même après avoir suffisamment regardé ce qui, en nous, relève de la nécessité. Non pas la nécessité sociale, non pas la peur, non pas l'imitation, mais cette direction intime où l'être sent : ceci correspond à ma nature profonde.",
    "La liberté devient alors autre chose qu'une fuite. Elle cesse d'être seulement une contrainte de moins. Elle devient consentement à ce que l'on est. Non pas résignation, mais reconnaissance. Le jour où je consens à ma nature propre, je ne suis plus seulement en train de m'arracher à ce qui m'enferme. Je commence à être déterminé par ce qui me constitue réellement.",
    "Cela demande de déshabiller les désirs. Les regarder un par un. Ce désir vient-il de moi, ou du regard des autres ? Vient-il d'une blessure, ou d'une nécessité ? Vient-il d'une comparaison, ou d'un appel intérieur ? Vient-il de la peur de manquer, ou d'une joie qui cherche sa forme ? Cette enquête peut prendre une vie entière, car le logiciel de l'incarnation a superposé en nous des couches de désirs appris, hérités, imités, compensatoires.",
    "Mais chaque désir vu clairement perd une partie de son pouvoir d'illusion. Chaque fois que je reconnais un désir qui ne m'appartient pas vraiment, je récupère une parcelle de présence. Chaque fois que je choisis un désir plus haut, plus juste, plus nécessaire, je ne supprime pas la boucle : je la réoriente.",
    "La liberté ne se cache donc pas derrière une contrainte de moins. Elle commence lorsque le désir cesse de courir au hasard et accepte d'être mis au service de ce qui, en nous, est vrai. Alors la question change. Ce n'est plus seulement : de quoi dois-je me libérer ? C'est : que vais-je faire de cette liberté ?",
    "Car la liberté sans orientation devient vite une nouvelle prison. Elle ouvre mille portes, mais aucune ne devient chemin. Elle permet tout, mais ne fonde rien. La liberté réelle demande une fidélité. Une loi intérieure. Une nécessité consentie. Elle n'est pas l'opposé de la contrainte ; elle est le choix conscient de la contrainte qui nous révèle au lieu de nous diminuer.",
    "Ainsi, se libérer du piège ne signifie pas devenir un être sans désir. Cela signifie devenir assez lucide pour ne plus laisser n'importe quel désir gouverner l'incarnation. La liberté véritable commence lorsque je ne suis plus seulement celui qui veut sortir, mais celui qui sait enfin vers quoi il accepte d'aller.",
]


def paragraph_text(paragraph):
    return "".join((node.text or "") for node in paragraph.findall(".//w:t", NS)).strip()


def paragraph_style(paragraph):
    style = paragraph.find("w:pPr/w:pStyle", NS)
    return style.get(f"{W}val") if style is not None else ""


def make_paragraph(text, style_name):
    paragraph = ET.Element(f"{W}p")
    ppr = ET.SubElement(paragraph, f"{W}pPr")
    ET.SubElement(ppr, f"{W}pStyle", {f"{W}val": style_name})
    run = ET.SubElement(paragraph, f"{W}r")
    text_node = ET.SubElement(run, f"{W}t")
    text_node.set(XML_SPACE, "preserve")
    text_node.text = text
    return paragraph


def remove_existing(body):
    paragraphs = list(body)
    start = None
    for index, paragraph in enumerate(paragraphs):
        if paragraph.tag == f"{W}p" and paragraph_text(paragraph) == TITLE:
            style = paragraph_style(paragraph)
            if style in {"Heading2", "TOC"}:
                start = index
                break
    if start is None:
        return

    style = paragraph_style(paragraphs[start])
    end = start + 1
    while end < len(paragraphs):
        item = paragraphs[end]
        if item.tag == f"{W}sectPr":
            break
        if item.tag == f"{W}p":
            next_style = paragraph_style(item)
            if style == "TOC" and next_style != "TOC":
                break
            if style == "Heading2" and next_style in {"Heading1", "Heading2"}:
                break
        end += 1
    for item in paragraphs[start:end]:
        body.remove(item)


def insert_after_toc(body):
    paragraphs = list(body)
    for index, paragraph in enumerate(paragraphs):
        if paragraph.tag == f"{W}p" and paragraph_style(paragraph) == "TOC" and paragraph_text(paragraph) == TOC_AFTER:
            body.insert(index + 1, make_paragraph(TITLE, "TOC"))
            return
    raise RuntimeError("Impossible de trouver l'entrée de table des matières cible.")


def insert_body_chapter(body):
    paragraphs = list(body)
    target_index = None
    for index, paragraph in enumerate(paragraphs):
        if paragraph.tag == f"{W}p" and paragraph_style(paragraph) == "Heading2" and paragraph_text(paragraph) == BODY_AFTER:
            target_index = index
            break
    if target_index is None:
        raise RuntimeError("Impossible de trouver le chapitre cible dans le corps du document.")

    insert_index = None
    for index in range(target_index + 1, len(paragraphs)):
        item = paragraphs[index]
        if item.tag == f"{W}p" and paragraph_text(item) == BODY_BEFORE and paragraph_style(item) == "Heading1":
            insert_index = index
            break
    if insert_index is None:
        raise RuntimeError("Impossible de trouver la partie suivante.")

    new_paragraphs = [make_paragraph(TITLE, "Heading2")]
    new_paragraphs.extend(make_paragraph(text, "Normal") for text in CHAPTER_PARAGRAPHS)
    for offset, paragraph in enumerate(new_paragraphs):
        body.insert(insert_index + offset, paragraph)


def write_docx(root, source, destination):
    document_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    temp = destination.with_suffix(".tmp.docx")
    with ZipFile(source, "r") as zin, ZipFile(temp, "w") as zout:
        for item in zin.infolist():
            data = document_xml if item.filename == "word/document.xml" else zin.read(item.filename)
            compression = ZIP_STORED if item.filename == "mimetype" else ZIP_DEFLATED
            zout.writestr(item, data, compress_type=compression)
    temp.replace(destination)


def write_txt_from_docx(docx_path, txt_path):
    with ZipFile(docx_path, "r") as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    body = root.find("w:body", NS)
    lines = []
    for paragraph in body.findall("w:p", NS):
        parts = []
        for node in paragraph.iter():
            if node.tag == f"{W}t":
                parts.append(node.text or "")
            elif node.tag == f"{W}tab":
                parts.append("\t")
            elif node.tag == f"{W}br":
                parts.append("\n")
        text = "".join(parts).strip()
        if text:
            lines.append(text)
    txt_path.write_text("\n\n".join(lines) + "\n", encoding="utf-8")


def main():
    if not DOCX.exists():
        raise SystemExit(f"Fichier introuvable: {DOCX}")

    with ZipFile(DOCX, "r") as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    body = root.find("w:body", NS)

    remove_existing(body)
    remove_existing(body)
    insert_after_toc(body)
    insert_body_chapter(body)

    backup = DOCX.with_suffix(".avant_boucle_liberte_backup.docx")
    if not backup.exists():
        shutil.copy2(DOCX, backup)

    write_docx(root, DOCX, DOCX)
    write_txt_from_docx(DOCX, TXT)

    with ZipFile(DOCX, "r") as archive:
        bad_file = archive.testzip()
    if bad_file:
        raise RuntimeError(f"DOCX invalide, premier fichier abîmé: {bad_file}")

    print(f"Chapitre ajouté: {TITLE}")
    print(f"DOCX mis à jour: {DOCX}")
    print(f"TXT synchronisé: {TXT}")


if __name__ == "__main__":
    main()
