from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED, ZIP_STORED
import xml.etree.ElementTree as ET


DOCX = Path("Le Logiciel d'incarnation - se liberer du piege.docx")
TXT = Path("Le Logiciel d'incarnation - se liberer du piege.txt")
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
W = f"{{{W_NS}}}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
ET.register_namespace("w", W_NS)

REPLACEMENTS = {
    "Ma mère me disait de faire attention, d'essuyer mes pieds et d'enlever mes chaussures avant d'entrer. Je m'exécutais. Puis je traversais l'appartement pour rejoindre ma chambre. Rien d'extraordinaire ne semblait devoir arriver. Pourtant, la scène recommençait.":
    "Ma mère me disait de faire attention, d'essuyer mes pieds et d'enlever mes chaussures avant d'entrer. Je m'exécutais. À cet instant, plusieurs gestes devenaient possibles. Je pouvais aller saluer, déjeuner, rejoindre ma chambre ou suivre une impulsion imprévue. Dans la première scène, je choisissais spontanément un comportement qui n'appartenait pas à la suite attendue. Ce geste minuscule ouvrait pourtant une bifurcation : à partir de lui, d'autres paroles, d'autres rencontres et un autre avenir pouvaient naître. Mais cette branche nouvelle n'était pas autorisée. Avant même qu'elle puisse se déployer, tout s'interrompait. La scène recommençait.",

    "Je revenais de l'extérieur. Ma mère et ma grand-mère se tenaient dans l'entrée. Mes semelles portaient la même terre. La même recommandation m'était adressée. J'essuyais mes chaussures, je les retirais, j'entrais. Puis tout recommençait encore, tandis que j'oubliais chaque répétition au moment même où elle s'achevait.":
    "Je revenais de l'extérieur. Ma mère et ma grand-mère se tenaient dans l'entrée. Mes semelles portaient la même terre. La même recommandation m'était adressée. J'essuyais mes chaussures, je les retirais, puis une nouvelle possibilité surgissait en moi. Chaque fois que je choisissais spontanément un comportement différent, ce choix produisait son propre futur potentiel. Mais un autre avenir existait déjà : un avenir principal, préparé par le système humain, dans lequel je devais accomplir les gestes prévus et reprendre ma place. Les deux futurs entraient alors en compétition. Le futur spontané menaçait la continuité du futur destiné ; le système tranchait en faveur de ce dernier. Il annulait la bifurcation, effaçait la scène et me ramenait au pas de la porte.",

    "Peu à peu, un détail fissura l'évidence du décor. En frottant mes semelles, j'aperçus leurs motifs carrés, réguliers, répétés comme les cases d'une grille. Des visions très brèves me montrèrent la même scène déjà vécue : mes gestes, le seuil, la terre, les deux femmes, l'ordre d'entrer proprement. D'abord, je n'y prêtai pas attention. Le système pouvait continuer parce que ma mémoire ne reliait pas encore les fragments.":
    "La répétition était parfaite parce que j'oubliais à chaque redémarrage. Je croyais toujours rentrer pour la première fois. Pourtant, un détail finit par fissurer l'évidence du décor. En frottant mes semelles, j'aperçus leurs motifs carrés, réguliers, répétés comme les cases d'une grille. Des visions très brèves me montrèrent la même scène déjà vécue : mes gestes, le seuil, la terre, les deux femmes, l'ordre d'entrer proprement, puis l'instant exact où mon comportement avait dévié. D'abord, je n'y prêtai pas attention. Tant que ma mémoire ne reliait pas les fragments, le système pouvait supprimer un avenir et me faire rejouer la scène sans que je comprenne ce qui avait été perdu.",

    "À la répétition suivante, les images se rassemblèrent. Je compris soudain que je rejouais la même séquence. Je compris aussi que cette boucle n'était pas un simple retour. Après avoir essuyé mes chaussures, un scénario précis devait se poursuivre : saluer, déjeuner, accomplir le geste attendu, suivre le chemin préparé. Si je bifurquais, si je choisissais un avenir que la scène n'avait pas prévu, le système supprimait cette possibilité, détruisait cette version de moi et me relançait quelques instants plus tôt.":
    "À la répétition suivante, les images se rassemblèrent. Je compris soudain que je rejouais la même séquence. La boucle n'était pas un simple retour du temps : elle était un mécanisme de correction. Après avoir essuyé mes chaussures, un scénario principal devait se poursuivre : saluer, déjeuner, accomplir le geste attendu, suivre le chemin préparé. Mais mon choix spontané créait un second scénario, réel dès sa naissance parce qu'il contenait déjà un potentiel d'événements à venir. Deux lignes d'avenir se faisaient face : celle que je venais d'ouvrir et celle que le système considérait comme destinée. La coexistence n'était pas permise. Dès que ma bifurcation menaçait le scénario principal, le système supprimait l'avenir concurrent, tuait la version de moi qui l'avait choisi et rebootait la scène quelques instants plus tôt.",

    "Chaque redémarrage annulait l'avenir que j'avais tenté de choisir. La scène antérieure disparaissait avec l'être qui l'avait vécue. Celui qui revenait au seuil lui ressemblait parfaitement, possédait presque tous ses souvenirs, croyait être le même, mais il était déjà une nouvelle version. Dans le rêve, cette connaissance me traversa en une fraction de seconde : je n'étais peut-être plus la version première de moi-même, mais la version 1865, ou une autre parmi des centaines, produite après autant de bifurcations refusées.":
    "Chaque redémarrage accomplissait donc une double suppression. Il tuait d'abord une réalité possible : les scènes, les rencontres et les conséquences qui auraient pu naître de mon choix. Il tuait ensuite le moi réel qui avait produit cette possibilité, car cette version de moi ne pouvait plus être admise dans la continuité du scénario destiné. Celui qui revenait au seuil lui ressemblait parfaitement et croyait être le même, mais il était un clone fonctionnel, une nouvelle version remise au point de départ pour tenter, cette fois, d'emprunter l'avenir principal. Dans le rêve, cette connaissance me traversa en une fraction de seconde : je n'étais peut-être plus la version première de moi-même, mais la version 1865, produite après autant de futurs spontanés supprimés et de bifurcations refusées.",

    "Le rêve s'arrête là où commence sa véritable portée philosophique. Il ne prouve évidemment pas que l'univers détruit littéralement nos futurs divergents ou fabrique des copies successives de notre personne. Mais il révèle avec une force exceptionnelle une vérité intérieure : nous pouvons vivre comme si chacune de nos adaptations avait remplacé une part de nous-mêmes.":
    "Le rêve s'arrête là où commence sa véritable portée philosophique. Il ne prouve évidemment pas que l'univers détruit littéralement nos futurs divergents ou fabrique matériellement des copies successives de notre personne. Mais il met en scène une intuition plus radicale qu'une simple adaptation : tout choix spontané ouvre une branche de réalité. Cette branche entre en compétition avec l'avenir principal que le système familial, social et humain cherche à maintenir. Lorsque la bifurcation n'est pas autorisée, le système ne se contente pas de corriger un geste ; il rend tout un futur impraticable et remplace progressivement l'être qui l'avait désiré par une version plus compatible avec la destinée collective.",

    "Le reboot du rêve ressemble ainsi au fonctionnement du logiciel d'incarnation. Lorsqu'un comportement menace l'équilibre du groupe, la réalité sociale ne nous tue pas physiquement ; elle corrige souvent notre trajectoire par le rejet, la honte, la culpabilité, la comparaison ou la peur de perdre l'amour. L'enfant essaie une voie, rencontre une sanction, puis revient au seuil. Il rejoue la scène autrement. Une version spontanée de lui-même s'efface, et une version plus compatible prend sa place.":
    "Le reboot du rêve ressemble ainsi au fonctionnement du logiciel d'incarnation. Le système humain porte un avenir principal : la continuité de la famille, la stabilité du groupe, la répétition des rôles, la prévisibilité des conduites. Lorsqu'un comportement spontané ouvre une scène nouvelle, il crée en même temps une ligne d'avenir concurrente. Si cette ligne menace l'équilibre collectif, la réalité sociale ne nous tue pas physiquement ; elle la neutralise par le rejet, la honte, la culpabilité, la comparaison ou la peur de perdre l'amour. L'enfant essaie une voie, rencontre la force corrective du groupe, puis revient symboliquement au seuil. Il rejoue la scène autrement. Le futur qu'il avait entrevu devient inaccessible, la version qui l'avait voulu se tait, et une copie plus compatible prend sa place dans le scénario principal.",
}


def paragraph_text(paragraph):
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS)).strip()


with ZipFile(DOCX, "r") as archive:
    root = ET.fromstring(archive.read("word/document.xml"))

found = set()
for paragraph in root.findall(".//w:body/w:p", NS):
    old = paragraph_text(paragraph)
    if old not in REPLACEMENTS:
        continue
    nodes = paragraph.findall(".//w:t", NS)
    nodes[0].text = REPLACEMENTS[old]
    nodes[0].set(XML_SPACE, "preserve")
    for node in nodes[1:]:
        node.text = ""
    found.add(old)

if len(found) != len(REPLACEMENTS):
    missing = set(REPLACEMENTS) - found
    raise RuntimeError(f"Paragraphes introuvables : {len(missing)}")

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

print(f"Paragraphes renforcés : {len(found)}")
