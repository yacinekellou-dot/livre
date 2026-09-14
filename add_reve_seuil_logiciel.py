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

TITLE = "Le rêve du seuil : à la recherche de la version originelle"
TOC_AFTER = "La boucle de la recherche de la liberté"
BODY_BEFORE = "Partie II - Les boucles collectives"

CHAPTER_PARAGRAPHS = [
    "Dans la nuit du dimanche 30 au lundi 31 août, j'ai fait un rêve dont la simplicité apparente dissimulait une architecture vertigineuse. Je rentrais chez moi. J'avais les clés de l'appartement et je revenais de l'extérieur, comme après une longue randonnée. Sous mes chaussures, la terre s'était accumulée. Sur le pas de la porte, dans le hall d'entrée, ma mère et ma grand-mère m'attendaient. Leur présence semblait coïncider exactement avec mon retour.",
    "Ma mère me disait de faire attention, d'essuyer mes pieds et d'enlever mes chaussures avant d'entrer. Je m'exécutais. Puis je traversais l'appartement pour rejoindre ma chambre. Rien d'extraordinaire ne semblait devoir arriver. Pourtant, la scène recommençait.",
    "Je revenais de l'extérieur. Ma mère et ma grand-mère se tenaient dans l'entrée. Mes semelles portaient la même terre. La même recommandation m'était adressée. J'essuyais mes chaussures, je les retirais, j'entrais. Puis tout recommençait encore, tandis que j'oubliais chaque répétition au moment même où elle s'achevait.",
    "Peu à peu, un détail fissura l'évidence du décor. En frottant mes semelles, j'aperçus leurs motifs carrés, réguliers, répétés comme les cases d'une grille. Des visions très brèves me montrèrent la même scène déjà vécue : mes gestes, le seuil, la terre, les deux femmes, l'ordre d'entrer proprement. D'abord, je n'y prêtai pas attention. Le système pouvait continuer parce que ma mémoire ne reliait pas encore les fragments.",
    "À la répétition suivante, les images se rassemblèrent. Je compris soudain que je rejouais la même séquence. Je compris aussi que cette boucle n'était pas un simple retour. Après avoir essuyé mes chaussures, un scénario précis devait se poursuivre : saluer, déjeuner, accomplir le geste attendu, suivre le chemin préparé. Si je bifurquais, si je choisissais un avenir que la scène n'avait pas prévu, le système supprimait cette possibilité, détruisait cette version de moi et me relançait quelques instants plus tôt.",
    "Chaque redémarrage annulait l'avenir que j'avais tenté de choisir. La scène antérieure disparaissait avec l'être qui l'avait vécue. Celui qui revenait au seuil lui ressemblait parfaitement, possédait presque tous ses souvenirs, croyait être le même, mais il était déjà une nouvelle version. Dans le rêve, cette connaissance me traversa en une fraction de seconde : je n'étais peut-être plus la version première de moi-même, mais la version 1865, ou une autre parmi des centaines, produite après autant de bifurcations refusées.",
    "Au commencement, il y aurait eu une version zéro : non clonée, non corrigée, encore entière. Puis, au fil de la vie, chaque écart aurait provoqué un ajustement. L'enfant originel aurait disparu sous les versions successives nécessaires pour s'accorder au scénario. Devenu adulte, je ne serais plus qu'un lointain descendant de moi-même, un clone suffisamment fidèle pour continuer à porter mon nom, mais trop éloigné pour reconnaître son origine.",
    "Alors je me suis effondré sur le pas de la porte, mes chaussures encore entre les mains. Une douleur déchirante m'a saisi. Je suppliais qu'on me rende ma version originelle. Je demandais pourquoi elle avait été perdue, pourquoi tant d'avenirs avaient dû être effacés. Dans le regard de ma mère et de ma grand-mère, je vis qu'elles comprenaient. Elles savaient. Elles avaient elles aussi traversé leurs propres répétitions et fait le deuil de ce qu'elles avaient été. Leur tristesse n'était pas celle de gardiennes cruelles, mais celle de femmes qui avaient appris à vivre à l'intérieur du scénario.",
    "Je leur en voulais pourtant. Leur silence me paraissait complice. Elles avaient continué à jouer leur rôle, à accueillir chacune de mes versions comme s'il s'agissait toujours du même enfant. Mais leur regard disait aussi autre chose : elles auraient voulu m'épargner cette connaissance. Mieux valait peut-être l'insouciance de celui qui se croit libre que la lucidité de celui qui découvre combien de fois il s'est abandonné pour rester compatible avec le monde.",
    "Le rêve s'arrête là où commence sa véritable portée philosophique. Il ne prouve évidemment pas que l'univers détruit littéralement nos futurs divergents ou fabrique des copies successives de notre personne. Mais il révèle avec une force exceptionnelle une vérité intérieure : nous pouvons vivre comme si chacune de nos adaptations avait remplacé une part de nous-mêmes.",
    "Le décor du rêve est un seuil. Je ne suis ni tout à fait dehors ni encore dedans. Le seuil est l'endroit où l'existence demande : qu'apportes-tu du monde, et que dois-tu abandonner pour être admis parmi les tiens ? La terre sous les chaussures est la trace du chemin personnel, de l'expérience sauvage, de ce qui a été touché loin de la maison. Avant de rentrer dans l'espace familial, il faut la retirer. Le geste est raisonnable dans la vie ordinaire ; symboliquement, il devient ambigu. S'agit-il seulement de préserver la maison, ou d'effacer sur soi les preuves du chemin parcouru ?",
    "Les semelles portent des carrés répétés. Ce motif transforme le chemin en matrice. La terre libre de la randonnée se trouve capturée par une géométrie régulière. À chaque pas, la grille imprime sa forme au sol ; à chaque retour, elle rappelle que même notre mouvement peut être programmé. Nous pensons avancer, mais nos pas reproduisent parfois un dessin installé avant nous : les habitudes de la famille, les règles de l'éducation, la morale du groupe, les attentes de la société, les automatismes du travail.",
    "Ma mère et ma grand-mère représentent alors davantage que deux personnes réelles. Elles forment une lignée. Elles sont la transmission vivante, celles par qui la vie arrive et celles par qui les règles se perpétuent. Elles se tiennent dans le hall parce que toute incarnation humaine entre dans un monde déjà habité. Avant même de pouvoir choisir, l'enfant rencontre des voix qui lui apprennent où poser les pieds, ce qu'il faut nettoyer, ce qui se fait et ce qui ne se fait pas.",
    "Cette transmission n'est pas nécessairement malveillante. La mère protège la maison ; la grand-mère porte une mémoire plus ancienne encore. Elles enseignent les conditions de l'appartenance parce qu'elles les ont elles-mêmes reçues. C'est précisément ce qui rend la boucle si difficile à voir : ceux qui nous programment ne sont pas toujours nos ennemis. Ils nous aiment souvent avec les outils que le système leur a donnés. Ils transmettent leurs peurs en croyant transmettre la prudence, leurs renoncements en croyant transmettre la sagesse, leurs scénarios en croyant transmettre la vie.",
    "Le reboot du rêve ressemble ainsi au fonctionnement du logiciel d'incarnation. Lorsqu'un comportement menace l'équilibre du groupe, la réalité sociale ne nous tue pas physiquement ; elle corrige souvent notre trajectoire par le rejet, la honte, la culpabilité, la comparaison ou la peur de perdre l'amour. L'enfant essaie une voie, rencontre une sanction, puis revient au seuil. Il rejoue la scène autrement. Une version spontanée de lui-même s'efface, et une version plus compatible prend sa place.",
    "Le clone n'est donc pas nécessairement un autre corps. Il est le moi adapté qui affirme encore « je » après avoir oublié ce qu'il a sacrifié. Il porte le même visage, mais son désir a changé de maître. Il appelle personnalité ce qui fut peut-être une stratégie de survie. Il appelle choix ce qui fut peut-être une correction. Il appelle maturité l'accumulation des gestes grâce auxquels il a appris à ne plus déranger le scénario.",
    "Le nombre 1865 ne désigne pas un compte exact. Il donne une mesure imaginaire à l'épaisseur de nos renoncements. Combien de fois avons-nous étouffé une parole avant qu'elle ne sorte ? Combien de fois avons-nous choisi la réponse acceptable plutôt que la réponse vraie ? Combien de désirs ont été déclarés impossibles avant même d'être examinés ? Chaque micro-renoncement peut être vécu comme la mort silencieuse d'un avenir.",
    "Voilà pourquoi la prise de conscience provoque un deuil. La lucidité n'est pas toujours lumineuse. Elle peut d'abord être une catastrophe intime. Voir la boucle, c'est apercevoir toutes les vies que nous n'avons pas vécues et tous les êtres que nous aurions pu devenir. C'est découvrir que l'adulte dont nous sommes fiers s'est parfois construit sur les ruines invisibles de l'enfant qu'il devait protéger.",
    "Mais le rêve contient aussi un piège philosophique : croire qu'il suffirait de retourner à la version zéro. Cette version originelle est peut-être introuvable, non parce qu'un système l'aurait détruite, mais parce qu'aucune existence humaine ne demeure intacte. Vivre, c'est être modifié. La conscience ne peut pas effacer le temps et redevenir celle qui précédait toute influence. Chercher l'origine comme une pureté perdue risque de produire une nouvelle boucle : le refus douloureux de tout ce que nous sommes devenus.",
    "La tâche n'est donc pas de revenir à une copie ancienne de soi. Elle est de retrouver, au milieu des versions accumulées, le principe vivant qui pouvait encore choisir. L'unité n'est pas derrière nous comme une photographie intacte ; elle doit être reconstruite. La version originelle n'est peut-être pas un personnage enfoui qu'il faudrait exhumer, mais une capacité : celle d'habiter pleinement l'acte présent, de reconnaître les programmes reçus et de décider lesquels méritent désormais notre fidélité.",
    "À cet instant, la solitude apparaît. Lorsque je cesse d'obéir automatiquement au scénario, personne ne peut choisir exactement à ma place. La famille peut m'aimer, la société peut me conseiller, la tradition peut me transmettre ses cartes, mais aucune ne peut accomplir mon incarnation. Cette solitude n'est pas l'absence des autres. Elle est l'irréductible responsabilité d'être celui qui répond de sa propre vie.",
    "Dans le rêve, Dieu lui-même prend la forme d'un abandon, d'une mélancolie douce et poignante. Le monde semble avoir été créé comme un système autonome, relié à la nature et capable de s'autoréguler selon ses lois. Il n'y a ni miracle commandé pour me soustraire aux conséquences, ni dérogation personnelle qui suspendrait la condition humaine. Cette image n'est pas nécessairement celle d'un Dieu absent. Elle peut être celle d'un Dieu qui ne confisque pas la création, qui lui laisse sa cohérence et nous laisse la dignité terrible de participer à son devenir.",
    "La liberté cesse alors d'être la promesse d'une protection absolue. Elle ressemble à la mer : elle reçoit les vents, les marées, la lune, les rivages et pourtant, dans ces déterminations, elle produit sa propre forme. Elle peut se lever en vague, se retirer, devenir calme ou ouragan. Elle ne choisit pas les lois de l'univers, mais elle ne se réduit pas à une immobilité. Sa puissance naît de la danse singulière qu'elle compose avec ce qui la traverse.",
    "De la même manière, devenir maître de son incarnation ne signifie pas sortir du système humain, abolir toute influence ou inventer sa vie depuis le néant. Cela signifie ne plus confondre les lois du monde avec l'obligation de répéter inconsciemment le même rôle. Nous ne choisissons ni notre naissance, ni notre lignée, ni toutes les forces qui nous façonnent. Mais nous pouvons apprendre à voir le moment du seuil : cette seconde où le programme attend notre geste habituel et où une conscience devenue présente peut répondre autrement.",
    "Répondre autrement ne garantit pas que le monde nous applaudira. Une bifurcation réelle peut coûter une appartenance, une image, un confort ou une ancienne identité. Elle peut donner l'impression qu'une version de nous meurt. Pourtant, cette mort symbolique n'est plus un reboot imposé ; elle devient une transformation consentie. Je ne suis plus supprimé parce que j'ai quitté le scénario. J'accepte de laisser mourir en moi ce qui ne peut plus porter ma vérité.",
    "Au réveil, le lundi 31 août, je me suis regardé dans le miroir avec le sentiment que quelque chose avait bougé pendant la nuit. Le visage était le même, mais la question ne l'était plus : quelle version de moi me regarde ? Est-elle le résultat passif de toutes les corrections reçues, ou le lieu depuis lequel une présence peut enfin se rassembler ?",
    "Peut-être ne retrouvons-nous jamais la version zéro. Peut-être est-ce même une chance, car elle était encore inconsciente de tout ce qu'elle aurait à traverser. Mais nous pouvons faire naître une version qui se souvient symboliquement de toutes les autres, qui accueille leurs blessures sans leur abandonner le gouvernement de sa vie. Une version non pas originale parce qu'elle serait première, mais originelle parce qu'elle redevient source.",
    "Le rêve du seuil ne me demande donc pas de fuir la maison, de condamner la mère, de rejeter la grand-mère ou de déclarer la guerre au destin. Il me demande d'essuyer mes chaussures en sachant pourquoi je le fais. D'entrer si je choisis d'entrer. De saluer si mon geste est vivant. De repartir si mon chemin l'exige. La liberté commence peut-être là : lorsque la scène peut se répéter, mais que la conscience, elle, n'est plus absente.",
    "Alors le système ne disparaît pas. Les lois naturelles demeurent, les humains restent interconnectés, les héritages continuent de parler. Mais leur mainmise cesse d'être totale. Je ne suis plus seulement la version que le monde a réussi à conserver. Je deviens celui qui voit les versions, recueille leur mémoire et choisit, au seuil de chaque instant, laquelle mérite de poursuivre l'incarnation.",
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
    for sought_style in ("TOC", "Heading2"):
        paragraphs = list(body)
        start = next((i for i, p in enumerate(paragraphs)
                      if p.tag == f"{W}p" and paragraph_style(p) == sought_style
                      and paragraph_text(p) == TITLE), None)
        if start is None:
            continue
        end = start + 1
        while end < len(paragraphs):
            item = paragraphs[end]
            if item.tag == f"{W}sectPr":
                break
            if item.tag == f"{W}p":
                next_style = paragraph_style(item)
                if sought_style == "TOC" and next_style != "TOC":
                    break
                if sought_style == "Heading2" and next_style in {"Heading1", "Heading2"}:
                    break
            end += 1
        for item in paragraphs[start:end]:
            body.remove(item)


def insert_toc(body):
    paragraphs = list(body)
    for index, paragraph in enumerate(paragraphs):
        if paragraph_style(paragraph) == "TOC" and paragraph_text(paragraph) == TOC_AFTER:
            body.insert(index + 1, make_paragraph(TITLE, "TOC"))
            return
    raise RuntimeError("Entrée cible introuvable dans la table des matières.")


def insert_chapter(body):
    paragraphs = list(body)
    for index, paragraph in enumerate(paragraphs):
        if paragraph_style(paragraph) == "Heading1" and paragraph_text(paragraph) == BODY_BEFORE:
            additions = [make_paragraph(TITLE, "Heading2")]
            additions.extend(make_paragraph(text, "Normal") for text in CHAPTER_PARAGRAPHS)
            for offset, addition in enumerate(additions):
                body.insert(index + offset, addition)
            return
    raise RuntimeError("Partie suivante introuvable dans le corps du document.")


def write_docx(root):
    xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    temp = DOCX.with_suffix(".tmp.docx")
    with ZipFile(DOCX, "r") as zin, ZipFile(temp, "w") as zout:
        for item in zin.infolist():
            data = xml if item.filename == "word/document.xml" else zin.read(item.filename)
            compression = ZIP_STORED if item.filename == "mimetype" else ZIP_DEFLATED
            zout.writestr(item, data, compress_type=compression)
    temp.replace(DOCX)


def sync_txt():
    with ZipFile(DOCX, "r") as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    lines = []
    for paragraph in root.findall(".//w:body/w:p", NS):
        text = paragraph_text(paragraph)
        if text:
            lines.append(text)
    TXT.write_text("\n\n".join(lines) + "\n", encoding="utf-8")


def main():
    with ZipFile(DOCX, "r") as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    body = root.find("w:body", NS)
    remove_existing(body)
    insert_toc(body)
    insert_chapter(body)

    backup = DOCX.with_suffix(".avant_reve_du_seuil_backup.docx")
    if not backup.exists():
        shutil.copy2(DOCX, backup)
    write_docx(root)
    sync_txt()

    with ZipFile(DOCX, "r") as archive:
        if archive.testzip() is not None:
            raise RuntimeError("L'archive DOCX est invalide.")
    print(TITLE)


if __name__ == "__main__":
    main()
