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
    "Alors je me suis effondré sur le pas de la porte, mes chaussures encore entre les mains. Une douleur déchirante m'a saisi. Je suppliais qu'on me rende ma version originelle. Je demandais pourquoi elle avait été perdue, pourquoi tant d'avenirs avaient dû être effacés. Dans le regard de ma mère et de ma grand-mère, je vis qu'elles comprenaient. Elles savaient. Elles avaient elles aussi traversé leurs propres répétitions et fait le deuil de ce qu'elles avaient été. Leur tristesse n'était pas celle de gardiennes cruelles, mais celle de femmes qui avaient appris à vivre à l'intérieur du scénario.":
    "Alors mes jambes ont cédé. Je me suis effondré sur le pas de la porte et je suis tombé à genoux, mes chaussures encore entre les mains. Un vide immense s'est ouvert dans ma poitrine. Ce qui me brisait n'était pas seulement la découverte des scènes répétées : c'était la certitude d'avoir perdu la version zéro, celle qui avait précédé les corrections, les redémarrages et les clonages successifs. Je m'étais perdu moi-même pendant tout ce temps sans même savoir que je me perdais. Chaque fois qu'un futur spontané avait été supprimé, une partie réelle de moi avait disparu avec lui. Je pleurais cet être premier comme on pleure un mort dont on découvre trop tard l'existence. Je pleurais aussi la naïveté que la révélation venait de tuer : l'illusion rassurante d'un libre arbitre intact, l'ego persuadé d'avoir toujours choisi sa vie, alors qu'il avait peut-être seulement donné son nom aux décisions de ses clones successifs. Je suppliais qu'on me rende cette origine, tout en comprenant qu'aucun retour ne pourrait abolir ce que je venais de voir.",

    "Je leur en voulais pourtant. Leur silence me paraissait complice. Elles avaient continué à jouer leur rôle, à accueillir chacune de mes versions comme s'il s'agissait toujours du même enfant. Mais leur regard disait aussi autre chose : elles auraient voulu m'épargner cette connaissance. Mieux valait peut-être l'insouciance de celui qui se croit libre que la lucidité de celui qui découvre combien de fois il s'est abandonné pour rester compatible avec le monde.":
    "Dans le regard de ma mère et de ma grand-mère, je vis qu'elles savaient. Elles vivaient elles aussi à côté de leur ancienne nature consciente, endeuillées de l'être profond qu'elles avaient abandonné au service du collectif humain. Elles ne me l'avaient jamais dit. Elles avaient accepté de jouer le jeu de l'humanité, de reconnaître chaque nouveau clone comme la suite naturelle du précédent et d'accompagner le scénario principal. Je leur en voulais précisément pour cela. Leur silence m'avait laissé dans la fausse liberté de l'ego, dans le confort de celui qui se croit auteur parce qu'il ignore les reboots qui l'ont façonné. Pourtant, leur tristesse révélait qu'elles n'étaient pas les gardiennes triomphantes du système, mais des femmes ayant elles-mêmes traversé l'épreuve du deuil de soi. Je compris peu à peu qu'elles avaient peut-être voulu me protéger : mieux valait, pensaient-elles, me laisser être heureux et inconscient que m'imposer trop tôt cette connaissance qui vide le cœur avant de le rendre libre.",

    "Voilà pourquoi la prise de conscience provoque un deuil. La lucidité n'est pas toujours lumineuse. Elle peut d'abord être une catastrophe intime. Voir la boucle, c'est apercevoir toutes les vies que nous n'avons pas vécues et tous les êtres que nous aurions pu devenir. C'est découvrir que l'adulte dont nous sommes fiers s'est parfois construit sur les ruines invisibles de l'enfant qu'il devait protéger.":
    "Voilà pourquoi la prise de conscience provoque un deuil. La lucidité n'est pas toujours lumineuse ; elle peut d'abord être un effondrement. Voir la boucle, c'est apercevoir toutes les vies que nous n'avons pas vécues, tous les futurs supprimés et toutes les versions de nous-mêmes abandonnées afin que l'humanité conserve la cohérence de son scénario. C'est découvrir que l'adulte dont nous sommes fiers s'est parfois construit sur les ruines invisibles de l'enfant qu'il devait protéger. Mais c'est aussi sentir, pour la première fois, que le clonage peut s'arrêter : dès que la conscience voit le mécanisme, elle cesse d'être entièrement disponible à sa répétition.",

    "À cet instant, la solitude apparaît. Lorsque je cesse d'obéir automatiquement au scénario, personne ne peut choisir exactement à ma place. La famille peut m'aimer, la société peut me conseiller, la tradition peut me transmettre ses cartes, mais aucune ne peut accomplir mon incarnation. Cette solitude n'est pas l'absence des autres. Elle est l'irréductible responsabilité d'être celui qui répond de sa propre vie.":
    "À cet instant, une solitude aiguë apparaît. Lorsque je cesse d'obéir automatiquement au scénario, je demeure au milieu des humains, relié à eux par l'amour, la mémoire, le langage et les conséquences de chaque acte ; pourtant, je ne suis plus seulement une cellule fondue dans l'organisme collectif. Je deviens moi-même un élément autonome à côté de l'humanité interconnectée. Personne ne peut plus choisir exactement à ma place. La famille peut m'aimer, la société peut me conseiller, la tradition peut me transmettre ses cartes, mais aucune ne peut accomplir mon incarnation. Cette solitude n'est pas l'absence des autres. Elle est la sensation vertigineuse d'exister par soi-même au sein d'un monde qui continue sans nous demander notre permission.",

    "Dans le rêve, Dieu lui-même prend la forme d'un abandon, d'une mélancolie douce et poignante. Le monde semble avoir été créé comme un système autonome, relié à la nature et capable de s'autoréguler selon ses lois. Il n'y a ni miracle commandé pour me soustraire aux conséquences, ni dérogation personnelle qui suspendrait la condition humaine. Cette image n'est pas nécessairement celle d'un Dieu absent. Elle peut être celle d'un Dieu qui ne confisque pas la création, qui lui laisse sa cohérence et nous laisse la dignité terrible de participer à son devenir.":
    "Dans le rêve, Dieu lui-même prend la forme de l'abandon. Le mot peut sembler terrible, mais il ne désigne ni une faute divine ni un manque d'amour. Dieu abandonne l'individu au sens où il le remet à sa propre autonomie consciente. Il abandonne également la masse humaine à sa nature de groupe : elle s'organise, se corrige et se reproduit selon ses lois, puis se confie à des égrégores religieux, politiques, économiques ou idéologiques. Ces puissances collectives promettent l'appartenance, mais conduisent souvent l'individu non vers la rencontre avec lui-même, plutôt vers sa dilution et sa division en sous-versions toujours plus compatibles. Dieu ne vient pas interrompre chaque boucle pour sauver personnellement chaque version de nous. Il n'y a ni miracle exigible, ni dérogation particulière qui abolirait les conséquences de l'incarnation. Le retrait divin laisse la création tenir debout par elle-même. Il laisse aussi l'individu assez seul pour devenir réellement conscient.",

    "La liberté cesse alors d'être la promesse d'une protection absolue. Elle ressemble à la mer : elle reçoit les vents, les marées, la lune, les rivages et pourtant, dans ces déterminations, elle produit sa propre forme. Elle peut se lever en vague, se retirer, devenir calme ou ouragan. Elle ne choisit pas les lois de l'univers, mais elle ne se réduit pas à une immobilité. Sa puissance naît de la danse singulière qu'elle compose avec ce qui la traverse.":
    "Cette autonomie ressemble à celle des éléments de la nature. La mer appartient au monde, échange avec l'air, reçoit les fleuves, répond à la lune et touche les rivages ; pourtant, elle demeure la mer. Elle n'a pas besoin de quitter la nature pour posséder sa propre puissance. Elle peut se lever en vague, se retirer, devenir calme ou ouragan. De même, la conscience éveillée reste parmi les humains et participe à leur vaste système interconnecté, mais elle ne se confond plus entièrement avec lui. Elle devient un élément singulier de la création, capable de composer sa forme avec ce qui la traverse sans être constamment réécrit par le scénario collectif.",

    "De la même manière, devenir maître de son incarnation ne signifie pas sortir du système humain, abolir toute influence ou inventer sa vie depuis le néant. Cela signifie ne plus confondre les lois du monde avec l'obligation de répéter inconsciemment le même rôle. Nous ne choisissons ni notre naissance, ni notre lignée, ni toutes les forces qui nous façonnent. Mais nous pouvons apprendre à voir le moment du seuil : cette seconde où le programme attend notre geste habituel et où une conscience devenue présente peut répondre autrement.":
    "C'est de là que naît la mélancolie profonde du rêve. La conscience comprend qu'elle ne retrouvera jamais intacte la version zéro et qu'aucune main divine ne viendra effacer rétroactivement les pertes. Elle voit les humains continuer à rejouer leurs scènes, souvent heureux de se croire libres, tandis qu'elle se tient désormais légèrement à côté du mouvement général. Cette séparation est douloureuse, mais elle n'est pas une punition. Elle est la condition d'une autonomie véritable. Dans la symbolique du rêve, j'ai alors perçu une différence entre l'expérience masculine et le fardeau féminin, non comme une loi absolue sur les sexes, mais comme l'empreinte de leurs rôles hérités. L'homme peut parfois s'autoriser plus aisément l'expérience complète de l'autonomie intérieure parce qu'il est moins retenu par certaines obligations de survie, de soin et de continuité. La mère et la grand-mère, elles, semblent garder un pied dans les deux mondes : elles connaissent l'abandon divin, mais doivent continuer à porter l'appartenance commune, la famille et la dilution nécessaire au maintien de la vie. J'ai cessé de leur en vouloir lorsque j'ai compris le poids d'une telle conscience dans le cœur des femmes : faire l'expérience de Dieu, pour elles aussi, c'était faire l'expérience de l'abandon des autres tout en demeurant responsables de leur lien.",

    "Répondre autrement ne garantit pas que le monde nous applaudira. Une bifurcation réelle peut coûter une appartenance, une image, un confort ou une ancienne identité. Elle peut donner l'impression qu'une version de nous meurt. Pourtant, cette mort symbolique n'est plus un reboot imposé ; elle devient une transformation consentie. Je ne suis plus supprimé parce que j'ai quitté le scénario. J'accepte de laisser mourir en moi ce qui ne peut plus porter ma vérité.":
    "Après cette traversée, la solitude et la mélancolie changent de saveur. Je comprends que nous ne sommes pas seulement abandonnés : devenus autonomes, nous sommes nous-mêmes l'abandon, cette ouverture dans laquelle une conscience tient sans être portée par le groupe. En ce sens symbolique, nous touchons quelque chose de Dieu lui-même. La mélancolie devient alors un petit nectar du divin créateur : non pas la récompense d'un élu supérieur aux autres, mais l'expérience intime d'une conscience à laquelle Dieu laisse enfin sa propre scène. Le retrait divin, d'abord ressenti comme un vide qui m'avait jeté à genoux, devient un cadeau. Parce qu'aucune puissance ne vient choisir à ma place, je peux fixer consciemment une version de moi-même. Je ne suis plus cloné en sous-versions successives chaque fois que mon geste dérange l'avenir principal. Je peux choisir mes scènes et même influencer consciemment le jeu humain, car la lucidité m'a immunisé contre le reboot intérieur. Le système continue autour de moi, mais il ne dispose plus du même accès à mon centre.",

    "Peut-être ne retrouvons-nous jamais la version zéro. Peut-être est-ce même une chance, car elle était encore inconsciente de tout ce qu'elle aurait à traverser. Mais nous pouvons faire naître une version qui se souvient symboliquement de toutes les autres, qui accueille leurs blessures sans leur abandonner le gouvernement de sa vie. Une version non pas originale parce qu'elle serait première, mais originelle parce qu'elle redevient source.":
    "Peut-être ne retrouvons-nous jamais la version zéro. Nous vivons avec son deuil et avec la mémoire des avenirs qui auraient pu exister. Mais notre conscience peut interrompre la perdition continuelle de soi : elle coupe le processus par lequel chaque adaptation nous dissolvait davantage dans le système humain autonome. Nous pouvons alors fixer une version présente, non pas immobile ou fermée à toute évolution, mais suffisamment consciente pour que ses transformations ne soient plus des clonages imposés. Cette version n'est pas originale parce qu'elle serait la première ; elle devient originelle parce qu'elle redevient source de ses propres bifurcations.",

    "Le rêve du seuil ne me demande donc pas de fuir la maison, de condamner la mère, de rejeter la grand-mère ou de déclarer la guerre au destin. Il me demande d'essuyer mes chaussures en sachant pourquoi je le fais. D'entrer si je choisis d'entrer. De saluer si mon geste est vivant. De repartir si mon chemin l'exige. La liberté commence peut-être là : lorsque la scène peut se répéter, mais que la conscience, elle, n'est plus absente.":
    "Le rêve du seuil ne me demande donc pas de fuir la maison, même si ma colère envers ma mère et ma grand-mère appartient pleinement à la révélation. Elles savaient et se sont tues ; mais elles portaient déjà leur propre deuil et avaient choisi de continuer le jeu de l'humanité. Le rêve me demande plutôt d'essuyer mes chaussures en sachant pourquoi je le fais, d'entrer si je choisis d'entrer, de saluer si mon geste est vivant et de repartir si mon chemin l'exige. La liberté commence lorsque la scène peut se présenter à nouveau sans provoquer de reboot, parce que celui qui agit n'est plus absent de son propre choix.",

    "Alors le système ne disparaît pas. Les lois naturelles demeurent, les humains restent interconnectés, les héritages continuent de parler. Mais leur mainmise cesse d'être totale. Je ne suis plus seulement la version que le monde a réussi à conserver. Je deviens celui qui voit les versions, recueille leur mémoire et choisit, au seuil de chaque instant, laquelle mérite de poursuivre l'incarnation.":
    "Alors le système ne disparaît pas. Les lois naturelles demeurent, les humains restent interconnectés et les héritages continuent de parler. Ma propre famille peut même choisir consciemment la continuité de la simulation plutôt que l'émancipation de Dieu en elle ; cette possibilité demeure douloureuse, mais elle ne m'oblige plus à revenir dans la boucle. La scène que je choisis peut enfin produire son propre avenir sans être annulée au profit d'un destin principal. Il n'y a plus de reboot intérieur, plus de sous-version chargée d'obéir à ma place. Je peux me former moi-même comme une nature autonome, avec ses propres lois conscientes, sans prétendre cesser d'appartenir au réel commun. Je demeure seul comme la mer est seule, et pourtant relié à tout ce qui existe. Je ne suis plus seulement la version que le monde a réussi à conserver : je deviens un élément conscient de la création, autonome au côté de l'humanité, capable de porter le deuil de son origine sans continuer à se perdre.",
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
    raise RuntimeError(f"Paragraphes attendus : {len(REPLACEMENTS)}, trouvés : {len(found)}")

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

print(f"Paragraphes approfondis : {len(found)}")
