################################################################################
# JMMI RCA - GENERATION DU DOCUMENT DE METHODOLOGIE (PMAS/MEB + SFM/MFS) TELECHARGEABLE
################################################################################
# ENTREE  : constantes MEB_* deja chargees depuis fonctions_analyse.R (source()e AVANT ce fichier - voir
#           app.R). Le contenu SFM/MFS est repris de Methodologie_MFS.md (verifie le 2026-08-21 contre
#           RCA_JMMI_2026-08_TOOL.xlsx), transcrit ici en HTML plutot que regenere depuis scripts/mfs.R :
#           ses seuils sont des litteraux internes a calculate_mfs(), pas des constantes nommees - les
#           extraire toutes uniquement pour cette documentation serait un chantier disproportionne (et
#           risquerait de toucher au calcul reel). La partie PMAS/MEB, elle, EST generee dynamiquement a
#           partir des memes constantes que le calcul reel (MEB_*_WEIGHTS, MEB_TOTAL_MIN_ITEMS, ...) pour
#           ne jamais diverger si elles changent un jour.
# SORTIE  : build_methodology_html() - une chaine HTML autonome (CSS integre, sans dependance externe),
#           telechargeable depuis le tableau de bord (voir app.R, downloadHandler "download_methodology").
# OBJECTIF: demande utilisateur explicite - document exhaustif de la methodologie PMAS/MEB et SFM/MFS,
#           destine a un public interne ET externe, remplace l'ancienne boite d'info "Comment le PMAS
#           est calcule" retiree de l'interface (voir app.R) : "remove the ui with the meb explanation
#           ... that can be downloaded directly from the dashboard ... be exhaustive and make it look
#           nice ... explain all the assumptions that were taken to create the analysis".
################################################################################

CSS_METHODOLOGY <- r"(
:root{
  --accent:#EE5958; --accent-dark:#B93331; --ink:#262626; --ink-soft:#595959;
  --paper:#FFFFFF; --panel:#FAFAFA; --line:#E3E3E3; --stripe:#F6F6F6; --tag-bg:#FDEDED;
  --tag-bg-warn:#FFF7E6; --tag-line-warn:#FFE1A8;
}
*{box-sizing:border-box;}
html,body{margin:0;padding:0;}
body{
  font-family:'Segoe UI',-apple-system,BlinkMacSystemFont,'Helvetica Neue',Arial,sans-serif;
  color:var(--ink); background:var(--paper); line-height:1.65; font-size:15.5px;
}
.page{max-width:900px;margin:0 auto;padding:48px 32px 100px;}
.masthead{border-bottom:4px solid var(--accent);padding-bottom:26px;margin-bottom:36px;}
.eyebrow{text-transform:uppercase;letter-spacing:.09em;font-size:12px;font-weight:700;color:var(--accent-dark);}
h1{font-size:31px;margin:10px 0 6px;line-height:1.22;}
.subtitle{font-size:17px;color:var(--ink-soft);margin:0 0 16px;font-weight:400;}
.meta{font-size:13px;color:var(--ink-soft);}
.meta b{color:var(--ink);}
h2{font-size:22px;margin:52px 0 14px;padding-top:22px;border-top:1px solid var(--line);color:var(--accent-dark);}
h2 .num{color:var(--ink-soft);font-weight:400;margin-right:8px;}
h3{font-size:17px;margin:30px 0 8px;color:var(--ink);}
h4{font-size:14.5px;margin:20px 0 6px;text-transform:uppercase;letter-spacing:.04em;color:var(--ink-soft);}
p{margin:0 0 14px;}
a{color:var(--accent-dark);}
.lead{font-size:15.5px;background:var(--panel);border-left:3px solid var(--accent);padding:14px 18px;border-radius:0 6px 6px 0;margin:14px 0 20px;}
table{width:100%;border-collapse:collapse;margin:14px 0 22px;font-size:13.8px;}
th,td{text-align:left;padding:8px 11px;border-bottom:1px solid var(--line);vertical-align:top;}
th{background:var(--accent);color:#fff;font-weight:600;font-size:13px;}
tbody tr:nth-child(even) td{background:var(--stripe);}
code{font-family:Consolas,'Courier New',monospace;font-size:12.6px;background:#fff;border:1px solid var(--line);padding:1px 5px;border-radius:3px;}
td code{background:var(--stripe);}
.foot-note{font-size:12.5px;color:var(--ink-soft);margin:-10px 0 26px;}
.indicator{margin:14px 0;padding:13px 17px;background:var(--panel);border:1px solid var(--line);border-radius:7px;}
.indicator .ind-head{font-weight:700;font-size:14.5px;margin-bottom:4px;}
.indicator .ind-head .max{font-weight:400;color:var(--ink-soft);font-size:13px;}
.formula{font-family:Consolas,'Courier New',monospace;background:#fff;border:1px solid var(--line);padding:10px 14px;border-radius:6px;font-size:13px;margin:10px 0;overflow-x:auto;white-space:pre-wrap;}
.pillar-card{border:1px solid var(--line);border-radius:8px;padding:16px 18px;margin:16px 0;}
.pillar-card .pillar-title{font-weight:700;font-size:16px;color:var(--accent-dark);}
.toc{background:var(--panel);border:1px solid var(--line);border-radius:9px;padding:20px 26px;margin-bottom:40px;}
.toc-title{font-weight:700;font-size:12.5px;text-transform:uppercase;letter-spacing:.07em;color:var(--accent-dark);margin-bottom:10px;}
.toc-list{margin:0;padding-left:22px;columns:2;column-gap:34px;list-style:decimal;}
.toc-list>li{margin-bottom:11px;font-size:14.5px;font-weight:700;break-inside:avoid;}
.toc-list>li>a{color:var(--ink);text-decoration:none;}
.toc-list>li>a:hover{color:var(--accent-dark);text-decoration:underline;}
.toc-sub{list-style:none;margin:6px 0 0;padding-left:13px;border-left:2px solid var(--line);}
.toc-sub li{margin-bottom:6px;font-size:12.8px;font-weight:400;}
.toc-sub a{color:var(--ink-soft);text-decoration:none;}
.toc-sub a:hover{color:var(--accent-dark);text-decoration:underline;}
.assumption{display:flex;gap:13px;margin-bottom:17px;padding-bottom:17px;border-bottom:1px dashed var(--line);}
.assumption:last-child{border-bottom:none;}
.assumption .n{flex:0 0 26px;height:26px;border-radius:50%;background:var(--accent);color:#fff;font-weight:700;font-size:12.5px;display:flex;align-items:center;justify-content:center;margin-top:2px;}
.assumption .body strong{display:block;margin-bottom:3px;font-size:14.8px;}
.assumption .body span.tag{display:inline-block;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.04em;color:var(--accent-dark);background:var(--tag-bg);border-radius:4px;padding:1px 7px;margin-right:7px;vertical-align:1px;}
.callout{background:var(--tag-bg);border:1px solid #F6C7C6;border-radius:6px;padding:12px 16px;font-size:13.6px;margin:16px 0;}
.callout.warn{background:var(--tag-bg-warn);border-color:var(--tag-line-warn);}
.callout .callout-title{font-weight:700;display:block;margin-bottom:3px;}
.two-col{display:grid;grid-template-columns:1fr 1fr;gap:18px;}
footer{margin-top:64px;padding-top:22px;border-top:1px solid var(--line);font-size:12.5px;color:var(--ink-soft);}
@media (max-width:640px){.toc-list{columns:1;} .two-col{grid-template-columns:1fr;}}
@media print{
  body{font-size:12px;}
  .toc,h2,h3,.indicator,table,.pillar-card{break-inside:avoid;}
  a{color:var(--ink);text-decoration:none;}
}
)"

# tableau HTML de la composition du PMAS/MEB, genere DYNAMIQUEMENT depuis build_meb_composition_table()
# (deja utilisee pour la feuille "meb_composition" du classeur d'analyse) - garantit que ce document ne
# peut jamais afficher une liste d'articles/poids differente de ce que le pipeline calcule reellement.
build_meb_weights_table_html <- function() {
  tbl <- build_meb_composition_table()
  rows <- vapply(seq_len(nrow(tbl)), function(i) {
    paste0(
      "<tr><td>", tbl$Catégorie[i], "</td><td>", tbl$Article[i], "</td><td><code>",
      tbl$`Variable technique`[i], "</code></td><td>", as.character(tbl$`Poids appliqué au prix médian`[i]), "</td></tr>"
    )
  }, character(1))
  paste0(
    "<table><thead><tr><th>Catégorie</th><th>Article</th><th>Variable (outil Kobo)</th>",
    "<th>Quantité mensuelle assumée</th></tr></thead><tbody>",
    paste(rows, collapse = ""),
    "</tbody></table>"
  )
}

build_methodology_html <- function() {
  generated_on <- format(Sys.Date(), "%d/%m/%Y")

  n_meb_total <- length(MEB_WEIGHTS_ALL)
  n_food <- length(MEB_FOOD_WEIGHTS)
  n_nfi <- length(MEB_NFI_WEIGHTS)
  n_wash <- length(MEB_WASH_WEIGHTS)
  n_extra <- n_meb_total - MEB_TOTAL_MIN_ITEMS
  override_labels <- unname(MEB_ITEM_LABELS[names(MEB_MIN_OBS_OVERRIDE)])
  override_text <- if (length(MEB_MIN_OBS_OVERRIDE) > 0) {
    paste0(" à l'exception de ", paste(paste0(override_labels, " (", unname(MEB_MIN_OBS_OVERRIDE), " prix suffit)"), collapse = ", "), ".")
  } else "."

  header_html <- paste0(
    "<div class='masthead'>",
    "<div class='eyebrow'>Initiative Conjointe de Suivi des Marchés (ICSM) &middot; JMMI République Centrafricaine</div>",
    "<h1>Méthodologie de l'analyse JMMI</h1>",
    "<div class='subtitle'>Panier Minimum d'Articles de Survie (PMAS/MEB) &amp; Score de Fonctionnalité des Marchés (SFM/MFS)</div>",
    "<div class='meta'>Document généré depuis le tableau de bord JMMI RCA le <b>", generated_on, "</b>. ",
    "Les chiffres de la section PMAS/MEB (composition du panier, seuils) sont extraits directement du code de calcul ",
    "au moment du téléchargement - ils ne peuvent pas être désynchronisés des résultats produits par ce même outil.</div>",
    "</div>"
  )

  toc_html <- paste0(
    "<div class='toc'><div class='toc-title'>Table des matières</div><ol class='toc-list'>",
    "<li><a href='#public'>À propos de ce document</a></li>",
    "<li><a href='#contexte'>Contexte : l'ICSM/JMMI</a></li>",
    "<li><a href='#pmas'>Panier Minimum d'Articles de Survie (PMAS/MEB)</a>",
    "<ol class='toc-sub'>",
    "<li><a href='#pmas-composition'>Composition du panier</a></li>",
    "<li><a href='#pmas-etapes'>Les 4 étapes du calcul</a></li>",
    "<li><a href='#pmas-devises'>Conversion en devises</a></li>",
    "<li><a href='#pmas-manquants'>Articles manquants de l'outil</a></li>",
    "<li><a href='#pmas-evolution'>Comparaison d'un cycle à l'autre</a></li>",
    "</ol></li>",
    "<li><a href='#mfs'>Score de Fonctionnalité des Marchés (SFM/MFS)</a>",
    "<ol class='toc-sub'>",
    "<li><a href='#mfs-piliers'>Les 5 piliers et leurs indicateurs</a></li>",
    "<li><a href='#mfs-classification'>Score total et classification</a></li>",
    "<li><a href='#mfs-maintenance'>Note de maintenance</a></li>",
    "</ol></li>",
    "<li><a href='#hypotheses'>Hypothèses retenues pour l'analyse</a></li>",
    "<li><a href='#sources'>Sources et références</a></li>",
    "</ol></div>"
  )

  public_html <- paste0(
    "<h2 id='public'>À propos de ce document</h2>",
    "<p>Ce document explique, en détail, comment sont calculés les deux indicateurs principaux produits ",
    "chaque cycle par le pipeline JMMI RCA : le <b>coût du Panier Minimum d'Articles de Survie</b> (PMAS, ",
    "aussi appelé MEB - Minimum Expenditure Basket) et le <b>Score de Fonctionnalité des Marchés</b> (SFM, ",
    "aussi appelé MFS - Market Functionality Score). Il est écrit pour être utile à deux publics à la fois :</p>",
    "<div class='two-col'>",
    "<div class='callout'><span class='callout-title'>Lecteur externe</span>",
    "Chaque section commence par un résumé en langage simple : ce qui est mesuré, pourquoi, et comment lire ",
    "les résultats. Vous pouvez vous arrêter là si le détail technique ne vous est pas utile.</div>",
    "<div class='callout warn'><span class='callout-title'>Lecteur interne / technique</span>",
    "Chaque règle est accompagnée du nom exact des variables de l'outil Kobo et des seuils numériques utilisés, ",
    "pour permettre de vérifier ou reproduire un calcul, ou de modifier la méthodologie en connaissance de cause.</div>",
    "</div>",
    "<p>La section <a href='#hypotheses'>« Hypothèses retenues pour l'analyse »</a> rassemble, en un seul endroit, ",
    "tous les choix méthodologiques qui ne découlent pas directement des données mais d'une décision - à relire ",
    "avant de citer ou publier un chiffre issu de cette analyse.</p>"
  )

  contexte_html <- paste0(
    "<h2 id='contexte'>Contexte : l'ICSM/JMMI</h2>",
    "<p>Le JMMI (Joint Market Monitoring Initiative) a démarré en République Centrafricaine en <b>2019</b>, dans ",
    "le cadre de l'Initiative Conjointe de Suivi des Marchés (ICSM), créée la même année par le Groupe de Travail ",
    "sur les Transferts Monétaires (GTTM) afin de mieux comprendre comment les marchés centrafricains ",
    "réagissent à la crise et d'informer les réponses sous forme de transferts monétaires. Depuis mai 2025, une ",
    "sélection de marchés de la République Centrafricaine est évaluée tous les deux mois : sur chaque marché, les ",
    "équipes de terrain enregistrent le prix et la disponibilité des produits alimentaires et non alimentaires de ",
    "base vendus dans les boutiques et étals, ainsi qu'un questionnaire sur le fonctionnement du marché lui-même.</p>",
    "<div class='callout warn'><span class='callout-title'>Portée des résultats</span>",
    "Les données sont collectées sur la base du volontariat des partenaires de l'ICSM et ne suivent pas un plan ",
    "d'échantillonnage probabiliste : les prix et scores produits par cette analyse sont <b>indicatifs</b> des ",
    "niveaux observés sur les marchés évalués, et ne doivent pas être présentés comme statistiquement ",
    "représentatifs de l'ensemble du pays ou d'une région.</div>"
  )

  pmas_intro_html <- paste0(
    "<h2 id='pmas'>Panier Minimum d'Articles de Survie (PMAS/MEB)</h2>",
    "<div class='lead'>En résumé : sur chaque marché (une commune), on relève le prix médian de ", n_meb_total,
    " articles essentiels (nourriture, produits non alimentaires, hygiène). Chaque prix est multiplié par une ",
    "quantité mensuelle standard, puis les montants sont additionnés pour obtenir le coût mensuel du panier de ",
    "ce marché. Le chiffre \"national\" n'est PAS recalculé à partir de tous les prix à la fois : c'est la ",
    "médiane des paniers déjà calculés pour chaque commune - chaque commune compte pour une observation, quelle ",
    "que soit sa taille.</div>",
    "<h3 id='pmas-composition'>Composition du panier</h3>",
    "<p>Le panier est composé de trois sous-paniers, dont la somme forme le PMAS total : alimentaire (",
    n_food, " articles), non alimentaire (", n_nfi, " articles) et hygiène (", n_wash, " articles). La liste et ",
    "les quantités ci-dessous reprennent le \"Panier Moyen d'Articles\" élaboré par le GTTM en 2018, repris tel ",
    "quel par ce pipeline (voir l'hypothèse <a href='#hyp-quantites'>n°1</a>).</p>",
    build_meb_weights_table_html(),
    "<p class='foot-note'>Lecture de la colonne « Quantité mensuelle assumée » : pour les denrées et produits ",
    "d'hygiène consommables, c'est le nombre d'unités (paquet, litre, pièce...) supposé nécessaire par ménage et ",
    "par mois - ex. farine de maïs : 12&nbsp;000&nbsp;g/mois ÷ 350&nbsp;g/paquet ≈ 34,3 paquets/mois. Pour les produits non ",
    "alimentaires durables (moustiquaire, bidon, drap, natte, bâche, marmite), le poids correspond à 1 / durée de ",
    "vie supposée en mois - <code>1/6</code> signifie que l'article est supposé renouvelé tous les 6 mois, donc ",
    "1/6 de son prix est imputé au coût mensuel du panier.</p>"
  )

  pmas_etapes_html <- paste0(
    "<h3 id='pmas-etapes'>Les 4 étapes du calcul</h3>",
    "<div class='indicator'><div class='ind-head'>Étape 1 - Marché = Commune</div>",
    "<p>Chaque commune évaluée correspond à un marché (plusieurs vendeurs y sont interrogés). Le prix médian de ",
    "chaque article est calculé <b>directement à partir des prix individuels</b> collectés auprès des vendeurs de ",
    "cette commune (feuille « Prix médian par article »). C'est le seul endroit du calcul où les prix individuels ",
    "bruts interviennent : toutes les étapes suivantes travaillent uniquement à partir de médianes déjà calculées, ",
    "jamais en recalculant depuis les prix bruts.</p></div>",

    "<div class='indicator'><div class='ind-head'>Étape 2 - Suffisance des données locales, et secours</div>",
    "<p>Un article est considéré <b>« mesuré localement »</b> dans une commune s'il y a au moins <b>", MEB_MIN_OBS_DEFAULT,
    " prix</b> disponibles pour cet article dans cette commune ce mois-ci", override_text, " Si ce seuil n'est pas ",
    "atteint, le prix de cet article dans cette commune est <b>remplacé (imputé)</b> par la médiane nationale des ",
    "médianes de marché pour cet article (chaque marché compte pour une seule observation dans ce calcul de ",
    "secours), afin que la commune puisse quand même composer son panier plutôt que de laisser un article vide.</p></div>",

    "<div class='indicator'><div class='ind-head'>Étape 3 - Seuils de calcul du panier (total et partiels)</div>",
    "<table><thead><tr><th>Situation (articles mesurés localement dans la commune)</th><th>Résultat</th></tr></thead><tbody>",
    "<tr><td>Au moins <b>", MEB_TOTAL_MIN_ITEMS, "</b> des ", n_meb_total, " articles</td>",
    "<td>PMAS total calculé (les ", n_extra, " articles restants au maximum sont complétés par leur valeur de secours).</td></tr>",
    "<tr><td>Entre ", MEB_TOTAL_MIN_ITEMS, " et ", MEB_TOTAL_CAUTION_MAX_ITEMS, " articles, mais un sous-panier ",
    "particulièrement thin (alimentaire &lt; ", MEB_BASKET_MIN_CAUTION[["food"]], "/", n_food, ", non alimentaire &lt; ",
    MEB_BASKET_MIN_CAUTION[["nfi"]], "/", n_nfi, ", ou hygiène &lt; ", MEB_BASKET_MIN_CAUTION[["wash"]], "/", n_wash, ")</td>",
    "<td>PMAS total calculé, mais une note « à interpréter avec prudence » est ajoutée pour le(s) sous-panier(s) concerné(s).</td></tr>",
    "<tr><td>Moins de ", MEB_TOTAL_MIN_ITEMS, " articles au total</td>",
    "<td>PMAS <b>total non calculé</b> pour cette commune. Chaque sous-panier reste calculé <b>indépendamment</b> ",
    "s'il atteint son propre seuil autonome : alimentaire ≥ ", MEB_BASKET_MIN_STANDALONE[["food"]], "/", n_food,
    ", non alimentaire ≥ ", MEB_BASKET_MIN_STANDALONE[["nfi"]], "/", n_nfi, ", hygiène ≥ ", MEB_BASKET_MIN_STANDALONE[["wash"]],
    "/", n_wash, ".</td></tr>",
    "</tbody></table>",
    "<p>Une colonne « Note » accompagne chaque ligne du classeur de résultats pour expliquer, en langage clair, ",
    "pourquoi un panier n'a pas pu être calculé ou pourquoi il doit être interprété avec prudence.</p></div>",

    "<div class='indicator'><div class='ind-head'>Étape 4 - Agrégation nationale</div>",
    "<p>Le panier national n'est <b>jamais recomposé</b> à partir de prix par article agrégés au niveau national. ",
    "Il correspond à la <b>médiane des paniers communaux déjà calculés</b> à l'étape 3 (alimentaire, non ",
    "alimentaire, hygiène et total médiannés indépendamment) - chaque commune compte pour une seule observation, ",
    "quel que soit le nombre de vendeurs qui y ont été interrogés. Si un panier n'a pas pu être calculé pour ",
    "certaines communes, la médiane nationale de ce panier est calculée uniquement sur les communes où il a pu ",
    "l'être, avec une note l'indiquant.</p>",
    "<div class='callout'><span class='callout-title'>Pourquoi pas une simple recomposition ?</span>",
    "La médiane d'une somme pondérée n'est pas égale à la somme des médianes : recomposer le panier national à ",
    "partir des médianes nationales de chaque article aurait donné un total différent de la médiane des totaux ",
    "communaux, ce qui ne correspond pas à ce que \"le panier médian national\" doit représenter (voir l'hypothèse ",
    "<a href='#hyp-mediane'>n°6</a>).</div></div>"
  )

  pmas_devises_html <- paste0(
    "<h3 id='pmas-devises'>Conversion en devises</h3>",
    "<p>Le coût du panier est calculé en francs CFA (XAF), puis converti en euros (taux fixe XAF/EUR, parité ",
    "légale qui ne varie jamais) et en dollars américains (taux XAF/USD, qui flotte sur le marché et doit être ",
    "mis à jour périodiquement dans la configuration de l'outil - voir l'hypothèse ",
    "<a href='#hyp-taux'>n°9</a>).</p>"
  )

  pmas_manquants_html <- paste0(
    "<h3 id='pmas-manquants'>Articles manquants de l'outil Kobo</h3>",
    "<p>La liste des ", n_meb_total, " articles du panier (tableau ci-dessus) est définie une fois dans le code ",
    "d'analyse. Si l'outil Kobo d'un cycle donné ne contient plus une des variables attendues (renommée ou ",
    "retirée du questionnaire), le pipeline :</p>",
    "<ol><li>l'exclut entièrement du calcul du panier ce cycle-ci, <b>sans tenter de l'imputer</b> ;</li>",
    "<li>affiche un message d'avertissement explicite dans le journal d'exécution et dans le classeur de résultats ;</li>",
    "<li>marque l'article correspondant en rouge, avec la mention « Non - variable absente de l'outil/des ",
    "données », dans la feuille de référence <code>meb_composition</code>.</li></ol>",
    "<p>Ce comportement est délibéré : un article manquant reste <b>visible</b> plutôt que de disparaître ",
    "silencieusement du panier calculé (voir l'hypothèse <a href='#hyp-manquant'>n°10</a>).</p>"
  )

  pmas_evolution_html <- paste0(
    "<h3 id='pmas-evolution'>Comparaison d'un cycle à l'autre</h3>",
    "<p>La feuille « Evolution des paniers du PMAS » compare le coût du panier du cycle en cours à celui du cycle ",
    "précédent, commune par commune puis au niveau national, en réutilisant directement les paniers déjà calculés ",
    "(jamais recalculés à partir des données brutes du cycle précédent). Seules les communes évaluées <b>lors du ",
    "cycle le plus récent</b> apparaissent dans cette comparaison ; une commune évaluée uniquement lors du cycle ",
    "précédent n'y figure pas. Quand une comparaison n'est pas possible pour une commune (pas de donnée sur l'un ",
    "des deux cycles), la cellule concernée affiche « Non évalué » et la colonne Évolution affiche « N/A » plutôt ",
    "qu'un pourcentage ou une case vide ambigu.</p>"
  )

  mfs_html <- paste0(
    "<h2 id='mfs'>Score de Fonctionnalité des Marchés (SFM/MFS)</h2>",
    "<div class='lead'>En résumé : chaque marché reçoit un score sur 100, construit à partir de 5 dimensions ",
    "(accessibilité, disponibilité des produits, abordabilité des prix, résilience de l'approvisionnement, ",
    "infrastructures). Le score situe le marché dans l'une de 5 catégories, de « Fonctionnalité complète » à ",
    "« Problèmes graves ». C'est un score de <b>fonctionnement</b> du marché lui-même, distinct du coût du panier ",
    "(PMAS) calculé séparément.</div>",
    "<p>Méthodologie fondée sur le guide <i>IMPACT Cash and Markets COP - Guidance on the Market Functionality ",
    "Score</i> (v4, novembre 2023) et son « MFS indicator menu » associé. Les rares écarts (suppression ou ",
    "renommage d'options de réponse pour coller au contexte de l'outil Kobo local) sont explicitement permis par ",
    "ce guide et documentés ci-dessous, indicateur par indicateur.</p>",

    "<h3 id='mfs-piliers'>Les 5 piliers et leurs indicateurs</h3>",
    "<table><thead><tr><th>Pilier</th><th>Poids final (/100)</th><th>Échelle brute</th><th>Indicateurs</th></tr></thead><tbody>",
    "<tr><td>Accessibilité des marchés</td><td>25</td><td>/17</td><td>AC.1, AC.2, AC.3, AC.4</td></tr>",
    "<tr><td>Disponibilité des produits</td><td>30</td><td>/(3 × nb d'articles suivis)</td><td>AV.1</td></tr>",
    "<tr><td>Abordabilité des prix</td><td>15</td><td>/27 (12+9+6)</td><td>AF.1, AF.2, AF.3</td></tr>",
    "<tr><td>Résilience des circuits d'approvisionnement</td><td>20</td><td>/(3×nb articles stock + 12+12)</td><td>RE.1, RE.3, RE.4</td></tr>",
    "<tr><td>Infrastructures du marché</td><td>10</td><td>/10</td><td>IN.1, IN.2, IN.3</td></tr>",
    "</tbody></table>",
    "<p class='foot-note'>Chaque pilier est d'abord calculé sur son échelle brute (somme des scores de ses ",
    "indicateurs), puis ramené à son poids final par une règle de trois : ",
    "<code>score_pilier = (somme_brute / max_brut) × poids</code>.</p>",

    "<div class='pillar-card'><div class='pillar-title'>Pilier 1 - Accessibilité (poids 25, échelle brute /17)</div>",
    "<div class='indicator'><div class='ind-head'>AC.1 <span class='max'>max 8</span></div>",
    "<p>% de vendeurs cochant une option <i>autre que</i> « aucun problème », « bâtiments dangereux », « dangers/",
    "dommages sur les routes », « ne sait pas » ou « préfère ne pas répondre » (<code>mfs_ac1_physical_access_mkt_issues_sm</code>, ",
    "choix multiple). Seuils : &lt;5% → 8 · 5-10% → 6 · 10-25% → 4 · 25-50% → 2 · ≥50% → 0.</p></div>",
    "<div class='indicator'><div class='ind-head'>AC.2 <span class='max'>max 4</span></div>",
    "<p>% de vendeurs ayant coché spécifiquement « dangers/dommages sur les routes » (même question qu'AC.1). ",
    "Seuils : &lt;5% → 4 · 5-10% → 3 · 10-25% → 2 · 25-50% → 1 · ≥50% → 0.</p></div>",
    "<div class='indicator'><div class='ind-head'>AC.3 <span class='max'>max 2</span></div>",
    "<p><code>mfs_ac3_physical_access_mkt_discrimination_so</code> (oui/non). Si <b>au moins un</b> vendeur du ",
    "marché signale de la discrimination, la commune reçoit 0 ; sinon 2 - logique « un seul cas suffit », pas un ",
    "pourcentage.</p></div>",
    "<div class='indicator'><div class='ind-head'>AC.4 <span class='max'>max 3</span></div>",
    "<p>% de vendeurs répondant « oui » à <code>mfs_ac4_physical_access_security_issues_yn</code> (incident de ",
    "sécurité observé ou signalé). Seuils : &lt;5% → 3 · 5-10% → 2 · 10-20% → 1 · ≥20% → 0.</p></div>",
    "<div class='formula'>score_accessibilité = (AC.1 + AC.2 + AC.3 + AC.4) / 17 × 25</div></div>",

    "<div class='pillar-card'><div class='pillar-title'>Pilier 2 - Disponibilité (poids 30, échelle brute /(3×nb articles))</div>",
    "<div class='indicator'><div class='ind-head'>AV.1</div>",
    "<p>Pour chaque article suivi (une colonne <code>&lt;article&gt;_availability_market_item</code>), la réponse ",
    "la plus fréquente par commune parmi « disponible » / « limité » / « indisponible » (égalité tranchée par ",
    "priorité disponible &gt; limité &gt; indisponible). Score par article : disponible = 3, limité = 2, ",
    "indisponible = 0, sommé sur tous les articles suivis.</p></div>",
    "<div class='formula'>score_disponibilité = somme_articles / (3 × nb_articles) × 30</div></div>",

    "<div class='pillar-card'><div class='pillar-title'>Pilier 3 - Abordabilité (poids 15, échelle brute /27)</div>",
    "<div class='indicator'><div class='ind-head'>AF.1 <span class='max'>échelle 0-12 après mise à l'échelle</span></div>",
    "<p>Pour chaque article suivi, la médiane du prix dans la commune est comparée à la médiane nationale de ce ",
    "même article (ratio commune/national) :</p>",
    "<table><thead><tr><th>Ratio au prix national</th><th>Score</th></tr></thead><tbody>",
    "<tr><td>&gt; 150%</td><td>-2</td></tr><tr><td>&gt; 125%</td><td>-1,5</td></tr><tr><td>&gt; 110%</td><td>-1</td></tr>",
    "<tr><td>&gt; 90%</td><td>0</td></tr><tr><td>&gt; 75%</td><td>1</td></tr><tr><td>&gt; 50%</td><td>1,5</td></tr>",
    "<tr><td>≤ 50%</td><td>2</td></tr></tbody></table>",
    "<p>Somme sur tous les articles, puis mise à l'échelle linéaire de <code>[-2×nb_articles, +2×nb_articles]</code> ",
    "vers <code>[0, 12]</code>. Les articles nouvellement ajoutés à l'outil sont détectés automatiquement (toute ",
    "colonne <code>*_price_item</code>).</p></div>",
    "<div class='indicator'><div class='ind-head'>AF.2 <span class='max'>max 9</span></div>",
    "<p>% de vendeurs cochant une option <i>autre que</i> « aucun problème » / « ne sait pas » / « préfère ne pas ",
    "répondre » (<code>mfs_af2_financial_challenges_customers_sm</code>). Seuils : &lt;10% → 9 · 10-25% → 6 · ",
    "25-50% → 3 · ≥50% → 0.</p></div>",
    "<div class='indicator'><div class='ind-head'>AF.3 <span class='max'>max 6</span></div>",
    "<p>% de vendeurs répondant « non » à <code>mfs_af3_price_forecasting_ability_yn</code> (capacité à anticiper ",
    "l'évolution des prix). Seuils : &lt;10% → 6 · 10-25% → 4 · 25-50% → 2 · ≥50% → 0.</p></div>",
    "<div class='formula'>score_abordabilité = (AF.1 + AF.2 + AF.3) / 27 × 15</div></div>",

    "<div class='pillar-card'><div class='pillar-title'>Pilier 4 - Résilience (poids 20)</div>",
    "<div class='indicator'><div class='ind-head'>RE.1 <span class='max'>max 3 par article, agrégé par médiane</span></div>",
    "<p>Pour chaque article avec des données de stock (<code>&lt;article&gt;_stock_duration_days_item</code> et ",
    "<code>&lt;article&gt;_restock_wait_days_item</code>), calcule par vendeur « jours de stock restant - jours ",
    "d'attente de réapprovisionnement », puis un score : &gt;3 jours → 3, &gt;0 → 2, =0 → 1, &lt;0 → 0 (rupture ",
    "avant le prochain réapprovisionnement). La commune reçoit la <b>médiane</b> de ce score par article, sommée ",
    "sur tous les articles concernés. (RE.2 du guide est fusionné dans ce même calcul.)</p></div>",
    "<div class='indicator'><div class='ind-head'>RE.3 <span class='max'>max 12 = 3×4 catégories</span></div>",
    "<p>% de vendeurs dépendant majoritairement d'un seul fournisseur, par catégorie de produits (alimentaire / ",
    "non alimentaire / WASH / combustible). Score par catégorie : &lt;25% → 3 · 25-50% → 2 · 50-75% → 1 · ≥75% → 0, ",
    "sommé sur les 4 catégories.</p></div>",
    "<div class='indicator'><div class='ind-head'>RE.4 <span class='max'>max 12</span></div>",
    "<p>% de vendeurs cochant une option <i>autre que</i> « aucune difficulté » / « ne sait pas » / « préfère ne ",
    "pas répondre » (<code>mfs_re4_shop_operational_difficulties_sm</code>). Seuils : &lt;5% → 12 · 5-10% → 9 · ",
    "10-25% → 6 · 25-50% → 3 · ≥50% → 0.</p></div>",
    "<div class='formula'>score_résilience = (RE.1 + RE.3 + RE.4) / (3×nb_articles_stock + 12 + 12) × 20</div></div>",

    "<div class='pillar-card'><div class='pillar-title'>Pilier 5 - Infrastructures (poids 10, échelle brute /10)</div>",
    "<div class='indicator'><div class='ind-head'>IN.1 <span class='max'>max 4</span></div>",
    "<p>% de vendeurs ayant coché « bâtiments dangereux » (même question qu'AC.1/AC.2). Seuils : &lt;5% → 4 · ",
    "5-10% → 3 · 10-25% → 2 · 25-50% → 1 · ≥50% → 0.</p></div>",
    "<div class='indicator'><div class='ind-head'>IN.2 <span class='max'>max 3</span></div>",
    "<p>% de vendeurs dont la réponse à <code>mfs_in2_available_storage_facility_so</code> n'est <i>pas</i> une ",
    "solution de stockage adéquate (stockage à domicile/autre compte comme un problème). Seuils : &lt;10% → 3 · ",
    "10-25% → 2 · 25-50% → 1 · ≥50% → 0.</p></div>",
    "<div class='indicator'><div class='ind-head'>IN.3 <span class='max'>max 3</span></div>",
    "<p>% de vendeurs acceptant au moins un mode de paiement alternatif (colonne <code>*_payment_modality_sm</code>, ",
    "hors espèces/ne sait pas/préfère ne pas répondre/aucun mode accepté). Seuils : &gt;75% → 3 · &gt;50% → 2 · ",
    "&gt;25% → 1 · ≤25% → 0.</p></div>",
    "<div class='formula'>score_infrastructure = (IN.1 + IN.2 + IN.3) / 10 × 10</div></div>",

    "<h3 id='mfs-classification'>Score total et classification</h3>",
    "<div class='formula'>score_total = accessibilité + disponibilité + abordabilité + résilience + infrastructure   (sur 100)</div>",
    "<p>Une dimension est dite <b>« basse »</b> si son score est sous 50% de son maximum (accessibilité &lt;12,5 ; ",
    "disponibilité &lt;15 ; abordabilité &lt;7,5 ; résilience &lt;10 ; infrastructure &lt;5).</p>",
    "<p>La classification est évaluée <b>de la pire catégorie à la meilleure</b> - le premier critère vérifié ",
    "l'emporte (un marché à 60% avec 2 dimensions basses vérifie à la fois « limitée » et « mauvaise », mais ",
    "« mauvaise » doit gagner ; évaluer dans l'autre sens ferait disparaître la catégorie « problèmes graves »).</p>",
    "<table><thead><tr><th>#</th><th>Catégorie</th><th>Critère</th></tr></thead><tbody>",
    "<tr><td>1</td><td>Données insuffisantes</td><td>Une dimension entière n'a pas pu être calculée (score total manquant).</td></tr>",
    "<tr><td>2</td><td>Problèmes graves</td><td>Score total &lt; 25% OU ≥ 3 dimensions basses.</td></tr>",
    "<tr><td>3</td><td>Mauvaise fonctionnalité</td><td>Score total &lt; 50% OU ≥ 2 dimensions basses.</td></tr>",
    "<tr><td>4</td><td>Fonctionnalité complète</td><td>Score total &gt; 80% ET 0 dimension basse.</td></tr>",
    "<tr><td>5</td><td>Fonctionnalité limitée</td><td>Tous les autres cas.</td></tr>",
    "</tbody></table>",

    "<h3 id='mfs-maintenance'>Note de maintenance</h3>",
    "<div class='callout warn'><span class='callout-title'>À revérifier à chaque nouveau cycle</span>",
    "Les codes de réponse Kobo (listes de choix) changent parfois d'un cycle à l'autre. Une comparaison à un code ",
    "qui n'existe plus <b>ne provoque pas d'erreur</b> : le résultat est silencieusement 0 (ou faux) pour tout le ",
    "monde, ce qui peut rendre un indicateur artificiellement « parfait » ou « au plus mauvais » sans aucun signal ",
    "visible. Voir l'hypothèse <a href='#hyp-codes'>n°12</a>. Dernière vérification contre l'outil du cycle : ",
    "2026-08-21 (contre <code>RCA_JMMI_2026-08_TOOL.xlsx</code>) - toutes les questions et tous les codes de choix ",
    "utilisés ci-dessus étaient bien présents dans l'outil à cette date.</div>"
  )

  hypotheses_html <- paste0(
    "<h2 id='hypotheses'>Hypothèses retenues pour l'analyse</h2>",
    "<p>Cette section rassemble tous les choix méthodologiques qui ne découlent pas mécaniquement des données, ",
    "mais d'une décision - de conception de l'enquête, de méthode statistique, ou de paramétrage de l'outil.</p>",

    "<h4>PMAS / MEB</h4>",
    "<div class='assumption' id='hyp-quantites'><div class='n'>1</div><div class='body'><span class='tag'>PMAS</span>",
    "<strong>Les quantités mensuelles assumées par article ne sont pas révisées chaque cycle.</strong>",
    "Le poids appliqué à chaque article (tableau ci-dessus) reprend le « Panier Moyen d'Articles » élaboré par le ",
    "GTTM en 2018, conservé sans modification. Le code source note explicitement qu'il faudrait vérifier auprès de ",
    "l'équipe si ces quantités doivent évoluer.</div></div>",

    "<div class='assumption'><div class='n'>2</div><div class='body'><span class='tag'>PMAS</span>",
    "<strong>Marché = Commune.</strong> L'enquête ne distingue pas plusieurs marchés physiques au sein d'une même ",
    "commune : tous les vendeurs interrogés dans une commune sont traités comme un seul marché.</div></div>",

    "<div class='assumption'><div class='n'>3</div><div class='body'><span class='tag'>PMAS</span>",
    "<strong>Les produits non alimentaires durables sont amortis sur 6 mois.</strong> Moustiquaire, bidon, drap, ",
    "natte, bâche et marmite reçoivent chacun un poids de 1/6, supposant un remplacement tous les 6 mois - une ",
    "durée de vie assumée, non mesurée par l'enquête.</div></div>",

    "<div class='assumption'><div class='n'>4</div><div class='body'><span class='tag'>PMAS</span>",
    "<strong>Seuil minimum de 3 prix locaux par article et par commune</strong> (1 seul pour la viande de bœuf, ",
    "article rare avec peu de bouchers par marché) pour considérer un prix « mesuré localement » plutôt qu'imputé.</div></div>",

    "<div class='assumption'><div class='n'>5</div><div class='body'><span class='tag'>PMAS</span>",
    "<strong>L'imputation utilise la médiane nationale des médianes de marché</strong>, pas une médiane calculée ",
    "sur tous les prix individuels bruts confondus - chaque marché compte pour une seule observation dans ce ",
    "secours, même s'il a interrogé beaucoup plus de vendeurs qu'un autre.</div></div>",

    "<div class='assumption' id='hyp-mediane'><div class='n'>6</div><div class='body'><span class='tag'>PMAS</span>",
    "<strong>L'agrégation nationale prend la médiane des paniers communaux déjà calculés</strong>, jamais une ",
    "recomposition à partir de prix par article agrégés au niveau national - ces deux méthodes donnent des ",
    "résultats mathématiquement différents, et seule la première est utilisée dans ce pipeline.</div></div>",

    "<div class='assumption'><div class='n'>7</div><div class='body'><span class='tag'>PMAS</span>",
    "<strong>Seuls les niveaux Commune et National sont publiés</strong> dans les feuilles « Coût médian du PMAS », ",
    "« Evolution des paniers du PMAS » et « Score de Fonctionnalité » - les niveaux intermédiaires (Sous-Préfecture, ",
    "Préfecture) ne sont plus calculés à ces endroits (ils restent disponibles dans les feuilles ",
    "<code>meb_analysis</code>/<code>mfs_analysis</code>, au format standard REACH multi-niveaux).</div></div>",

    "<div class='assumption'><div class='n'>8</div><div class='body'><span class='tag'>PMAS</span>",
    "<strong>La comparaison d'un cycle à l'autre est limitée aux marchés du cycle le plus récent.</strong> Un ",
    "marché évalué uniquement lors d'un cycle antérieur n'apparaît pas dans la feuille « Evolution » du cycle ",
    "suivant.</div></div>",

    "<div class='assumption' id='hyp-taux'><div class='n'>9</div><div class='body'><span class='tag'>PMAS</span>",
    "<strong>Le taux de change XAF/USD est un paramètre fixe de l'outil</strong>, pas resaisi à chaque cycle - il ",
    "doit être mis à jour manuellement dans la configuration du tableau de bord lorsqu'il s'écarte trop du taux de ",
    "marché. Le taux XAF/EUR est une parité fixe qui ne varie jamais.</div></div>",

    "<div class='assumption' id='hyp-manquant'><div class='n'>10</div><div class='body'><span class='tag'>PMAS</span>",
    "<strong>Un article dont la variable a disparu de l'outil Kobo est exclu du panier sans imputation</strong>, ",
    "plutôt que remplacé par une valeur estimée - le panier calculé ce cycle-là repose alors sur un article de ",
    "moins que d'habitude, ce qui est signalé explicitement plutôt que masqué.</div></div>",

    "<h4>Score de Fonctionnalité des Marchés (SFM/MFS)</h4>",
    "<div class='assumption'><div class='n'>11</div><div class='body'><span class='tag'>MFS</span>",
    "<strong>L'indicateur de discrimination (AC.3) suit une logique « un seul cas suffit »</strong>, pas un seuil ",
    "de pourcentage comme les autres indicateurs : un seul vendeur signalant de la discrimination sur un marché ",
    "suffit à mettre le score de cet indicateur à 0 pour toute la commune.</div></div>",

    "<div class='assumption' id='hyp-codes'><div class='n'>12</div><div class='body'><span class='tag'>MFS</span>",
    "<strong>Les codes de réponse Kobo utilisés par le calcul du MFS doivent être revérifiés à chaque cycle.</strong> ",
    "Un code renommé ou retiré de l'outil n'entraîne pas d'erreur visible : l'indicateur concerné devient ",
    "silencieusement 0 (ou faux) pour tous les marchés, ce qui peut fausser un score sans qu'aucun message ",
    "n'apparaisse.</div></div>",

    "<div class='assumption'><div class='n'>13</div><div class='body'><span class='tag'>MFS</span>",
    "<strong>La classification finale retient la pire catégorie applicable</strong>, pas la première rencontrée ",
    "dans un ordre arbitraire - un marché qui remplirait à la fois les critères « limitée » et « mauvaise » est ",
    "classé « mauvaise ».</div></div>",

    "<h4>Général</h4>",
    "<div class='assumption'><div class='n'>14</div><div class='body'><span class='tag'>Général</span>",
    "<strong>Le mois et l'année du cycle sont détectés automatiquement</strong> à partir de la date la plus ",
    "ancienne de la colonne « date » du classeur importé, jamais ressaisis manuellement.</div></div>",

    "<div class='assumption'><div class='n'>15</div><div class='body'><span class='tag'>Général</span>",
    "<strong>L'échantillon n'est pas un échantillon probabiliste.</strong> Les marchés évalués sont sélectionnés ",
    "sur la base du volontariat des partenaires de l'ICSM ; les résultats sont indicatifs des niveaux observés sur ",
    "les marchés évalués, et ne doivent pas être présentés comme statistiquement représentatifs d'une zone plus ",
    "large.</div></div>"
  )

  sources_html <- paste0(
    "<h2 id='sources'>Sources et références</h2>",
    "<ul>",
    "<li>IMPACT Cash and Markets Community of Practice - <i>Guidance on the Market Functionality Score</i> ",
    "(v4, novembre 2023) et le « MFS indicator menu » associé.</li>",
    "<li>GTTM (Groupe de Travail sur les Transferts Monétaires) - « Panier Moyen d'Articles », 2018 (base de la ",
    "composition et des quantités du PMAS/MEB).</li>",
    "<li>Pour toute question de méthodologie ou de mise à jour de ce document : les fichiers sources sont ",
    "<code>rca_jmmi_dashboard/scripts/fonctions_analyse.R</code> (PMAS/MEB) et ",
    "<code>rca_jmmi_dashboard/scripts/mfs.R</code> (SFM/MFS), accompagnés de leurs commentaires détaillés.</li>",
    "</ul>",
    "<footer>Document généré automatiquement par le tableau de bord JMMI RCA - reflète le code de calcul en ",
    "vigueur à la date indiquée en en-tête. En cas de doute entre ce document et le code source, le code source ",
    "fait foi.</footer>"
  )

  paste0(
    "<!doctype html><html lang='fr'><head><meta charset='utf-8'>",
    "<meta name='viewport' content='width=device-width, initial-scale=1'>",
    "<title>Méthodologie JMMI RCA - PMAS/MEB &amp; SFM/MFS</title>",
    "<style>", CSS_METHODOLOGY, "</style></head><body><div class='page'>",
    header_html, toc_html, public_html, contexte_html,
    pmas_intro_html, pmas_etapes_html, pmas_devises_html, pmas_manquants_html, pmas_evolution_html,
    mfs_html, hypotheses_html, sources_html,
    "</div></body></html>"
  )
}
