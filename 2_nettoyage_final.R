################################################################################
# JMMI CAR - NETTOYAGE FINAL (A LA FIN DE LA COLLECTE)
################################################################################
# ENTREE      : jeu de donnees brut, complet - la
#               collecte est fermee, ressources/RCA_JMMI_2026-08_TOOL.xlsx, ressources/check_list_
#               august.xlsx, et (a la 2eme execution) le classeur "suivi a verifier" rempli par
#               l'equipe de suivi dans Sys.getenv("suivi_final_remplis_path")
# SORTIE      : 1ere execution : output/2_a_verifier/RCA_JMMI_<mois>_suivi_a_verifier_<date>_<heure>.xlsx
#               (meme convention de nommage que le suivi quotidien - feuille "A_verifier"
#               - TOUS les problemes detectes dans un seul tableau simple, plus feuille
#               "reponses_autres" pour le recodage des reponses "autre"). 2eme execution :
#               output/4_final_donnees_nettoyees/REACH_CAR_Dataset_JMMI_<Mois>-<Annee>.xlsx (classeur
#               livrable unique).
# DESCRIPTION : CE SCRIPT UNIQUE SE RELANCE DEUX FOIS, A LA FIN DE LA COLLECTE (il n'y a plus de
#               verifications pendant la collecte - voir 1_suivi_quotidien.R) :
#                 1) 1ere execution : fait TOUTES les verifications d'un coup (doublons, consentement,
#                    valeurs bizarres/sentinelles, verifications logiques de la liste, valeurs
#                    aberrantes, quasi-doublons, pourcentage de valeurs manquantes), les regroupe dans
#                    UNE SEULE feuille consolidee ecrite en langage simple (pas de jargon technique),
#                    avec une liste deroulante pour dire quoi faire, et exporte ce classeur.
#                 2) L'equipe de suivi remplit la colonne "Que faire ?" (et recode les reponses
#                    "autre" dans l'autre feuille du meme classeur), enregistre le fichier dans
#                    Sys.getenv("suivi_final_remplis_path").
#                 3) 2eme execution : le script detecte ce classeur rempli, applique toutes les
#                    corrections, recalcule les prix en euros, retire les informations personnelles
#                    identifiables, et assemble le classeur final.
# OBJECTIF    : Un seul point d'entree pour tout le nettoyage final, avec un seul fichier a remplir par
#               l'equipe de suivi (au lieu de plusieurs classeurs a des etapes differentes) - pense pour
#               une equipe avec un faible niveau de formation technique.
# CONTEXTE    : Remplace 2_application_suivis_et_verification_aberrantes.R et
#               3_finalisation_et_classeur_final.R, qui separaient les verifications logiques et les
#               valeurs aberrantes en deux etapes/classeurs distincts. Comme il n'y a plus de
#               verifications pendant la collecte, il n'y a plus besoin de ce decoupage : tout se fait
#               en une fois sur le jeu de donnees complet, a la fin de la collecte. Inspire du
#               data-editing.R unique de l'ETH, lui-meme concu pour etre execute en deux sessions.
################################################################################

rm(list=ls())
source("scripts/0_configuration.R", local = TRUE)

################################################################################
# CHARGEMENT DU JEU DE DONNEES BRUT (jeu de donnees complet - la collecte est fermee)
################################################################################

raw_data_all <- read_raw_data(Sys.getenv("raw_dataset")) %>%
  add_admin_org_labels()

################################################################################
# CONVERSION AUTOMATIQUE DES REPONSES SANS INFORMATION UTILE EN VIDE
################################################################################
# Les champs TEXTE libres contiennent souvent des reponses qui ne portent aucune information utile :
# "RAS"/"R.A.S" (casse/ponctuation/espacement variables), mais aussi "Rs", toute variante de "rien a
# signaler/dire/ajouter/partager", "non"/"non merci"/"non rien", "aucun(e) probleme/incident/
# difficulte/souci", "je ne sais pas", "pas de probleme/souci/remarque/incident/information", et les
# remerciements ("merci", "merci pour l'approche", "remerciements") - demande explicite de l'equipe de
# suivi, avec des exemples reels tires du terrain, fautes de frappe incluses ("Riena signale", "Pas
# de.probleme", "Je n.ai paq constate", "J'avais rien a partage"). Converties en vide ICI, AVANT toute
# verification (donc avant meme la creation du classeur "A verifier" en 1ere passe), pour qu'elles ne
# soient jamais signalees a tort comme une reponse "autre" a recoder ou une valeur aberrante.
# raw_data_all_avant_ras (avant conversion) est conservee uniquement pour la verification
# review_cleaning() plus bas (2eme passe) - le reste du script utilise raw_data_all (converti).
raw_data_all_avant_ras <- raw_data_all

# normalisation tolerante aux fautes de frappe : accents retires, minuscules, toute ponctuation (y
# compris une apostrophe tapee par erreur comme un point, ex. "n.ai" pour "n'ai") remplacee par un
# espace, espaces multiples reduits. "riena" (faute de frappe frequente pour "rien a", sans espace) est
# explicitement recollee en "rien a" avant le decoupage en mots.
normalize_empty_text <- function(x) {
  x <- stringi::stri_trans_general(x, "Latin-ASCII")
  x <- str_to_lower(x)
  x <- str_replace_all(x, "\\briena\\b", "rien a")
  x <- str_replace_all(x, "[^a-z0-9]+", " ")
  x <- str_squish(x)
  # "RAS" saisi avec des separateurs (points/espaces : "R.A.S", "R A S"...) se retrouve decoupe en 3
  # lettres isolees ("r", "a", "s") par le retrait de ponctuation ci-dessus - recollee ici en un seul
  # mot "ras" pour rester reconnue par EMPTY_CONTENT_STEMS plus bas (sinon seule la forme deja collee
  # "ras"/"RAS" sans aucun separateur serait detectee).
  str_replace_all(x, "\\br a s\\b", "ras")
}

# connecteurs grammaticaux retires avant de verifier le contenu (articles, negation, "je n'ai (pas)"...
# "paq" = faute de frappe frequente pour "pas", "ja" = faute de frappe frequente pour "je")
EMPTY_STOPWORDS <- c("a", "de", "du", "des", "le", "la", "les", "l", "d", "je", "j", "ja", "n", "ne",
                     "pas", "paq", "ai", "pour", "avais", "avait", "vous")
# racines "sans contenu utile" une fois les connecteurs retires - une reponse n'est videe QUE si TOUS
# les mots restants appartiennent a cette liste (pas une simple sous-chaine) : "Rien a signaler
# concernant les prix" garde "concernant"/"prix" (hors liste) et n'est donc PAS videe, contrairement a
# "Rien a signaler" seule.
EMPTY_CONTENT_STEMS <- c(
  "ras", "rs", "rien", "non", "aucun", "aucune", "merci", "remerciement", "remerciements",
  "probleme", "problemes", "problems", "souci", "soucis", "remarque", "remarques", "incident",
  "incidents", "difficulte", "difficultes", "information", "informations", "dire", "dit", "tout",
  "signaler", "signale", "signalee", "ajoute", "ajoutee", "ajouter", "constate", "constatee", "sais",
  "savoir", "approche", "partage", "partager", "commentaire", "choses"
)
is_empty_content <- function(x) {
  if (is.na(x)) return(FALSE)
  norm <- normalize_empty_text(x)
  if (!nzchar(norm)) return(TRUE)
  # prefixe special : "pas de probleme..." est videe quel que soit ce qui suit ("pas de probleme pour
  # le moment", "pas de probleme de ma part"...) - contrairement au reste du classificateur, qui exige
  # que TOUS les mots restants soient des racines connues, "de ma part"/"pour le moment" etc. ne le
  # seraient jamais tous individuellement (demande explicite de l'utilisateur : "tout ce qui commence
  # par pas de probleme").
  if (str_starts(norm, "pas de probleme")) return(TRUE)
  words <- setdiff(str_split(norm, "\\s+")[[1]], EMPTY_STOPWORDS)
  length(words) == 0 || all(words %in% EMPTY_CONTENT_STEMS)
}

ras_text_cols <- intersect(kobo_survey$name[kobo_survey$type == "text"], colnames(raw_data_all))

# toutes les colonnes finales sont definies des la construction (meme a 0 ligne) - sinon, quand aucune
# reponse videe n'est trouvee (cas frequent), ras_log resterait limite a uuid/question/old_value et le
# select(..., change_type, new_value, issue, comment) plus bas (fusion dans combined_cleaning_log)
# plangerait avec "Column `change_type` doesn't exist".
ras_log <- tibble(uuid = character(0), question = character(0), old_value = character(0),
                  change_type = character(0), new_value = character(0), issue = character(0), comment = character(0))
for (col in ras_text_cols) {
  vals <- raw_data_all[[col]]
  is_empty <- vapply(vals, is_empty_content, logical(1))
  if (any(is_empty)) {
    ras_log <- bind_rows(ras_log, tibble(
      uuid = raw_data_all$uuid[is_empty], question = col, old_value = vals[is_empty],
      change_type = "blank_response", new_value = NA_character_,
      issue = "Reponse sans information utile (RAS, rien a signaler, aucun probleme, remerciement...)",
      comment = "Convertie automatiquement en vide avant verification"
    ))
    raw_data_all[[col]][is_empty] <- NA_character_
  }
}
if (nrow(ras_log) > 0) {
  cat("--> ", nrow(ras_log), " reponse(s) sans information utile converties automatiquement en vide.\n", sep = "")
}

period <- get_assessment_period(raw_data_all, date_col = "date")
has_logical_checks <- nrow(logical_checks_list) > 0

items_list <- kobo_survey %>%
  filter(type == "calculate" & str_ends(name, "_price_item")) %>%
  pull(name)
items_other_unit <- kobo_survey %>%
  filter(q_type %in% c("decimal", "integer") & str_ends(name, "_other_unit_item")) %>%
  pull(name)
days_db <- kobo_survey %>%
  filter(type %in% c("integer") & str_detect(name, "_days_")) %>%
  pull(name)

# --- Bloc de variables "brutes" par article (voir "detection des valeurs aberrantes" plus bas) : chaque
# <item>_price_item est un champ CALCULATE de l'outil kobo, jamais saisi directement - le signaler tel
# quel dans le classeur "A verifier" ne donnerait rien de concret a corriger a l'equipe de suivi. On
# affiche a la place les variables REELLEMENT saisies qui alimentent son calcul : 1 ou 2 variantes
# "_price_unit_item" (le prix brut saisi, et pour la plupart des articles, une version deja convertie a
# la taille standard - toutes les deux se terminent par "_price_unit_item", d'ou le besoin de les
# distinguer explicitement) ainsi que "_std_unit_item" (l'unite saisie est-elle la norme ?) et
# "_other_unit_item" (quelle unite/quantite sinon) quand ils existent pour cet article. Recherche par
# PREFIXE de nom (pas par analyse de la formule "calculate" - plus simple et tout aussi fiable ici,
# contrairement a recompute_price_items() plus bas qui doit, elle, suivre exactement la formule) - un
# article dont le nommage ne suit pas cette convention (ex. "egg", dont la structure reelle des colonnes
# ne correspond ni a cette convention ni a l'outil "Aout 2026" actuellement utilise ici - decalage
# constate entre l'outil et les donnees reelles, l'outil sera standardise a un prochain cycle, demande
# utilisateur de ne rien adapter cote pipeline pour ce cycle-ci) n'aura simplement aucune variable
# trouvee, auquel cas la ligne _price_item d'origine est conservee telle quelle (repli explicite, voir
# expand_price_item_row() plus bas).
item_prefixes <- sub("_price_item$", "", items_list)
# certains articles ont un champ INTERMEDIAIRE qui se termine lui aussi par "_price_unit_item" (ex.
# "bucket_20l_price_unit_item", "salt_150g_price_unit_item") tout en etant LUI-MEME un champ CALCULATE
# de l'outil kobo (recalcule par recompute_price_items() a partir des VRAIS champs saisis, ex.
# "bucket_price_unit_item"/"bucket_other_unit_item"/"bucket_std_unit_item") - le suffixe seul ne les
# distingue pas des champs reellement saisis. calculated_fields (ci-dessous) marque ces variables :
# write_review_sheet() verrouille leur cellule "Nouvelle reponse" dans Verification_prix (voir plus bas)
# pour eviter qu'une correction y soit tapee sans jamais prendre effet - recompute_price_items() les
# ECRASE INCONDITIONNELLEMENT depuis les VRAIS champs saisis, quelle que soit la valeur du journal de
# nettoyage pour ce champ precis (bug constate : le journal affirmait la correction appliquee alors que
# clean_data gardait l'ancienne valeur, silencieusement recalculee depuis les champs source inchanges).
calculated_field_names <- kobo_survey$name[kobo_survey$type == "calculate"]
get_item_block <- function(prefix) {
  cand <- kobo_survey$name[str_starts(kobo_survey$name, fixed(paste0(prefix, "_")))]
  price_unit <- intersect(cand[str_ends(cand, "_price_unit_item")], colnames(raw_data_all))
  std_unit <- intersect(cand[str_ends(cand, "_std_unit_item")], colnames(raw_data_all))
  other_unit <- intersect(cand[str_ends(cand, "_other_unit_item")], colnames(raw_data_all))
  list(
    price_unit = price_unit,
    std_unit = std_unit,
    other_unit = other_unit,
    calculated = intersect(c(price_unit, std_unit, other_unit), calculated_field_names)
  )
}
item_blocks <- setNames(lapply(item_prefixes, get_item_block), item_prefixes)

INFO_COLS <- c("date", "submission_date", "admin1_code", "admin1_label", "admin2_code", "admin2_label",
              "admin3_code", "admin3_label", "marketplace_type", "enum_id", "shop_name",
              "organisation_name_label")

################################################################################
# DETECTION DU MODE : classeur "a verifier" deja rempli, ou premiere execution ?
################################################################################

filled_files <- list.files(Sys.getenv("suivi_final_remplis_path"),
                           pattern = Sys.getenv("suivi_final_remplis_pattern"),
                           recursive = TRUE, full.names = TRUE)

# noms d'onglets numerotes (ordre de lecture explicite pour l'equipe de suivi) - definis ICI (avant le
# if/else PASSE 1 / PASSE 2 ci-dessous) et non a l'interieur de la PASSE 1 : un bloc if/else ne cree PAS
# un nouvel environnement en R au niveau d'un script source()e - des constantes definies seulement dans
# la branche PASSE 1 n'existeraient tout simplement pas quand la PASSE 2 (l'autre branche) s'execute.
SHEET_LISEZ_MOI <- "1_Lisez_moi"
SHEET_DONNEES_BRUTES <- "2_Donnees_brutes"
SHEET_A_VERIFIER <- "3_A_verifier"
SHEET_VERIFICATION_PRIX <- "4_Verification_prix"
SHEET_REPONSES_AUTRES <- "5_Reponses_autres"
SHEET_JOURNAL_ADDITIONNEL <- "6_Journal_additionnel"
SHEET_JOURNAL_SUPPRESSION <- "7_Journal_suppression"

if (length(filled_files) == 0) {

################################################################################
# PASSE 1 : TOUTES LES VERIFICATIONS, PUIS EXPORT DU CLASSEUR "A VERIFIER"
################################################################################

cat("\n--> Aucun classeur rempli trouve dans '", Sys.getenv("suivi_final_remplis_path"),
    "' - execution de toutes les verifications (1ere passe)...\n", sep = "")

# 1) doublons stricts (meme identifiant d'enquete) : signale un probleme technique d'export/pipeline
#    (jamais quelque chose que l'equipe de suivi terrain peut resoudre) - on arrete plutot que
#    d'ajouter cela a la feuille de verification.
dup_check <- raw_data_all %>%
  check_duplicate(uuid_column = "uuid", log_name = "duplicate_log")
if (nrow(dup_check$duplicate_log) > 0) {
  stop("Doublons detectes (identifiant d'enquete identique) --> necessite une intervention technique avant de continuer.")
}

# 2) consentement + valeurs bizarres (sentinelles) + verifications logiques de la liste
chain <- raw_data_all %>%
  check_logical(
    uuid_column = "uuid",
    information_to_add = NULL,
    check_id = "logical_reject",
    check_to_perform = "consent == \"non\"",
    columns_to_clean = c("consent"),
    description = "Pas de consentement - a supprimer"
  ) %>%
  check_value_modified(uuid_column = "uuid",
                       element_name = "checked_dataset",
                       columns_not_to_check = c("enum_id", "_index"),
                       values_to_look = c(99, 999, -999, 88, 888, -888))

# "days_db" (stock_duration_days/restock_wait_days, voir plus haut) : PAS de detection statistique
# d'aberrations pour ces variables (exclues de check_outliers() plus bas) - demande utilisateur, seule
# une valeur EXPLICITEMENT superieure a 365 jours (plus d'un an) est signalee. Une valeur juste
# "inhabituelle" par rapport a la mediane (ex. 40 jours quand la plupart des vendeurs en ont 10) reste
# plausible sur le terrain et n'a pas besoin d'etre signalee pour ces variables. as.numeric(...) dans la
# condition car raw_data_all est entierement en texte (read_raw_data(), col_types="text") - une
# comparaison ">" sans cela serait lexicographique, pas numerique ("45" > "365" vaudrait TRUE).
if (length(days_db) > 0) {
  days_threshold_checks <- tibble(
    id_verification = paste0("days_365_", days_db),
    condition_verification = paste0("as.numeric(", days_db, ") > 365"),
    colonnes_a_nettoyer = days_db,
    description = "Valeur superieure a 365 jours - a verifier"
  )
  chain <- chain %>%
    check_logical_with_list(uuid_column = "uuid",
                            list_of_check = days_threshold_checks,
                            check_id_column = "id_verification",
                            check_to_perform_column = "condition_verification",
                            columns_to_clean_column = "colonnes_a_nettoyer",
                            description_column = "description")
}

if (has_logical_checks) {
  chain <- chain %>%
    check_logical_with_list(uuid_column = "uuid",
                            list_of_check = logical_checks_list,
                            check_id_column = "id_verification",
                            check_to_perform_column = "condition_verification",
                            columns_to_clean_column = "colonnes_a_nettoyer",
                            description_column = "description")
} else {
  cat("--> Aucune verification logique definie dans ressources/check_list_august.xlsx pour le moment - ignore.\n")
}

log_principal <- create_combined_log(list_of_log = c(chain), dataset_name = "checked_dataset") %>%
  add_info_to_cleaning_log(dataset = "checked_dataset", cleaning_log = "cleaning_log",
                           dataset_uuid_column = "uuid", cleaning_log_uuid_column = "uuid",
                           information_to_add = INFO_COLS)
df_principal <- log_principal$cleaning_log %>% mutate(issue = translate_issue_fr(issue))
# check_id (identifiant de la verification elle-meme, ex. "logical_reject"/"days_365_<variable>"/le
# id_verification du checklist) est deja pose par check_logical()/check_logical_with_list() -
# create_combined_log() ne le touche pas (uniquement check_binding) - garde-fou de robustesse ici au cas
# ou une future verification ajoutee a ce chainage ne le renseignerait pas : sert a construire row_id
# (identifiant STABLE d'une ligne d'un tour du pipeline a l'autre, voir plus bas et
# carry_forward_decisions() dans fonctions_nettoyage.R) - sans repli, une ligne sans check_id
# produirait un row_id contenant "NA" invariable, qui matcherait a tort N'IMPORTE QUELLE autre ligne
# elle aussi depourvue de check_id.
if (!"check_id" %in% colnames(df_principal)) df_principal$check_id <- NA_character_
df_principal$check_id[is.na(df_principal$check_id)] <- "verification_logique"

# 3) valeurs statistiquement aberrantes - UNIQUEMENT sur les <item>_price_item (prix normalise pour une
#    quantite comparable d'un marche a l'autre, ce qui rend la comparaison statistique pertinente) ;
#    "_price_unit_item"/"_std_unit_item"/"_other_unit_item" en sont explicitement exclus (prix/unite
#    "bruts", pas comparables entre repondants puisque l'unite saisie varie) - directement sur les
#    donnees brutes, il n'y a plus de seconde passe apres application des verifications logiques.
#    "days_db" (stock_duration_days/restock_wait_days) est egalement exclu ici (demande utilisateur) :
#    ces variables sont desormais signalees uniquement via un seuil fixe (> 365 jours, voir plus haut),
#    pas par comparaison statistique a la mediane.
outliers_log <- raw_data_all %>%
  check_outliers(
    uuid_column = "uuid", kobo_survey = NULL, kobo_choices = NULL, cols_to_add_cleaning_log = NULL,
    strongness_factor = 3, minimum_unique_value_of_variable = 2, remove_choice_multiple = TRUE,
    sm_separator = "/",
    columns_not_to_check = setdiff(names(raw_data_all), c("uuid", items_list))
  )
log_outliers <- create_combined_log(list_of_log = outliers_log) %>%
  add_info_to_cleaning_log(dataset = "checked_dataset", cleaning_log = "cleaning_log",
                           dataset_uuid_column = "uuid", cleaning_log_uuid_column = "uuid",
                           information_to_add = INFO_COLS)
df_outliers <- log_outliers$cleaning_log %>%
  # check_outliers() ne renseigne pas check_id (contrairement a check_logical()/check_logical_with_list()
  # ci-dessus) - valeur fixe ici (une seule verification statistique dans ce chainage, donc pas besoin de
  # differencier davantage) pour que row_id (voir plus bas) reste stable d'un tour du pipeline a l'autre.
  mutate(issue = translate_issue_fr(issue), check_id = "outlier") %>%
  # ne pas signaler deux fois le meme champ (deja signale ci-dessus par les verifications logiques)
  anti_join(df_principal, by = c("uuid", "question"))

# 3bis) prix depassant un seuil ABSOLU (demande utilisateur : un prix normalise superieur a 20000 ne
#       devrait JAMAIS arriver, quelle que soit la population de l'echantillon). Contrairement a la
#       verification precedente (3, RELATIVE a la distribution de l'echantillon - moyenne/ecart-type),
#       ce seuil ne depend d'aucune distribution : il signale donc meme le cas extreme ou une colonne
#       entiere serait anormalement elevee (ex. erreur d'unite systematique sur tout un article), que la
#       methode statistique relative ne peut par nature detecter (rien n'y "depasse" le reste de sa
#       propre colonne).
PRICE_ABSOLUTE_THRESHOLD <- 20000
absolute_outlier_rows <- list()
for (item in items_list) {
  x <- suppressWarnings(as.numeric(raw_data_all[[item]]))
  is_outlier <- !is.na(x) & x > PRICE_ABSOLUTE_THRESHOLD
  if (!any(is_outlier)) next
  absolute_outlier_rows[[item]] <- data.frame(
    uuid = raw_data_all$uuid[is_outlier],
    question = item,
    old_value = as.character(x[is_outlier]),
    issue = paste0("Prix superieur au seuil absolu de ", PRICE_ABSOLUTE_THRESHOLD, " - a verifier"),
    check_id = "prix_seuil_absolu", check_binding = NA_character_,
    change_type = NA_character_, new_value = NA_character_,
    stringsAsFactors = FALSE
  )
}
df_outliers_absolute <- if (length(absolute_outlier_rows) > 0) {
  bind_rows(absolute_outlier_rows) %>%
    left_join(raw_data_all %>% select(uuid, all_of(INFO_COLS)), by = "uuid") %>%
    # ne pas signaler deux fois le meme champ (deja signale par les verifications logiques, ou par
    # check_outliers() ci-dessus quand les deux methodes s'accordent)
    anti_join(df_principal, by = c("uuid", "question")) %>%
    anti_join(df_outliers, by = c("uuid", "question"))
} else {
  df_outliers[0, ]
}
df_outliers <- bind_rows(df_outliers, df_outliers_absolute)

# les signalements sur un _price_item sont isoles dans leur PROPRE onglet ("Verification_prix", voir
# plus bas) plutot que melanges au reste de "A_verifier" - eclates chacun en plusieurs lignes (une par
# variable "brute" du bloc de l'article, voir item_blocks ci-dessus) : les variantes "_price_unit_item"
# en premier (role "prix" - ce sont elles qui expliquent la valeur aberrante), puis "_std_unit_item"/
# "_other_unit_item" en dessous (role "contexte"). block_id (partage par toutes les lignes d'un meme
# bloc) sert au regroupement visuel (bordure epaisse, voir write_review_sheet()) dans ce nouvel onglet.
# Le reste des signalements (ex. days_db) suit le flux normal, integre a "A_verifier" comme avant. Un
# article dont le nommage ne suit pas la convention "_price_unit_item"/"_std_unit_item"/"_other_unit_item"
# (ex. "egg", voir item_blocks) n'a aucune variable de bloc trouvee : sa ligne _price_item d'origine part
# quand meme vers "Verification_prix" (isolee comme les autres articles), mais sans eclatement.
is_price_item_outlier <- df_outliers$question %in% items_list
df_outliers_prix_raw <- df_outliers[is_price_item_outlier, ]
df_outliers <- df_outliers[!is_price_item_outlier, ]

expand_price_item_row <- function(row) {
  item_prefix <- item_prefixes[paste0(item_prefixes, "_price_item") == row$question[1]]
  block <- if (length(item_prefix) == 1) item_blocks[[item_prefix]] else NULL
  # le repli "ligne _price_item d'origine conservee telle quelle" (voir commentaire item_blocks plus
  # haut) doit se declencher des qu'il n'y a AUCUNE variable "_price_unit_item" (role "prix", celle qui
  # explique concretement la valeur aberrante) - pas seulement quand le bloc entier est vide. Un article
  # comme "egg" a bien des variables "_std_unit_item"/"_other_unit_item" (role "contexte") MAIS aucune
  # "_price_unit_item" : sans ce garde specifique, le bloc aurait ete quand meme eclate en lignes
  # "contexte" uniquement (aucune ligne "prix" au-dessus pour les expliquer), toutes affichant a tort
  # "Variable liee au prix ci-dessus".
  block_vars <- if (!is.null(block) && length(block$price_unit) > 0) {
    c(block$price_unit, block$std_unit, block$other_unit)
  } else {
    character(0)
  }
  if (length(block_vars) == 0) {
    row$block_id <- NA_character_
    row$block_role <- NA_character_
    # repli : la ligne _price_item d'origine elle-meme, TOUJOURS un champ CALCULATE par definition
    # (voir items_list) - la verrouiller aussi (voir calculated_field_names/is_calculated_field plus
    # haut), pour la meme raison que les variables intermediaires du bloc.
    row$is_calculated_field <- TRUE
    return(row)
  }
  raw_row <- raw_data_all[raw_data_all$uuid == row$uuid[1], ]
  out <- row[rep(1, length(block_vars)), ]
  out$question <- block_vars
  out$old_value <- vapply(block_vars, function(v) {
    val <- raw_row[[v]][1]
    if (is.na(val)) "" else as.character(val)
  }, character(1))
  out$block_id <- paste0(row$uuid[1], "__", item_prefix)
  out$block_role <- ifelse(block_vars %in% block$price_unit, "prix", "contexte")
  out$issue <- ifelse(out$block_role == "prix", row$issue[1], "Variable liee au prix ci-dessus")
  out$is_calculated_field <- block_vars %in% block$calculated
  out
}

if (nrow(df_outliers_prix_raw) > 0) {
  expanded_rows <- vector("list", nrow(df_outliers_prix_raw))
  for (i in seq_len(nrow(df_outliers_prix_raw))) {
    expanded_rows[[i]] <- expand_price_item_row(df_outliers_prix_raw[i, ])
  }
  df_outliers_prix <- bind_rows(expanded_rows)
} else {
  df_outliers_prix <- df_outliers_prix_raw
  df_outliers_prix$block_id <- character(0)
  df_outliers_prix$block_role <- character(0)
  df_outliers_prix$is_calculated_field <- logical(0)
}

# 4) quasi-doublons (enquetes tres semblables entre elles) - un probleme au niveau de l'enquete
#    entiere, pas d'un champ en particulier (question = NA)
soft_dup <- raw_data_all %>%
  check_soft_duplicates(kobo_survey = kobo_survey, uuid_column = "uuid", idnk_value = "ne_sait_pas",
                        sm_separator = "/", log_name = "soft_duplicate_log", threshold = 5,
                        return_all_results = FALSE)
# check_soft_duplicates() ne renvoie la colonne "id_most_similar_survey" QUE s'il y a au moins un
# quasi-doublon trouve - avec 0 quasi-doublon (cas reel, pas juste theorique, sur un petit jeu de
# donnees propre), soft_duplicate_log n'a que uuid/issue, et la reference a cette colonne plus bas
# ferait planter le mutate(). L'ajouter vide dans ce cas la rend inoffensive (0 lignes de toute facon).
soft_dup_log <- soft_dup$soft_duplicate_log
if (!"id_most_similar_survey" %in% colnames(soft_dup_log)) {
  soft_dup_log$id_most_similar_survey <- character(0)
}
df_soft_dup <- soft_dup_log %>%
  mutate(question = NA_character_,
        old_value = paste0("Tres semblable a l'enquete ", id_most_similar_survey),
        issue = translate_issue_fr(issue),
        # valeur fixe (pas NA) pour row_id (voir plus bas) - un seul type de verification "quasi-doublon"
        # dans ce chainage, uuid seul suffit deja a l'identifier de facon unique par ligne.
        check_id = "quasi_doublon", check_binding = NA_character_,
        change_type = NA_character_, new_value = NA_character_) %>%
  select(uuid, question, old_value, issue, check_id, check_binding, change_type, new_value) %>%
  left_join(raw_data_all %>% select(uuid, all_of(INFO_COLS)), by = "uuid")

# 5) pourcentage de valeurs manquantes anormalement eleve - egalement un probleme au niveau de
#    l'enquete entiere (question = NA)
raw_pm <- raw_data_all %>%
  add_percentage_missing(kobo_survey = kobo_survey,
                         type_to_include = c("integer", "decimal", "select_one", "select_multiple"))
pm <- list(checked_dataset = raw_pm) %>%
  check_percentage_missing(uuid_column = "uuid", column_to_check = "percentage_missing",
                           strongness_factor = 3, log_name = "percentage_missing_log")
df_pct_missing <- pm$percentage_missing_log %>%
  mutate(question = NA_character_,
        old_value = paste0(round(as.numeric(old_value) * 100, 1), "% de reponses manquantes"),
        issue = translate_issue_fr(issue),
        # valeur fixe (pas NA), meme raison que check_id = "quasi_doublon" ci-dessus.
        check_id = "pct_missing", check_binding = NA_character_,
        change_type = NA_character_, new_value = NA_character_) %>%
  select(uuid, question, old_value, issue, check_id, check_binding, change_type, new_value) %>%
  left_join(raw_data_all %>% select(uuid, all_of(INFO_COLS)), by = "uuid")

# toutes les verifications combinees dans un seul tableau
all_issues <- bind_rows(df_principal, df_outliers, df_soft_dup, df_pct_missing)

# anonymisation du nom de la boutique/du vendeur (DPI, voir remove_cols plus bas ou shop_name est
# retire du classeur final) : un identifiant numerique stable remplace le nom en clair dans la colonne
# "Boutique / vendeur" du classeur "A verifier" - stable pour une meme boutique (numerotee une seule
# fois, par ordre alphabetique du nom, pour que l'equipe de suivi puisse quand meme repérer une meme
# boutique signalee sur plusieurs lignes) plutot qu'un numero de ligne qui romprait ce lien.
shop_anon_map <- raw_data_all %>%
  distinct(shop_name) %>%
  filter(!is.na(shop_name) & nzchar(shop_name)) %>%
  arrange(shop_name) %>%
  mutate(shop_anon_id = paste0("Boutique ", row_number())) %>%
  { setNames(.$shop_anon_id, .$shop_name) }

# colonnes simples/lisibles pour l'equipe de suivi (voir REVIEW_HEADERS_FR / write_review_sheet() dans
# scripts/fonctions_nettoyage.R)
review_df <- all_issues %>%
  rowwise() %>%
  mutate(question_label = get_question_label(question)) %>%
  ungroup() %>%
  transmute(
    uuid = uuid,
    # identifiant STABLE d'une ligne (uuid + variable signalee + verification a l'origine du
    # signalement) - le meme d'une execution a l'autre du pipeline tant que les valeurs brutes de
    # l'enquete n'ont pas change (voir carry_forward_decisions() dans fonctions_nettoyage.R). "ENQUETE"
    # en repli pour les verifications au niveau de l'enquete entiere (quasi-doublon/% manquant), ou
    # question vaut NA.
    row_id = paste(uuid, coalesce(question, "ENQUETE"), check_id, sep = "::"),
    commune = admin3_label,
    partenaire = organisation_name_label,
    date = date,
    enqueteur = enum_id,
    boutique = unname(shop_anon_map[shop_name]),
    variable = question,
    question_label = question_label,
    reponse_actuelle = old_value,
    probleme = issue,
    # ecrase par carry_forward_decisions() en "Repris (deja traite)" pour les lignes qui reprennent une
    # decision d'un classeur precedent (voir plus bas, apres la construction de df_other).
    statut = "Nouveau",
    action = NA_character_,
    nouvelle_reponse = NA_character_,
    commentaire = NA_character_
  ) %>%
  arrange(partenaire, date)

if (interactive()) View(review_df)

cat("\n--> ", nrow(review_df), " probleme(s) a verifier detecte(s).\n", sep = "")

# meme mise en forme que review_df ci-dessus, pour les signalements sur un _price_item eclates en bloc
# (voir df_outliers_prix plus haut) - ecrits dans leur PROPRE onglet "Verification_prix" (voir plus bas)
# plutot que melanges a "A_verifier", avec block_id/block_role conserves pour le regroupement visuel
# (bordure epaisse - voir write_review_sheet()).
review_df_prix <- df_outliers_prix %>%
  rowwise() %>%
  mutate(question_label = get_question_label(question)) %>%
  ungroup() %>%
  transmute(
    uuid = uuid,
    # question a ce stade est deja la variable "brute" ECLATEE (ex. egg_price_unit_item), pas le
    # <item>_price_item d'origine - row_id reste donc unique par ligne meme si check_id ("outlier") est
    # identique pour toutes les lignes d'un meme bloc (voir expand_price_item_row() plus haut).
    row_id = paste(uuid, question, check_id, sep = "::"),
    commune = admin3_label,
    partenaire = organisation_name_label,
    date = date,
    enqueteur = enum_id,
    boutique = unname(shop_anon_map[shop_name]),
    variable = question,
    question_label = question_label,
    reponse_actuelle = old_value,
    probleme = issue,
    statut = "Nouveau",
    action = NA_character_,
    nouvelle_reponse = NA_character_,
    commentaire = NA_character_,
    block_id = block_id,
    block_role = block_role,
    is_calculated_field = is_calculated_field
  ) %>%
  # uuid ajoute au tri (en plus de partenaire/date, deja identiques pour toutes les lignes d'un meme
  # bloc) pour garantir que les lignes d'un meme bloc restent bien ADJACENTES apres le tri, quel que
  # soit l'ordre de construction - necessaire pour que le regroupement visuel de write_review_sheet()
  # (base sur des lignes CONSECUTIVES partageant le meme block_id) soit correct.
  arrange(partenaire, date, uuid)

cat("--> ", nrow(review_df_prix), " ligne(s) de verification de prix (onglet 'Verification_prix').\n", sep = "")

################################################################################
# REPONSES "AUTRE"
################################################################################

var_other_raw <- other_db$name[other_db$name %in% colnames(raw_data_all)]

other_responses <- raw_data_all %>%
  select(c("uuid", all_of(var_other_raw))) %>%
  pivot_longer(cols = all_of(var_other_raw), names_to = "question_name", values_to = "response") %>%
  filter(!is.na(response)) %>%
  select(uuid, question_name, response)

# memes colonnes de contexte (commune, partenaire, date, enqueteur, boutique) et memes sources
# (admin3_label, organisation_name_label, enum_id, shop_anon_map[shop_name]) que review_df/
# review_df_prix plus haut ("3_A_verifier"/"4_Verification_prix") - demande utilisateur, pour que
# l'equipe retrouve les MEMES reperes d'une feuille a l'autre du classeur.
df_other <- other_responses %>%
  left_join(select(raw_data_all, uuid, enum_id, organisation_name_label, date, shop_name, admin3_label),
            by = "uuid") %>%
  arrange(question_name) %>%
  left_join(select(other_db, name, full_label, q_type, list_name), by = c("question_name" = "name")) %>%
  transmute(
    uuid = uuid,
    # identifiant stable (voir row_id dans review_df plus haut) - un seul type de ligne possible ici
    # (une reponse "autre" a classer), uuid + question_name suffit deja a l'identifier de facon unique.
    row_id = paste(uuid, question_name, sep = "::"),
    commune = admin3_label,
    partenaire = organisation_name_label,
    date = date,
    enqueteur = enum_id,
    boutique = unname(shop_anon_map[shop_name]),
    q_type = q_type, list_name = list_name, question_name = question_name, full_label = full_label,
    response = response, selected_choices = NA,
    statut = "Nouveau",
    action = NA_character_,
    new_value_text = NA_character_, new_value_choice = NA_character_
  )

# available_choices ("Choix possibles") a ete retire (demande utilisateur - redondant avec la liste
# deroulante de "Nouvelle reponse - choix existant", qui montre deja les memes choix - voir
# add_other_responses_sheet()) : ne reste que selected_choices (choix DEJA coches pour une question
# select_multiple, pas une reference des choix disponibles).
for (r in seq_len(nrow(df_other))) {
  ref_name <- other_db[["ref_question"]][other_db$name == df_other$question_name[r]]
  q_type <- other_db[["q_type"]][other_db$name %in% df_other$question_name[r]]
  if (q_type %in% "select_multiple") {
    choices <- str_split(get_value_from_uuid(df_other$uuid[r], ref_name), " ")[[1]]
    labels <- unlist(lapply(choices, function(c) get_label_from_name(df_other$list_name[r], c)))
    df_other$selected_choices[r] <- paste0(labels, collapse = ";\n")
  }
}
df_other <- relocate(df_other, "selected_choices", .before = "response")

if (interactive()) View(df_other)

cat("--> ", nrow(df_other), " reponse(s) autre a recoder.\n", sep = "")

################################################################################
# REPRISE DES DECISIONS D'UN CLASSEUR "SUIVI A VERIFIER" PRECEDENT (optionnel)
################################################################################
# Si un classeur "suivi a verifier" precedent est fourni (previous_review_path - voir Etape 1 du tableau
# de bord app.R, "Un nettoyage precedent existe-t-il ?") : reprend automatiquement, ligne par ligne, les
# decisions ("Que faire ?"/"Nouvelle reponse"/commentaire) deja prises pour les lignes qui existent
# encore, a l'identique (meme row_id ET meme valeur d'origine), dans les 3 feuilles ci-dessus - evite a
# l'equipe de suivi de tout reprendre a zero ou de copier-coller a la main entre deux classeurs quand la
# collecte se poursuit et que le pipeline est relance sur des donnees qui se sont agrandies. Les
# anciennes lignes elles-memes ne sont jamais modifiees (hypothese de l'utilisateur : une enquete deja
# exportee une fois ne change plus) - voir carry_forward_decisions() dans fonctions_nettoyage.R pour le
# detail (dont le garde-fou "reponse actuelle inchangee", qui empeche de reprendre une decision perimee
# si cette hypothese s'averait fausse dans un cas particulier).
# previous_journal_additionnel/previous_journal_suppression initialises AVANT le if/else ci-dessous (pas
# seulement dans la branche executee) : necessaires plus bas, apres la fin du if/else, au moment
# d'ecrire le classeur (add_additional_log_sheets()) - un bloc if/else ne cree pas un nouvel
# environnement en R au niveau d'un script source()e, une variable assignee uniquement dans la branche
# "if" n'existerait simplement pas si la branche "else" (ou aucune branche) s'executait (meme bug de
# fond que celui deja rencontre avec les constantes SHEET_* plus haut dans ce fichier).
previous_journal_additionnel <- NULL
previous_journal_suppression <- NULL
previous_review_path <- Sys.getenv("previous_review_path")
if (nzchar(previous_review_path) && file.exists(previous_review_path)) {
  cat("\n--> Classeur 'suivi a verifier' precedent fourni (", previous_review_path, ") - reprise des ",
      "decisions deja prises...\n", sep = "")

  previous_a_verifier <- tryCatch(
    read_review_sheet(previous_review_path, sheet_name = SHEET_A_VERIFIER, translate_choices = FALSE) %>%
      rename(variable = question, reponse_actuelle = old_value, nouvelle_reponse = new_value,
            probleme = issue, commentaire = comment),
    error = function(e) {
      warning("Lecture de '", SHEET_A_VERIFIER, "' dans le classeur precedent : ", conditionMessage(e))
      NULL
    }
  )
  review_df <- carry_forward_decisions(review_df, previous_a_verifier, id_col = "row_id",
                                       action_col = "action",
                                       carry_cols = c("nouvelle_reponse", "commentaire"),
                                       old_value_col = "reponse_actuelle")

  previous_verification_prix <- tryCatch(
    read_review_sheet(previous_review_path, sheet_name = SHEET_VERIFICATION_PRIX, translate_choices = FALSE) %>%
      rename(variable = question, reponse_actuelle = old_value, nouvelle_reponse = new_value,
            probleme = issue, commentaire = comment),
    error = function(e) {
      warning("Lecture de '", SHEET_VERIFICATION_PRIX, "' dans le classeur precedent : ", conditionMessage(e))
      NULL
    }
  )
  review_df_prix <- carry_forward_decisions(review_df_prix, previous_verification_prix, id_col = "row_id",
                                            action_col = "action",
                                            carry_cols = c("nouvelle_reponse", "commentaire"),
                                            old_value_col = "reponse_actuelle")

  previous_reponses_autres <- tryCatch(
    read_other_responses_sheet(previous_review_path, sheet_name = SHEET_REPONSES_AUTRES),
    error = function(e) {
      warning("Lecture de '", SHEET_REPONSES_AUTRES, "' dans le classeur precedent : ", conditionMessage(e))
      NULL
    }
  )
  df_other <- carry_forward_decisions(df_other, previous_reponses_autres, id_col = "row_id",
                                      action_col = "action",
                                      carry_cols = c("new_value_text", "new_value_choice"),
                                      old_value_col = "response")

  # Journal_additionnel/Journal_suppression : saisie 100% manuelle, jamais regeneree automatiquement -
  # "reprendre" signifie ici recopier telles quelles les lignes deja tapees (voir le commentaire sur
  # add_additional_log_sheets() dans fonctions_nettoyage.R), pas les faire correspondre par row_id comme
  # pour les 3 feuilles ci-dessus. Meme filtre "au moins une cellule non vide" que
  # prepare_additional_cleaning_log()/prepare_additional_deletion_log() (fonctions_nettoyage.R), pour
  # ignorer les lignes vides restantes en bas de la feuille.
  previous_journal_additionnel <- tryCatch({
    df <- read_excel(previous_review_path, sheet = SHEET_JOURNAL_ADDITIONNEL, col_types = "text")
    df[rowSums(!is.na(df)) > 0, , drop = FALSE]
  }, error = function(e) {
    warning("Lecture de '", SHEET_JOURNAL_ADDITIONNEL, "' dans le classeur precedent : ", conditionMessage(e))
    NULL
  })
  previous_journal_suppression <- tryCatch({
    df <- read_excel(previous_review_path, sheet = SHEET_JOURNAL_SUPPRESSION, col_types = "text")
    df[rowSums(!is.na(df)) > 0, , drop = FALSE]
  }, error = function(e) {
    warning("Lecture de '", SHEET_JOURNAL_SUPPRESSION, "' dans le classeur precedent : ", conditionMessage(e))
    NULL
  })

  n_repris <- sum(review_df$statut == "Repris (deja traite)") +
    sum(review_df_prix$statut == "Repris (deja traite)") +
    sum(df_other$statut == "Repris (deja traite)")
  n_journal_repris <- (if (!is.null(previous_journal_additionnel)) nrow(previous_journal_additionnel) else 0) +
    (if (!is.null(previous_journal_suppression)) nrow(previous_journal_suppression) else 0)
  cat("--> ", n_repris, " ligne(s) reprise(s) automatiquement du classeur precedent ",
      "(statut \"Repris (deja traite)\"), plus ", n_journal_repris,
      " ligne(s) recopiee(s) depuis Journal_additionnel/Journal_suppression.\n", sep = "")
} else if (nzchar(previous_review_path)) {
  warning("Classeur 'suivi a verifier' precedent introuvable (", previous_review_path, ") - reprise ",
         "des decisions ignoree, toutes les lignes sont traitees comme nouvelles.")
}

################################################################################
# TRI "NOUVEAU" EN PREMIER (demande utilisateur - le plus visible en haut de chaque feuille)
################################################################################
# Sans effet quand aucun classeur precedent n'est fourni (statut vaut "Nouveau" partout, tri stable ->
# ordre inchange) : applique donc inconditionnellement, pas seulement dans la branche de reprise
# ci-dessus. "statut != 'Nouveau'" vaut FALSE (0) pour "Nouveau" et TRUE (1) pour "Repris (deja
# traite)" - trier dessus en premier fait remonter "Nouveau" en tete tout en gardant l'ordre existant
# (partenaire/date[/uuid]) comme critere secondaire, un tri etant stable en R (ties non reordonnees).
review_df <- review_df %>% arrange(statut != "Nouveau", partenaire, date)

# review_df_prix : chaque bloc (prix + contexte d'un meme article, voir expand_price_item_row() plus
# haut) DOIT rester un groupe de lignes CONSECUTIVES pour que le regroupement visuel de
# write_review_sheet() (fond alterne + bordure) reste correct - trier directement sur "statut" ligne par
# ligne casserait un bloc dont une seule ligne serait (a tort ou a raison) differente des autres. Calcule
# donc un statut AU NIVEAU DU BLOC (le bloc entier remonte en "a traiter" des qu'AU MOINS une de ses
# lignes est "Nouveau" - plus utile pour l'equipe qu'un bloc a moitie remonte, a moitie enterre) et trie
# dessus : toutes les lignes d'un meme bloc partagent alors la MEME cle de tri primaire, donc restent
# groupees ensemble (uuid/partenaire/date, deja identiques au sein d'un bloc, achevent de les regrouper).
# Lignes sans bloc (block_id NA - article sans variables "_price_unit_item", voir plus haut) : chacune
# traitee comme son propre groupe (cle synthetique unique), pour ne jamais les faire influencer entre
# elles alors qu'elles n'ont rien a voir les unes avec les autres.
block_group_key <- ifelse(is.na(review_df_prix$block_id),
                          paste0("__ligne_", seq_len(nrow(review_df_prix))),
                          review_df_prix$block_id)
review_df_prix <- review_df_prix %>%
  mutate(.groupe_bloc = block_group_key) %>%
  group_by(.groupe_bloc) %>%
  mutate(.bloc_a_traiter = any(statut == "Nouveau")) %>%
  ungroup() %>%
  arrange(desc(.bloc_a_traiter), partenaire, date, uuid) %>%
  select(-.groupe_bloc, -.bloc_a_traiter)

df_other <- df_other %>% arrange(statut != "Nouveau")

################################################################################
# ECRITURE DU CLASSEUR "A VERIFIER" (verifications + reponses autres, meme fichier)
################################################################################

wb <- createWorkbook()
# 1ere feuille (openxlsx ordonne les onglets selon l'ordre des addWorksheet() ci-dessous) : mode
# d'emploi du classeur, voir write_suivi_intro_sheet().
wb <- write_suivi_intro_sheet(wb, sheet_name = SHEET_LISEZ_MOI)

# donnees brutes, anonymisees de la meme maniere que raw_data dans 1_suivi_quotidien.R/le classeur
# final - juste apres "Lisez_moi" (demande utilisateur) pour que l'equipe puisse tout de suite
# retrouver le detail complet d'une enquete signalee dans les onglets suivants, sans devoir rouvrir
# l'export kobo brut separement. raw_data_all a ce stade porte deja les corrections automatiques
# appliquees plus haut (labels admin, conversion RAS/reponses sans information utile en vide) - sert
# aussi de reference a jour pour "Journal_additionnel" (voir write_suivi_intro_sheet()).
remove_cols_raw <- c(
  intersect(colnames(raw_data_all), filter(kobo_survey, type %in% c("geopoint"))$name),
  "audit", "audit_URL", "deviceid", "enum_id", "shop_name", "organisation_name_label"
)
remove_cols_raw <- intersect(remove_cols_raw, colnames(raw_data_all))
raw_data_export <- select(raw_data_all, -all_of(remove_cols_raw))
if ("_index" %in% colnames(raw_data_export) && !"id" %in% colnames(raw_data_export)) {
  raw_data_export <- rename(raw_data_export, id = `_index`)
}
write_labeled_data_sheet(wb, SHEET_DONNEES_BRUTES, apply_var_labels(raw_data_export))

wb <- write_review_sheet(wb, review_df, sheet_name = SHEET_A_VERIFIER)
# onglet dedie aux signalements de prix (voir review_df_prix plus haut), isole du reste des
# verifications - demande utilisateur, pour ne pas alourdir "A_verifier" avec les blocs de variables
# liees (prix + std/other unit) qui ont chacun besoin d'etre lus/corriges ensemble.
wb <- write_review_sheet(wb, review_df_prix, sheet_name = SHEET_VERIFICATION_PRIX)
wb <- add_other_responses_sheet(wb, df_other, sheet_name = SHEET_REPONSES_AUTRES)
# journal de nettoyage additionnel, integre directement dans ce classeur plutot qu'un modele separe a
# telecharger/reimporter (voir add_additional_log_sheets() et prepare_additional_cleaning_log()/
# prepare_additional_deletion_log() qui le relisent, plus bas, en passe 2).
wb <- add_additional_log_sheets(wb, sheet_name = SHEET_JOURNAL_ADDITIONNEL, sheet_name2 = SHEET_JOURNAL_SUPPRESSION,
                                previous_additional = previous_journal_additionnel,
                                previous_deletion = previous_journal_suppression)

a_remplir_dir <- Sys.getenv("suivi_final_a_remplir_dir")
if (!dir.exists(a_remplir_dir)) dir.create(a_remplir_dir, recursive = TRUE)
# meme convention de nommage que le suivi quotidien (RCA_JMMI_<mois>_<type>_<date>_<heure>) - voir
# 1_suivi_quotidien.R
out_path <- paste0(a_remplir_dir, "RCA_JMMI_", period$assessment_month,
                   "_suivi_a_verifier_", format(Sys.time(), "%Y-%m-%d_%Hh%M"), ".xlsx")
saveWorkbook(wb, out_path, overwrite = TRUE)
# convertit les listes deroulantes (format x14, non reconnu par Google Sheets/LibreOffice) au format
# classique - voir apply_classic_dropdowns()/queue_list_validation() dans fonctions_nettoyage.R
apply_classic_dropdowns(out_path)

cat("\n--> Classeur exporte dans :", out_path, "\n")
cat("--> Faire remplir la colonne 'Que faire ?' (feuille A_verifier) et recoder les reponses autres\n",
    "    (feuille reponses_autres) par l'equipe de suivi, enregistrer le classeur (meme nom) dans '",
    Sys.getenv("suivi_final_remplis_path"), "'\n",
    "    (ou executer scripts/TEST_simuler_classeurs_remplis.R pour un test a blanc),\n",
    "    puis relancer 2_nettoyage_final.R pour construire le classeur final.\n", sep = "")

} else {

################################################################################
# PASSE 2 : APPLICATION DES CORRECTIONS, PUIS CONSTRUCTION DU CLASSEUR FINAL
################################################################################

cat("\n--> Classeur rempli trouve - construction du classeur final (2eme passe)...\n")

filled <- lapply(filled_files, function(f) read_review_sheet(f, sheet_name = SHEET_A_VERIFIER))
# "4_Verification_prix" (voir PASSE 1 ci-dessus) lue exactement comme "3_A_verifier" - meme structure de
# colonnes (REVIEW_HEADERS_FR), read_review_sheet() est directement reutilisable. Garde-fou sur les
# feuilles presentes (excel_sheets()) pour rester compatible avec un classeur "a verifier" produit AVANT
# l'ajout de cet onglet (aucune ligne a lire dans ce cas, pas une erreur).
filled_prix <- lapply(filled_files, function(f) {
  if (SHEET_VERIFICATION_PRIX %in% readxl::excel_sheets(f)) read_review_sheet(f, sheet_name = SHEET_VERIFICATION_PRIX) else NULL
})
filled <- do.call(plyr::rbind.fill, c(filled, filled_prix))
# "source" trace d'ou vient chaque ligne (classeur suivi a verifier rempli, ou l'un des deux onglets du
# journal de nettoyage additionnel) - utilisee plus bas (test de doublons enquete/variable) pour que le
# message d'erreur precise si le probleme vient du classeur suivi a verifier, du journal additionnel, ou
# d'une contradiction ENTRE les deux, plutot que de laisser l'utilisateur deviner.
filled$source <- "classeur suivi à vérifier"

# journal de nettoyage additionnel (optionnel) - desormais integre DIRECTEMENT dans le classeur "suivi
# a verifier" (onglets "Journal_additionnel"/"Journal_suppression", voir add_additional_log_sheets()
# dans fonctions_nettoyage.R et l'ecriture du classeur en PASSE 1 plus haut) plutot qu'un modele separe
# a telecharger/reimporter - complement manuel au classeur "A_verifier" pour des corrections que la
# revue automatique n'a pas signalees (voir prepare_additional_cleaning_log()/
# prepare_additional_deletion_log() pour la validation ligne par ligne). Lu depuis les MEMES fichiers
# que A_verifier/Verification_prix (filled_files, potentiellement plusieurs) - tolerant a un classeur
# genere par une version anterieure du pipeline sans ces onglets (prepare_*() renvoient alors NULL, pas
# une erreur). Concatene ici, avant arrange(), pour que tout le reste de cette passe (filled_field_level,
# deletion_log, combined_cleaning_log, ...) le traite exactement comme des lignes de plus du classeur
# rempli.
additional_logs <- lapply(filled_files, function(f) prepare_additional_cleaning_log(f, raw_data_all))
additional_logs <- additional_logs[!vapply(additional_logs, is.null, logical(1))]
if (length(additional_logs) > 0) {
  additional_log <- do.call(plyr::rbind.fill, additional_logs)
  cat("--> ", nrow(additional_log), " ligne(s) valide(s) dans le journal additionnel (onglet Journal_additionnel) - fusionnees avec le classeur rempli.\n", sep = "")
  additional_log$source <- "classeur suivi a verifier (onglet Journal_additionnel)"
  filled <- plyr::rbind.fill(filled, additional_log)
} else {
  cat("--> Aucune ligne dans le journal de nettoyage additionnel (onglet Journal_additionnel) - ignore.\n")
}

additional_deletion_logs <- lapply(filled_files, function(f) prepare_additional_deletion_log(f, raw_data_all))
additional_deletion_logs <- additional_deletion_logs[!vapply(additional_deletion_logs, is.null, logical(1))]
if (length(additional_deletion_logs) > 0) {
  additional_deletion_log <- do.call(plyr::rbind.fill, additional_deletion_logs)
  cat("--> ", nrow(additional_deletion_log), " enquete(s) supplementaire(s) a supprimer (onglet Journal_suppression) - fusionnees avec le classeur rempli.\n", sep = "")
  additional_deletion_log$source <- "classeur suivi a verifier (onglet Journal_suppression)"
  filled <- plyr::rbind.fill(filled, additional_deletion_log)
} else {
  cat("--> Onglet 'Journal_suppression' vide ou absent de tous les classeurs remplis - ignore.\n")
}

filled <- filled %>% arrange(uuid, question)

# les champs "calculate" (<item>_price_item et leurs eventuels intermediaires, voir
# get_calculated_price_field_names()/recompute_price_items() dans fonctions_nettoyage.R) sont EXCLUS
# ici, AVANT filled_field_level/filled_no_deletion (utilisees respectivement par create_clean_data() et
# par la construction de combined_cleaning_log plus bas) : recompute_price_items() les ecrase de toute
# facon INCONDITIONNELLEMENT plus bas a partir des vraies variables source - appliquer/journaliser quand
# meme une correction saisie dessus (Verification_prix) n'aurait aucun effet final ET creerait une
# entree de journal en double avec celle du recalcul automatique (bug constate : review_cleaning()
# signalait a la fois "Changes were not applied" et "Duplicated entry" pour le meme champ). La note
# ajoutee sur la cellule "Que faire ?" de ces lignes (voir write_review_sheet()) previent deja l'equipe
# de suivi que toute saisie ici sera ignoree - ce filtre en est la contrepartie cote traitement. Filtre
# applique sur `filled` (pas seulement filled_field_level) pour que combined_cleaning_log (base sur
# filled_no_deletion, derivee separement de `filled`) herite lui aussi de cette exclusion.
calculated_price_fields <- get_calculated_price_field_names(kobo_survey)
n_ignored_calculated <- sum(filled$question %in% calculated_price_fields, na.rm = TRUE)
if (n_ignored_calculated > 0) {
  cat("--> ", n_ignored_calculated, " signalement(s) sur un champ calcule automatiquement (prix) ignore(s) ",
     "- ces champs sont de toute facon recalcules plus bas (voir journal de nettoyage).\n", sep = "")
}
filled <- filled %>% filter(is.na(question) | !question %in% calculated_price_fields)

# les lignes "toute l'enquete" (quasi-doublon, valeurs manquantes) ont variable/question = NA - elles
# ne correspondent a aucune colonne reelle du jeu de donnees et ne peuvent donc etre appliquees qu'au
# niveau de l'enquete entiere (via l'action "Supprimer toute l'enquete"), jamais au niveau d'un champ.
filled_field_level <- filled %>% filter(!is.na(question))

test <- filled_field_level %>%
  select(uuid, question, old_value, change_type, new_value, source) %>% distinct() %>%
  group_by(uuid, question) %>%
  summarise(n = n(), sources = paste(unique(source), collapse = " ET "), .groups = "drop") %>%
  filter(n > 1)
if (nrow(test) > 0) {
  # "sources" indique precisement d'ou vient le conflit : le classeur suivi a verifier seul (deux lignes
  # differentes pour la meme enquete/variable dans A_verifier), le journal de nettoyage additionnel seul,
  # ou une CONTRADICTION entre le classeur suivi a verifier et le journal additionnel (les deux se
  # retrouvent dans "sources", separes par " ET ").
  detail <- paste0("uuid=", test$uuid, ", variable=", test$question, " (signalee ", test$n,
                   " fois avec des valeurs differentes - source(s) : ", test$sources, ")")
  stop("Certaines combinaisons enquete/variable sont signalees plus d'une fois, avec des valeurs ",
      "differentes - a resoudre avant de continuer :\n- ", paste(detail, collapse = "\n- "))
}

# review_cleaning_log() exige qu'au moins une ligne ait change_type == "change_response" - sur un test
# a blanc (scripts/TEST_simuler_classeurs_remplis.R), chaque signalement est rempli automatiquement
# avec "Ne rien changer" (no_action), il se peut donc qu'il n'y en ait aucun.
if ("change_response" %in% filled_field_level$change_type) {
  check_results <- review_cleaning_log(raw_dataset = raw_data_all, raw_data_uuid_column = "uuid",
                                       cleaning_log = filled_field_level, cleaning_log_uuid_column = "uuid",
                                       cleaning_log_question_column = "question",
                                       cleaning_log_new_value_column = "new_value",
                                       cleaning_log_change_type_column = "change_type",
                                       change_response_value = "change_response")
  print(check_results)
} else {
  cat("--> Aucune entree 'Corriger la reponse' dans le classeur rempli pour le moment - review_cleaning_log() ignore.\n")
}

# cleaningtools::create_clean_data() ne gere pas un cleaning_log a 0 ligne (leve une erreur "You have
# NAs in change_type option(s)" au lieu de ne rien faire) - situation reelle (pas seulement theorique)
# quand tous les signalements sont soit "Ne rien changer", soit des signalements au niveau de
# l'enquete entiere (quasi-doublon, % de valeurs manquantes - question = NA, donc exclus de
# filled_field_level) marques "Supprimer toute l'enquete". Dans ce cas il n'y a rien a corriger au
# niveau du champ : repartir simplement des donnees brutes (la suppression elle-meme est appliquee
# plus bas, explicitement, via deletion_log - voir REVUE DU NETTOYAGE).
if (nrow(filled_field_level) > 0) {
  clean_data <- create_clean_data(raw_dataset = raw_data_all, raw_data_uuid_column = "uuid",
                                  cleaning_log = filled_field_level, cleaning_log_uuid_column = "uuid",
                                  cleaning_log_question_column = "question",
                                  cleaning_log_new_value_column = "new_value",
                                  cleaning_log_change_type_column = "change_type",
                                  change_response_value = "change_response",
                                  NA_response_value = "blank_response",
                                  no_change_value = "no_action",
                                  remove_survey_value = "remove_survey")
} else {
  cat("--> Aucune correction au niveau du champ dans le classeur rempli - create_clean_data() ignore.\n")
  clean_data <- raw_data_all
}

################################################################################
# APPLIQUER LES MODIFICATIONS DES REPONSES "AUTRE"
################################################################################

# les reponses "autre" vivent dans la feuille "reponses_autres" du meme classeur "a verifier" que les
# verifications (voir add_other_responses_sheet() / 2_nettoyage_final.R passe 1) - memes fichiers que
# filled_files ci-dessus, feuille differente.
cleaning_log_other <- data.frame(uuid = character(), question = character(), change_type = character(),
                                 old_value = character(), new_value = character())

or.files <- filled_files

if (length(or.files) > 0) {
  # traduction "francais affiche -> nom technique" factorisee dans read_other_responses_sheet()
  # (fonctions_nettoyage.R), reutilisee ici et par la reprise des decisions d'un classeur precedent en
  # PASSE 1 ci-dessus - derivee DIRECTEMENT de OTHER_HEADERS_FR plutot qu'une correspondance maintenue
  # separement a la main, qui se desynchroniserait silencieusement des que les libelles affiches changent.
  or <- lapply(or.files, function(f) read_other_responses_sheet(f, sheet_name = SHEET_REPONSES_AUTRES))
  or <- do.call(plyr::rbind.fill, or)

  # filtrer "or" selon les uuid du jeu de donnees nettoye
  or <- or %>% filter(uuid %in% clean_data$uuid)

  or <- or %>%
    left_join(select(other_db, name, ref_question, ref_type = q_type), by = c("question_name" = "name")) %>%
    rename(name = "question_name", ref_name = "ref_question")

  # "Que faire ?" (voir OTHER_ACTION_LABELS_FR dans fonctions_nettoyage.R) <- code interne. Une valeur
  # non reconnue (faute de frappe, texte tape au lieu de choisi dans la liste deroulante) donnerait un
  # code NA silencieux sans ce garde-fou - la ligne serait ensuite ignoree plus loin sans que personne
  # ne s'en apercoive.
  other_action_rev <- setNames(names(OTHER_ACTION_LABELS_FR), unname(OTHER_ACTION_LABELS_FR))
  or <- or %>% mutate(action_code = unname(other_action_rev[action]))
  invalid_other_action <- or %>% filter(!is.na(action) & nzchar(action) & is.na(action_code))
  if (nrow(invalid_other_action) > 0) {
    detail <- paste0("uuid=", invalid_other_action$uuid, ", variable=", invalid_other_action$name,
                     " : valeur '", invalid_other_action$action, "' non reconnue pour \"Que faire ?\"")
    stop("Valeur(s) invalide(s) dans la colonne \"Que faire ?\" du classeur suivi a verifier (feuille ",
        "reponses_autres) - utilisez uniquement les options de la liste deroulante ('",
        paste(unname(OTHER_ACTION_LABELS_FR), collapse = "', '"), "') :\n- ", paste(detail, collapse = "\n- "))
  }

  # deux colonnes distinctes cote "Nouvelle reponse" (texte libre + choix existant - voir
  # OTHER_HEADERS_FR/add_other_responses_sheet() dans fonctions_nettoyage.R : Excel ne pouvant pas
  # verrouiller l'une des deux dynamiquement selon l'action choisie, les deux restent modifiables et
  # sont reconstituees ici en une seule valeur effective "new_value" par ligne). "Reattribuer..."
  # utilise en priorite "choix existant" (liste deroulante, le cas courant a un seul choix), avec repli
  # sur "texte libre" pour un select_multiple ou plusieurs choix sont necessaires a la fois (limitation
  # de la liste deroulante, qui n'en insere qu'un seul a la fois). Un avertissement (pas une erreur)
  # signale les rares lignes ou les deux colonnes sont remplies avec un contenu DIFFERENT - meme principe
  # que le conflit choix/INVALIDE deja gere ailleurs dans ce script (priorite + avertissement plutot que
  # bloquer tout le classeur pour une ambiguite sur une seule ligne).
  choice_empty <- is.na(or$new_value_choice) | !nzchar(or$new_value_choice)
  text_empty <- is.na(or$new_value_text) | !nzchar(or$new_value_text)
  both_filled_differ <- or$action_code == "reassign" & !choice_empty & !text_empty &
    or$new_value_choice != or$new_value_text
  if (any(both_filled_differ, na.rm = TRUE)) {
    conflict_rows <- or[which(both_filled_differ), ]
    detail <- paste0("uuid=", conflict_rows$uuid, ", variable=", conflict_rows$name,
                     " (choix existant='", conflict_rows$new_value_choice, "', texte libre='",
                     conflict_rows$new_value_text, "')")
    warning("\"Reattribuer a un choix existant\" : colonnes \"choix existant\" et \"texte libre\" ",
           "remplies avec un contenu different (classeur suivi a verifier, feuille reponses_autres) - ",
           "la colonne \"choix existant\" est utilisee en priorite, verifiez le journal de nettoyage ",
           "pour ces lignes :\n- ", paste(detail, collapse = "\n- "))
  }
  or <- or %>%
    mutate(new_value = case_when(
      action_code == "rewrite" ~ new_value_text,
      action_code == "reassign" & !(is.na(new_value_choice) | !nzchar(new_value_choice)) ~ new_value_choice,
      action_code == "reassign" ~ new_value_text,
      TRUE ~ NA_character_
    ))

  # "Reecrire" requiert un texte de remplacement, "Reattribuer..." requiert un choix existant - meme
  # verification que la colonne "new_value"/action_code == "change_response" dans A_verifier.
  missing_new_value <- or %>% filter(action_code %in% c("rewrite", "reassign") & (is.na(new_value) | !nzchar(new_value)))
  if (nrow(missing_new_value) > 0) {
    detail <- paste0("uuid=", missing_new_value$uuid, ", variable=", missing_new_value$name,
                     " (action='", OTHER_ACTION_LABELS_FR[missing_new_value$action_code], "')")
    stop("'Nouvelle reponse' est obligatoire quand \"Que faire ?\" = 'Reecrire' ou 'Reattribuer a un ",
        "choix existant' (classeur suivi a verifier, feuille reponses_autres) :\n- ", paste(detail, collapse = "\n- "))
  }

  # extraire les reponses "Reecrire" les plus recurrentes et sauvegarder le fichier - utile pour reperer
  # une reponse "autre" frequente qui merite d'etre ajoutee comme choix formel a l'outil kobo.
  if (any(or$action_code == "rewrite", na.rm = TRUE)) {
    res <- or %>%
      filter(action_code == "rewrite") %>%
      group_by(list_name) %>%
      mutate(n_list = n()) %>%
      group_by(list_name, new_value) %>%
      summarise(n_list = n_list[1], n_option = n()) %>%
      arrange(-n_list, -n_option) %>%
      select(-n_list)

    frequentes_dir <- Sys.getenv("final_workbook_dir")
    if (!dir.exists(frequentes_dir)) dir.create(frequentes_dir, recursive = TRUE)
    write_xlsx(res, paste0(frequentes_dir, Sys.Date(), "_reponses_autres_frequentes.xlsx"))
  }

  or_remove <- filter(or, action_code == "remove")
  # add_to_cleaning_log_other_recode()/_one()/_multiple() (fonctions_nettoyage.R) attendent x$existing_other
  # (nom historique) - renomme ici plutot que de toucher a ces fonctions deja testees.
  or_reassign <- filter(or, action_code == "reassign") %>% rename(existing_other = new_value)
  or_rewrite <- filter(or, action_code == "rewrite")

  print(paste("Nombre de reponses a supprimer :", nrow(or_remove)))
  if (nrow(or_remove) > 0) for (r in 1:nrow(or_remove)) add_to_cleaning_log_other_remove(or_remove[r, ])

  print(paste("Nombre de reponses a reattribuer a un choix existant :", nrow(or_reassign)))
  if (nrow(or_reassign) > 0) for (r in 1:nrow(or_reassign)) add_to_cleaning_log_other_recode(or_reassign[r, ])

  print(paste("Nombre de reponses reecrites :", nrow(or_rewrite)))
  if (nrow(or_rewrite) > 0) {
    cl_or_rewrite <- or_rewrite %>%
      select(uuid, name, response, new_value) %>%
      rename(question = name, old_value = response) %>%
      mutate(change_type = "Rewriting other response")
    cleaning_log_other <- rbind(cleaning_log_other, cl_or_rewrite)
  }

  # distinct() AVANT le comptage (pas apres) - meme ordre que le controle equivalent sur
  # filled_field_level plus haut. Sans cela, une ligne parfaitement IDENTIQUE ajoutee deux fois par
  # erreur (aucune vraie contradiction) declenche quand meme le stop() ci-dessous. Cas reel : une
  # reponse "autre" dont la question de reference n'a pas de "relevant" exploitable retombe sur name ==
  # ref_name (repli auto-referentiel, voir get_other_labels()/other_db dans fonctions_nettoyage.R) - si
  # ref_type vaut alors "text", add_to_cleaning_log_other_remove() ajoute deux lignes avec exactement le
  # meme question/old_value/new_value (une pour "retirer le texte de la reponse", une pour "retirer
  # l'entree liee" - la meme colonne dans ce cas precis), un doublon inoffensif, pas une contradiction.
  cleaning_log_other <- cleaning_log_other %>% distinct()

  t <- cleaning_log_other %>% group_by(uuid, question) %>% summarise(n = n(), .groups = "drop") %>% filter(n > 1)
  if (nrow(t) > 0) {
    detail <- paste0("uuid=", t$uuid, ", variable=", t$question, " (", t$n, " modifications contradictoires)")
    stop("Doublons detectes dans le recodage des reponses autre (classeur suivi a verifier, feuille ",
        "reponses_autres) - la meme variable recoit plusieurs modifications differentes pour ces lignes, ",
        "verifiez qu'une seule action a ete remplie par reponse 'autre' :\n- ", paste(detail, collapse = "\n- "))
  }

  # modification -- appliquer les changements aux donnees nettoyees
  if (nrow(cleaning_log_other) > 0) {
    for (r in 1:nrow(cleaning_log_other)) {
      b <- clean_data[[cleaning_log_other$question[r]]][clean_data$uuid == cleaning_log_other$uuid[r]]
      clean_data[[cleaning_log_other$question[r]]][clean_data$uuid == cleaning_log_other$uuid[r]] <- cleaning_log_other$new_value[r]
      a <- clean_data[[cleaning_log_other$question[r]]][clean_data$uuid == cleaning_log_other$uuid[r]]
      print(paste0(r, "-", cleaning_log_other$uuid[r], ": ", b, " --> ", a))
    }
  }
} else {
  cat("--> Aucun classeur de reponses autre trouve - recodage des reponses autre ignore.\n")
}

################################################################################
# REVUE DU NETTOYAGE
################################################################################

deletion_log <- filled %>% filter(change_type == "remove_survey")

# create_clean_data() (ci-dessus) ne supprime une enquete de clean_data que si sa ligne de
# suppression lui a ete transmise via filled_field_level - or les signalements au niveau de
# l'enquete entiere (question = NA : quasi-doublon, % de valeurs manquantes) en sont TOUJOURS exclus,
# meme quand ils sont marques "Supprimer toute l'enquete". Sans ce filtre explicite, une telle
# enquete resterait presente dans clean_data alors meme que deletion_log (et donc le classeur final)
# affirme qu'elle a ete supprimee. Filtre applique explicitement ici, quelle que soit l'origine du
# signalement, plutot que de dependre du mecanisme interne (partiel) de create_clean_data().
clean_data <- clean_data %>% filter(!uuid %in% deletion_log$uuid)

# recalcule <item>_price_item (et son eventuel intermediaire <item>_<taille>_price_unit_item) DIRECTEMENT
# depuis les variables source de clean_data (std_unit_item/other_unit_item/price_unit_item - voir "A
# verifier" plus haut, ou ces variables sont desormais affichees ensemble comme un bloc modifiable pour
# chaque article), qui peuvent avoir ete corrigees par le journal de nettoyage. Necessaire car
# create_clean_data() (ci-dessus) applique les corrections champ par champ mais ne re-execute PAS les
# formules "calculate" de l'outil kobo - _price_item resterait sinon celui calcule par Kobo au moment de
# la collecte, jamais mis a jour. Voir recompute_price_items() dans scripts/fonctions_nettoyage.R ; les
# colonnes EUR/USD ci-dessous (RECALCUL DES PRIX EN EUROS/DOLLARS) en dependent directement, et se
# recalculent donc elles aussi correctement a partir des valeurs a jour.
price_recompute <- recompute_price_items(clean_data, kobo_survey)
clean_data <- price_recompute$df
price_recalc_log <- price_recompute$log
if (nrow(price_recalc_log) > 0) {
  cat("--> ", nrow(price_recalc_log), " valeur(s) de champ \"calculate\" (prix) recalculee(s) suite a une ",
     "correction d'une variable source (journalisee(s)).\n", sep = "")
}

# "days_db" (stock_duration_days/restock_wait_days, voir plus haut) : une valeur de 0 jour est
# techniquement valide (rupture immediate) mais ambigue une fois combinee au calcul RE.1 du SFM
# (jours de stock restant - jours d'attente de reapprovisionnement, voir scripts/mfs.R) - convertie en 1
# ici, demande utilisateur. Journalisee comme les autres corrections automatiques (voir ras_log plus
# haut) pour rester tracable dans le classeur final, fusionnee dans combined_cleaning_log ci-dessous.
days_zero_cols <- intersect(days_db, colnames(clean_data))
days_zero_log <- tibble(uuid = character(0), question = character(0), old_value = character(0),
                        change_type = character(0), new_value = character(0), issue = character(0),
                        comment = character(0))
for (col in days_zero_cols) {
  numeric_vals <- suppressWarnings(as.numeric(clean_data[[col]]))
  is_zero <- !is.na(numeric_vals) & numeric_vals == 0
  if (any(is_zero)) {
    # as.character() explicite sur old_value : "days_db" fait partie de cols_numeric
    # (0_configuration.R) et a deja ete converti en numerique par read_raw_data() - sans cette
    # conversion, bind_rows() plus bas (fusion avec ras_log/cleaning_log_other, tous en texte) echoue
    # ("Can't combine ..$old_value <character> and ..$old_value <integer>").
    days_zero_log <- bind_rows(days_zero_log, tibble(
      uuid = clean_data$uuid[is_zero], question = col,
      old_value = as.character(clean_data[[col]][is_zero]),
      change_type = "change_response", new_value = "1",
      issue = "Valeur de 0 jour convertie en 1",
      comment = "Convertie automatiquement dans le classeur final"
    ))
    clean_data[[col]][is_zero] <- "1"
  }
}
if (nrow(days_zero_log) > 0) {
  cat("--> ", nrow(days_zero_log), " valeur(s) de 0 jour converties automatiquement en 1.\n", sep = "")
}

filled_no_deletion <- filled %>%
  filter(change_type != "remove_survey") %>%
  filter(!uuid %in% deletion_log$uuid) %>%
  filter(!is.na(question)) %>%
  mutate(new_value = ifelse(change_type == "no_action", old_value, new_value))

cleaning_log_other <- cleaning_log_other %>%
  mutate(issue = "Reponse autre - a verifier", comment = NA_character_)

combined_cleaning_log <- filled_no_deletion %>%
  select(uuid, question, change_type, old_value, new_value, issue, comment) %>%
  rbind(cleaning_log_other) %>%
  rbind(ras_log %>% select(uuid, question, change_type, old_value, new_value, issue, comment)) %>%
  rbind(days_zero_log %>% select(uuid, question, change_type, old_value, new_value, issue, comment)) %>%
  rbind(price_recalc_log %>% select(uuid, question, change_type, old_value, new_value, issue, comment))

# raw_data_all_avant_ras (avant la conversion automatique RAS -> vide, voir plus haut) est utilise ici,
# pas raw_data_all (deja converti) - sinon review_cleaning() ne verrait plus aucune difference pour ces
# champs (raw ET clean auraient deja la meme valeur vide) et ne pourrait pas confirmer que les entrees
# ras_log ci-dessus correspondent bien a un changement reellement survenu.
review_of_cleaning <- review_cleaning(raw_dataset = raw_data_all_avant_ras,
                                      raw_dataset_uuid_column = "uuid",
                                      clean_dataset = clean_data,
                                      clean_dataset_uuid_column = "uuid",
                                      cleaning_log = combined_cleaning_log,
                                      cleaning_log_uuid_column = "uuid",
                                      cleaning_log_question_column = "question",
                                      cleaning_log_new_value_column = "new_value",
                                      cleaning_log_change_type_column = "change_type",
                                      cleaning_log_old_value_column = "old_value",
                                      deletion_log = deletion_log,
                                      deletion_log_uuid_column = "uuid")

print(review_of_cleaning)

################################################################################
# RECALCUL DES PRIX EN EUROS
################################################################################

eur_rate <- as.numeric(Sys.getenv("eur_xrate"))

# Le champ "calculate" kobo de chaque article pour son prix converti en euros est recalcule ici a
# partir de la formule "calculation" de l'outil kobo lui-meme (ex. round(${maize_price_item} div
# 655.95)), plutot que d'une liste d'articles codee en dur - voir le commentaire detaille dans
# scripts/fonctions_nettoyage.R (ou l'ancien 3_finalisation_et_classeur_final.R). Un nouvel article
# ajoute a l'outil kobo est repris ici automatiquement, sans modifier ce script.
eur_calc_fields <- kobo_survey %>%
  filter(type == "calculate", str_detect(calculation, "_price_item\\}")) %>%
  mutate(source_item = str_match(calculation, "\\$\\{(\\w+_price_item)\\}")[, 2]) %>%
  filter(!is.na(source_item)) %>%
  select(source_item, target_field = name)

# libelle (row 1 de la feuille clean_data - voir apply_var_labels() dans fonctions_nettoyage.R) de
# chaque colonne EUR/USD nouvellement calculee : reprend le libelle francais de la question kobo
# d'origine (ex. "Le prix de 1 kg de maïs en euro :"), avec "USD" substitue a "euro"/"eur" pour la
# colonne dollar - stocke dans un vecteur nomme lu par apply_var_labels() (meme environnement
# d'execution que ce script, voir scripts/0_configuration.R), plutot qu'invente separement pour chaque
# devise.
price_conversion_labels <- setNames(character(0), character(0))

clean_data_updated <- clean_data
# journalise CHAQUE prix en euro recalcule (demande utilisateur) - le champ EUR lui-meme n'est jamais
# corrige directement par l'equipe de suivi (ce n'est pas une colonne editable des feuilles de revue),
# mais sa valeur est desormais TOUJOURS recalculee ici avec un arrondi a 2 decimales (round(..., 2) plus
# bas), alors que la valeur d'origine calculee par l'outil kobo au moment de la collecte n'avait pas
# forcement cette meme precision (constat utilisateur) - la quasi-totalite des lignes changent donc
# reellement de valeur de ce seul fait, meme sans aucune correction du prix source. Toutes ces lignes
# sont journalisees ici, sans filtrage sur l'origine du changement (arrondi ou correction reelle) :
# old_value = valeur EUR telle que stockee avant ce recalcul, new_value = valeur EUR fraichement
# recalculee. Aucune ligne ajoutee uniquement quand les deux sont deja rigoureusement identiques.
eur_recalc_log <- tibble(uuid = character(0), question = character(0), old_value = character(0),
                         change_type = character(0), new_value = character(0), issue = character(0),
                         comment = character(0))
for (i in seq_len(nrow(eur_calc_fields))) {
  source_item <- eur_calc_fields$source_item[i]
  target_field <- eur_calc_fields$target_field[i]
  if (source_item %in% colnames(clean_data_updated)) {
    new_vals <- round(as.numeric(clean_data_updated[[source_item]]) / eur_rate, 2)
    if (target_field %in% colnames(clean_data_updated)) {
      old_vals_chr <- as.character(clean_data_updated[[target_field]])
      new_vals_chr <- as.character(new_vals)
      changed <- (is.na(old_vals_chr) != is.na(new_vals_chr)) |
        (!is.na(old_vals_chr) & !is.na(new_vals_chr) & old_vals_chr != new_vals_chr)
      if (any(changed)) {
        eur_recalc_log <- bind_rows(eur_recalc_log, tibble(
          uuid = clean_data_updated$uuid[changed], question = target_field,
          old_value = old_vals_chr[changed], change_type = "change_response",
          new_value = new_vals_chr[changed],
          issue = "Prix en euro recalcule (arrondi a 2 decimales)",
          comment = NA_character_
        ))
      }
    }
    clean_data_updated[[target_field]] <- new_vals
  }
}
# ajoutee A LA SUITE du reste (demande utilisateur : "au bas du journal de nettoyage") - combined_cleaning_log
# n'est plus relu que par cleaning_log_final plus bas (deja utilise par review_cleaning() avant ce point,
# donc aucun risque de fausser cette verification avec des lignes ajoutees apres coup).
combined_cleaning_log <- bind_rows(combined_cleaning_log, eur_recalc_log)
if (nrow(eur_recalc_log) > 0) {
  cat("--> ", nrow(eur_recalc_log), " prix en euro recalcule(s) suite a une correction du prix source ",
      "(journalise(s) au bas du journal de nettoyage).\n", sep = "")
}

################################################################################
# RECALCUL DES PRIX EN DOLLARS US (USD)
################################################################################
# meme principe que la conversion EUR ci-dessus (reutilise la meme liste eur_calc_fields), mais l'outil
# kobo ne definit PAS de champ "calculate" dedie au dollar (seul l'euro y est prevu) : on cree donc une
# NOUVELLE colonne par article, en reprenant le nom du champ EUR de CET article et en y substituant
# "usd" a "eur"/"euro" (ex. "maize_eur_xrate_official" -> "maize_usd_xrate_official",
# "jerrycan_price_euro_xrate_official" -> "jerrycan_price_usd_xrate_official" - la convention de nommage
# EUR n'est pas la meme pour tous les articles, donc on la reprend telle quelle plutot que d'en inventer
# une nouvelle). Placee juste apres sa colonne EUR correspondante (relocate). Contrairement au taux EUR/
# XAF (parite fixe, ne change jamais), le taux USD/XAF flotte sur le marche - usd_xrate doit etre mis a
# jour periodiquement (voir .Renviron / configuration du tableau de bord).
usd_rate <- as.numeric(Sys.getenv("usd_xrate"))
for (i in seq_len(nrow(eur_calc_fields))) {
  source_item <- eur_calc_fields$source_item[i]
  eur_target_field <- eur_calc_fields$target_field[i]
  usd_target_field <- sub("(^|_)euro?(_|$)", "\\1usd\\2", eur_target_field)
  if (source_item %in% colnames(clean_data_updated) && eur_target_field %in% colnames(clean_data_updated)) {
    clean_data_updated[[usd_target_field]] <- round(as.numeric(clean_data_updated[[source_item]]) / usd_rate, 2)
    clean_data_updated <- clean_data_updated %>% relocate(all_of(usd_target_field), .after = all_of(eur_target_field))

    eur_label <- kobo_survey$`label::french`[kobo_survey$name == eur_target_field]
    if (length(eur_label) > 0 && !is.na(eur_label[1])) {
      usd_label <- if (str_detect(eur_label[1], regex("eur", ignore_case = TRUE))) {
        str_replace(eur_label[1], regex("euro?", ignore_case = TRUE), "USD")
      } else {
        paste0(eur_label[1], " (USD)")
      }
      price_conversion_labels[usd_target_field] <- usd_label
    }
  }
}

clean_data_updated <- clean_data_updated %>% relocate(uuid)

################################################################################
# RETRAIT DES INFORMATIONS SENSIBLES
################################################################################

# "_id" (colonne systeme kobo optionnelle, absente des exports sans historique de soumissions
# modifiees) a deja ete renommee "id" dans read_raw_data() si elle est presente (voir fonctions_
# nettoyage.R). "_index" est conservee et renommee "id" a son tour SEULEMENT si ce n'est pas deja le
# cas - sinon un vrai export Kobo qui a A LA FOIS "_id" ET "_index" (constate sur un export telecharge
# directement depuis Kobo) ferait planter le rename() (deux colonnes "id"). Si "id" existe deja,
# "_index" est simplement supprimee comme toute autre colonne systeme prefixee par "_".
# --- Colonnes ORPHELINES de l'outil actuellement charge (absentes de kobo_survey$name) ET
# ENTIEREMENT VIDES sur tout le jeu de donnees - demande utilisateur : retirer TOUTE question de type
# "note" (texte d'introduction Kobo, jamais de saisie possible - donc TOUJOURS entierement vide, par
# construction meme du type "note", pas seulement dans ce jeu de donnees precis), y compris celles que
# l'outil actuellement charge ne reconnait plus/pas encore (renommees/retirees/ajoutees depuis une
# version differente du formulaire - le contenu, pas le nom de colonne, est alors le seul signal fiable
# disponible, puisqu'aucune metadonnee de type n'existe pour une variable absente de kobo_survey).
# Exclut les colonnes systeme ("_" en tete, deja gerees separement ci-dessous) et les colonnes binaires
# select_multiple ("/" dans le nom, deja rattachees a leur question parente - jamais a retirer meme si
# 100% des repondants ont laisse un choix rare de cote ce mois-ci, ce qui n'a rien d'anormal).
orphan_empty_cols <- setdiff(colnames(clean_data_updated), kobo_survey$name)
orphan_empty_cols <- orphan_empty_cols[!str_starts(orphan_empty_cols, "_") & !str_detect(orphan_empty_cols, "/")]
orphan_empty_cols <- orphan_empty_cols[vapply(orphan_empty_cols, function(v) {
  x <- clean_data_updated[[v]]
  all(is.na(x) | !nzchar(as.character(x)))
}, logical(1))]

rename_index_to_id <- "_index" %in% colnames(clean_data_updated) && !"id" %in% colnames(clean_data_updated)
remove_cols <- c(
  # "decimal" retire de cette liste (bug constate) : les 2 seuls champs "decimal" de l'outil sont
  # vegetable_oil_other_unit_item/water_other_unit_item - precisement les champs "_other_unit_item" que
  # le commentaire ci-dessous dit explicitement CONSERVER, silencieusement retires malgre cette intention.
  intersect(colnames(clean_data_updated), filter(kobo_survey, type %in% c("note", "geopoint"))$name),
  "audit", "audit_URL", "deviceid", "enum_id", "shop_name", "organisation_name_label",
  # consent est CONSERVEE dans clean_data (demande utilisateur) - ne plus la retirer ici.
  # admin1_label/admin2_label/admin3_label sont conservees (nouvelles variables, consignees dans
  # variable_tracker ci-dessous) - c'est le mapping PCode -> nom depuis la feuille choices du kobo,
  # pas une colonne technique.
  if (rename_index_to_id) {
    setdiff(colnames(clean_data_updated)[str_starts(colnames(clean_data_updated), "_")], "_index")
  } else {
    colnames(clean_data_updated)[str_starts(colnames(clean_data_updated), "_")]
  },
  orphan_empty_cols
)
# NOTE : les colonnes "_std_unit_item" (ex. maize_std_unit_item) ET "_other_unit_item" (ex.
# maize_other_unit_item) sont conservees (demande utilisateur - plus retirees comme avant) - elles
# indiquent l'unite/quantite "autre" declaree pour le prix collecte, ce sont des donnees pertinentes et
# non un champ technique.
remove_cols <- intersect(remove_cols, colnames(clean_data_updated))

clean_data_final <- select(clean_data_updated, -all_of(remove_cols))
if (rename_index_to_id) clean_data_final <- rename(clean_data_final, id = `_index`)

# donnees brutes pour validation - anonymisees de la meme maniere, sauf les colonnes qui n'existent
# qu'apres le nettoyage (admin1_label/admin2_label/admin3_label sont conservees - voir remove_cols ci-dessus)
remove_cols_raw <- c(
  intersect(colnames(raw_data_all), filter(kobo_survey, type %in% c("geopoint"))$name),
  "audit", "audit_URL", "deviceid", "enum_id", "shop_name", "organisation_name_label"
)
remove_cols_raw <- intersect(remove_cols_raw, colnames(raw_data_all))

raw_data_all_validation <- select(raw_data_all, -all_of(remove_cols_raw))
# voir rename_index_to_id ci-dessus (clean_data_final) - meme garde contre la collision "_id"/"_index"
if ("_index" %in% colnames(raw_data_all_validation) && !"id" %in% colnames(raw_data_all_validation)) {
  raw_data_all_validation <- rename(raw_data_all_validation, id = `_index`)
}

################################################################################
# CONSTRUCTION DU CLASSEUR FINAL SELON LE MODELE DE LA CHECKLIST
################################################################################

enum_lookup <- raw_data_all %>% select(uuid, enum_id, deviceid) %>% distinct()

# reason_deletion vient de la colonne "Commentaire" de la feuille A_verifier (voir REVIEW_HEADERS_FR/
# read_review_sheet() dans fonctions_nettoyage.R, qui la lit deja sous le nom interne "comment") -
# c'est la ou l'equipe de suivi explique pourquoi elle a choisi "Supprimer toute l'enquete". Si une
# enquete est signalee plusieurs fois, distinct(uuid, .keep_all = TRUE) garde le commentaire de la
# premiere ligne rencontree.
deletion_log_final <- deletion_log %>%
  distinct(uuid, .keep_all = TRUE) %>%
  left_join(enum_lookup, by = "uuid") %>%
  rename(reason_deletion = comment) %>%
  select(uuid, reason_deletion, enum_id)

cleaning_log_final <- combined_cleaning_log %>%
  filter(!uuid %in% deletion_log_final$uuid) %>%
  left_join(enum_lookup, by = "uuid") %>%
  rename(variable = question, feedback = comment) %>%
  select(uuid, variable, issue, feedback, change_type, old_value, new_value, enum_id) %>%
  distinct() %>%
  arrange(uuid, variable)

justify_removed_variable <- function(v) {
  case_when(
    v %in% c("deviceid", "enum_id", "shop_name") ~ "Confidentialite / DPI",
    v %in% filter(kobo_survey, type == "geopoint")$name ~ "Coordonnees GPS - DPI",
    v %in% orphan_empty_cols ~ paste0("Texte d'introduction Kobo (note) ou variable similaire, entierement ",
                                      "vide et absente de l'outil actuellement charge - orpheline d'une ",
                                      "version differente du formulaire, non pertinente pour l'analyse"),
    v %in% filter(kobo_survey, type == "note")$name ~ "Texte d'introduction, non pertinent pour l'analyse",
    TRUE ~ "Variable retiree lors du nettoyage final"
  )
}

# suivies meme quand absentes de kobo_survey$name (ex. orphan_empty_cols, orphelines d'une version
# differente du formulaire, retirees ci-dessus mais jamais definies dans l'outil ACTUELLEMENT charge) -
# sinon intersect(remove_cols, kobo_survey$name) les laisserait silencieusement hors du suivi.
variable_tracker_removed <- data.frame(variable = union(intersect(remove_cols, kobo_survey$name),
                                                         orphan_empty_cols)) %>%
  mutate(action = "Supprime", justification = justify_removed_variable(variable))

# variables presentes dans clean_data_final qui ne sont PAS des questions de l'outil kobo (calculee
# dynamiquement, comme variable_tracker_removed ci-dessus, plutot que maintenue a la main - un nouvel
# ajout futur, ex. les colonnes USD, est ainsi automatiquement repris ici sans modifier ce script).
# Exclut : colonnes systeme kobo standard, toujours presentes dans tout export (renommees depuis
# _uuid/_index/__version__/_submission_time/admin1/admin2/admin3, ou derivees comme submission_date -
# pas des "variables ajoutees" specifiques a ce pipeline) et les colonnes binaires select_multiple
# (suffixe "/code", deja rattachees a leur question kobo parente).
kobo_system_cols <- c("uuid", "id", "version", "submission_time", "submission_date",
                      "admin1_code", "admin2_code", "admin3_code")
added_vars <- setdiff(colnames(clean_data_final), c(kobo_survey$name, kobo_system_cols))
added_vars <- added_vars[!str_detect(added_vars, "/")]

justify_added_variable <- function(v) {
  case_when(
    v %in% c("admin1_label", "admin2_label", "admin3_label") ~
      paste0("Nom correspondant au PCode ", sub("_label$", "_code", v),
            " (mapping depuis la feuille choices du questionnaire Kobo)"),
    str_ends(v, "_usd_xrate_official") ~
      "Prix converti en dollars US (USD) - calcule a partir du prix final (apres corrections du journal de nettoyage), sur le meme principe que la conversion en euro deja prevue par l'outil kobo (voir 'RECALCUL DES PRIX EN DOLLARS US' dans 2_nettoyage_final.R)",
    TRUE ~ "Variable ajoutee lors du nettoyage final"
  )
}

variable_tracker_added <- data.frame(variable = added_vars) %>%
  mutate(action = "Ajoute", justification = justify_added_variable(variable))

variable_tracker_final <- bind_rows(variable_tracker_removed, variable_tracker_added)

final_survey_count <- nrow(clean_data_final)
# la periode de collecte provient de la variable kobo "date" (voir `period` ci-dessus), pas de submission_date
collection_start <- period$collection_start
collection_end <- period$collection_end
assessment_year <- substr(period$assessment_month, 1, 4)
admin1_covered <- unique(raw_data_all$admin1_label)
admin2_covered <- unique(raw_data_all$admin2_label)
partners_covered <- unique(raw_data_all$organisation_name_label)

output_dir <- Sys.getenv("final_workbook_dir")
if (!dir.exists(output_dir)) dir.create(output_dir, recursive = TRUE)
# convention de nommage INITIATIVE_CODEPAYS3LETTRES_Type-Produit_Nom-Produit_Mois-Annee de la checklist
output_path <- paste0(output_dir, "REACH_CAR_Dataset_JMMI_", period$assessment_month_label, "-", assessment_year, ".xlsx")

build_final_workbook(
  raw_data_export = raw_data_all_validation,
  clean_data_export = clean_data_final,
  cleaning_log_final = cleaning_log_final,
  deletion_log_final = deletion_log_final,
  variable_tracker_final = variable_tracker_final,
  output_path = output_path,
  final_survey_count = final_survey_count,
  collection_start = collection_start,
  collection_end = collection_end,
  admin1_covered = admin1_covered,
  admin2_covered = admin2_covered,
  partners_covered = partners_covered,
  assessment_month_label = period$assessment_month_label,
  assessment_year = assessment_year
)

cat("\n--> Classeur combine final ecrit dans :", output_path, "\n")

}
