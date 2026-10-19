################################################################################
# RCA JMMI - FONCTIONS PARTAGEES POUR L'ANALYSE MENSUELLE
################################################################################
# Fonctions utilisees par 1_analyse_complete.R : detection automatique de la periode d'evaluation,
# etiquetage du jeu de donnees, generation automatique de la liste d'analyse (LOA) a partir de l'outil
# kobo, et calculs du Panier Minimum de Aliments et Services (PMAS/MEB).
################################################################################

##############################################################################################################
# PERIODE D'EVALUATION (annee/mois)
##############################################################################################################
# Determine l'annee et le mois du cycle directement depuis la colonne "date" du classeur nettoye importe
# (le jour de la visite du marche), plutot que de demander a l'utilisateur de les retaper - cette
# information est deja presente dans le classeur (nom de fichier ET feuille README), la deriver
# directement des donnees est plus fiable qu'un re-parsing du nom de fichier ou du texte libre du README.
# Meme logique que get_assessment_period() dans scripts/fonctions_nettoyage.R du pipeline de nettoyage,
# dupliquee ici (plutot que sourcee depuis l'autre pipeline) pour que ce script d'analyse reste autonome.
get_assessment_period_from_data <- function(df, date_col = "date") {
  mois_fr <- c("Janvier", "Fevrier", "Mars", "Avril", "Mai", "Juin",
              "Juillet", "Aout", "Septembre", "Octobre", "Novembre", "Decembre")

  d <- as.Date(df[[date_col]])
  d <- d[!is.na(d)]
  if (length(d) == 0) {
    stop("Impossible de determiner automatiquement l'annee/le mois du cycle : la colonne '", date_col,
        "' du classeur importe ne contient aucune date valide.")
  }
  start_date <- min(d)

  list(
    assessment_month_label = mois_fr[as.integer(format(start_date, "%m"))],
    assessment_year = format(start_date, "%Y")
  )
}

##############################################################################################################
# LECTURE DE L'OUTIL KOBO
##############################################################################################################
# feuille "choices" : colonne de libelle "label" ou "label::french" ("label::french" retenue si les deux
# existent) - meme logique que standardise_choices_label() dans scripts/fonctions_nettoyage.R, dupliquee
# ici (plutot que sourcee depuis l'autre pipeline) pour que ce script d'analyse reste autonome.
standardise_choices_label <- function(choices) {
  if ("label::french" %in% names(choices)) {
    choices <- choices %>% select(-any_of("label")) %>% rename(label = `label::french`)
  }
  if (!"label" %in% names(choices)) {
    stop("Feuille \"choices\" de l'outil kobo : aucune colonne \"label\" ni \"label::french\" trouvee.")
  }
  choices
}

##############################################################################################################
# ETIQUETAGE (name -> label, code -> label) - "Donnees nettoyees"
##############################################################################################################

# convertit une colonne "calculate"/"integer"/"decimal" en numerique (les select_one/select_multiple
# restent du texte, pour etre recodes en libelles par recode_column_labels() ci-dessous)
convert_column_type <- function(df, col, tool_survey) {
  if (col %in% tool_survey$name) {
    q <- tool_survey[tool_survey$name == col, ]
    if (q$q_type[1] %in% c("calculate", "integer", "decimal")) return(suppressWarnings(as.numeric(df[[col]])))
  }
  df[[col]]
}

# transforme un nom de colonne technique (et son eventuel suffixe "/code" pour les binaires
# select_multiple) en libelle francais lisible - reprend le meme decoupage "/" que le reste du pipeline
# de nettoyage (voir REACH_CAR_JMMI_scripts/rca_jmmi_cleaning)
name2label_question <- function(tool_survey, tool_choices, col) {
  if (str_detect(col, "/")) {
    name <- str_split(col, "/")[[1]][1]
    c_name <- paste0(tail(str_split(col, "/")[[1]], -1), collapse = "/")
  } else {
    name <- col
    c_name <- NA
  }
  if (name %in% tool_survey$name) {
    q <- tool_survey[tool_survey$name == name, ][1, ]
    label <- q$`label::french`
    if (is.na(label) || q$q_type == "note") label <- name
    if (!is.na(c_name)) {
      c_label <- tool_choices$label[tool_choices$list_name == q$list_name & tool_choices$name == c_name]
      label <- if (length(c_label) == 0 || is.na(c_label[1])) label else paste0(label, "/", c_label[1])
    }
  } else {
    label <- col
  }
  label
}

# recode les valeurs d'une question select_one/select_multiple (code -> libelle francais) - les
# colonnes binaires des select_multiple (suffixe "/code") sont laissees telles quelles (0/1), seul le
# NOM de colonne est retraduit en libelle par name2label_question() ci-dessus
recode_column_labels <- function(df, tool_survey, tool_choices) {
  for (col in colnames(df)) {
    if (str_detect(col, "/")) next  # colonne binaire select_multiple : valeur 0/1, pas de recodage
    if (!col %in% tool_survey$name) next
    q <- tool_survey[tool_survey$name == col, ][1, ]
    if (q$q_type != "select_one" || is.na(q$list_name)) next
    choices <- tool_choices %>% filter(list_name == q$list_name) %>% select(name, label)
    df[[col]] <- choices$label[match(as.character(df[[col]]), choices$name)] %>%
      { ifelse(is.na(.) & !is.na(df[[col]]), df[[col]], .) }  # code inconnu -> garder tel quel plutot que perdre l'info
  }
  df
}

# construit le jeu de donnees etiquete ("Donnees nettoyees") : valeurs recodees en libelles + noms de
# colonnes traduits en libelles - a partir du clean_data brut (codes techniques)
build_labeled_dataset <- function(clean_data, kobo_survey, kobo_choices) {
  df <- clean_data
  for (col in colnames(df)) df[[col]] <- convert_column_type(df, col, kobo_survey)
  df <- recode_column_labels(df, kobo_survey, kobo_choices)
  new_names <- vapply(colnames(df), function(col) name2label_question(kobo_survey, kobo_choices, col), character(1))
  colnames(df) <- unname(new_names)
  df
}

##############################################################################################################
# GENERATION AUTOMATIQUE DE LA LISTE D'ANALYSE (LOA)
##############################################################################################################
# Construit la LOA directement depuis l'outil kobo (survey/choices) plutot que depuis un fichier DAP
# maintenu a la main (qui, comme le reste des scripts d'analyse decouverts ce mois-ci, referencait des
# variables qui n'existent plus) - chaque question select_one/select_multiple pertinente de l'outil
# devient une ligne de la LOA, plus une ligne "median" pour chaque variable de `price_vars` (prix des
# articles - absents de l'outil kobo lui-meme, calcules par le nettoyage, voir PRICE_ITEM_CATALOG dans
# 4_analyse_complete.R). Voir analysistools::create_analysis().
# noms des questions telles que nommees dans l'outil kobo (kobo_survey$name) - le pipeline de nettoyage
# les renomme en admin1_code/admin2_code/admin3_code dans clean_data, mais la LOA doit utiliser les noms
# D'ORIGINE puisqu'elle est generee depuis kobo_survey
LOA_EXCLUDE_VARS <- c(
  "admin1", "admin2", "admin3",  # deja representes par admin*_label (le group_var)
  "consent",                      # filtre de nettoyage, pas un indicateur de marche
  "shop_category_other", "infrastructure_shop_other"  # texte libre "autre" - pas de choices a analyser
)

# chaque variable est desagregee a DEUX niveaux dans la MEME loa (comme dans RCA_JMMI_loa.xlsx, qui
# melange des lignes group_var=NA - passe nationale - et group_var=<niveau choisi> - voir son contenu :
# 124 lignes group_var=NA suivies de 124 lignes group_var="Quelle est la commune?") : la passe nationale
# est placee EN PREMIER pour chaque variable, ce qui fait apparaitre le niveau national en premiere ligne
# du tableau pivote (create_table_group_x_variable() respecte l'ordre des lignes de la loa pour ordonner
# les groupes - constate empiriquement, aucun parametre de tri explicite disponible).
build_loa <- function(kobo_survey, group_var = "admin2_label", level = 0.95, exclude_vars = LOA_EXCLUDE_VARS,
                      price_vars = character(0)) {
  eligible <- kobo_survey %>%
    filter(q_type %in% c("select_one", "select_multiple")) %>%
    filter(!name %in% exclude_vars) %>%
    filter(!str_detect(name, "_other(_tx)?$")) %>%  # champs texte libre "si autre, precisez" (type="text" de toute facon, garde-fou)
    distinct(name, .keep_all = TRUE)

  all_vars <- bind_rows(
    tibble(
      analysis_type = if_else(eligible$q_type == "select_one", "prop_select_one", "prop_select_multiple"),
      analysis_var = eligible$name
    ),
    tibble(analysis_type = "median", analysis_var = price_vars)
  )

  grp <- group_var  # copie locale - evite toute ambiguite entre le parametre et la colonne du meme nom cree ci-dessous
  bind_rows(
    all_vars %>% mutate(group_var = NA_character_, level = level),
    all_vars %>% mutate(group_var = grp, level = level)
  )
}

##############################################################################################################
# CALCULS DU PANIER MINIMUM DE ALIMENTS ET SERVICES (PMAS / MEB)
##############################################################################################################
# Poids/formules repris des scripts d'analyse existants (archive/analysis.Rmd) - seuls les noms de
# colonnes ont ete mis a jour pour correspondre a l'outil kobo actuel (voir les memes renommages que
# dans scripts/mfs.R : maize_flour_350g_... -> maize_350g_..., etc.), les poids eux-memes n'ont pas ete
# modifies. A verifier aupres de l'equipe si les quantites mensuelles supposees doivent evoluer.
MEB_FOOD_WEIGHTS <- list(
  maize_350g_price_unit_item = 12000 / 350,
  manioc_500g_price_unit_item = 30000 / 500,
  rice_500g_price_unit_item = 15000 / 500,
  bean_500g_price_unit_item = 18000 / 500,
  peanut_150g_price_unit_item = 6000 / 150,
  beef_meat_price_item = 2,
  vegetable_oil_price_item = 5,
  sugar_200g_price_unit_item = 5000 / 200,
  salt_150g_price_unit_item = 1000 / 150
)
MEB_NFI_WEIGHTS <- list(
  mosquito_net_price_item = 1 / 6,
  jerrycan_20l_price_unit_item = 1 / 6,
  bed_sheets_1piece_price_unit_item = 1 / 6,
  carpet_price_item = 1 / 6,
  tarpaulin_20m2_price_unit_item = 1 / 6,
  cooking_pot_5l_price_unit_item = 1 / 6
)
MEB_WASH_WEIGHTS <- list(
  soap_200g_price_unit_item = 10,
  bucket_20l_price_unit_item = 15 / 20 / 6
)
MEB_WEIGHTS_ALL <- c(MEB_FOOD_WEIGHTS, MEB_NFI_WEIGHTS, MEB_WASH_WEIGHTS)

# renvoie les variables PMAS/MEB attendues (MEB_WEIGHTS_ALL) mais absentes de `data` - typiquement une
# variable renommee/retiree de l'outil kobo d'un cycle a l'autre. AVANT ce garde-fou, une telle variable
# disparaissait silencieusement du panier (calculate_pmas_by_level() filtrait deja par
# `items %in% colnames(data)`, sans avertissement) : le panier se recalculait avec un article en moins,
# sans note ni erreur nulle part, et "meb_composition" continuait de lister l'article comme si de rien
# n'etait - demande utilisateur explicite (suite a l'audit de robustesse) : rendre ce cas visible.
check_meb_items_missing <- function(data) {
  setdiff(names(MEB_WEIGHTS_ALL), colnames(data))
}

##############################################################################################################
# SUFFISANCE DES DONNEES POUR LE CALCUL DU PMAS
##############################################################################################################
# Nombre minimum d'observations LOCALES (a un niveau donne - commune, sous-prefecture, ...) requises pour
# qu'un article soit considere "mesure localement" plutot qu'impute avec la mediane nationale de cet
# article. La viande est une exception (1 seule observation suffit) : article rare, peu de bouchers dans
# beaucoup de marches, exiger 3 observations le rendrait quasi toujours impute.
MEB_MIN_OBS_DEFAULT <- 3
MEB_MIN_OBS_OVERRIDE <- c(beef_meat_price_item = 1)

# seuils de suffisance (nombre d'articles mesures LOCALEMENT, pas imputes) pour decider quoi calculer :
#  - PMAS complet (17 articles au total) : calcule uniquement si au moins MEB_TOTAL_MIN_ITEMS articles
#    sont mesures localement (les autres restent imputes avec la mediane nationale).
#  - En-dessous de ce seuil, le total n'est PAS calcule, mais chaque panier peut quand meme l'etre
#    INDIVIDUELLEMENT si son propre seuil "autonome" (MEB_BASKET_MIN_STANDALONE) est atteint.
#  - Si le total EST calcule mais avec peu de marge (entre MEB_TOTAL_MIN_ITEMS et 14 articles mesures),
#    un panier est signale "a interpreter avec prudence" s'il n'atteint pas le seuil, plus souple,
#    MEB_BASKET_MIN_CAUTION.
MEB_TOTAL_MIN_ITEMS <- 12
# borne haute de la "zone de prudence" (entre MEB_TOTAL_MIN_ITEMS et cette valeur, le total EST calcule
# mais avec peu de marge - MEB_BASKET_MIN_CAUTION signale alors les paniers individuels concernes). Etait
# un "14" ecrit en dur directement dans la condition de calculate_pmas_by_level() ci-dessous - nommee ici
# pour que le texte explicatif (PMAS_EXPLANATION, 4_analyse_complete.R) puisse la citer dynamiquement au
# lieu de repeter "14" a la main (demande utilisateur, suite a l'audit de robustesse).
MEB_TOTAL_CAUTION_MAX_ITEMS <- 14
MEB_BASKET_MIN_STANDALONE <- c(food = 6, nfi = 4, wash = 2)
MEB_BASKET_MIN_CAUTION <- c(food = 6, nfi = 4, wash = 1)

# etape 1 (commune = marche) : mediane de PRIX INDIVIDUELS par article, au niveau du marche - "marche"
# et "commune" (admin3_label) coincident dans ce dispositif d'enquete, faute d'identifiant de marche
# distinct (plusieurs detaillants du MEME marche sont enquetes par commune, voir l'echantillonnage
# decrit dans write_suivi_intro_sheet() de fonctions_nettoyage.R). C'EST LE SEUL ENDROIT ou les prix
# INDIVIDUELS bruts interviennent - tout le reste (secours d'un marche a court de donnees, niveau
# national) est construit a partir de medianes deja calculees, jamais recalcule sur les prix individuels
# (demande utilisateur explicite). Secours sur la MEDIANE DES MEDIANES DE MARCHE (chaque marche compte
# pour une seule observation dans ce secours, quel que soit son nombre de vendeurs enquetes) si un marche
# a moins de min_obs prix locaux pour un article donne. Renvoie une valeur ET un indicateur "mesure
# localement" par (marche, article).
calculate_market_medians <- function(data, items, min_obs_vec) {
  vars <- data %>%
    select(admin3_label, all_of(items)) %>%
    mutate(across(all_of(items), as.numeric))

  market_median <- vars %>% group_by(admin3_label) %>%
    summarise(across(all_of(items), ~median(., na.rm = TRUE)), .groups = "drop")
  market_count <- vars %>% group_by(admin3_label) %>%
    summarise(across(all_of(items), ~sum(!is.na(.))), .groups = "drop")
  national_median_of_markets <- market_median %>% summarise(across(all_of(items), ~median(., na.rm = TRUE)))

  market_is_local <- market_count
  market_value <- market_median
  for (it in items) {
    market_is_local[[it]] <- market_count[[it]] >= min_obs_vec[[it]]
    market_value[[it]] <- if_else(market_is_local[[it]], market_median[[it]], national_median_of_markets[[it]][1])
  }
  list(value = market_value, is_local = market_is_local)
}

# calcule le PMAS (paniers alimentaire/non alimentaire/hygiene + total) au niveau Commune UNIQUEMENT -
# c'est le SEUL niveau ou les prix INDIVIDUELS bruts interviennent (etape 1, calculate_market_medians()
# ci-dessus : mediane de marche calculee directement a partir des prix individuels collectes). Le niveau
# National (calculate_pmas_national() ci-dessous) N'APPELLE PLUS cette fonction avec un niveau
# d'agregation different pour recomposer un panier a partir de prix par article agreges : il prend
# directement la MEDIANE des paniers communaux DEJA CALCULES par cette fonction (demande utilisateur
# explicite : "the original base data needs to be used only to compute the median at the market level ..
# all else is medians of already calculated medians") - la mediane d'une somme ponderee n'etant pas egale
# a la somme des medianes, recomposer le panier national a partir de medianes par article donnait un
# total different de la mediane des totaux communaux, ce qui ne correspondait pas a ce qui etait attendu.
calculate_pmas_by_level <- function(data, level_col = "admin3_label", min_obs_default = MEB_MIN_OBS_DEFAULT,
                                    min_obs_override = MEB_MIN_OBS_OVERRIDE) {
  stopifnot(identical(level_col, "admin3_label"))
  items <- names(MEB_WEIGHTS_ALL)
  missing_meb_items <- check_meb_items_missing(data)
  if (length(missing_meb_items) > 0) {
    cat("  ATTENTION : ", length(missing_meb_items), " article(s) PMAS/MEB absent(s) des donnees (variable(s) ",
       "introuvable(s) dans clean_data, probablement renommee(s)/retiree(s) de l'outil kobo) - EXCLU(S) DU ",
       "CALCUL DU PANIER sans imputation possible : ", paste(missing_meb_items, collapse = ", "), "\n", sep = "")
  }
  items <- items[items %in% colnames(data)]
  food_items <- intersect(names(MEB_FOOD_WEIGHTS), items)
  nfi_items <- intersect(names(MEB_NFI_WEIGHTS), items)
  wash_items <- intersect(names(MEB_WASH_WEIGHTS), items)

  min_obs_vec <- setNames(rep(min_obs_default, length(items)), items)
  for (nm in names(min_obs_override)) if (nm %in% items) min_obs_vec[nm] <- min_obs_override[[nm]]

  mkt <- calculate_market_medians(data, items, min_obs_vec)
  level_value <- mkt$value
  level_is_local <- mkt$is_local
  group_keys <- level_value$admin3_label

  w_food <- unlist(MEB_FOOD_WEIGHTS[food_items])
  w_nfi <- unlist(MEB_NFI_WEIGHTS[nfi_items])
  w_wash <- unlist(MEB_WASH_WEIGHTS[wash_items])

  results <- vector("list", length(group_keys))
  for (i in seq_along(group_keys)) {
    is_local <- vapply(items, function(it) isTRUE(level_is_local[[it]][i]), logical(1))
    value <- vapply(items, function(it) level_value[[it]][i], numeric(1))
    names(value) <- items

    n_food <- sum(is_local[food_items])
    n_nfi <- sum(is_local[nfi_items])
    n_wash <- sum(is_local[wash_items])
    n_total <- n_food + n_nfi + n_wash

    compute_total <- n_total >= MEB_TOTAL_MIN_ITEMS
    compute_food <- compute_total || n_food >= MEB_BASKET_MIN_STANDALONE[["food"]]
    compute_nfi <- compute_total || n_nfi >= MEB_BASKET_MIN_STANDALONE[["nfi"]]
    compute_wash <- compute_total || n_wash >= MEB_BASKET_MIN_STANDALONE[["wash"]]

    meb_food_val <- sum(value[food_items] * w_food)
    meb_nfi_val <- sum(value[nfi_items] * w_nfi)
    meb_wash_val <- sum(value[wash_items] * w_wash)

    note_parts <- character(0)
    if (!compute_total) {
      note_parts <- c(note_parts, paste0(
        "PMAS total non calculé : données insuffisantes (", n_total, "/", length(items),
        " articles avec assez de prix locaux dans ce marché, ", MEB_TOTAL_MIN_ITEMS, " requis)."
      ))
    } else if (n_total >= MEB_TOTAL_MIN_ITEMS && n_total <= MEB_TOTAL_CAUTION_MAX_ITEMS) {
      caution <- character(0)
      if (n_food < MEB_BASKET_MIN_CAUTION[["food"]]) caution <- c(caution, "alimentaire")
      if (n_nfi < MEB_BASKET_MIN_CAUTION[["nfi"]]) caution <- c(caution, "non alimentaire")
      if (n_wash < MEB_BASKET_MIN_CAUTION[["wash"]]) caution <- c(caution, "hygiène")
      if (length(caution) > 0) {
        note_parts <- c(note_parts, paste0(
          "À interpréter avec prudence : le panier ", paste(caution, collapse = " et "),
          " a été calculé avec peu de prix locaux dans ce marché."
        ))
      }
    }

    results[[i]] <- tibble(
      !!level_col := group_keys[i],
      meb_food = if (compute_food) meb_food_val else NA_real_,
      meb_nfi = if (compute_nfi) meb_nfi_val else NA_real_,
      meb_wash = if (compute_wash) meb_wash_val else NA_real_,
      meb_total = if (compute_total) meb_food_val + meb_nfi_val + meb_wash_val else NA_real_,
      note = if (length(note_parts) > 0) paste(note_parts, collapse = " ") else NA_character_
    )
  }
  bind_rows(results)
}

# niveau National : MEDIANE DES PANIERS COMMUNAUX DEJA CALCULES par calculate_pmas_by_level() (demande
# utilisateur explicite - voir commentaire au-dessus de calculate_pmas_by_level()), PAS un panier
# recompose a partir de medianes de prix par article. Chaque panier (alimentaire/non alimentaire/hygiene/
# total) est medianne INDEPENDAMMENT sur les communes ou il a pu etre calcule (na.rm = TRUE - une commune
# sans assez d'articles mesures localement pour un panier donne n'entre simplement pas dans la mediane de
# ce panier, meme si ses autres paniers sont calculables).
calculate_pmas_national <- function(data, min_obs_default = MEB_MIN_OBS_DEFAULT,
                                    min_obs_override = MEB_MIN_OBS_OVERRIDE) {
  p_commune <- calculate_pmas_by_level(data, "admin3_label", min_obs_default, min_obs_override)
  n_communes <- nrow(p_commune)

  note_missing <- function(label, n_ok) {
    if (n_ok < n_communes) {
      paste0("Panier ", label, " national calculé sur ", n_ok, "/", n_communes,
            " communes (les autres n'ont pas assez d'articles mesurés localement pour ce panier).")
    } else {
      NA_character_
    }
  }
  note_parts <- c(
    note_missing("alimentaire", sum(!is.na(p_commune$meb_food))),
    note_missing("non alimentaire", sum(!is.na(p_commune$meb_nfi))),
    note_missing("hygiène", sum(!is.na(p_commune$meb_wash))),
    note_missing("total", sum(!is.na(p_commune$meb_total)))
  )
  note_parts <- note_parts[!is.na(note_parts)]

  tibble(
    meb_food = median(p_commune$meb_food, na.rm = TRUE),
    meb_nfi = median(p_commune$meb_nfi, na.rm = TRUE),
    meb_wash = median(p_commune$meb_wash, na.rm = TRUE),
    meb_total = median(p_commune$meb_total, na.rm = TRUE),
    note = if (length(note_parts) > 0) paste(note_parts, collapse = " ") else NA_character_
  )
}

# libelles francais lisibles pour les articles du PMAS/MEB (voir MEB_*_WEIGHTS ci-dessus) - variables
# techniques a taille/unite specifique (ex. maize_350g_price_unit_item), differentes des variables
# utilisees par la feuille "Prix médian par article" (ex. maize_price_item, sans taille), donc pas
# directement reutilisables depuis PRICE_ITEM_CATALOG (4_analyse_complete.R) malgre les noms proches.
MEB_ITEM_LABELS <- c(
  maize_350g_price_unit_item = "Farine de maïs (350g)",
  manioc_500g_price_unit_item = "Manioc (500g)",
  rice_500g_price_unit_item = "Riz (500g)",
  bean_500g_price_unit_item = "Haricot (500g)",
  peanut_150g_price_unit_item = "Arachide (150g)",
  beef_meat_price_item = "Viande de bœuf",
  vegetable_oil_price_item = "Huile végétale",
  sugar_200g_price_unit_item = "Sucre (200g)",
  salt_150g_price_unit_item = "Sel (150g)",
  mosquito_net_price_item = "Moustiquaire",
  jerrycan_20l_price_unit_item = "Bidon (20L)",
  bed_sheets_1piece_price_unit_item = "Drap (1 pièce)",
  carpet_price_item = "Natte",
  tarpaulin_20m2_price_unit_item = "Bâche (20m²)",
  cooking_pot_5l_price_unit_item = "Marmite (5L)",
  soap_200g_price_unit_item = "Savon (200g)",
  bucket_20l_price_unit_item = "Seau (20L)"
)

# construit le tableau de reference documentant la composition du PMAS/MEB (categorie, article, variable
# technique, poids applique au prix median de l'article pour obtenir sa contribution au panier) - pour
# la feuille "meb_composition" (voir 4_analyse_complete.R), afin que quelqu'un qui n'a jamais lu le
# script comprenne exactement ce qui entre dans chaque panier et avec quel poids. `missing_items`
# (voir check_meb_items_missing()) ajoute une colonne "Disponible ce mois-ci" - AVANT cet ajout, un
# article dont la variable avait disparu de l'outil kobo (renommee/retiree) restait liste ici comme les
# autres, sans aucune indication qu'il n'est en realite plus applique au panier calcule ce cycle-ci
# (demande utilisateur explicite, suite a l'audit de robustesse).
build_meb_composition_table <- function(missing_items = character(0)) {
  build_block <- function(weights, categorie) {
    tibble(
      Catégorie = categorie,
      Article = unname(MEB_ITEM_LABELS[names(weights)]),
      `Variable technique` = names(weights),
      `Poids appliqué au prix médian` = round(unlist(weights, use.names = FALSE), 4)
    )
  }
  out <- bind_rows(
    build_block(MEB_FOOD_WEIGHTS, "Panier alimentaire"),
    build_block(MEB_NFI_WEIGHTS, "Panier non alimentaire"),
    build_block(MEB_WASH_WEIGHTS, "Panier des produits d'hygiène")
  )
  out$`Disponible ce mois-ci` <- if_else(out$`Variable technique` %in% missing_items,
                                        "Non - variable absente de l'outil/des données, article exclu du panier",
                                        "Oui")
  out
}

# calcule le PMAS a un niveau d'agregation donne (admin3_label, admin2_label ou admin1_label) : mediane
# DIRECTEMENT a partir des prix individuels groupes a ce niveau (pas une mediane de medianes imbriquee -
# plus robuste), avec secours sur la mediane nationale si moins de `min_obs` observations a ce niveau
# pour un article donne. Fonction generique reutilisee par calculate_meb() (niveau commune) et par
# 4_analyse_complete.R pour les niveaux sous-prefecture/prefecture de la feuille "Coût médian du PMAS".
calculate_meb_by_level <- function(data, weights_list, level_col, min_obs = 3) {
  items <- names(weights_list)
  items <- items[items %in% colnames(data)]
  if (length(items) == 0) {
    empty <- tibble(x = character(0), MEB = numeric(0))
    return(rename(empty, !!level_col := x))
  }

  vars <- data %>%
    select(all_of(level_col), all_of(items)) %>%
    mutate(across(all_of(items), as.numeric))

  national_median <- vars %>% summarise(across(all_of(items), ~median(., na.rm = TRUE)))

  level_median <- vars %>% group_by(across(all_of(level_col))) %>%
    summarise(across(all_of(items), ~median(., na.rm = TRUE)), .groups = "drop")
  level_count <- vars %>% group_by(across(all_of(level_col))) %>%
    summarise(across(all_of(items), ~sum(!is.na(.))), .groups = "drop")

  level_imputed <- level_median %>%
    mutate(across(all_of(items), ~if_else(level_count[[cur_column()]] < min_obs, national_median[[cur_column()]], .)))

  w <- unlist(weights_list[items])
  level_imputed %>%
    mutate(MEB = rowSums(across(all_of(items)) * matrix(w, nrow = n(), ncol = length(items), byrow = TRUE), na.rm = TRUE)) %>%
    select(all_of(level_col), MEB)
}

# calcule le PMAS au niveau national (mediane sur l'ensemble des prix individuels, tous marches
# confondus) pour un panier de poids donne - une seule valeur, pas de secours necessaire (c'est deja le
# niveau le plus agrege).
calculate_meb_national <- function(data, weights_list) {
  items <- names(weights_list)
  items <- items[items %in% colnames(data)]
  if (length(items) == 0) return(tibble(MEB = NA_real_))

  vars <- data %>% select(all_of(items)) %>% mutate(across(everything(), as.numeric))
  national_median <- vars %>% summarise(across(everything(), ~median(., na.rm = TRUE)))
  w <- unlist(weights_list[items])
  tibble(MEB = sum(unlist(national_median) * w, na.rm = TRUE))
}

# calcule le PMAS (median par admin3, secours sur la mediane nationale si moins de 3 observations dans
# cet admin3) pour un panier de poids donne - renvoie une ligne par admin3 (+ 1 ligne "National"). Voir
# calculate_meb_by_level()/calculate_meb_national() pour les niveaux intermediaires (sous-prefecture/
# prefecture), utilises separement par 4_analyse_complete.R pour la feuille "Coût médian du PMAS".
calculate_meb <- function(data, weights_list) {
  res_admin3 <- calculate_meb_by_level(data, weights_list, "admin3_label")
  res_national <- calculate_meb_national(data, weights_list) %>%
    mutate(admin3_label = "National") %>%
    select(admin3_label, MEB)
  bind_rows(res_admin3, res_national)
}

##############################################################################################################
# HAUTEUR DE LIGNE DEPENDANTE DU CONTENU (feuille READ_ME)
##############################################################################################################
# meme logique que estimate_wrapped_row_height() dans scripts/fonctions_nettoyage.R (NETTOYAGE - remplit
# la feuille READ_ME initialement) - dupliquee ici cote ANALYSE (pas de sibling fonctions_nettoyage.R a
# sourcer specifiquement pour cette seule fonction) pour que les lignes ajoutees a la feuille READ_ME
# depuis 4_analyse_complete.R (description des feuilles d'analyse) aient la meme hauteur dependante du
# contenu que les lignes deja remplies par le nettoyage, plutot qu'une hauteur fixe qui detonnerait dans
# le meme tableau.
estimate_wrapped_row_height <- function(text, col_width_chars, line_height = 15, min_height = 15, padding = 6) {
  if (is.na(text) || !nzchar(text)) return(min_height)
  explicit_lines <- max(lengths(gregexpr("\n", text))) + 1
  chars_per_line <- max(1, floor(col_width_chars))
  wrapped_lines <- ceiling(nchar(text) / chars_per_line)
  n_lines <- max(explicit_lines, wrapped_lines)
  max(min_height, n_lines * line_height + padding)
}
