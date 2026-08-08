################################################################################
# JMMI RELABELLING - FONCTIONS DE RE-ETIQUETAGE DES CHOIX
################################################################################
# Remplace, dans un jeu de donnees nettoye (clean_data), le CODE de chaque reponse a choix
# (select_one / select_multiple) par son LIBELLE tel que defini dans la feuille "choices" de l'outil
# kobo correspondant - uniquement pour les colonnes presentes A LA FOIS dans les donnees et dans
# l'outil (survey$name). Ne touche JAMAIS aux noms de colonnes/variables (l'en-tete) - uniquement aux
# valeurs qu'elles contiennent.
#
# GENERIQUE PAR CONCEPTION : aucune de ces fonctions ne suppose quoi que ce soit sur le contenu
# specifique d'un jeu de donnees JMMI ou d'un outil kobo en particulier (aucun nom de variable, de
# liste de choix, de pays code en dur) - seule la structure standard XLSForm (colonnes "type"/"name"
# dans survey, "list_name"/"name"/"label..." dans choices) est utilisee. Peut donc etre reutilisee tel
# quel avec n'importe quel jeu JMMI nettoye + son outil kobo.
#
# NOTE : copie autonome pour rca_jmmi_dashboard (voir en-tete de app.R) - source d'origine dans
# jmmi_relabelling/scripts/fonctions_relabel.R, a recopier ici si elle evolue.
################################################################################

#-------------------------------------------------------------------------------------------------
# Detecte la colonne de libelle a utiliser dans une feuille kobo (survey ou choices) : prefere une
# colonne "label::<langue>" si plusieurs langues sont presentes, sinon la colonne "label" simple.
#-------------------------------------------------------------------------------------------------
detect_label_column <- function(df, prefer_lang = "french") {
  label_cols <- names(df)[str_detect(tolower(names(df)), "^label")]
  if (length(label_cols) == 0) {
    stop("Aucune colonne de libelle (label / label::<langue>) trouvee dans cette feuille.")
  }
  preferred <- label_cols[str_detect(tolower(label_cols), tolower(prefer_lang))]
  if (length(preferred) > 0) return(preferred[1])
  if ("label" %in% label_cols) return("label")
  label_cols[1]
}

#-------------------------------------------------------------------------------------------------
# A partir de la feuille "survey" (colonnes standard XLSForm "type"/"name"), construit la
# correspondance nom de variable -> (q_type, list_name) pour les questions a choix
# (select_one/select_multiple) uniquement.
#-------------------------------------------------------------------------------------------------
build_choice_variable_map <- function(kobo_survey) {
  kobo_survey %>%
    mutate(
      .type_clean = str_squish(as.character(.data[["type"]])),
      q_type = word(.type_clean, 1),
      list_name = word(.type_clean, 2)
    ) %>%
    filter(q_type %in% c("select_one", "select_multiple"), !is.na(list_name), nzchar(list_name)) %>%
    select(name, q_type, list_name) %>%
    filter(!is.na(name), nzchar(name)) %>%
    distinct(name, .keep_all = TRUE)
}

#-------------------------------------------------------------------------------------------------
# A partir de la feuille "choices" (colonnes standard XLSForm "list_name"/"name"/"label..."),
# construit la correspondance (liste, code) -> libelle.
#-------------------------------------------------------------------------------------------------
build_choice_label_map <- function(kobo_choices) {
  label_col <- detect_label_column(kobo_choices, prefer_lang = "french")
  kobo_choices %>%
    filter(!is.na(.data[["list_name"]]), !is.na(.data[["name"]])) %>%
    transmute(
      list_name = as.character(.data[["list_name"]]),
      code = as.character(.data[["name"]]),
      label = as.character(.data[[label_col]])
    ) %>%
    distinct(list_name, code, .keep_all = TRUE)
}

#-------------------------------------------------------------------------------------------------
# Remplace les codes par les libelles dans UNE valeur de colonne select_one/select_multiple. Pour un
# select_multiple, les codes sont separes par un espace dans les donnees (convention kobo/ODK
# standard) ; les libelles trouves sont rejoints avec `out_sep` (et non un espace, puisque les
# libelles peuvent eux-memes contenir des espaces). Un code sans correspondance dans la feuille
# choices est laisse tel quel (jamais d'erreur, jamais de perte de donnee).
#-------------------------------------------------------------------------------------------------
relabel_value <- function(value, q_type, list_name, label_lookup, sep = " ", out_sep = "; ") {
  if (is.na(value) || !nzchar(trimws(value))) return(value)
  codes <- if (identical(q_type, "select_multiple")) str_split(trimws(value), sep)[[1]] else value
  codes <- codes[nzchar(codes)]
  if (length(codes) == 0) return(value)
  labels <- vapply(codes, function(code) {
    lbl <- label_lookup$label[label_lookup$list_name == list_name & label_lookup$code == code]
    if (length(lbl) > 0 && !is.na(lbl[1]) && nzchar(lbl[1])) lbl[1] else code
  }, character(1))
  paste(labels, collapse = out_sep)
}

#-------------------------------------------------------------------------------------------------
# Fonction principale. Re-etiquette toutes les colonnes de `clean_data` qui sont A LA FOIS presentes
# dans les donnees et definies comme select_one/select_multiple dans l'outil kobo (survey$name). Les
# colonnes absentes de l'outil, ou qui ne sont pas des questions a choix (texte, nombre, date, notes,
# champs "calculate", colonnes ajoutees par le nettoyage comme les *_label admin, etc.), sont
# renvoyees inchangees. Les NOMS de colonnes ne sont jamais modifies.
#-------------------------------------------------------------------------------------------------
relabel_choices <- function(clean_data, kobo_survey, kobo_choices, verbose = TRUE) {
  var_map <- build_choice_variable_map(kobo_survey)
  label_lookup <- build_choice_label_map(kobo_choices)

  vars_to_relabel <- intersect(colnames(clean_data), var_map$name)

  if (verbose) {
    cat("--> ", length(vars_to_relabel), " colonne(s) a choix re-etiquetee(s) sur ",
       ncol(clean_data), " colonne(s) au total (les autres colonnes ne sont pas modifiees).\n", sep = "")
  }

  for (v in vars_to_relabel) {
    q_type <- var_map$q_type[var_map$name == v][1]
    list_name <- var_map$list_name[var_map$name == v][1]
    clean_data[[v]] <- vapply(
      clean_data[[v]], relabel_value, character(1),
      q_type = q_type, list_name = list_name, label_lookup = label_lookup
    )
  }

  clean_data
}

#-------------------------------------------------------------------------------------------------
# A partir de la feuille "survey", construit la correspondance nom de variable -> libelle de la
# question (toutes questions confondues, pas seulement les questions a choix).
#-------------------------------------------------------------------------------------------------
build_question_label_map <- function(kobo_survey) {
  label_col <- detect_label_column(kobo_survey, prefer_lang = "french")
  kobo_survey %>%
    filter(!is.na(.data[["name"]]), nzchar(.data[["name"]])) %>%
    transmute(name = as.character(.data[["name"]]), label = as.character(.data[[label_col]])) %>%
    distinct(name, .keep_all = TRUE)
}

#-------------------------------------------------------------------------------------------------
# Les questions select_multiple sont exportees par kobo/ODK sous forme d'une colonne "parent" (les
# codes selectionnes, separes par un espace) PLUS une colonne binaire par choix possible, nommee
# "parent/code" (convention standard). Le nom de ces colonnes binaires reste technique (ex.
# "shop_category/food_item") meme apres nettoyage - cette fonction calcule, pour la LIGNE D'ETIQUETTE
# uniquement (jamais le nom de colonne lui-meme), un libelle complet "libelle question/libelle choix"
# (ex. "Quelles sont les categories de produits...?/Produits alimentaires...") en detectant, pour
# chaque colonne, si son prefixe (avant "/") correspond a une question select_multiple de l'outil et
# si son suffixe correspond a un code de choix valide pour cette liste. Les colonnes qui ne
# correspondent pas a ce schema (pas de "/", ou prefixe/suffixe non reconnus) conservent leur
# etiquette d'origine, inchangee.
#-------------------------------------------------------------------------------------------------
relabel_multiple_choice_subcolumn_labels <- function(colnames, label_row, kobo_survey, kobo_choices) {
  sm_map <- build_choice_variable_map(kobo_survey) %>% filter(q_type == "select_multiple")
  question_labels <- build_question_label_map(kobo_survey)
  choice_labels <- build_choice_label_map(kobo_choices)

  for (i in seq_along(colnames)) {
    col <- colnames[i]
    slash_pos <- regexpr("/", col, fixed = TRUE)
    if (slash_pos < 0) next

    parent <- substr(col, 1, slash_pos - 1)
    code <- substr(col, slash_pos + 1, nchar(col))

    vm <- sm_map[sm_map$name == parent, ]
    if (nrow(vm) == 0) next
    list_name <- vm$list_name[1]

    q_label <- question_labels$label[question_labels$name == parent]
    c_label <- choice_labels$label[choice_labels$list_name == list_name & choice_labels$code == code]
    if (length(q_label) == 0 || length(c_label) == 0) next
    if (is.na(q_label[1]) || is.na(c_label[1]) || !nzchar(q_label[1]) || !nzchar(c_label[1])) next

    label_row[i] <- paste0(q_label[1], "/", c_label[1])
  }

  label_row
}

#-------------------------------------------------------------------------------------------------
# Lit une feuille avec un en-tete SUR DEUX LIGNES (ligne 1 = libelle de la question, ligne 2 = nom de
# colonne technique, donnees a partir de la ligne 3 - convention utilisee par raw_data/clean_data du
# classeur final JMMI) en conservant les deux lignes d'en-tete separement, pour pouvoir les reecrire
# telles quelles apres re-etiquetage des valeurs.
#-------------------------------------------------------------------------------------------------
read_two_row_header_sheet <- function(path, sheet) {
  label_row <- suppressMessages(
    readxl::read_excel(path, sheet = sheet, range = readxl::cell_rows(1), col_names = FALSE, col_types = "text")
  )
  data <- suppressMessages(
    readxl::read_excel(path, sheet = sheet, skip = 1, col_names = TRUE, col_types = "text")
  )
  list(label_row = as.character(unlist(label_row[1, ], use.names = FALSE)), data = data)
}

#-------------------------------------------------------------------------------------------------
# Remplace les VALEURS (ligne 1 = libelles, ligne 2 = noms de colonnes, donnees a partir de la ligne
# 3) d'une feuille EXISTANTE dans un classeur openxlsx deja charge (via loadWorkbook() sur le classeur
# d'entree) - sans creer de nouvelle feuille, ce qui preserve la mise en forme d'origine (styles,
# largeurs de colonnes, volet fige) ainsi que toutes les autres feuilles du classeur et leur ordre.
#-------------------------------------------------------------------------------------------------
overwrite_two_row_header_sheet <- function(wb, sheet_name, data, label_row) {
  openxlsx::writeData(wb, sheet_name, as.list(label_row), startRow = 1, colNames = FALSE)
  openxlsx::writeData(wb, sheet_name, data, startRow = 2, colNames = TRUE)
}
