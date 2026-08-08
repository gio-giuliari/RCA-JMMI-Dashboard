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
source("scripts/0_configuration.R")

################################################################################
# CHARGEMENT DU JEU DE DONNEES BRUT (jeu de donnees complet - la collecte est fermee)
################################################################################

raw_data_all <- read_raw_data(Sys.getenv("raw_dataset")) %>%
  add_admin_org_labels()

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

INFO_COLS <- c("date", "submission_date", "admin1_code", "admin1_label", "admin2_code", "admin2_label",
              "admin3_code", "admin3_label", "marketplace_type", "enum_id", "shop_name",
              "organisation_name_label")

################################################################################
# DETECTION DU MODE : classeur "a verifier" deja rempli, ou premiere execution ?
################################################################################

filled_files <- list.files(Sys.getenv("suivi_final_remplis_path"),
                           pattern = Sys.getenv("suivi_final_remplis_pattern"),
                           recursive = TRUE, full.names = TRUE)

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

# 3) valeurs statistiquement aberrantes (directement sur les donnees brutes - il n'y a plus de
#    seconde passe apres application des verifications logiques)
outliers_log <- raw_data_all %>%
  check_outliers(
    uuid_column = "uuid", kobo_survey = NULL, kobo_choices = NULL, cols_to_add_cleaning_log = NULL,
    strongness_factor = 3, minimum_unique_value_of_variable = 2, remove_choice_multiple = TRUE,
    sm_separator = "/",
    columns_not_to_check = setdiff(names(raw_data_all), c("uuid", items_other_unit, items_list, days_db))
  )
log_outliers <- create_combined_log(list_of_log = outliers_log) %>%
  add_info_to_cleaning_log(dataset = "checked_dataset", cleaning_log = "cleaning_log",
                           dataset_uuid_column = "uuid", cleaning_log_uuid_column = "uuid",
                           information_to_add = INFO_COLS)
df_outliers <- log_outliers$cleaning_log %>%
  mutate(issue = translate_issue_fr(issue)) %>%
  # ne pas signaler deux fois le meme champ (deja signale ci-dessus par les verifications logiques)
  anti_join(df_principal, by = c("uuid", "question"))

# 4) quasi-doublons (enquetes tres semblables entre elles) - un probleme au niveau de l'enquete
#    entiere, pas d'un champ en particulier (question = NA)
soft_dup <- raw_data_all %>%
  check_soft_duplicates(kobo_survey = kobo_survey, uuid_column = "uuid", idnk_value = "ne_sait_pas",
                        sm_separator = "/", log_name = "soft_duplicate_log", threshold = 5,
                        return_all_results = FALSE)
df_soft_dup <- soft_dup$soft_duplicate_log %>%
  mutate(question = NA_character_,
        old_value = paste0("Tres semblable a l'enquete ", id_most_similar_survey),
        issue = translate_issue_fr(issue),
        check_id = NA_character_, check_binding = NA_character_,
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
        check_id = NA_character_, check_binding = NA_character_,
        change_type = NA_character_, new_value = NA_character_) %>%
  select(uuid, question, old_value, issue, check_id, check_binding, change_type, new_value) %>%
  left_join(raw_data_all %>% select(uuid, all_of(INFO_COLS)), by = "uuid")

# toutes les verifications combinees dans un seul tableau
all_issues <- bind_rows(df_principal, df_outliers, df_soft_dup, df_pct_missing)

# colonnes simples/lisibles pour l'equipe de suivi (voir REVIEW_HEADERS_FR / write_review_sheet() dans
# scripts/fonctions_nettoyage.R)
review_df <- all_issues %>%
  rowwise() %>%
  mutate(question_label = get_question_label(question)) %>%
  ungroup() %>%
  transmute(
    uuid = uuid,
    commune = admin3_label,
    partenaire = organisation_name_label,
    date = date,
    enqueteur = enum_id,
    boutique = shop_name,
    variable = question,
    question_label = question_label,
    reponse_actuelle = old_value,
    probleme = issue,
    action = NA_character_,
    nouvelle_reponse = NA_character_,
    commentaire = NA_character_
  ) %>%
  arrange(partenaire, date)

if (interactive()) View(review_df)

cat("\n--> ", nrow(review_df), " probleme(s) a verifier detecte(s).\n", sep = "")

################################################################################
# REPONSES "AUTRE"
################################################################################

var_other_raw <- other_db$name[other_db$name %in% colnames(raw_data_all)]

other_responses <- raw_data_all %>%
  select(c("uuid", all_of(var_other_raw))) %>%
  pivot_longer(cols = all_of(var_other_raw), names_to = "question_name", values_to = "response") %>%
  filter(!is.na(response)) %>%
  select(uuid, question_name, response)

df_other <- other_responses %>%
  left_join(select(raw_data_all, uuid, enum_id,
                   admin1_code, admin1_label, admin2_code, admin2_label, admin3_code, admin3_label),
            by = "uuid") %>%
  arrange(question_name) %>%
  left_join(select(other_db, name, full_label, q_type, list_name), by = c("question_name" = "name")) %>%
  select(uuid, enum_id, admin1_code, admin1_label, admin2_code, admin2_label, admin3_code, admin3_label,
        q_type, list_name, question_name, full_label, response) %>%
  mutate(selected_choices = NA,
        "TRUE other (copy response or provide a better translation)" = NA,
        "EXISTING other 1 (select the most appropriate choice)" = NA,
        "EXISTING other 2 (select the most appropriate choice)" = NA,
        "EXISTING other 3 (select the most appropriate choice)" = NA,
        "INVALID other (select yes or leave blank)" = NA,
        "FOLLOW-UP message (what is unclear about this response?)" = NA,
        "Explanation" = NA)

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
# ECRITURE DU CLASSEUR "A VERIFIER" (verifications + reponses autres, meme fichier)
################################################################################

wb <- createWorkbook()
wb <- write_review_sheet(wb, review_df, sheet_name = "A_verifier")
wb <- add_other_responses_sheet(wb, df_other, sheet_name = "reponses_autres")

a_remplir_dir <- Sys.getenv("suivi_final_a_remplir_dir")
if (!dir.exists(a_remplir_dir)) dir.create(a_remplir_dir, recursive = TRUE)
# meme convention de nommage que le suivi quotidien (RCA_JMMI_<mois>_<type>_<date>_<heure>) - voir
# 1_suivi_quotidien.R
out_path <- paste0(a_remplir_dir, "RCA_JMMI_", period$assessment_month,
                   "_suivi_a_verifier_", format(Sys.time(), "%Y-%m-%d_%Hh%M"), ".xlsx")
saveWorkbook(wb, out_path, overwrite = TRUE)

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

filled <- lapply(filled_files, read_review_sheet)
filled <- do.call(plyr::rbind.fill, filled) %>% arrange(uuid, question)

# les lignes "toute l'enquete" (quasi-doublon, valeurs manquantes) ont variable/question = NA - elles
# ne correspondent a aucune colonne reelle du jeu de donnees et ne peuvent donc etre appliquees qu'au
# niveau de l'enquete entiere (via l'action "Supprimer toute l'enquete"), jamais au niveau d'un champ.
filled_field_level <- filled %>% filter(!is.na(question))

test <- filled_field_level %>%
  select(uuid, question, old_value, change_type, new_value) %>% distinct() %>%
  group_by(uuid, question) %>% summarise(n = n(), .groups = "drop") %>% filter(n > 1)
if (nrow(test) > 0) stop("Certaines combinaisons enquete/variable sont signalees plus d'une fois dans le classeur rempli - a resoudre avant de continuer")

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

clean_data <- create_clean_data(raw_dataset = raw_data_all, raw_data_uuid_column = "uuid",
                                cleaning_log = filled_field_level, cleaning_log_uuid_column = "uuid",
                                cleaning_log_question_column = "question",
                                cleaning_log_new_value_column = "new_value",
                                cleaning_log_change_type_column = "change_type",
                                change_response_value = "change_response",
                                NA_response_value = "blank_response",
                                no_change_value = "no_action",
                                remove_survey_value = "remove_survey")

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
  or <- lapply(or.files, function(f) read_excel(f, col_types = "text", sheet = "reponses_autres"))
  or <- do.call(plyr::rbind.fill, or)

  # Les classeurs de reponses autre portent des en-tetes francais pour l'equipe de suivi (voir
  # add_other_responses_sheet()) - on les retraduit ici vers les noms internes utilises par le reste
  # de ce script, plutot que de changer toute la logique ci-dessous.
  header_en <- c(type_question = "q_type", liste_choix = "list_name", nom_question = "question_name",
                libelle_complet = "full_label", reponse = "response", choix_selectionnes = "selected_choices")
  match_idx <- match(colnames(or), names(header_en))
  colnames(or)[!is.na(match_idx)] <- header_en[match_idx[!is.na(match_idx)]]

  colnames(or)[str_starts(colnames(or), "VRAI")] <- "true_other"
  colnames(or)[str_starts(colnames(or), "EXISTANT autre 1")] <- "existing_other_1"
  colnames(or)[str_starts(colnames(or), "EXISTANT autre 2")] <- "existing_other_2"
  colnames(or)[str_starts(colnames(or), "EXISTANT autre 3")] <- "existing_other_3"
  colnames(or)[str_starts(colnames(or), "INVALIDE")] <- "invalid_other"
  colnames(or)[str_starts(colnames(or), "Message de SUIVI")] <- "fu_message"

  # filtrer "or" selon les uuid du jeu de donnees nettoye
  or <- or %>% filter(uuid %in% clean_data$uuid)

  # extraire les "TRUE other" les plus recurrents et sauvegarder le fichier
  if (any(!is.na(or$true_other))) {
    res <- or %>%
      filter(!is.na(true_other)) %>%
      group_by(list_name) %>%
      mutate(n_list = n()) %>%
      group_by(list_name, true_other) %>%
      summarise(n_list = n_list[1], n_option = n()) %>%
      arrange(-n_list, -n_option) %>%
      select(-n_list)

    frequentes_dir <- Sys.getenv("final_workbook_dir")
    if (!dir.exists(frequentes_dir)) dir.create(frequentes_dir, recursive = TRUE)
    write_xlsx(res, paste0(frequentes_dir, Sys.Date(), "_reponses_autres_frequentes.xlsx"))
  }

  # preparer les trois types de recodage des "reponses autre"
  or <- or %>%
    unite(existing_other, c(existing_other_1, existing_other_2, existing_other_3),
          sep = ";", remove = T, na.rm = T) %>%
    mutate(existing_other = ifelse(existing_other == "", NA, existing_other)) %>%
    left_join(select(other_db, name, ref_question, ref_type = q_type), by = c("question_name" = "name")) %>%
    rename(name = "question_name", ref_name = "ref_question")

  or$check <- rowSums(is.na(select(or, true_other, existing_other, invalid_other)))
  t <- filter(or, check != 2)
  if (nrow(t) > 0) warning("Entrees manquantes ou plusieurs colonnes selectionnees dans le recodage des reponses autre")

  or_recode <- filter(or, !is.na(existing_other))
  or_remove <- filter(or, !is.na(invalid_other))
  or_true <- filter(or, !is.na(true_other))

  print(paste("Nombre de reponses a supprimer :", nrow(or_remove)))
  if (nrow(or_remove) > 0) for (r in 1:nrow(or_remove)) add_to_cleaning_log_other_remove(or_remove[r, ])

  print(paste("Nombre de reponses a recoder :", nrow(or_recode)))
  if (nrow(or_recode) > 0) for (r in 1:nrow(or_recode)) add_to_cleaning_log_other_recode(or_recode[r, ])

  print(paste("Nombre de 'true other' :", nrow(or_true)))
  if (nrow(or_true) > 0) {
    cl_or_true <- or_true %>%
      select(uuid, name, response, true_other) %>%
      rename(question = name, old_value = response, new_value = true_other) %>%
      mutate(change_type = "Translation of true other")
    cleaning_log_other <- rbind(cleaning_log_other, cl_or_true)
  }

  t <- cleaning_log_other %>% group_by(uuid, question) %>% summarise(n = n()) %>% filter(n > 1)
  if (nrow(t) > 0) stop("Doublons detectes dans le journal de nettoyage des reponses autre")

  cleaning_log_other <- cleaning_log_other %>% distinct()

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

filled_no_deletion <- filled %>%
  filter(change_type != "remove_survey") %>%
  filter(!uuid %in% deletion_log$uuid) %>%
  filter(!is.na(question)) %>%
  mutate(new_value = ifelse(change_type == "no_action", old_value, new_value))

cleaning_log_other <- cleaning_log_other %>%
  mutate(issue = "Reponse autre - a verifier", comment = NA_character_)

combined_cleaning_log <- filled_no_deletion %>%
  select(uuid, question, change_type, old_value, new_value, issue, comment) %>%
  rbind(cleaning_log_other)

review_of_cleaning <- review_cleaning(raw_dataset = raw_data_all,
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

clean_data_updated <- clean_data
for (i in seq_len(nrow(eur_calc_fields))) {
  source_item <- eur_calc_fields$source_item[i]
  target_field <- eur_calc_fields$target_field[i]
  if (source_item %in% colnames(clean_data_updated)) {
    clean_data_updated[[target_field]] <- round(as.numeric(clean_data_updated[[source_item]]) / eur_rate)
  }
}
clean_data_updated <- clean_data_updated %>% relocate(uuid)

################################################################################
# RETRAIT DES INFORMATIONS SENSIBLES
################################################################################

remove_cols <- c(
  intersect(colnames(clean_data_updated), filter(kobo_survey, type %in% c("note", "geopoint", "decimal"))$name),
  "audit", "audit_URL", "deviceid", "enum_id", "shop_name", "consent", "organisation_name_label",
  # admin1_label/admin2_label/admin3_label sont conservees (nouvelles variables, consignees dans
  # variable_tracker ci-dessous) - c'est le mapping PCode -> nom depuis la feuille choices du kobo,
  # pas une colonne technique.
  # "_index" est conservee (renommee "_id" ci-dessous) suivant la recommandation de la checklist de
  # garder un identifiant unique independant de Kobo en plus de uuid - toute autre colonne systeme
  # prefixee par "_" est supprimee.
  setdiff(colnames(clean_data_updated)[str_starts(colnames(clean_data_updated), "_")], "_index"),
  colnames(clean_data_updated)[str_ends(colnames(clean_data_updated), "_other_unit_item")]
)
# NOTE : les colonnes "_std_unit_item" (ex. maize_std_unit_item) sont conservees - elles indiquent
# l'unite standard utilisee pour le prix collecte, ce sont des donnees pertinentes et non un champ
# technique.
remove_cols <- intersect(remove_cols, colnames(clean_data_updated))

clean_data_final <- select(clean_data_updated, -all_of(remove_cols))
if ("_index" %in% colnames(clean_data_final)) clean_data_final <- rename(clean_data_final, id = `_index`)

# donnees brutes pour validation - anonymisees de la meme maniere, sauf les colonnes qui n'existent
# qu'apres le nettoyage (admin1_label/admin2_label/admin3_label sont conservees - voir remove_cols ci-dessus)
remove_cols_raw <- c(
  intersect(colnames(raw_data_all), filter(kobo_survey, type %in% c("geopoint"))$name),
  "audit", "audit_URL", "deviceid", "enum_id", "shop_name", "organisation_name_label"
)
remove_cols_raw <- intersect(remove_cols_raw, colnames(raw_data_all))

raw_data_all_validation <- select(raw_data_all, -all_of(remove_cols_raw))
if ("_index" %in% colnames(raw_data_all_validation)) raw_data_all_validation <- rename(raw_data_all_validation, id = `_index`)

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
  distinct()

justify_removed_variable <- function(v) {
  case_when(
    v %in% c("deviceid", "enum_id", "shop_name", "consent") ~ "Confidentialite / DPI",
    v %in% filter(kobo_survey, type == "geopoint")$name ~ "Coordonnees GPS - DPI",
    v %in% filter(kobo_survey, type == "note")$name ~ "Texte d'introduction, non pertinent pour l'analyse",
    str_ends(v, "_other_unit_item") ~ "Variable technique utilisee uniquement pour calculer le prix standardise",
    TRUE ~ "Variable retiree lors du nettoyage final"
  )
}

variable_tracker_removed <- data.frame(variable = intersect(remove_cols, kobo_survey$name)) %>%
  mutate(action = "Supprime", justification = justify_removed_variable(variable))

# admin1_code/admin2_code/admin3_code sont les PCodes kobo (renommes depuis admin1/admin2/admin3) ;
# admin1_label/admin2_label/admin3_label sont de nouvelles variables qui font correspondre ces PCodes
# aux noms reels de prefecture/sous-prefecture/commune, depuis la feuille "choices" de l'outil kobo.
variable_tracker_added <- data.frame(
  variable = c("admin1_label", "admin2_label", "admin3_label"),
  action = "Ajoute",
  justification = c(
    "Nom de la prefecture correspondant au PCode admin1_code (mapping depuis la feuille choices du questionnaire Kobo)",
    "Nom de la sous-prefecture correspondant au PCode admin2_code (mapping depuis la feuille choices du questionnaire Kobo)",
    "Nom de la commune correspondant au PCode admin3_code (mapping depuis la feuille choices du questionnaire Kobo)"
  )
)

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
