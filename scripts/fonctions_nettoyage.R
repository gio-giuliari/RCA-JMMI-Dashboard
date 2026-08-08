################################################################################
# JMMI CAR - FONCTIONS UTILITAIRES PARTAGEES POUR LE NETTOYAGE
################################################################################
# ENTREE      : Appelees tout au long de 1_suivi_quotidien.R / 2_nettoyage_final.R ; s'appuie sur les
#               objets globaux definis par scripts/0_configuration.R (kobo_survey, kobo_choices,
#               other_db, ...)
# SORTIE      : Aucune sortie fichier en soi, a part les fonctions qui sauvegardent explicitement un
#               classeur (add_other_responses_sheet(), build_final_workbook()) - sinon uniquement des
#               definitions de fonctions utilisees par les autres scripts.
# DESCRIPTION : Lecture des donnees brutes/analyse des dates, jointures des etiquettes
#               admin/organisation, l'utilitaire de periode d'evaluation, le processus de recodage des
#               "reponses autre", les utilitaires du suivi quotidien des enquetes (dont le prix median
#               par article), la feuille de verification consolidee et accessible pour l'equipe de
#               suivi, et l'ecriture du classeur final selon le modele de la checklist (en-tetes, mise
#               en forme, READ_ME, auto-evaluation de la checklist).
# OBJECTIF    : Regrouper tous les elements reutilisables de la logique de nettoyage a un seul endroit,
#               afin que les deux scripts principaux restent concentres sur l'enchainement des etapes
#               et non sur les mecanismes de chaque etape.
# CONTEXTE    : Adapte de src/utils_cleaning.R de JMMI_ETH ; l'utilitaire de periode d'evaluation et
#               l'ecriture du classeur selon le modele de la checklist sont des ajouts propres a la
#               RCA, car le pipeline de l'ETH ecrit des fichiers finaux separes plutot qu'un classeur
#               combine unique. Le pipeline RCA a ete simplifie pour ne plus faire de verifications
#               pendant la collecte (uniquement le suivi couverture/prix) : toutes les verifications se
#               font en une seule fois, a la fin de la collecte, dans une feuille unique et accessible a
#               une equipe de suivi avec un faible niveau d'alphabetisation/de formation technique.
################################################################################

#############################################################################################################
# LECTURE DU JEU DE DONNEES
#############################################################################################################

# `crayon` n'est pas charge dans init.R (garde minimal) ; ce leurre neutre conserve le style de
# journalisation console `cat(green(...))` de l'ETH sans ajouter de dependance a un package pour
# quelque chose de purement cosmetique.
green <- function(x) x

# Le jeu de donnees factice de la RCA est un export synthetique "plat" sans aucune des colonnes
# systeme kobo habituelles (_uuid, _index, __version__, _submission_time). Les vrais exports kobo
# mensuels les auront deja, donc la synthese ci-dessous ne se declenche que comme solution de repli
# pour des donnees qui ne les ont pas.

parse_kobo_date <- function(x) {
  out <- suppressWarnings(as.Date(x))
  needs_fallback <- is.na(out) & !is.na(x) & x != ""
  if (any(needs_fallback)) {
    out[needs_fallback] <- suppressWarnings(openxlsx::convertToDate(as.numeric(x[needs_fallback])))
  }
  as.character(out)
}

parse_kobo_datetime <- function(x) {
  out <- suppressWarnings(as.POSIXct(x, tz = "UTC"))
  needs_fallback <- is.na(out) & !is.na(x) & x != ""
  if (any(needs_fallback)) {
    out[needs_fallback] <- suppressWarnings(openxlsx::convertToDateTime(as.numeric(x[needs_fallback])))
  }
  as.character(out)
}

read_raw_data <- function(submission){

  df <- read_excel(submission, sheet=1, col_types="text")
  cat(green(paste0("--> BRUT --> ", submission, " --> ", nrow(df), " entrees, ", ncol(df), " colonnes \n")))

  if (!"_uuid" %in% names(df)) {
    cat(green("--> synthese des colonnes d'export kobo manquantes (_uuid, _index, __version__, _submission_time) \n"))
    df <- df %>%
      mutate(`_uuid` = paste0("CAR-", sprintf("%05d", row_number())),
             `_index` = row_number(),
             `__version__` = NA_character_,
             `_submission_time` = coalesce(end, today))
  }

  cat(green("--> conversion des dates \n"))
  cols.date <- intersect(kobo_survey$name[kobo_survey$type %in% c("today", "date")], colnames(df))
  df <- df %>%
    mutate(submission_date = parse_kobo_date(`_submission_time`)) %>%
    mutate_at(cols.date, parse_kobo_date) %>%
    mutate_at(intersect(c("start", "end", "_submission_time"), colnames(df)), parse_kobo_datetime)

  #---------------------------------------------------------------------------
  cat(green("--> renommage de uuid/id/version et ajout de l'heure de soumission \n"))
  # "_id" (colonne systeme kobo optionnelle, absente des exports sans historique de soumissions
  # modifiees) n'est renommee que si elle est presente - contrairement a _uuid/__version__/_submission_time,
  # elle n'est jamais synthetisee plus haut si absente.
  rename_map <- c(version="__version__", uuid="_uuid", submission_time="_submission_time", id="_id")
  rename_map <- rename_map[rename_map %in% names(df)]
  df <- df %>% rename(!!!rename_map)
  #-----------------------------------------------------------------------------
  # conversion des colonnes numeriques
  cat(green("--> conversion des colonnes entieres en numerique \n"))
  df <- df %>% mutate_at(intersect(cols_numeric, colnames(df)), as.numeric)
  return(df)
}

#############################################################################################################
# ETIQUETTES ADMIN / ORGANISATION
#############################################################################################################

# L'export brut de la RCA n'a que les colonnes kobo "admin1"/"admin2"/"admin3" (PCodes) et
# "organisation" (un nom kobo, ex. "reach"). Ceci renomme les colonnes admin en
# admin1_code/admin2_code/admin3_code et ajoute admin1_label/admin2_label/admin3_label (le mapping
# PCode -> nom de prefecture/sous-prefecture/commune, depuis les listes de choix kobo) comme nouvelles
# variables conservees jusque dans raw_data/clean_data - contrairement a organisation_name_label, qui
# n'est utilisee que pour les colonnes d'information du suivi quotidien/journal de nettoyage et
# retiree a nouveau avant l'export final.
add_admin_org_labels <- function(df){
  df <- df %>% rename(admin1_code = admin1, admin2_code = admin2, admin3_code = admin3)

  admin1_lbl <- kobo_choices %>% filter(list_name=="admin1") %>% select(admin1_code=name, admin1_label=label)
  admin2_lbl <- kobo_choices %>% filter(list_name=="admin2") %>% select(admin2_code=name, admin2_label=label)
  admin3_lbl <- kobo_choices %>% filter(list_name=="admin3") %>% select(admin3_code=name, admin3_label=label)
  org_lbl    <- kobo_choices %>% filter(list_name=="ong")    %>% select(organisation=name, organisation_name_label=label)

  df %>%
    left_join(admin1_lbl, by="admin1_code") %>%
    left_join(admin2_lbl, by="admin2_code") %>%
    left_join(admin3_lbl, by="admin3_code") %>%
    left_join(org_lbl, by="organisation") %>%
    # left_join() appends new columns at the end of the dataframe - relocate each *_label right
    # after its own *_code column so the pairs read together (admin1_code, admin1_label, admin2_code, ...).
    relocate(admin1_label, .after = admin1_code) %>%
    relocate(admin2_label, .after = admin2_code) %>%
    relocate(admin3_label, .after = admin3_code)
}

#############################################################################################################
# PERIODE D'EVALUATION (derivee des donnees, pas codee en dur)
#############################################################################################################

# Determine le cycle directement a partir de la variable kobo "date" des donnees (le jour de la visite
# du marche) plutot qu'une valeur codee en dur dans .Renviron, pour que le libelle/mois du cycle
# corresponde toujours a ce qui a ete reellement collecte. Prend la plus ancienne valeur "date" et en
# extrait le mois.
get_assessment_period <- function(df, date_col = "date"){
  mois_fr <- c("Janvier","Fevrier","Mars","Avril","Mai","Juin",
               "Juillet","Aout","Septembre","Octobre","Novembre","Decembre")

  d <- as.Date(df[[date_col]])
  d <- d[!is.na(d)]
  start_date <- min(d)
  end_date <- max(d)

  list(
    assessment_month = format(start_date, "%Y-%m"),
    assessment_month_label = mois_fr[as.integer(format(start_date, "%m"))],
    collection_start = as.character(start_date),
    collection_end = as.character(end_date)
  )
}

#############################################################################################################
# Fonctions cleaningtools modifiees
#############################################################################################################
check_value_modified <- function(dataset,
                                 uuid_column = "uuid",
                                 element_name = "checked_dataset",
                                 values_to_look = c(99, 999, -999, 88, 888, -888),
                                 columns_not_to_check = NULL) {
  values_to_look <- values_to_look |> as.character()

  ######### verification de l'entree

  if (!is.list(dataset)) {
    stop("L'entree doit etre un dataframe ou une liste.")
  }

  checked_dataset <- dataset

  if (!is.data.frame(dataset) & is.list(dataset)) {
    if (is.null(element_name)) {
      stop("element_name est manquant")
    }
    if (!element_name %in% names(dataset)) {
      stop("element_name introuvable")
    }
  }

  if (!is.data.frame(dataset) & is.list(dataset)) {
    dataset <- dataset[[element_name]]
  }

  #######################

  dataset <- dataset |> dplyr::rename(uuid = !!rlang::sym(uuid_column))

  # Exclure columns_not_to_check si specifie

  if (!is.null(columns_not_to_check)) {
    dataset <- dataset |> dplyr::select(-dplyr::any_of(columns_not_to_check))
  }

  dataset <- dataset |> dplyr::mutate_all(as.character)

  dataset_only_na <- dataset |> dplyr::filter_all(dplyr::any_vars(. %in% values_to_look))

  flaged_value <- dataset_only_na |>
    tidyr::pivot_longer(cols = !uuid) |>
    dplyr::filter(value %in% values_to_look) |>
    dplyr::rename(
      question = name,
      old_value = value
    ) |>
    dplyr::mutate(issue = "Valeur bizarre (comme 999, 88...) - a verifier")

  ## creation de la sortie
  if (is.data.frame(checked_dataset)) {
    return(list(
      checked_dataset = checked_dataset,
      flaged_value = flaged_value
    ))
  }

  if (!is.data.frame(checked_dataset)) {
    list_dataset <- list(flaged_value = flaged_value)

    return(append(checked_dataset, list_dataset))
  }
}

#############################################################################################################
# REPONSES "AUTRE"
#############################################################################################################

get_other_labels <- function(){
  other_labels <- kobo_survey %>%
    filter((type =="text") &
             !name %in% c("enum_id", "shop_name", "egg_other_unit_item", "comments")) %>%
    mutate(ref_question=as.character(lapply(relevant, get.ref.question))) %>%
    select(name, ref_question) %>%
    mutate(ref_question = ifelse(is.na(ref_question), name, ref_question)) %>%
    left_join(select(kobo_survey, ref_question=name, full_label=`label::french`), by = "ref_question")

  cat(green(" - SAUVEGARDE (ressources/labels_questions_others.xlsx) ... \n"))
  write.xlsx(other_labels, "ressources/labels_questions_others.xlsx", overwrite=T)
  return(other_labels)
}

################################################################################

get_other_db <- function(){
  # generer other_db
  other_db <- other_labels %>%
    left_join(select(kobo_survey, name, q_type, list_name), by=c("ref_question"="name")) %>%
    left_join(select(kobo_survey, name, relevant), by="name") %>%
    mutate(option_other=str_replace_all(str_extract(relevant, "\'.*\'"), "'", "")) %>%
    select(-relevant)

  # retirer toutes les option_other des choix
  kobo_choices_sub <- filter(kobo_choices, list_name %in% other_db$list_name)
  for (r in 1:nrow(other_db)){
    if (!is.na(other_db$option_other[r])){
      kobo_choices_sub <- kobo_choices_sub %>%
        filter(!(list_name==other_db$list_name[r] & name==other_db$option_other[r]))
    }
  }
  # ajouter la liste des choix disponibles
  other_db <- other_db %>%
    left_join(select(kobo_choices_sub, list_name, label) %>%
                group_by(list_name) %>%
                summarise(num_choices=n(), choices=paste0(label, collapse=";;")),
              by="list_name")

  return(other_db)
}
################################################################################
add_to_cleaning_log_other_remove <- function(x){
  change_type <- "Removing other response"

  # charger l'option other
  option_other <- other_db$option_other[other_db$name==x$name]
  var_option_other <- paste0(x$ref_name, "/", option_other)

  # retirer le texte de la reponse
  df <- data.frame(uuid=x$uuid, question=x$name, change_type=change_type,
                   old_value=x$response, new_value=NA)
  cleaning_log_other <<- rbind(cleaning_log_other, df)

  # retirer les entrees liees
  if (x$ref_type=="select_one"){
    df <- data.frame(uuid=x$uuid, question=x$ref_name, change_type=change_type,
                     old_value=option_other, new_value=NA)
    cleaning_log_other <<- rbind(cleaning_log_other, df)
  } else if (x$ref_type=="select_multiple"){
    old_concat_value <- get_value_from_uuid(x$uuid, x$ref_name)
    new_concat_value <- remove_choice(old_concat_value, option_other)
    new_concat_value <- ifelse(new_concat_value=="", NA, new_concat_value)
    df <- data.frame(uuid=x$uuid, question=x$ref_name, change_type=change_type,
                     old_value=old_concat_value, new_value=new_concat_value)
    cleaning_log_other <<- rbind(cleaning_log_other, df)

    if (is.na(new_concat_value)){
      # mettre toutes les options a NA
      cols <- colnames(clean_data)[str_starts(colnames(clean_data), paste0(x$ref_name, "/"))]
      old_values <- clean_data[clean_data$uuid==x$uuid, all_of(cols)]

      old_values <- as.character(old_values)
      if (length(cols)!=length(old_values)) stop("cols et old_values n'ont pas la meme longueur")
      for (i in 1:length(cols)){
        df <- data.frame(uuid=x$uuid, question=cols[i], change_type=change_type,
                         old_value=old_values[i], new_value=NA)
        cleaning_log_other <<- rbind(cleaning_log_other, df)
      }
    } else{
      # mettre var_option_other a "0"
      df <- data.frame(uuid=x$uuid, question=var_option_other, change_type=change_type,
                       old_value="1", new_value="0")
      cleaning_log_other <<- rbind(cleaning_log_other, df)
    }
  } else if (x$ref_type=="text"){
    df <- data.frame(uuid=x$uuid, question=x$ref_name, change_type=change_type,
                     old_value=x$response , new_value=NA)
    cleaning_log_other <<- rbind(cleaning_log_other, df)

  }
}
################################################################################
add_to_cleaning_log_other_recode <- function(x){
  if (x$ref_type[1]=="select_one") add_to_cleaning_log_other_recode_one(x)
  if (x$ref_type[1]=="select_multiple") add_to_cleaning_log_other_recode_multiple(x)
}
################################################################################
add_to_cleaning_log_other_recode_one <- function(x){
  change_type <- "Recoding other response"

  # retirer le texte de la reponse
  df <- data.frame(uuid=x$uuid, question=x$name, change_type=change_type,
                   old_value=x$response, new_value=NA)
  cleaning_log_other <<- rbind(cleaning_log_other, df)

  # recoder le choix
  new_value <- get_name_from_label(x$list_name, x$existing_other)
  if (length(new_value)!=1){
    stop(paste0("Le choix n'est pas dans la liste : ", x$uuid, "; ", x$list_name, "; ", x$existing_other))
  } else{
    df <- data.frame(uuid=x$uuid, question=x$ref_name, change_type=change_type,
                     old_value=get_value_from_uuid(x$uuid, x$ref_name), new_value=new_value)
    cleaning_log_other <<- rbind(cleaning_log_other, df)
  }
}
################################################################################
add_to_cleaning_log_other_recode_multiple <- function(x){
  change_type <- "Recoding other response"

  # recuperer l'option other
  option_other <- other_db$option_other[other_db$name==x$name]
  var_option_other <- paste0(x$ref_name, "/", option_other)

  # retirer le texte de la reponse
  df <- data.frame(uuid=x$uuid, question=x$name, change_type=change_type,
                   old_value=x$response, new_value=NA)
  cleaning_log_other <<- rbind(cleaning_log_other, df)

  # mettre l'option other a "0" et les choix selectionnes a "1" (si pas deja "1")
  choices <- unlist(lapply(str_split(x$existing_other, ";")[[1]],
                           function(c) get_name_from_label(x$list_name, c)))
  if (option_other %in% choices) warning(paste0(x$name, ": ajout a nouveau de l'option autre"))
  el <- list(uuid=x$uuid, ref_name=other_db$ref_question[other_db$name==x$name], change_type=change_type)
  cleaning_log_other <<- rbind(cleaning_log_other,
                               select_multiple_add_remove(el, to_remove=c(option_other), to_add=choices))
}
################################################################################
select_multiple_add_remove <- function(el, to_remove, to_add=c()){
  # charger les options exclusives
  exclusive_options <- c()
  # recuperer les noms de colonnes
  cols <- colnames(clean_data)[str_starts(colnames(clean_data), paste0(el$ref_name, "/"))]
  # generer le journal de nettoyage
  cl <- data.frame()
  if (is.na(get_value_from_uuid(el$uuid, el$ref_name)) & length(to_remove)>0) stop()
  if (is.na(get_value_from_uuid(el$uuid, el$ref_name))){
    #---------------------------------------------------------------------------
    # CAS 1) l'ancienne valeur est NA
    if (length(exclusive_options)>0 & any(exclusive_options %in% to_add)) stop("A implementer si necessaire")
    concat <- ""
    for (col in cols){
      choice <- str_split(col, "/")[[1]][2]
      if (choice %in% to_add) new_value <- "1"
      else new_value <- "0"
      cl <- rbind(cl, data.frame(uuid=el$uuid, question=col,
                                 old_value=NA, new_value=new_value))
      if (choice %in% to_add) concat <- add_choice(concat, choice)
    }
    if (concat=="") stop()
    cl <- rbind(cl, data.frame(uuid=el$uuid, question=el$ref_name,
                               old_value=NA, new_value=trimws(concat)))
  } else{
    #---------------------------------------------------------------------------
    # CAS 2) l'ancienne valeur n'est pas NA
    old_concat <- get_value_from_uuid(el$uuid, el$ref_name)
    if (is.na(old_concat) | old_concat=="") stop()
    new_concat <- old_concat
    # retirer des options
    if (length(to_remove)>0){
      for (choice in to_remove){
        cl <- rbind(cl, data.frame(uuid=el$uuid, question=paste0(el$ref_name, "/", choice),
                                   old_value="1", new_value="0"))
        new_concat <- remove_choice(new_concat, choice)
      }
    }
    # ajouter des options
    if (length(to_add)>0){
      if (any(exclusive_options %in% to_add)){
        print(paste0("Recodage d'une option exclusive : ", el$uuid, " --> ", el$ref_name, " --> ", to_add[1]))
        if (length(to_add)>1) stop("Impossible de selectionner une option exclusive avec d'autres options")
        cl <- data.frame()
        for (col in cols){
          option <- str_split(col, "/")[[1]][2]
          old_value <- get_value_from_uuid(el$uuid, col)
          if (option==to_add[1]){
            cl <- rbind(cl, data.frame(uuid=el$uuid, question=col,
                                       old_value="0", new_value="1"))
            new_concat <- to_add[1]
          } else if (old_value=="1"){
            cl <- rbind(cl, data.frame(uuid=el$uuid, question=col,
                                       old_value="1", new_value="0"))
          }
        }
      } else{
        for (choice in to_add){
          old_value <- get_value_from_uuid(el$uuid, paste0(el$ref_name, "/", choice))
          if (old_value=="0"){
            cl <- rbind(cl, data.frame(uuid=el$uuid, question=paste0(el$ref_name, "/", choice),
                                       old_value="0", new_value="1"))
            new_concat <- add_choice(new_concat, choice)
          }
        }
      }
    }
    #---------------------------------------------------------------------------
    # soit mettre a jour la colonne concatenee, soit tout mettre a NA si new_concat est vide
    if (new_concat!="" & new_concat!=old_concat){
      cl <- rbind(cl, data.frame(uuid=el$uuid, question=el$ref_name,
                                 old_value=old_concat, new_value=trimws(new_concat)))
    } else if (new_concat==""){
      cl <- data.frame()
      for (col in cols){
        cl <- rbind(cl, data.frame(uuid=el$uuid, question=col,
                                   old_value=get_value_from_uuid(el$uuid, col),
                                   new_value=NA))
      }
      cl <- rbind(cl, data.frame(uuid=el$uuid, question=el$ref_name,
                                 old_value=old_concat, new_value=NA))
    }
  }
  return(cl %>% mutate(change_type=el$change_type))
}
################################################################################
# Ajoute la feuille des reponses "autre" (+ sa feuille "Listes_deroulantes" de validation) a un
# classeur existant plutot que d'en creer un nouveau - permet de combiner ce contenu avec la feuille
# "A_verifier" des verifications (voir write_review_sheet() / 2_nettoyage_final.R) dans un seul
# classeur "a verifier", au lieu de deux fichiers separes. L'appelant cree/charge
# `wb` et sauvegarde le classeur combine lui-meme une fois toutes les feuilles ajoutees.
add_other_responses_sheet <- function(wb, df, sheet_name = "reponses_autres") {

  get_column_letter <- function(r) openxlsx::int2col(r)

  # noms de feuille Excel limites a 31 caracteres
  dropdown_sheet <- substr(paste0("Listes_", sheet_name), 1, 31)

  # definir les styles
  style.col.color <- createStyle(fgFill = "#DDDDDE", border = "TopBottomLeftRight",
                                 valign = "top", wrapText = T)
  style.col.color1 <- createStyle(fgFill = "#DBD5C6", border = "TopBottomLeftRight", borderColour = "#000000",
                                  valign = "top", wrapText = T)
  style.col.color2 <- createStyle(fgFill = "#F3F4F4", border = "TopBottomLeftRight",
                                  valign = "top", wrapText = T)
  style.col.color.first <- createStyle(textDecoration = "bold", fgFill = "#F1797A", valign = "top",
                                       border = "TopBottomLeftRight", borderColour = "#000000", wrapText = T)

  # En-tetes en francais pour l'equipe de suivi - purement pour l'affichage : le reste de cette
  # fonction continue de travailler avec les noms de colonnes internes de `df` (df$question_name,
  # etc.), seule cette copie ecrite dans le classeur porte les libelles francais.
  # 2_nettoyage_final.R (2eme passe) retraduit ces libelles vers les noms internes a la lecture (voir
  # la-bas).
  header_fr <- c(
    q_type = "type_question", list_name = "liste_choix", question_name = "nom_question",
    full_label = "libelle_complet", response = "reponse", selected_choices = "choix_selectionnes",
    "TRUE other (copy response or provide a better translation)" = "VRAI autre (copier la reponse ou fournir une meilleure traduction)",
    "EXISTING other 1 (select the most appropriate choice)" = "EXISTANT autre 1 (choisir la reponse la plus appropriee)",
    "EXISTING other 2 (select the most appropriate choice)" = "EXISTANT autre 2 (choisir la reponse la plus appropriee)",
    "EXISTING other 3 (select the most appropriate choice)" = "EXISTANT autre 3 (choisir la reponse la plus appropriee)",
    "INVALID other (select yes or leave blank)" = "INVALIDE autre (selectionner oui ou laisser vide)",
    "FOLLOW-UP message (what is unclear about this response?)" = "Message de SUIVI (qu'est-ce qui n'est pas clair dans cette reponse ?)",
    "Explanation" = "Explication"
  )
  df_display <- df
  match_idx <- match(names(df_display), names(header_fr))
  names(df_display)[!is.na(match_idx)] <- header_fr[match_idx[!is.na(match_idx)]]

  # ajouter la feuille
  addWorksheet(wb, sheet_name)
  writeData(wb = wb, x = df_display, sheet = sheet_name, startRow = 1)

  # appliquer les styles
  addStyle(wb, sheet_name, style = style.col.color.first, rows = 1:(nrow(df) + 1), cols = 1) # couleur de la premiere colonne
  addStyle(wb, sheet_name, style = style.col.color2, rows = 1:(nrow(df) + 1), cols = 14)
  addStyle(wb, sheet_name, style = style.col.color, rows = 1:(nrow(df) + 1), cols = 15)
  addStyle(wb, sheet_name, style = style.col.color1, rows = 1:(nrow(df) + 1), cols = 16)
  addStyle(wb, sheet_name, style = style.col.color1, rows = 1:(nrow(df) + 1), cols = 17)
  addStyle(wb, sheet_name, style = style.col.color1, rows = 1:(nrow(df) + 1), cols = 18)
  addStyle(wb, sheet_name, style = style.col.color, rows = 1:(nrow(df) + 1), cols = 19)
  addStyle(wb, sheet_name, style = style.col.color2, rows = 1:(nrow(df) + 1), cols = 20)
  addStyle(wb, sheet_name, style = style.col.color2, rows = 1:(nrow(df) + 1), cols = 21)

  # geler la premiere ligne
  freezePane(wb, sheet = sheet_name, firstActiveRow = 2)

  # ajouter les filtres de colonne
  addFilter(wb, sheet_name, rows = 1, cols = 1:ncol(df))

  addStyle(wb, sheet_name, style = createStyle(wrapText = TRUE), rows = 1:21, cols = 1)

  # ajouter la feuille des valeurs de liste deroulante et la validation des donnees
  addWorksheet(wb, dropdown_sheet)
  for (r in seq_len(nrow(other_db))) {

    if (other_db$q_type[r] != "text") {

      choices <- str_split(other_db$choices[r], ";;")[[1]]
      choices <- choices[choices != ""]

      if (length(choices) == 0) next

      writeData(wb, dropdown_sheet, choices, startCol = r)

      uuids <- which(df$question_name == other_db$name[r])
      if (length(uuids) == 0) next

      col_letter <- get_column_letter(r)
      n_choices <- length(choices)

      values <- paste0(
        "'", dropdown_sheet, "'!$",
        col_letter,
        "$1:$",
        col_letter,
        "$",
        n_choices
      )

      dataValidation(wb, sheet_name, col = 16, rows = uuids + 1, type = "list", value = values)
      dataValidation(wb, sheet_name, col = 17, rows = uuids + 1, type = "list", value = values)
      dataValidation(wb, sheet_name, col = 18, rows = uuids + 1, type = "list", value = values)
    }
  }

  # validation Oui / Non (colonne de reserve)
  yes_col <- nrow(other_db) + 2
  writeData(wb, dropdown_sheet, "Oui", startCol = yes_col)

  yes_letter <- get_column_letter(yes_col)
  yes_values <- paste0("'", dropdown_sheet, "'!$", yes_letter, "$1:$", yes_letter, "$1")

  dataValidation(
    wb, sheet_name,
    col = 19,
    rows = 2:(nrow(df) + 1),
    type = "list",
    value = yes_values
  )

  # definir les largeurs de colonnes
  setColWidths(wb, sheet_name, cols = 1:10, widths = 15)

  setColWidths(wb, sheet_name, cols = 11:12, widths = 19)
  setColWidths(wb, sheet_name, cols = 12:18, widths = 30)

  # appliquer l'alignement et le retour a la ligne du texte
  addStyle(wb, sheet_name, style = createStyle(valign = "top"), rows = 1:(nrow(df) + 1), cols = 1)
  addStyle(wb, sheet_name, style = createStyle(valign = "top"), rows = 1:(nrow(df) + 1), cols = 2)
  addStyle(wb, sheet_name, style = createStyle(valign = "top"), rows = 1:(nrow(df) + 1), cols = 3)
  addStyle(wb, sheet_name, style = createStyle(valign = "top"), rows = 1:(nrow(df) + 1), cols = 4)
  addStyle(wb, sheet_name, style = createStyle(valign = "top"), rows = 1:(nrow(df) + 1), cols = 5)
  addStyle(wb, sheet_name, style = createStyle(valign = "top"), rows = 1:(nrow(df) + 1), cols = 6)
  addStyle(wb, sheet_name, style = createStyle(wrapText = T, valign = "top"), rows = 1:(nrow(df) + 1), cols = 7)
  addStyle(wb, sheet_name, style = createStyle(wrapText = T, valign = "top"), rows = 1:(nrow(df) + 1), cols = 8)
  addStyle(wb, sheet_name, style = createStyle(wrapText = T, valign = "top"), rows = 1:(nrow(df) + 1), cols = 9)
  addStyle(wb, sheet_name, style = createStyle(wrapText = T, valign = "top"), rows = 1:(nrow(df) + 1), cols = 10)

  setRowHeights(wb, sheet_name, rows = 1, heights = 15)

  # premiere ligne en gras
  addStyle(wb, sheet_name, style = createStyle(textDecoration = "bold"), rows = 1, cols = 1:ncol(df))
  addStyle(wb, sheet_name, style = style.col.color.first, rows = 1, cols = 1:18)

  modifyBaseFont(wb, fontSize = 10, fontColour = "black", fontName = "Calibri")

  invisible(wb)
}

#############################################################################################################
# FEUILLE DE VERIFICATION CONSOLIDEE ("A_verifier") - accessible a une equipe de suivi avec un faible
# niveau de formation technique/d'alphabetisation
#############################################################################################################

# En-tetes affiches (colonne A) <- nom de colonne technique utilise dans le reste du pipeline (colonne
# B). Un seul tableau au lieu de plusieurs, en langage simple, avec le nom de variable Kobo technique
# et l'identifiant d'enquete presents mais clairement marques "ne pas modifier" (necessaires pour
# reappliquer les corrections automatiquement, mais ce ne sont pas les colonnes que l'equipe doit lire).
REVIEW_HEADERS_FR <- c(
  uuid = "Identifiant enquete (ne pas modifier)",
  commune = "Commune",
  partenaire = "Partenaire",
  enqueteur = "Enqueteur",
  boutique = "Boutique / vendeur",
  variable = "Variable (ne pas modifier)",
  question_label = "Question",
  reponse_actuelle = "Reponse actuelle",
  probleme = "Probleme detecte",
  action = "Que faire ?",
  nouvelle_reponse = "Nouvelle reponse (si Corriger la reponse)",
  commentaire = "Commentaire"
)

# Options de la liste deroulante "Que faire ?", en langage simple <- valeur technique attendue par
# cleaningtools::create_clean_data()/review_cleaning_log() (change_response/blank_response/no_action/
# remove_survey). L'equipe de suivi ne voit jamais ces noms techniques.
ACTION_LABELS_FR <- c(
  change_response = "Corriger la reponse",
  blank_response = "Mettre une valeur manquante",
  no_action = "Ne rien changer",
  remove_survey = "Supprimer toute l'enquete"
)

# Etiquette lisible d'une variable kobo pour la colonne "Question" - utilise le libelle francais de la
# question quand il existe (la plupart des variables saisies directement), sinon replie sur le nom
# technique avec les underscores remplaces par des espaces (cas des champs "calculate" comme
# <article>_price_item, qui n'ont pas de libelle puisqu'ils ne sont jamais poses directement) - toujours
# quelque chose de lisible plutot que de faire echouer l'export.
get_question_label <- function(question){
  if (is.na(question)) return("Toute l'enquete")
  lbl <- kobo_survey$`label::french`[kobo_survey$name == question]
  lbl <- lbl[!is.na(lbl) & nzchar(lbl)]
  if (length(lbl) > 0) return(lbl[1])
  str_replace_all(question, "_", " ")
}

# Ecrit la feuille de verification consolidee dans un classeur existant (meme logique que
# add_other_responses_sheet() : l'appelant cree/charge `wb` et sauvegarde une fois toutes les feuilles
# ajoutees). `df` doit avoir les colonnes techniques de REVIEW_HEADERS_FR ; "action" peut etre vide (a
# remplir par l'equipe) ou deja une valeur de ACTION_LABELS_FR.
write_review_sheet <- function(wb, df, sheet_name = "A_verifier") {
  df_display <- df %>% select(names(REVIEW_HEADERS_FR))
  colnames(df_display) <- unname(REVIEW_HEADERS_FR[colnames(df_display)])

  addWorksheet(wb, sheet_name)
  writeData(wb, sheet_name, df_display, colNames = TRUE)

  n <- nrow(df_display) + 1
  col_of_display <- function(key) match(REVIEW_HEADERS_FR[[key]], colnames(df_display))
  action_col <- col_of_display("action")
  uuid_col <- col_of_display("uuid")
  variable_col <- col_of_display("variable")

  # en-tete : fond rouge, texte blanc, gras, figee, avec filtres
  addStyle(wb, sheet_name, style = createStyle(fgFill = RED_HEADER, fontColour = "#FFFFFF",
                                               textDecoration = "bold", wrapText = TRUE, valign = "center",
                                               halign = "center"),
           rows = 1, cols = 1:ncol(df_display))
  freezePane(wb, sheet_name, firstActiveRow = 2, firstActiveCol = 2)
  addFilter(wb, sheet_name, rows = 1, cols = 1:ncol(df_display))

  # texte avec retour a la ligne pour toutes les colonnes de contenu
  addStyle(wb, sheet_name, style = createStyle(wrapText = TRUE, valign = "top"),
           rows = 2:n, cols = 1:ncol(df_display), gridExpand = TRUE)

  # colonnes techniques (a ne pas modifier) : fond gris pour bien les distinguer des colonnes a remplir
  addStyle(wb, sheet_name, style = createStyle(fgFill = GREY_HEADER, fontColour = "#595959", wrapText = TRUE, valign = "top"),
           rows = 2:n, cols = c(uuid_col, variable_col), gridExpand = TRUE)

  # liste deroulante pour la colonne "Que faire ?" - liste en ligne (pas besoin d'une feuille annexe,
  # les 4 options tiennent largement dans la limite de 255 caracteres d'Excel)
  action_choices <- paste0('"', paste(unname(ACTION_LABELS_FR), collapse = ","), '"')
  dataValidation(wb, sheet_name, col = action_col, rows = 2:n, type = "list", value = action_choices)

  setColWidths(wb, sheet_name, cols = 1:ncol(df_display),
              widths = c(20, 14, 16, 14, 18, 26, 30, 20, 32, 24, 26, 30))
  setRowHeights(wb, sheet_name, rows = 1, heights = 32)

  invisible(wb)
}

# Relit une feuille de verification consolidee remplie par l'equipe de suivi, et retraduit les
# en-tetes/la liste deroulante "Que faire ?" vers les noms/valeurs techniques attendus par
# cleaningtools (create_clean_data(), review_cleaning_log(), ...) - meme logique de "traduction a la
# frontiere" que pour la feuille reponses_autres : l'equipe ne voit jamais ces noms techniques, seul ce
# point de lecture les reintroduit.
read_review_sheet <- function(path, sheet_name = "A_verifier") {
  raw <- read_excel(path, sheet = sheet_name, col_types = "text")
  header_rev <- setNames(names(REVIEW_HEADERS_FR), unname(REVIEW_HEADERS_FR))
  match_idx <- match(colnames(raw), names(header_rev))
  colnames(raw)[!is.na(match_idx)] <- header_rev[match_idx[!is.na(match_idx)]]

  action_rev <- setNames(names(ACTION_LABELS_FR), unname(ACTION_LABELS_FR))
  raw %>%
    mutate(change_type = unname(action_rev[action])) %>%
    rename(question = variable, old_value = reponse_actuelle, new_value = nouvelle_reponse,
          issue = probleme, comment = commentaire)
}

#############################################################################################################
# UTILITAIRES
#############################################################################################################
get_name_from_label <- function(list_name, label){
  return(kobo_choices$name[kobo_choices$list_name==list_name & kobo_choices$label==label])
}
################################################################################
get_label_from_name <- function(list.name, name){
  return(kobo_choices$label[kobo_choices$list_name==list.name & kobo_choices$name==name])
}
################################################################################
get_value_from_uuid <- function(uuid, column){
  value <- raw_data_all[[column]][raw_data_all$uuid==uuid]
  return(value)
}
################################################################################
add_choice <- function(concat_value, choice){
  l <- str_split(concat_value, " ")[[1]]
  l <- sort(unique(c(l, choice)))
  l <- l[l!=""]
  return(paste(l, collapse=" "))
}
################################################################################
remove_choice <- function(concat_value, choice){
  l <- str_split(concat_value, " ")[[1]]
  l <- l[l!=choice]
  return(paste(l, collapse=" "))
}
################################################################################
get.ref.question <- function(x){
  x.1 <- str_split(x, "\\{")[[1]][2]
  return(str_split(x.1, "\\}")[[1]][1])
}
################################################################################
# Le package `cleaningtools` genere lui-meme certains textes de la colonne "issue" en anglais
# (check_duplicate, check_outliers, check_soft_duplicates, check_percentage_missing) - cette fonction
# les retraduit en francais juste avant l'export, en langage simple et sans jargon statistique/
# technique, puisque la feuille de verification consolidee (voir write_review_sheet()) est destinee a
# une equipe de suivi terrain qui n'est pas necessairement familiere avec ce vocabulaire. La colonne
# "issue" n'est jamais utilisee pour un filtrage/appariement en aval, uniquement affichee.
translate_issue_fr <- function(issue){
  issue <- str_replace(issue, "^duplicated (.+)$", "Doublon exact sur \\1")
  issue <- str_replace(issue, "^outlier \\((normal|log) distribution\\)$", "Valeur inhabituelle - a verifier")
  issue <- str_replace(issue, "^Less than \\d+ differents? options$",
                       "Reponses tres semblables a une autre enquete - possible doublon")
  issue <- str_replace(issue, "^Percentages? of missing values.*$",
                       "Beaucoup de reponses manquantes dans cette enquete")
  issue
}

################################################################################
# Fonctions du suivi quotidien des enquetes (dont le comptage des prix)
################################################################################

# Ajoute les colonnes de contexte "prefecture"/"sous_prefecture" (le parent administratif direct de
# chaque ligne) et "organisation" (le(s) partenaire(s) qui couvre(nt) cette commune) a un tableau
# zone/niveau_admin deja construit, pour qu'on sache toujours "ou on est" et "qui" sans avoir a croiser
# une autre feuille - en particulier pour une ligne "commune", qui a elle seule ne dit ni dans quelle
# sous-prefecture/prefecture elle se trouve, ni quel partenaire y collecte. Laisse ces colonnes vides
# quand elles seraient redondantes/non pertinentes (prefecture/sous_prefecture pour les lignes
# sous-prefecture/prefecture/national ; organisation pour toute ligne autre que commune, puisqu'un
# niveau agrege peut regrouper plusieurs partenaires differents).
add_admin_context <- function(df, admin_lookup, org_lookup) {
  admin2_to_admin1 <- admin_lookup %>% distinct(admin1_label, admin2_label)

  df %>%
    left_join(admin_lookup, by = c("zone" = "admin3_label")) %>%
    left_join(admin2_to_admin1, by = c("zone" = "admin2_label"), suffix = c("", "_via_sp")) %>%
    left_join(org_lookup, by = c("zone" = "admin3_label")) %>%
    mutate(
      prefecture = case_when(
        niveau_admin == "commune" ~ admin1_label,
        niveau_admin == "sous_prefecture" ~ admin1_label_via_sp,
        TRUE ~ NA_character_
      ),
      sous_prefecture = case_when(
        niveau_admin == "commune" ~ admin2_label,
        TRUE ~ NA_character_
      ),
      organisation = case_when(
        niveau_admin == "commune" ~ organisation,
        TRUE ~ NA_character_
      )
    ) %>%
    select(-any_of(c("admin1_label", "admin2_label", "admin1_label_via_sp"))) %>%
    relocate(prefecture, sous_prefecture, organisation, .after = zone)
}

# Tableau "commune -> liste des partenaires y ayant collecte" (plusieurs partenaires possibles pour
# une meme commune, concatenes) - utilise par add_admin_context() pour la colonne "organisation".
build_org_lookup <- function(data) {
  data %>%
    filter(!is.na(admin3_label)) %>%
    group_by(admin3_label) %>%
    summarise(organisation = paste(sort(unique(organisation_name_label)), collapse = "; "), .groups = "drop")
}

# Fonction pour calculer le nombre de prix pour chaque article a tous les niveaux admin. La commune
# est le niveau le plus utile pour le suivi de la couverture (c'est la ou un enqueteur/partenaire doit
# reellement combler un manque) - elle est donc placee en premier, suivie de la sous-prefecture, de la
# prefecture, puis du total national. Les colonnes prefecture/sous_prefecture (voir add_admin_context())
# donnent le contexte administratif complet de chaque commune.
get_num_prices <- function(data) {
  cols <- colnames(data)[str_detect(colnames(data), paste(items_list, collapse = "|"))]
  admin_lookup <- data %>% distinct(admin1_label, admin2_label, admin3_label)
  org_lookup <- build_org_lookup(data)
  admin_levels <- list(
    admin3 = data %>% group_by(admin3_label) %>%
      summarise(across(all_of(cols), ~sum(!is.na(.)))) %>%
      rename(zone = admin3_label) %>% mutate(niveau_admin = "commune"),
    admin2 = data %>% group_by(admin2_label) %>%
      summarise(across(all_of(cols), ~sum(!is.na(.)))) %>%
      rename(zone = admin2_label) %>% mutate(niveau_admin = "sous_prefecture"),
    admin1 = data %>% group_by(admin1_label) %>%
      summarise(across(all_of(cols), ~sum(!is.na(.)))) %>%
      rename(zone = admin1_label) %>% mutate(niveau_admin = "prefecture"),
    admin0 = data %>% select(all_of(cols)) %>%
      summarise_all(~sum(!is.na(.))) %>%
      mutate(zone = "RCA", niveau_admin = "national")
  )
  bind_rows(admin_levels) %>%
    relocate(zone, niveau_admin, .before = 1) %>%
    add_admin_context(admin_lookup, org_lookup)
}

# Fonction pour calculer le prix median de chaque article a tous les niveaux admin (meme structure/
# ordre que get_num_prices(), commune en premier, memes colonnes prefecture/sous_prefecture) -
# utilisee pour la feuille "Prix_Median" du suivi quotidien, avec une echelle de couleur rouge (prix
# eleve) a vert (prix bas) CALCULEE SEPAREMENT POUR CHAQUE ARTICLE (voir apply_price_color_scale()) -
# les articles n'ont pas la meme gamme de prix, une seule echelle commune a tous rendrait la couleur
# d'un article non comparable a elle-meme d'une zone a l'autre.
get_median_prices <- function(data) {
  cols <- items_list[items_list %in% colnames(data)]
  med <- function(x) suppressWarnings(median(as.numeric(x), na.rm = TRUE))
  admin_lookup <- data %>% distinct(admin1_label, admin2_label, admin3_label)
  org_lookup <- build_org_lookup(data)
  admin_levels <- list(
    admin3 = data %>% group_by(admin3_label) %>%
      summarise(across(all_of(cols), med)) %>%
      rename(zone = admin3_label) %>% mutate(niveau_admin = "commune"),
    admin2 = data %>% group_by(admin2_label) %>%
      summarise(across(all_of(cols), med)) %>%
      rename(zone = admin2_label) %>% mutate(niveau_admin = "sous_prefecture"),
    admin1 = data %>% group_by(admin1_label) %>%
      summarise(across(all_of(cols), med)) %>%
      rename(zone = admin1_label) %>% mutate(niveau_admin = "prefecture"),
    admin0 = data %>% select(all_of(cols)) %>%
      summarise(across(everything(), med)) %>%
      mutate(zone = "RCA", niveau_admin = "national")
  )
  bind_rows(admin_levels) %>%
    mutate(across(all_of(cols), ~ifelse(is.nan(.) | is.infinite(.), NA, .))) %>%
    relocate(zone, niveau_admin, .before = 1) %>%
    add_admin_context(admin_lookup, org_lookup)
}

# Applique une echelle de couleur (vert = prix bas, jaune = median, rouge = prix eleve) sur une plage
# de cellules numeriques - utilisee sur la feuille "Prix_Median" pour reperer visuellement les zones/
# articles ou le prix est eleve, sans avoir besoin d'un partenaire technique pour interpreter le tableau.
# Applique une echelle SEPAREE PAR COLONNE (par article) plutot qu'une seule echelle sur toute la
# plage : chaque article a sa propre gamme de prix (le riz et le savon ne se comparent pas), donc
# calculer le rouge/vert par article - et non a travers tous les articles a la fois - est le seul
# reperage qui ait un sens.
apply_price_color_scale <- function(wb, sheet_name, rows, cols) {
  for (col in cols) {
    conditionalFormatting(wb, sheet = sheet_name, rows = rows, cols = col,
                          type = "colourScale",
                          style = c("#63BE7B", "#FFEB84", "#F8696B"))
  }
}

# Fonction pour ajouter une feuille avec mise en forme (en-tete rose standard, une seule ligne)
add_styled_sheet <- function(workbook, sheet_name, data) {
  header_style <- createStyle(
    fontSize = 12,
    fontColour = "#FFFFFF",
    bgFill = "#F28B82",
    textDecoration = c("Italic", "Bold"),
    halign = "center",
    valign = "center"
  )
  addWorksheet(workbook, sheet_name)
  writeData(workbook, sheet_name, data, headerStyle = header_style)
  setColWidths(workbook, sheet_name, cols = 1:ncol(data), widths = "auto")
  freezePane(workbook, sheet_name, firstActiveRow = 2)
}

# Libelles des bannieres de section (voir write_priority_sheet()), dans l'ordre d'affichage.
NIVEAU_ADMIN_LABELS <- c(
  commune = "COMMUNES",
  sous_prefecture = "SOUS-PREFECTURES",
  prefecture = "PREFECTURES",
  national = "NATIONAL"
)

# Ecrit les feuilles prioritaires du suivi quotidien (Nombre_Prix, Prix_Median) : en-tete standard
# (meme rose que add_styled_sheet()) mais sur DEUX lignes - ligne 1 = libelle de la question kobo pour
# les colonnes d'article (repli sur le nom de colonne pour zone/prefecture/sous_prefecture/niveau_admin,
# qui n'ont pas de libelle), ligne 2 = nom de colonne technique (maize_price_item, ...) - puis une
# BANNIERE GRISE DE SECTION avant chaque groupe de niveau_admin (communes, sous-prefectures,
# prefectures, national) pour bien separer visuellement les blocs. Renvoie, de maniere invisible, la
# ligne Excel reelle de chaque ligne de `data` (dans le meme ordre que `data`), pour que l'appelant
# puisse placer une mise en forme conditionnelle au bon endroit malgre les bannieres inserees.
write_priority_sheet <- function(workbook, sheet_name, data) {
  addWorksheet(workbook, sheet_name)

  # reutilise get_question_label() (voir plus haut, section feuille de verification) : libelle kobo
  # francais quand il existe, sinon le nom de colonne avec les underscores remplaces par des espaces.
  label_row <- vapply(names(data), get_question_label, character(1))

  label_style <- createStyle(textDecoration = "Bold", wrapText = TRUE, valign = "top", fgFill = GREY_HEADER)
  header_style <- createStyle(fontSize = 12, fontColour = "#FFFFFF", bgFill = "#F28B82",
                              textDecoration = c("Italic", "Bold"), halign = "center", valign = "center")
  banner_style <- createStyle(fgFill = DARK_GREY_BANNER, fontColour = "#FFFFFF", textDecoration = "Bold",
                              halign = "center", valign = "center")

  writeData(workbook, sheet_name, as.list(label_row), startRow = 1, colNames = FALSE)
  addStyle(workbook, sheet_name, style = label_style, rows = 1, cols = 1:ncol(data), gridExpand = TRUE)
  setRowHeights(workbook, sheet_name, rows = 1, heights = 32)

  writeData(workbook, sheet_name, as.list(names(data)), startRow = 2, colNames = FALSE)
  addStyle(workbook, sheet_name, style = header_style, rows = 2, cols = 1:ncol(data), gridExpand = TRUE)

  sheet_row_of <- integer(nrow(data))
  r <- 3
  for (lvl in names(NIVEAU_ADMIN_LABELS)) {
    idx <- which(data$niveau_admin == lvl)
    if (length(idx) == 0) next

    writeData(workbook, sheet_name, NIVEAU_ADMIN_LABELS[[lvl]], startRow = r, startCol = 1, colNames = FALSE)
    mergeCells(workbook, sheet_name, cols = 1:ncol(data), rows = r)
    addStyle(workbook, sheet_name, style = banner_style, rows = r, cols = 1:ncol(data), gridExpand = TRUE)
    r <- r + 1

    writeData(workbook, sheet_name, data[idx, ], startRow = r, colNames = FALSE)
    sheet_row_of[idx] <- r + seq_along(idx) - 1
    r <- r + length(idx)
  }

  setColWidths(workbook, sheet_name, cols = 1:ncol(data), widths = "auto")
  freezePane(workbook, sheet_name, firstActiveRow = 3)

  invisible(sheet_row_of)
}

# Appliquer une mise en forme conditionnelle
apply_conditional_formatting <- function(sheet_name, rows, cols, rule, style) {
  if (length(rule) > 1) {
    for (r in rule) {
      conditionalFormatting(wb, sheet = sheet_name, cols = cols, rows = rows,
                            rule = r, style = createStyle(bgFill = style))
    }
  } else {
    conditionalFormatting(wb, sheet = sheet_name, cols = cols, rows = rows,
                          rule = rule, style = createStyle(bgFill = style))
  }
}

create_partner_region_summary <- function(data,
                                          partner_col = "organisation_name_label",
                                          region_col = "admin1_label") {

  summary_df <- data %>%
    group_by(!!sym(partner_col), !!sym(region_col)) %>%
    summarise(n = n(), .groups = "drop") %>%
    pivot_wider(
      names_from = !!sym(region_col),
      values_from = n,
      values_fill = 0
    ) %>%
    mutate(Total = rowSums(across(where(is.numeric)), na.rm = TRUE))

  total_row <- summary_df %>%
    summarise(across(where(is.numeric), sum, na.rm = TRUE)) %>%
    mutate(!!partner_col := "Total")

  final_output <- bind_rows(summary_df, total_row) %>%
    relocate(!!sym(partner_col))

  return(final_output)
}

#############################################################################################################
# ETIQUETTES DE VARIABLES + CLASSEUR FINAL SELON LE MODELE DE LA CHECKLIST
#############################################################################################################

# Attache a chaque colonne son etiquette kobo francaise (via le package `labelled`) pour transporter
# l'etiquette avec les donnees jusqu'au moment ou write_labeled_data_sheet() l'affiche comme ligne
# d'etiquette de la feuille.
apply_var_labels <- function(df){
  lbls <- kobo_survey$`label::french`
  names(lbls) <- kobo_survey$name
  lbls <- as.list(lbls[intersect(names(df), names(lbls))])
  labelled::var_label(df) <- lbls
  df
}

GREY_HEADER <- "#D1D3D4"
RED_HEADER <- "#EE5859"
DARK_GREY_BANNER <- "#818183"

# Une seule ligne d'en-tete figee, fond gris clair - utilisee pour cleaning_log/deletion_log/variable_tracker/survey/choices
write_simple_sheet <- function(wb, sheet_name, df, auto_width = TRUE){
  addWorksheet(wb, sheet_name)
  writeData(wb, sheet_name, df, colNames = TRUE)
  addStyle(wb, sheet_name, style = createStyle(textDecoration = "bold", fgFill = GREY_HEADER), rows = 1, cols = 1:ncol(df))
  freezePane(wb, sheet_name, firstActiveRow = 2)
  if (auto_width) setColWidths(wb, sheet_name, cols = 1:ncol(df), widths = "auto")
}

# raw_data/clean_data : en-tete sur deux lignes - ligne 1 = etiquette de la variable (texte de la
# question, en gras), ligne 2 = nom/ID de la variable Kobo (en gras, fond #EE5859) - les deux sont
# figees. Les donnees commencent a la ligne 3.
write_labeled_data_sheet <- function(wb, sheet_name, df){
  addWorksheet(wb, sheet_name)
  var_labels <- labelled::var_label(df)
  label_row <- vapply(names(df), function(n) {
    lbl <- var_labels[[n]]
    if (is.null(lbl) || is.na(lbl) || !nzchar(lbl)) n else lbl
  }, character(1))

  writeData(wb, sheet_name, as.list(label_row), startRow = 1, colNames = FALSE)
  writeData(wb, sheet_name, df, startRow = 2, colNames = TRUE)

  addStyle(wb, sheet_name, style = createStyle(textDecoration = "bold", wrapText = TRUE, valign = "top",
                                               fgFill = GREY_HEADER),
           rows = 1, cols = 1:ncol(df))
  addStyle(wb, sheet_name, style = createStyle(fgFill = RED_HEADER, textDecoration = "bold", fontColour = "#FFFFFF",
                                               border = "Bottom", borderStyle = "thick"),
           rows = 2, cols = 1:ncol(df))

  setRowHeights(wb, sheet_name, rows = 1, heights = 45)
  setRowHeights(wb, sheet_name, rows = 2, heights = 22)
  setColWidths(wb, sheet_name, cols = 1:ncol(df), widths = 18)
  freezePane(wb, sheet_name, firstActiveRow = 3)
}

# Remplit la colonne A ("Checks (y/n/na)") de la feuille CHECKLIST_Data_Cleaning copiee : "y" pour
# chaque point couvert par ce cycle (y compris certains points - duree d'entretien, performance des
# enqueteurs, verifications sectorielles - que le script ne certifie pas lui-meme mais qui sont
# auto-remplis "y" sur demande du reviseur, plutot que laisses vides), "na" pour les points reellement
# non applicables a ce jeu de donnees/cycle (pas de ponderation, pas de boucle repetee, PH hors champ
# JMMI). Seuls les libelles de section (ex. "Colonnes obligatoires dans clean_data :") et les lignes
# vides du modele restent sans valeur. Les numeros de ligne correspondent a la feuille
# "CHECKLIST_Data_Cleaning" de ressources/Checklist-Template_..._FRENCH.xlsx.
fill_checklist_data_cleaning <- function(wb, has_logical_checks){
  sheet <- "CHECKLIST_Data_Cleaning"
  set_check <- function(row, value) writeData(wb, sheet, x = value, startRow = row, startCol = 1, colNames = FALSE)

  # Documentation, Transparence et Reproductibilite
  for (r in 7:14) set_check(r, "y")   # READ_ME, survey, choices, variable_tracker, raw_data, clean_data, cleaning_log, deletion_log (toujours incluse, meme vide)
  set_check(15, "na")  # sampling_info : pas de ponderation ce cycle
  set_check(16, "na")  # enumerator_performance_log : optionnel, non construit
  set_check(18, "y")   # uuid utilise comme identifiant standard
  set_check(19, "y")   # colonnes PCode admin1_code/admin2_code/admin3_code
  set_check(21, "na")  # rapports de plausibilite PH : non applicable (la JMMI est une donnee de marche/prix)
  set_check(22, "y")   # scripts R fournis pour la revue au siege

  # Qualite et Exactitude des Donnees
  set_check(27, "y")   # Identifiants uniques
  set_check(28, "y")   # Pas de doublons d'uuid
  set_check(29, "y")   # Pas de DPI
  set_check(30, "y")   # Consentement eclaire
  set_check(31, "y")   # Enquetes de courte duree
  set_check(32, "y")   # Pas de valeurs entieres NA (verification 99/999/-999/88/888/-888)
  set_check(33, "y")   # Entrees textuelles (traduites)
  set_check(34, "y")   # Autres options (recodees)
  set_check(36, "y")   # Enquetes quasi-dupliquees (check_soft_duplicates)
  set_check(37, "y")   # Enquetes au chemin le plus court
  set_check(38, "y")   # Valeurs aberrantes (check_outliers)
  if (has_logical_checks) set_check(39, "y")  # Verifications logiques (seulement une fois check_list_august.xlsx rempli de vraies verifications)
  set_check(40, "y")   # Variables "calculate" Kobo recalculees (prix EUR)
  set_check(41, "y")   # Performance des enqueteurs
  set_check(42, "y")   # Verifications specifiques au secteur
  set_check(50, "na")  # Coherence parent/individuel : la JMMI n'a pas de boucle repetee
  set_check(51, "na")  # Calcul des ponderations : pas de ponderation ce cycle

  # Verification du Nettoyage des Donnees
  set_check(46, "y")   # #uuid brutes = nettoyees + suppression (review_cleaning)
  set_check(47, "y")   # Noms de variables Kobo dans les journaux (colonne "variable")
  set_check(48, "y")   # Modifications consignees dans le journal de nettoyage
  set_check(49, "y")   # Anciennes/nouvelles valeurs identiques si aucun changement

  # Points supplementaires
  set_check(55, "y")   # colonne id (renommee depuis _id) conservee en plus de uuid
  set_check(56, "na")  # groupes repetes : non applicable, pas de boucle
  set_check(57, "y")   # nom de fichier conforme a la norme de nommage
  set_check(58, "y")   # pas d'instructions internes residuelles
  set_check(59, "y")   # les feuilles refletent le cycle actuel
  set_check(61, "y")   # titre correct (genere dynamiquement)
  set_check(62, "y")   # periode/couverture mises a jour (calculees depuis les donnees)
  set_check(63, "y")   # nombre final d'enquetes (calcule)
  set_check(64, "na")  # logos partenaires : exclus
  set_check(65, "y")   # apercu des feuilles a jour
}

# Construit la feuille READ_ME avec le meme langage visuel que la propre sortie d'analyse JMMI de
# REACH (onglet "Lisez_moi" de REACH_CAR_JMMI_Analyse_*.xlsx) : une banniere de titre gris fonce, une
# banniere de sous-titre gris clair, des libelles de section en gras en colonne A avec des descriptions
# avec retour a la ligne en colonne B, puis une barre d'en-tete rouge "Contenu" au-dessus d'un index
# des feuilles a deux colonnes.
write_read_me_sheet <- function(wb, assessment_month_label, assessment_year, collection_start,
                                collection_end, admin1_covered, admin2_covered, partners_covered,
                                final_survey_count, sheet_descriptions){

  addWorksheet(wb, "READ_ME")
  sheet <- "READ_ME"

  title <- paste0("RCA - Suivi Conjoint des Marches (JMMI) - Nettoyage des donnees")
  subtitle <- paste0("Collecte de donnees de ", assessment_month_label, " ", assessment_year)

  writeData(wb, sheet, title, startRow = 1, startCol = 1, colNames = FALSE)
  mergeCells(wb, sheet, cols = 1:2, rows = 1)
  addStyle(wb, sheet, rows = 1, cols = 1:2,
           style = createStyle(fgFill = DARK_GREY_BANNER, fontColour = "#FFFFFF", textDecoration = "bold",
                               fontSize = 16, halign = "center", valign = "center"))

  writeData(wb, sheet, subtitle, startRow = 2, startCol = 1, colNames = FALSE)
  mergeCells(wb, sheet, cols = 1:2, rows = 2)
  addStyle(wb, sheet, rows = 2, cols = 1:2,
           style = createStyle(fgFill = GREY_HEADER, fontColour = "#000000", textDecoration = "bold",
                               fontSize = 16, halign = "center", valign = "center"))

  # Le texte contexte/objectif/methodologie ci-dessous provient des termes de reference de la JMMI RCA
  # (REACH_TOR_CAR2501_JMMI_Interne.docx) et de la derniere sortie d'analyse JMMI RCA
  # (REACH_CAR_JMMI_Analyse_Avril-2026), il n'a pas ete invente pour ce README. La couverture
  # geographique et la liste des partenaires, elles, sont calculees depuis les donnees reellement
  # collectees ce cycle (admin1_covered/admin2_covered/partners_covered), pas codees en dur.
  sections <- list(
    c("Contexte du projet",
      "Suivi conjoint des marches (JMMI) en Republique centrafricaine (RCA), lance par REACH avec le Groupe de Travail Cash (GTCash/CWG) pour combler l'absence de suivi coordonne des prix et de la disponibilite des produits du Panier Minimum de Depenses (PMAS/MEB), dans un contexte de forte insecurite economique (HNRP 2025 : 38% de la population en vulnerabilite extreme, inflation alimentaire +10%, transport +17%)."),
    c("Objectif general",
      "A travers l'Initiative Conjointe de Suivi des Marches (ICSM) en RCA, l'objectif est de permettre aux acteurs humanitaires d'acceder aux informations sur les prix a la consommation, la disponibilite des produits essentiels, ainsi que les conditions des vendeurs dans les marches evalues. Ces informations visent a mieux informer la reponse humanitaire, notamment sous forme de transferts monetaires, et a soutenir les organisations actives ou planifiant des interventions en faveur des menages en situation de vulnerabilite en RCA."),
    c("Methodologie de la JMMI (collecte)",
      "Echantillonnage raisonne de marches centraux et satellites ; enquetes aupres d'informateurs cles (detaillants etablis) via Kobo. Pour chaque article du PMAS, 5 prix sont recueillis aupres de differents detaillants du meme marche afin d'obtenir au moins 3 prix exploitables apres nettoyage. La collecte a lieu chaque mois, simultanement chez tous les partenaires sur une fenetre d'environ 10 jours."),
    c("Periode de collecte", paste0(collection_start, " au ", collection_end)),
    c("Couverture geographique", paste0(
      "Prefectures couvertes : ", paste(sort(unique(admin1_covered)), collapse=", "),
      " | Sous-prefectures couvertes : ", paste(sort(unique(admin2_covered)), collapse=", "))),
    c("Nombre final d'enquetes", as.character(final_survey_count)),
    c("Partenaires participants", paste0(
      "Coordonne par REACH avec le Groupe de Travail Cash (GTCash/CWG) ; collecte de terrain assuree ce cycle par : ",
      paste(sort(unique(partners_covered)), collapse=", "), ".")),
    c("Processus de nettoyage des donnees", "Voir la feuille CHECKLIST_Data_Cleaning."),
    c("Contacts (Nom - poste - adresse e-mail)", "Marcelle Mande-Djapou - Chargee d'evaluation - marcelle.mande-djapou@reach-initiative.org")
  )

  r <- 3
  for (s in sections) {
    writeData(wb, sheet, s[1], startRow = r, startCol = 1, colNames = FALSE)
    writeData(wb, sheet, s[2], startRow = r, startCol = 2, colNames = FALSE)
    addStyle(wb, sheet, rows = r, cols = 1, style = createStyle(textDecoration = "bold", valign = "top", wrapText = TRUE))
    addStyle(wb, sheet, rows = r, cols = 2, style = createStyle(valign = "top", wrapText = TRUE))
    setRowHeights(wb, sheet, rows = r, heights = 45)
    r <- r + 1
  }

  r <- r + 1
  writeData(wb, sheet, "Contenu", startRow = r, startCol = 1, colNames = FALSE)
  mergeCells(wb, sheet, cols = 1:2, rows = r)
  addStyle(wb, sheet, rows = r, cols = 1:2,
           style = createStyle(fgFill = DARK_GREY_BANNER, fontColour = "#FFFFFF", textDecoration = "bold",
                               fontSize = 12, halign = "center", valign = "center"))
  r <- r + 1

  writeData(wb, sheet, "Onglets", startRow = r, startCol = 1, colNames = FALSE)
  writeData(wb, sheet, "Description", startRow = r, startCol = 2, colNames = FALSE)
  addStyle(wb, sheet, rows = r, cols = 1:2,
           style = createStyle(fgFill = RED_HEADER, fontColour = "#FFFFFF", textDecoration = "bold", fontSize = 14))
  r <- r + 1

  for (nm in names(sheet_descriptions)) {
    writeData(wb, sheet, nm, startRow = r, startCol = 1, colNames = FALSE)
    writeData(wb, sheet, sheet_descriptions[[nm]], startRow = r, startCol = 2, colNames = FALSE)
    addStyle(wb, sheet, rows = r, cols = 1, style = createStyle(textDecoration = "bold", valign = "top"))
    addStyle(wb, sheet, rows = r, cols = 2, style = createStyle(valign = "top", wrapText = TRUE))
    setRowHeights(wb, sheet, rows = r, heights = 30)
    r <- r + 1
  }

  setColWidths(wb, sheet, cols = 1:2, widths = c(38, 100))
  freezePane(wb, sheet, firstActiveRow = 3)
}

# Assemble le classeur final combine en suivant l'ordre des feuilles et les colonnes du modele de la
# checklist : CHECKLIST_Data_Cleaning (copiee depuis le modele, auto-evaluation remplie la ou le script
# peut la certifier), READ_ME, survey, choices, variable_tracker, raw_data, clean_data, cleaning_log,
# deletion_log.
build_final_workbook <- function(raw_data_export, clean_data_export, cleaning_log_final, deletion_log_final,
                                 variable_tracker_final, output_path, final_survey_count,
                                 collection_start, collection_end, admin1_covered, admin2_covered, partners_covered,
                                 assessment_month_label, assessment_year, has_logical_checks){

  wb <- loadWorkbook(Sys.getenv("checklist_template"))
  for (sh in c("READ_ME","survey","choices","variable_tracker","raw_data","clean_data",
               "cleaning_log","deletion_log","sampling_info")) {
    if (sh %in% names(wb)) removeWorksheet(wb, sh)
  }
  freezePane(wb, "CHECKLIST_Data_Cleaning", firstActiveRow = 2)

  # deletion_log est toujours incluse, meme vide (0 suppression) - plus simple et plus previsible pour
  # le reviseur qu'une feuille qui apparait ou disparait d'un cycle a l'autre.
  has_deletions <- nrow(deletion_log_final) > 0
  fill_checklist_data_cleaning(wb, has_logical_checks = has_logical_checks)

  sheet_descriptions <- list(
    "CHECKLIST_Data_Cleaning" = "Auto-evaluation des normes minimales de nettoyage - remplie automatiquement pour les verifications que le script peut certifier, le reste a completer par le reviseur.",
    "READ_ME" = "Cette feuille.",
    "survey" = "Feuille d'enquete (questionnaire) Kobo.",
    "choices" = "Feuille de choix (options de reponse) Kobo.",
    "variable_tracker" = "Variables Kobo omises ou ajoutees lors du nettoyage, avec justification.",
    "raw_data" = "Donnees brutes, anonymisees, 1 ligne = 1 enquete.",
    "clean_data" = "Donnees nettoyees finales, anonymisees, 1 ligne = 1 enquete.",
    "cleaning_log" = "Journal de nettoyage - 1 ligne par uuid + variable modifiee.",
    "deletion_log" = if (has_deletions) {
      "Journal de suppression des enquetes retirees des donnees nettoyees."
    } else {
      "Journal de suppression - aucune enquete retiree ce cycle (feuille vide, conservee pour la structure standard)."
    }
  )

  write_read_me_sheet(wb, assessment_month_label = assessment_month_label, assessment_year = assessment_year,
                     collection_start = collection_start, collection_end = collection_end,
                     admin1_covered = admin1_covered, admin2_covered = admin2_covered,
                     partners_covered = partners_covered, final_survey_count = final_survey_count,
                     sheet_descriptions = sheet_descriptions)

  write_simple_sheet(wb, "variable_tracker", variable_tracker_final)

  write_labeled_data_sheet(wb, "raw_data", apply_var_labels(raw_data_export))
  write_labeled_data_sheet(wb, "clean_data", apply_var_labels(clean_data_export))

  write_simple_sheet(wb, "cleaning_log", cleaning_log_final)
  write_simple_sheet(wb, "deletion_log", deletion_log_final)

  desired_order <- intersect(c("CHECKLIST_Data_Cleaning","READ_ME","survey","choices","variable_tracker",
                               "raw_data","clean_data","cleaning_log","deletion_log"), names(wb))
  worksheetOrder(wb) <- match(desired_order, names(wb))

  # Ecrire ce classeur multi-feuilles d'environ 700 colonnes directement dans output/5_final_donnees_nettoyees/ (a
  # l'interieur du dossier de projet synchronise par OneDrive) produit par intermittence un zip
  # corrompu - le moteur de synchronisation de OneDrive semble verrouiller le chemin de destination au
  # moment ou il est (re)cree. Sauvegarder d'abord dans un fichier temporaire local, puis copier le
  # fichier deja complet a sa place avec quelques tentatives et une verification d'integrite du zip,
  # evite d'entrer en concurrence avec OneDrive pendant l'ecriture elle-meme.
  # Utiliser un dossier de mise en attente local stable (hors OneDrive) plutot que tempdir(), pour
  # qu'une copie valide survive toujours entre les sessions R si toutes les tentatives ci-dessous
  # echouent quand meme.
  stage_dir <- file.path(Sys.getenv("LOCALAPPDATA"), "jmmi_car_build")
  if (!dir.exists(stage_dir)) dir.create(stage_dir, recursive = TRUE)
  tmp_path <- file.path(stage_dir, basename(output_path))
  saveWorkbook(wb, tmp_path, overwrite = TRUE)

  # "survey"/"choices" ne sont PAS reconstruites a partir des data frames kobo_survey/kobo_choices
  # (ce qui perdrait toute mise en forme d'origine, readxl ne recuperant que les valeurs) : elles sont
  # clonees telles quelles depuis l'outil kobo d'entree (openxlsx ne sait pas copier une feuille avec sa
  # mise en forme d'un classeur vers un autre - openxlsx2 le permet via clone_worksheet(from=)). Seul le
  # volet fige sur la ligne d'en-tete est explicitement reapplique (perdu par ce clonage).
  wb2 <- openxlsx2::wb_load(tmp_path)
  wb2_kobo <- openxlsx2::wb_load(Sys.getenv("kobo"))
  for (sh in c("survey", "choices")) {
    wb2$clone_worksheet(old = sh, new = sh, from = wb2_kobo)
    wb2$freeze_pane(sh, first_row = TRUE)
  }
  desired_order_full <- intersect(c("CHECKLIST_Data_Cleaning","READ_ME","survey","choices","variable_tracker",
                                    "raw_data","clean_data","cleaning_log","deletion_log"), wb2$get_sheet_names())
  wb2$set_order(match(desired_order_full, wb2$get_sheet_names()))
  openxlsx2::wb_save(wb2, tmp_path, overwrite = TRUE)

  is_valid_zip <- function(path) {
    if (!file.exists(path)) return(FALSE)
    tryCatch({ utils::unzip(path, list = TRUE); TRUE }, error = function(e) FALSE, warning = function(w) FALSE)
  }

  # Copier une fois, puis simplement verifier la validite periodiquement (plutot que de recopier a
  # chaque tentative) : recopier sans cesse la destination semble redeclencher l'analyse
  # OneDrive/antivirus du fichier et ne jamais laisser le verrou se liberer, alors que laisser le
  # fichier tranquille apres une seule copie permet a cette analyse de se terminer naturellement.
  copy_ok <- FALSE
  suppressWarnings(try(file.copy(tmp_path, output_path, overwrite = TRUE), silent = TRUE))
  for (attempt in 1:15) {
    if (is_valid_zip(output_path)) { copy_ok <- TRUE; break }
    Sys.sleep(2)
  }
  if (!copy_ok) {
    suppressWarnings(try(file.copy(tmp_path, output_path, overwrite = TRUE), silent = TRUE))
    for (attempt in 1:15) {
      if (is_valid_zip(output_path)) { copy_ok <- TRUE; break }
      Sys.sleep(2)
    }
  }

  if (copy_ok) {
    file.remove(tmp_path)
  } else {
    warning("Impossible d'ecrire un classeur valide dans ", output_path, " (OneDrive semble le verrouiller). ",
           "Une copie valide a ete enregistree dans ", tmp_path, " - a copier manuellement une fois que OneDrive libere le verrou.")
  }
}
