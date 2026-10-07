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

## as.Date()/as.POSIXct() sur une chaine qui RESSEMBLE a un nombre (ex. "46239.488067129627", le
## numero de serie Excel d'une date exportee par Kobo en XLS - readxl col_types="text" renvoie ce
## numero brut plutot qu'une date formatee pour toute colonne mise en forme "date" dans le fichier
## Excel source) ne renvoie PAS juste NA + warning comme pour une chaine illisible ordinaire : elles
## LEVENT UNE ERREUR ("character string is not in a standard unambiguous format"), que
## suppressWarnings() ne peut pas intercepter puisque ce n'est pas un warning. Cette erreur casse tout
## le mutate() vectorise (une seule valeur en cause suffit a faire echouer toute la colonne, ligne par
## ligne). tryCatch() ici traite ce cas comme un echec total du parsing "normal", pour que TOUTES les
## valeurs de la colonne passent alors par le repli "numero de serie Excel" (openxlsx::convertToDate/
## convertToDateTime) - le cas reel observe sur un export Kobo XLS telecharge directement (vu sur
## _submission_time/start/end/today/date).
parse_kobo_date <- function(x) {
  out <- tryCatch(suppressWarnings(as.Date(x)), error = function(e) rep(as.Date(NA), length(x)))
  needs_fallback <- is.na(out) & !is.na(x) & x != ""
  if (any(needs_fallback)) {
    out[needs_fallback] <- suppressWarnings(openxlsx::convertToDate(as.numeric(x[needs_fallback])))
  }
  as.character(out)
}

parse_kobo_datetime <- function(x) {
  out <- tryCatch(suppressWarnings(as.POSIXct(x, tz = "UTC")), error = function(e) rep(as.POSIXct(NA), length(x)))
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
# RECALCUL DES CHAMPS "CALCULATE" KOBO LIES AU PRIX (<item>_price_item)
#############################################################################################################
# create_clean_data() applique les corrections du journal de nettoyage champ par champ (ex.
# <item>_price_unit_item, <item>_std_unit_item, <item>_other_unit_item - voir "A verifier" / 2eme partie
# de 2_nettoyage_final.R, ou ces variables sont desormais affichees ensemble comme un bloc modifiable
# pour chaque article), mais ne re-execute PAS les formules "calculate" de l'outil kobo : la valeur de
# <item>_price_item resterait donc celle calculee par Kobo au moment de la collecte, jamais mise a jour
# si l'une de ses variables source est corrigee. recompute_price_items() (appelee depuis
# 2_nettoyage_final.R juste apres l'application du journal de nettoyage) resout ce probleme en
# reevaluant directement les formules kobo elles-memes - dynamiquement, depuis kobo_survey$calculation,
# plutot qu'une liste d'articles codee en dur, pour qu'un nouvel article ajoute a l'outil soit pris en
# compte automatiquement.

# Traduit une formule "calculate" XLSForm (syntaxe de l'outil kobo, ex. "round(if(${x}='non',
# (${y} div ${z}),${y}),0)") en expression R evaluable : ${var} -> reference de colonne (df[["var"]]),
# div -> /, = (comparaison XLSForm, jamais une assignation dans ce contexte) -> ==, 'texte' -> "texte",
# if(...) -> ifelse(...). Couvre exactement les constructions utilisees par les formules _price_item/
# _price_unit_item de cet outil, pas un analyseur XLSForm general.
translate_xlsform_formula <- function(formula) {
  f <- formula
  f <- str_replace_all(f, "\\$\\{(\\w+)\\}", 'df[["\\1"]]')
  f <- str_replace_all(f, "\\bdiv\\b", "/")
  f <- str_replace_all(f, "(?<![=<>!])=(?!=)", "==")
  f <- str_replace_all(f, "'", '"')
  f <- str_replace_all(f, "\\bif\\(", "ifelse(")
  f
}

# Reevalue un seul champ "calculate" (ex. <item>_price_item, ou son eventuel intermediaire
# <item>_<taille>_price_unit_item) a partir de sa formule kobo et des donnees ACTUELLES de df (qui
# peuvent avoir ete corrigees par le journal de nettoyage). Renvoie NA partout si une variable
# referencee par la formule est absente de df (garde-fou, plutot que de faire planter tout le script).
recompute_calculate_field <- function(df, formula) {
  if (is.na(formula) || !nzchar(formula)) return(rep(NA_real_, nrow(df)))
  vars_used <- str_remove_all(unlist(str_extract_all(formula, "\\$\\{\\w+\\}")), "[${}]")
  if (length(setdiff(vars_used, colnames(df))) > 0) return(rep(NA_real_, nrow(df)))
  expr <- translate_xlsform_formula(formula)
  as.numeric(eval(parse(text = expr), envir = list(df = df)))
}

# Recalcule TOUS les <item>_price_item (et leurs eventuels intermediaires <item>_<taille>_
# price_unit_item, ex. soap_200g_price_unit_item) directement depuis les variables source ACTUELLES de
# df. 2 passes (les intermediaires d'abord, puis les _price_item qui peuvent les referencer) plutot
# qu'une resolution generale de dependances - aucun intermediaire observe dans l'outil ne depend d'un
# AUTRE intermediaire, une chaine a 2 niveaux suffit pour la structure actuelle.
empty_price_recalc_log <- function() {
  tibble(uuid = character(0), question = character(0), old_value = character(0),
        change_type = character(0), new_value = character(0), issue = character(0),
        comment = character(0))
}

# journalise une valeur recalculee par rapport a sa valeur precedente dans `df` (avant ecrasement) -
# meme logique de comparaison que le recalcul des prix en euros (2_nettoyage_final.R, RECALCUL DES PRIX
# EN EUROS) : NA != NA doit valoir FALSE (sinon chaque champ jamais renseigne serait signale comme
# "change"), une seule ligne par uuid dont la valeur a reellement change.
log_recalculated_field <- function(df, fname, new_vals) {
  old_vals_chr <- as.character(df[[fname]])
  new_vals_chr <- as.character(new_vals)
  changed <- (is.na(old_vals_chr) != is.na(new_vals_chr)) |
    (!is.na(old_vals_chr) & !is.na(new_vals_chr) & old_vals_chr != new_vals_chr)
  if (!any(changed)) return(empty_price_recalc_log())
  tibble(uuid = df$uuid[changed], question = fname, old_value = old_vals_chr[changed],
        change_type = "change_response", new_value = new_vals_chr[changed],
        issue = paste0("Valeur recalculee automatiquement (champ \"calculate\" de l'outil kobo) suite a ",
                       "une modification d'une des variables source du bloc prix"),
        comment = NA_character_)
}

# Noms des champs "calculate" (<item>_price_item + leurs eventuels intermediaires <item>_<taille>_
# price_unit_item) qui seront de toute facon ECRASES par recompute_price_items(), quelle que soit la
# valeur presente dans clean_data au moment ou elle tourne - utilisee a la fois par
# recompute_price_items() (ci-dessous) et par 2_nettoyage_final.R pour exclure ces champs de
# create_clean_data() (une correction saisie dessus dans Verification_prix n'a de toute facon aucun
# effet final, et l'appliquer quand meme produirait une entree de journal en double avec le recalcul
# automatique - voir price_recalc_log). Factorisee ici pour que les deux endroits identifient
# exactement le meme ensemble de champs, sans jamais diverger.
get_calculated_price_field_names <- function(kobo_survey) {
  price_item_fields <- kobo_survey %>% filter(type == "calculate", str_ends(name, "_price_item"))
  if (nrow(price_item_fields) == 0) return(character(0))
  referenced_vars <- str_remove_all(unlist(str_extract_all(price_item_fields$calculation, "\\$\\{\\w+\\}")), "[${}]")
  intermediate_fields <- kobo_survey %>%
    filter(type == "calculate", name %in% referenced_vars, str_ends(name, "_price_unit_item"))
  unique(c(price_item_fields$name, intermediate_fields$name))
}

# Renvoie list(df=..., log=...) plutot que seulement df (voir ci-dessus) : le recalcul des champs
# "calculate" ecrase directement clean_data sans jamais passer par create_clean_data()/le journal de
# nettoyage habituel - sans ce journal separe, une correction sur une variable source (ex.
# water_price_unit_item) qui fait changer <item>_price_item n'apparaissait nulle part dans le classeur
# final, alors que la valeur avait bel et bien change (bug constate : "Entry missing in cleaning log"
# signale par review_cleaning() a l'execution). Fusionne dans combined_cleaning_log par l'appelant (voir
# 2_nettoyage_final.R) - qui exclut par ailleurs ces memes champs de create_clean_data() (voir
# get_calculated_price_field_names() ci-dessus), donc `df` arrive ici SANS correction manuelle
# deja appliquee sur ces champs : old_value ci-dessous est donc toujours la vraie valeur brute/precedente,
# jamais une valeur intermediaire tapee par l'equipe de suivi et vouee a etre ecrasee.
recompute_price_items <- function(df, kobo_survey) {
  price_item_fields <- kobo_survey %>% filter(type == "calculate", str_ends(name, "_price_item"))
  if (nrow(price_item_fields) == 0) return(list(df = df, log = empty_price_recalc_log()))

  referenced_vars <- str_remove_all(unlist(str_extract_all(price_item_fields$calculation, "\\$\\{\\w+\\}")), "[${}]")
  intermediate_fields <- kobo_survey %>%
    filter(type == "calculate", name %in% referenced_vars, str_ends(name, "_price_unit_item"))
  intermediate_referenced <- str_remove_all(unlist(str_extract_all(intermediate_fields$calculation, "\\$\\{\\w+\\}")), "[${}]")

  # coercion numerique de toute variable REFERENCEE par une formule (sauf les select_one, type
  # std_unit_item, comparees a une chaine 'oui'/'non' - jamais utilisees dans un calcul arithmetique)
  select_one_vars <- kobo_survey$name[str_starts(kobo_survey$type, "select_one")]
  numeric_vars <- setdiff(unique(c(referenced_vars, intermediate_referenced, price_item_fields$name)), select_one_vars)
  for (v in intersect(numeric_vars, colnames(df))) df[[v]] <- suppressWarnings(as.numeric(df[[v]]))

  recalc_log <- empty_price_recalc_log()
  for (i in seq_len(nrow(intermediate_fields))) {
    fname <- intermediate_fields$name[i]
    if (fname %in% colnames(df)) {
      # arrondi a l'entier le plus proche (demande utilisateur) - certaines formules kobo (ex. "egg")
      # n'appellent pas round() elles-memes, produisant des decimales sur ~16 chiffres qui n'ont aucun
      # sens pour un prix/une quantite en CFA ; applique ici plutot que de dependre de chaque formule
      # kobo pour rester correct meme sur un article dont la formule oublierait round().
      new_vals <- round(recompute_calculate_field(df, intermediate_fields$calculation[i]))
      recalc_log <- bind_rows(recalc_log, log_recalculated_field(df, fname, new_vals))
      df[[fname]] <- new_vals
    }
  }
  for (i in seq_len(nrow(price_item_fields))) {
    fname <- price_item_fields$name[i]
    if (fname %in% colnames(df)) {
      new_vals <- round(recompute_calculate_field(df, price_item_fields$calculation[i]))
      recalc_log <- bind_rows(recalc_log, log_recalculated_field(df, fname, new_vals))
      df[[fname]] <- new_vals
    }
  }
  list(df = df, log = recalc_log)
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

  # une question select_one ne peut avoir qu'UNE seule valeur : si l'equipe de suivi a rempli plusieurs
  # colonnes EXISTANT autre par erreur (x$existing_other est alors plusieurs libelles joints par ";",
  # voir 2_nettoyage_final.R), on ne bloque pas tout le traitement pour autant - seul le PREMIER choix
  # est retenu (VRAI autre reste sans effet ici, il n'est jamais utilise par cette fonction) et les
  # autres sont juste signales dans un avertissement, pas une erreur.
  existing_parts <- str_split(x$existing_other, ";")[[1]]
  if (length(existing_parts) > 1) {
    warning(paste0("Cette variable est une question a choix UNIQUE (select_one) mais plusieurs choix ",
                   "EXISTANT autre ont ete selectionnes en meme temps (classeur suivi a verifier, feuille ",
                   "reponses_autres) - uuid=", x$uuid, ", variable=", x$ref_name,
                   " : seul le premier est retenu ('", trimws(existing_parts[1]),
                   "'), ignore(s) : '", paste(trimws(existing_parts[-1]), collapse = "', '"), "'."))
  }
  existing_choice <- trimws(existing_parts[1])

  # retirer le texte de la reponse
  df <- data.frame(uuid=x$uuid, question=x$name, change_type=change_type,
                   old_value=x$response, new_value=NA)
  cleaning_log_other <<- rbind(cleaning_log_other, df)

  # recoder le choix (uniquement le premier choix retenu, voir ci-dessus)
  new_value <- get_name_from_label(x$list_name, existing_choice)
  if (length(new_value)!=1){
    stop(paste0("Le choix '", existing_choice, "' n'existe pas dans la liste de reponses '", x$list_name,
               "' : uuid=", x$uuid, ", variable=", x$ref_name, ", question 'autre'=", x$name,
               " (classeur suivi a verifier, feuille reponses_autres) - verifiez l'orthographe exacte ",
               "du libelle choisi."))
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
# LISTES DEROULANTES (post-traitement openxlsx2 - voir apply_classic_dropdowns() ci-dessous)
################################################################################
# openxlsx::dataValidation() ecrit TOUJOURS les validations de type "list" au format d'extension
# Excel 2010+ ("x14", range dans <extLst>) des qu'une plage de cellules est referencee - meme pour une
# reference sur la MEME feuille, et meme pour une liste litterale en ligne (verifie empiriquement :
# aucune des 3 formes n'obtient le <dataValidation> classique avec cette version d'openxlsx). Google
# Sheets et certaines versions de LibreOffice ignorent silencieusement ce format x14 : le fichier
# s'ouvre normalement mais aucune liste deroulante n'apparait, alors que le code semble correct et
# qu'Excel 2010+ les affiche sans probleme. openxlsx2, en revanche, ecrit le format <dataValidation>
# classique (verifie empiriquement, y compris pour des references inter-feuilles et des plages non
# contigues) - pris en charge par tous les tableurs. Plutot que de reconstruire toute la feuille avec
# openxlsx2 (API de mise en forme totalement differente, risque de perdre la protection/le figeage/les
# styles qu'openxlsx sait deja faire correctement), les listes deroulantes sont enregistrees ici (au
# lieu d'appeler openxlsx::dataValidation() directement) puis ajoutees en un seul post-traitement,
# apres la sauvegarde complete du classeur par openxlsx - voir apply_classic_dropdowns(), appelee apres
# saveWorkbook() dans 2_nettoyage_final.R.
PENDING_DROPDOWN_VALIDATIONS <- list()

# `rows` n'est pas necessairement contigue (ex. write_review_sheet() : les lignes signalees pour une
# meme liste de choix peuvent etre dispersees dans le classeur, pas triees par variable) - mais quand
# elles LE SONT (cas courant de add_other_responses_sheet() : df_other est trie par question_name, donc
# toutes les lignes d'une meme question se suivent), les regrouper en une seule plage "colX:colY" au
# lieu d'une reference "colX:colX" par ligne individuelle jointes par des virgules - un sqref avec des
# dizaines de plages a une seule cellule est inutilement long/fragile (risque reel qu'Excel ignore
# silencieusement la validation plutot que de lever une erreur claire), pour un resultat strictement
# equivalent.
queue_list_validation <- function(sheet_name, col, rows, list_range) {
  if (length(rows) == 0) return(invisible(NULL))
  col_letter <- openxlsx::int2col(col)
  rows_sorted <- sort(unique(rows))
  run_id <- cumsum(c(TRUE, diff(rows_sorted) != 1))
  run_start <- tapply(rows_sorted, run_id, min)
  run_end <- tapply(rows_sorted, run_id, max)
  # attribut "sqref" de type ST_Sqref (xsd:list) dans le schema OOXML : plusieurs plages doivent etre
  # separees par un ESPACE, jamais par une virgule (regle generale xsd:list, pas specifique a Excel) -
  # confirme via openpyxl (parseur independant d'openxlsx/openxlsx2, qui rejette explicitement un sqref
  # joint par des virgules alors qu'un XML generique comme xml2::read_xml()/openxlsx2::wb_load() ne
  # detecte rien, puisque c'est syntaxiquement du XML bien forme, juste invalide vis-a-vis du schema).
  # C'etait la cause reelle des classeurs "a reparer" sous Excel : le seul endroit de ce script ou
  # plusieurs plages A UNE SEULE CELLULE (donc necessairement jointes, jamais fusionnables en une plage
  # colX:colY) sont systematiquement generees est la liste deroulante "_std_unit_item" (oui/non) de
  # write_review_sheet() pour "Verification_prix" - une ligne "contexte" dispersee par bloc de prix, PAS
  # triee/regroupee comme add_other_responses_sheet() peut l'etre - d'ou un classeur systematiquement
  # invalide pour cet onglet specifiquement, jamais pour "A_verifier".
  dims <- paste0(col_letter, run_start, ":", col_letter, run_end, collapse = " ")
  PENDING_DROPDOWN_VALIDATIONS[[length(PENDING_DROPDOWN_VALIDATIONS) + 1]] <<- list(
    sheet = sheet_name, dims = dims, value = list_range
  )
  invisible(NULL)
}

# PENDING_COMMENTS (voir queue_comment() ci-dessous) suit exactement le meme principe que
# PENDING_DROPDOWN_VALIDATIONS : les notes Excel (write_review_sheet(), lignes is_calculated_field) NE
# SONT PAS ecrites directement via openxlsx::writeComment() - constate empiriquement que apply_classic_
# dropdowns() (qui recharge/re-sauvegarde deja le classeur avec openxlsx2 pour corriger les listes
# deroulantes, voir plus haut) fait alors disparaitre silencieusement les commentaires ecrits par
# openxlsx (avertissement "file expected to be in output is missing: /xl/commentsN.xml" a la
# sauvegarde openxlsx2 - les deux bibliotheques ne serialisent pas les commentaires de la meme facon,
# openxlsx2 ne retrouve pas la partie qu'openxlsx a ecrite et la perd au re-enregistrement). Les notes
# sont donc, comme les listes deroulantes, mises en attente ici et ajoutees par openxlsx2 lui-meme,
# dans la MEME session de rechargement/sauvegarde - jamais ecrites par openxlsx en premier lieu, donc
# rien a perdre au passage par openxlsx2.
PENDING_COMMENTS <- list()

queue_comment <- function(sheet_name, col, row, text) {
  col_letter <- openxlsx::int2col(col)
  PENDING_COMMENTS[[length(PENDING_COMMENTS) + 1]] <<- list(
    sheet = sheet_name, dims = paste0(col_letter, row), text = text
  )
  invisible(NULL)
}

# A appeler UNE FOIS, apres saveWorkbook() (openxlsx) - recharge le fichier deja sauvegarde avec
# openxlsx2 (sans toucher a la mise en forme/protection deja ecrite par openxlsx, voir plus haut),
# ajoute toutes les listes deroulantes enregistrees via queue_list_validation() au format classique
# ainsi que les notes enregistrees via queue_comment(), et re-sauvegarde au meme chemin.
apply_classic_dropdowns <- function(path) {
  if (length(PENDING_DROPDOWN_VALIDATIONS) == 0 && length(PENDING_COMMENTS) == 0) return(invisible(NULL))
  wb2 <- openxlsx2::wb_load(path)
  for (spec in PENDING_DROPDOWN_VALIDATIONS) {
    wb2 <- wb2$add_data_validation(sheet = spec$sheet, dims = spec$dims, type = "list", value = spec$value)
  }
  for (spec in PENDING_COMMENTS) {
    wb2 <- wb2$add_comment(sheet = spec$sheet, dims = spec$dims,
                           comment = openxlsx2::wb_comment(text = spec$text, author = "Pipeline JMMI", visible = FALSE))
  }
  wb2$save(path)
  invisible(NULL)
}

################################################################################
# En-tetes affiches (colonne A) <- nom de colonne technique - memes colonnes de contexte (commune,
# partenaire, date, enqueteur, boutique) que REVIEW_HEADERS_FR pour "3_A_verifier"/"4_Verification_prix"
# (demande utilisateur, pour les memes reperes d'une feuille a l'autre). DEUX colonnes distinctes pour la
# nouvelle reponse plutot qu'une seule (demande explicite de l'utilisateur) : "texte libre" (aucune
# validation, pour "Reecrire" - accepte aussi plusieurs choix separes par ";" si "Reattribuer" porte sur
# une question a choix multiples, la liste deroulante ne pouvant selectionner qu'un seul choix a la
# fois) et "choix existant" (liste deroulante, pour "Reattribuer" a un seul choix - le cas courant).
# Excel ne permet pas de verrouiller/deverrouiller une cellule DYNAMIQUEMENT selon la valeur d'une autre
# cellule sans macro VBA (ecartees ici, comme partout ailleurs dans ce classeur, pour eviter les
# avertissements de securite/l'incompatibilite avec Google Sheets/LibreOffice) - les deux colonnes
# restent donc techniquement modifiables en permanence, mais une mise en forme conditionnelle (voir
# add_other_responses_sheet()) grise celle qui ne correspond pas a l'action choisie, pour qu'il reste
# sans ambiguite laquelle remplir.
# row_id/statut : voir carry_forward_decisions() plus bas - meme principe que dans REVIEW_HEADERS_FR
# (identifiant stable d'une ligne + indicateur "deja traite au tour precedent" ou "nouveau"), colonnes
# techniques (ne pas modifier).
OTHER_HEADERS_FR <- c(
  uuid = "Identifiant enquete (ne pas modifier)",
  row_id = "ID (ne pas modifier)",
  commune = "Commune",
  partenaire = "Partenaire",
  date = "Date",
  enqueteur = "Enqueteur",
  boutique = "Boutique / vendeur",
  q_type = "Type de question (ne pas modifier)",
  list_name = "Liste de choix (ne pas modifier)",
  question_name = "Variable (ne pas modifier)",
  full_label = "Question (ne pas modifier)",
  selected_choices = "Choix deja selectionnes (ne pas modifier)",
  response = "Reponse \"autre\" actuelle (ne pas modifier)",
  statut = "Statut (ne pas modifier)",
  action = "Que faire ? (obligatoire)",
  new_value_text = "Nouvelle reponse - texte libre (si Reecrire ; ou Reattribuer a plusieurs choix, separes par ;)",
  new_value_choice = "Nouvelle reponse - choix existant (si Reattribuer a un seul choix)"
)

# Options de la liste deroulante "Que faire ?" (feuille reponses_autres) <- valeur technique interne -
# meme principe que ACTION_LABELS_FR pour "A_verifier". "Reattribuer..." remplace l'ancien systeme a 3
# colonnes "EXISTANT autre 1/2/3" (un seul champ "Nouvelle reponse" suffit desormais) ; "Supprimer"
# remplace "INVALIDE autre". "Ne rien changer" (comme no_action dans ACTION_LABELS_FR) : une reponse
# "autre" flaguee n'a pas forcement besoin d'etre recodee/reecrite/supprimee - elle peut deja etre
# valable telle quelle. Aucun traitement PASSE 2 special requis : les lignes dont action_code vaut
# "no_action" (ou est encore vide) ne correspondent a aucun des filtres or_remove/or_reassign/or_rewrite
# dans 2_nettoyage_final.R, donc ne produisent naturellement aucune entree de cleaning_log - la seule
# difference est que l'option est desormais EXPLICITE (choix delibere plutot qu'une ligne juste oubliee).
OTHER_ACTION_LABELS_FR <- c(
  no_action = "Ne rien changer",
  rewrite = "Reecrire",
  reassign = "Reattribuer a un choix existant",
  remove = "Supprimer"
)

# Ajoute la feuille des reponses "autre" (+ sa feuille "Listes_..." de validation) a un classeur
# existant - meme principe que write_review_sheet()/A_verifier : feuille consolidee en langage simple,
# une seule liste deroulante "Que faire ?", colonnes techniques grisees verrouillees.
add_other_responses_sheet <- function(wb, df, sheet_name = "5_Reponses_autres") {
  # capture AVANT le select()/rename ci-dessous (list_name n'est pas forcement affiche sous ce nom une
  # fois passe dans df_display) - reutilise plus bas pour regrouper les lignes par liste de choix kobo,
  # meme principe que block_id dans write_review_sheet().
  list_name_raw <- df$list_name

  df_display <- df %>% select(names(OTHER_HEADERS_FR))
  colnames(df_display) <- unname(OTHER_HEADERS_FR[colnames(df_display)])
  df_display <- sanitize_df_text(df_display)

  addWorksheet(wb, sheet_name)
  writeData(wb, sheet_name, df_display, colNames = TRUE)

  # n vaut 1 (donc data_rows une plage vide) quand df a 0 ligne - cas reel (aucune reponse "autre" a
  # recoder), pas seulement theorique. Sans ce garde-fou, 2:1 (sequence DESCENDANTE) retoucherait aussi
  # la ligne d'en-tete.
  n <- nrow(df_display) + 1
  data_rows <- if (n >= 2) 2:n else integer(0)
  col_of_display <- function(key) match(OTHER_HEADERS_FR[[key]], colnames(df_display))
  action_col <- col_of_display("action")
  new_value_text_col <- col_of_display("new_value_text")
  new_value_choice_col <- col_of_display("new_value_choice")
  editable_cols <- c(action_col, new_value_text_col, new_value_choice_col)
  technical_cols <- setdiff(seq_len(ncol(df_display)), editable_cols)

  addStyle(wb, sheet_name, style = createStyle(fgFill = RED_HEADER, fontColour = "#FFFFFF",
                                               textDecoration = "bold", wrapText = TRUE, valign = "center",
                                               halign = "center"), rows = 1, cols = 1:ncol(df_display))
  freezePane(wb, sheet_name, firstActiveRow = 2, firstActiveCol = 2)
  addFilter(wb, sheet_name, rows = 1, cols = 1:ncol(df_display))

  addStyle(wb, sheet_name, style = createStyle(wrapText = TRUE, valign = "top", locked = FALSE),
          rows = data_rows, cols = editable_cols, gridExpand = TRUE)
  addStyle(wb, sheet_name, style = createStyle(fgFill = GREY_HEADER, fontColour = "#595959", wrapText = TRUE,
                                               valign = "top", locked = TRUE),
          rows = data_rows, cols = technical_cols, gridExpand = TRUE)

  # liste deroulante "Que faire ?" - liste en ligne (4 options, largement sous la limite de 255
  # caracteres d'Excel pour une liste litterale).
  action_choices <- paste0('"', paste(unname(OTHER_ACTION_LABELS_FR), collapse = ","), '"')
  queue_list_validation(sheet_name, col = action_col, rows = data_rows, list_range = action_choices)

  # liste deroulante pour "Nouvelle reponse - choix existant" : chaque ligne recoit la liste des choix
  # VALIDES DE SA PROPRE question - meme principe que "Nouvelle reponse" dans write_review_sheet()/
  # A_verifier (feuille cachee "Listes_...", une colonne par liste de choix kobo, une plage par liste
  # jointe en une seule liste deroulante par regroupement de lignes contigues), mais sans avoir besoin de
  # get_variable_choice_list() : list_name est deja une colonne technique de other_db (voir
  # list_name_raw ci-dessus). "Nouvelle reponse - texte libre" (aucune validation) reste le champ a
  # utiliser pour "Reecrire", ou pour "Reattribuer" une question a choix multiples a PLUSIEURS choix a la
  # fois (la liste deroulante n'en insere qu'un seul, limitation Excel) - voir OTHER_HEADERS_FR plus haut.
  choice_lists_sheet <- substr(paste0("Listes_", sheet_name), 1, 31)
  addWorksheet(wb, choice_lists_sheet, visible = FALSE)

  if (length(data_rows) > 0) {
    unique_lists <- unique(list_name_raw[!is.na(list_name_raw) & nzchar(list_name_raw)])
    for (li in seq_along(unique_lists)) {
      list_name <- unique_lists[li]
      choice_labels <- kobo_choices$label[kobo_choices$list_name == list_name]
      choice_labels <- choice_labels[!is.na(choice_labels) & nzchar(choice_labels)]
      if (length(choice_labels) == 0) next

      writeData(wb, choice_lists_sheet, choice_labels, startCol = li)
      col_letter <- openxlsx::int2col(li)
      list_values <- paste0("'", choice_lists_sheet, "'!$", col_letter, "$1:$", col_letter,
                            "$", length(choice_labels))

      target_rows <- which(list_name_raw == list_name) + 1  # +1 pour la ligne d'en-tete
      queue_list_validation(sheet_name, col = new_value_choice_col, rows = target_rows,
                            list_range = list_values)
    }
  }

  # Excel ne sait pas verrouiller/deverrouiller une cellule DYNAMIQUEMENT selon la valeur d'une autre
  # sans macro VBA (voir commentaire sur OTHER_HEADERS_FR plus haut) : mise en forme conditionnelle a la
  # place, qui grise (fond + police) la colonne "texte libre" tant que l'action n'est PAS "Reecrire", et
  # la colonne "choix existant" tant que l'action n'est PAS "Reattribuer..." - y compris quand "Que
  # faire ?" est encore vide (les deux colonnes apparaissent grisees par defaut avant tout choix). Dans
  # les deux cas la cellule reste techniquement modifiable : un dxf (mise en forme conditionnelle OOXML)
  # ne peut de toute facon porter aucun attribut de protection/verrouillage, ce n'est donc qu'un repere
  # visuel de laquelle remplir, jamais un veritable verrou.
  if (length(data_rows) > 0) {
    action_col_letter <- openxlsx::int2col(action_col)
    anchor_row <- data_rows[1]
    rewrite_label <- unname(OTHER_ACTION_LABELS_FR[["rewrite"]])
    reassign_label <- unname(OTHER_ACTION_LABELS_FR[["reassign"]])
    grey_out_style <- createStyle(fgFill = "#D9D9D9", fontColour = "#A6A6A6")

    openxlsx::conditionalFormatting(wb, sheet_name, cols = new_value_text_col, rows = data_rows,
                          rule = paste0("$", action_col_letter, anchor_row, '<>"', rewrite_label, '"'),
                          style = grey_out_style, type = "expression")
    openxlsx::conditionalFormatting(wb, sheet_name, cols = new_value_choice_col, rows = data_rows,
                          rule = paste0("$", action_col_letter, anchor_row, '<>"', reassign_label, '"'),
                          style = grey_out_style, type = "expression")
  }

  setColWidths(wb, sheet_name, cols = 1:ncol(df_display),
              widths = c(22, 30, 18, 16, 14, 16, 20, 18, 22, 26, 44, 30, 34, 16, 22, 38, 34))
  setRowHeights(wb, sheet_name, rows = 1, heights = 40)

  # protection de la feuille : ligne d'en-tete et colonnes techniques verrouillees, "Que faire ?" et les
  # deux colonnes "Nouvelle reponse" deverrouillees ci-dessus - meme logique que write_review_sheet()/
  # A_verifier. Aucun mot de passe defini : protection uniquement contre les modifications accidentelles.
  protectWorksheet(wb, sheet_name, protect = TRUE)

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
# row_id/statut : voir carry_forward_decisions() plus bas - row_id est un identifiant STABLE d'une
# ligne (uuid + variable + type de verification), le meme d'une execution a l'autre du pipeline tant que
# les donnees a l'origine de la ligne n'ont pas change ; statut indique si la ligne reprend deja une
# decision d'un classeur "suivi a verifier" precedent (voir Etape 1 du tableau de bord) ou si elle est
# nouvelle. Les deux sont des colonnes techniques (ne pas modifier), jamais utilisees pour decider quoi
# que ce soit cote equipe de suivi - uniquement pour permettre au pipeline de se souvenir du travail deja
# fait la prochaine fois qu'il tourne sur des donnees qui se sont agrandies.
REVIEW_HEADERS_FR <- c(
  uuid = "Identifiant enquete (ne pas modifier)",
  row_id = "ID (ne pas modifier)",
  commune = "Commune",
  partenaire = "Partenaire",
  date = "Date",
  enqueteur = "Enqueteur",
  boutique = "Boutique / vendeur",
  variable = "Variable (ne pas modifier)",
  question_label = "Question",
  reponse_actuelle = "Reponse actuelle",
  probleme = "Probleme detecte",
  statut = "Statut (ne pas modifier)",
  action = "Que faire ? (obligatoire)",
  nouvelle_reponse = "Nouvelle reponse (si 'Corriger la reponse') - si plusieurs choix, separer par ;",
  # ce commentaire est aussi repris comme raison de suppression dans le classeur final (deletion_log)
  # quand "Que faire ?" = "Supprimer toute l'enquete" - voir deletion_log_final dans 2_nettoyage_final.R
  commentaire = "Commentaire (raison de suppression si l'enquete entiere est supprimee)"
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
  # length(question) == 0 (pas seulement NA) : cas reel quand all_issues a 0 ligne (aucun probleme
  # detecte du tout - jeu de donnees parfaitement propre) - rowwise()/mutate() sur un tibble a 0 ligne
  # appelle quand meme cette fonction une fois avec question = character(0) pour deduire le type de
  # colonne resultat, et if(is.na(character(0))) leve "argument is of length zero" sans ce garde-fou.
  if (length(question) == 0 || is.na(question)) return("Toute l'enquete")
  lbl <- kobo_survey$`label::french`[kobo_survey$name == question]
  lbl <- lbl[!is.na(lbl) & nzchar(lbl)]
  if (length(lbl) > 0) return(lbl[1])
  str_replace_all(question, "_", " ")
}

# Resout une variable technique (nom direct, ou colonne indicatrice binaire "question/choix" pour les
# select_multiple) vers son q_type et sa liste de choix kobo - utilise pour proposer une liste
# deroulante des choix valides dans la colonne "Nouvelle reponse" (voir write_review_sheet()) plutot
# que du texte libre. Renvoie NULL si la variable n'est pas une question a choix (ou n'existe pas dans
# l'outil), auquel cas la cellule reste en texte libre.
get_variable_choice_list <- function(variable) {
  if (length(variable) == 0 || is.na(variable) || !nzchar(variable)) return(NULL)
  base_name <- if (variable %in% kobo_survey$name) variable else str_split(variable, "/")[[1]][1]
  row <- kobo_survey[kobo_survey$name == base_name, ]
  if (nrow(row) == 0 || !row$q_type[1] %in% c("select_one", "select_multiple")) return(NULL)
  list(q_type = row$q_type[1], list_name = row$list_name[1])
}

# Retire les caracteres de controle bruts (hors tabulation/retour a la ligne/retour chariot) illegaux
# dans le contenu XML (norme XML 1.0, production Char) - openxlsx ne les filtre/echappe pas de maniere
# fiable avant ecriture (confirme empiriquement : un octet de controle isole comme 0x01, parfaitement
# valide en UTF-8, est ecrit tel quel dans sharedStrings.xml) ce qui rend le classeur illisible pour
# Excel ("reparation" au chargement, erreur XML generique du type "Line 1, column 0") alors
# qu'openxlsx2::wb_load() ne detecte rien du tout (parseur plus permissif - pas fiable comme test de
# validite ici, verifier plutot avec un parseur XML strict sur le fichier reellement sauvegarde). Ce
# genre de caractere peut apparaitre dans n'importe quel champ texte libre issu de kobo (reponse,
# commentaire...) ou tape par l'equipe de suivi en corrigeant le classeur - improbable mais jamais
# totalement exclu vu la variete des claviers/telephones/copier-coller utilises sur le terrain.
sanitize_xml_text <- function(x) {
  if (!is.character(x)) return(x)
  gsub("[\x01-\x08\x0b\x0c\x0e-\x1f]", "", x, perl = TRUE)
}

# Applique sanitize_xml_text() a toutes les colonnes texte d'un data.frame avant ecriture (write_*_sheet
# ci-dessous) - les colonnes non-texte (numeriques, dates, facteurs...) traversent inchangees.
sanitize_df_text <- function(df) {
  df[] <- lapply(df, sanitize_xml_text)
  df
}

# Feuille d'accueil du classeur "suivi a verifier" - un tableau "Onglet -> Quoi faire", une ligne par
# onglet DANS LE MEME ORDRE que les onglets eux-memes (voir les noms numerotes SHEET_* dans
# 2_nettoyage_final.R), pour que l'equipe de suivi sache en un coup d'oeil quoi faire dans chacun.
# Toujours la 1ere feuille du classeur (appelee en 1er dans 2_nettoyage_final.R, avant les autres
# addWorksheet() - openxlsx ordonne les feuilles selon l'ordre d'appel). Texte fixe (pas de donnees) :
# autonome, n'utilise que openxlsx comme add_other_responses_sheet()/write_review_sheet() ci-dessous.
write_suivi_intro_sheet <- function(wb, sheet_name = "1_Lisez_moi") {
  addWorksheet(wb, sheet_name)

  intro_text <- paste(
    "Ce classeur reprend les enquetes signalees par la revue automatique du mois JMMI.",
    "Certaines corrections evidentes (reponses \"RAS\"/\"rien a signaler\"/remerciements sans",
    "information utile, reponse \"autre\" incoherente avec la question...) ont deja ete converties",
    "en vide automatiquement - inutile d'y revenir.",
    "Si la collecte est encore en cours et que ce classeur a ete regenere sur des donnees qui se sont",
    "agrandies depuis une premiere revue : la colonne \"Statut\" indique \"Repris (deja traite)\" pour",
    "les lignes qui reprennent automatiquement une decision deja prise precedemment (rien a refaire,",
    "sauf pour changer d'avis), et \"Nouveau\" pour les lignes qui necessitent une premiere revue.",
    "Les lignes deja tapees dans Journal_additionnel/Journal_suppression sont recopiees telles quelles",
    "(pas de colonne \"Statut\" pour ces 2 onglets, saisis manuellement).",
    "Une fois tous les onglets ci-dessous remplis, enregistrez ce classeur (meme nom) et",
    "transmettez-le pour la construction du classeur final."
  )

  guide <- data.frame(
    Onglet = c("2_Donnees_brutes", "3_A_verifier", "4_Verification_prix", "5_Reponses_autres",
              "6_Journal_additionnel", "7_Journal_suppression"),
    "Quoi faire" = c(
      "Donnees brutes de reference (deja anonymisees, corrections automatiques deja appliquees) - pour retrouver le detail complet d'une enquete sans rouvrir l'export Kobo. Lecture seule.",
      "Une ligne par probleme signale (valeur aberrante, incoherence logique, quasi-doublon...). Colonnes grisees = ne pas modifier. Pour chaque ligne, choisir une action dans \"Que faire ? (obligatoire)\" : Ne rien changer, Corriger la reponse (remplir alors \"Nouvelle reponse\"), Mettre une valeur manquante, ou Supprimer toute l'enquete.",
      "Memes principes que 3_A_verifier, uniquement pour les prix des articles suivis. Les lignes d'un meme article (prix saisi(s) + unite standard + autre unite) sont regroupees visuellement (fond de couleur alterne, bordure) : elles forment UN SEUL bloc a traiter ensemble. Le prix final (et sa conversion en euros/dollars) est recalcule automatiquement une fois les corrections appliquees.",
      "Reponses \"autre\" (texte libre) a reclasser. Pour chaque ligne, choisir une action dans \"Que faire ? (obligatoire)\" : Reecrire (taper le texte corrige/traduit dans \"Nouvelle reponse - texte libre\"), Reattribuer a un choix existant (choisir dans la liste deroulante \"Nouvelle reponse - choix existant\" ; si plusieurs choix sont necessaires a la fois, les taper plutot dans \"texte libre\", separes par ;), ou Supprimer (la reponse est mise a vide). Excel ne pouvant pas verrouiller une colonne selon l'action choisie, les deux colonnes \"Nouvelle reponse\" restent toujours modifiables mais celle qui ne correspond pas a l'action choisie apparait grisee - ne remplir que celle qui s'applique.",
      "Pour ajouter une correction que la revue automatique n'a PAS signalee (absente de 3_A_verifier/4_Verification_prix) : une ligne par correction, avec l'identifiant d'enquete (uuid, voir 2_Donnees_brutes) et le nom technique de la variable concernee. change_type : change_response (remplir alors new_value), blank_response (mettre a vide), ou no_action.",
      "Pour supprimer une enquete ENTIERE qui n'a pas ete signalee automatiquement : une ligne par enquete, avec l'identifiant d'enquete (uuid) et la raison de la suppression."
    ),
    check.names = FALSE, stringsAsFactors = FALSE
  )

  # banniere de titre (meme rouge que l'entete de 3_A_verifier/4_Verification_prix/5_Reponses_autres)
  writeData(wb, sheet_name, "Comment remplir ce classeur", startRow = 1, startCol = 1, colNames = FALSE)
  mergeCells(wb, sheet_name, cols = 1:2, rows = 1)
  addStyle(wb, sheet_name, style = createStyle(fgFill = RED_HEADER, fontColour = "#FFFFFF",
                                               textDecoration = "bold", fontSize = 16, valign = "center",
                                               halign = "left"), rows = 1, cols = 1:2)
  setRowHeights(wb, sheet_name, rows = 1, heights = 30)

  writeData(wb, sheet_name, intro_text, startRow = 2, startCol = 1, colNames = FALSE)
  mergeCells(wb, sheet_name, cols = 1:2, rows = 2)
  addStyle(wb, sheet_name, style = createStyle(wrapText = TRUE, valign = "top"), rows = 2, cols = 1)
  # hauteur calculee comme pour le tableau "guide" plus bas (~130 caracteres de large une fois les
  # colonnes 1+2 fusionnees, ~16 unites de hauteur par ligne enveloppee) plutot que fixe - un texte
  # d'introduction plus long (ex. ajout d'un paragraphe) ne doit pas se retrouver tronque a l'ouverture.
  setRowHeights(wb, sheet_name, rows = 2, heights = max(62, ceiling(nchar(intro_text) / 130) * 16))

  table_row1 <- 4
  writeData(wb, sheet_name, guide, startRow = table_row1, colNames = TRUE)
  addStyle(wb, sheet_name, style = createStyle(fgFill = GREY_HEADER, textDecoration = "bold",
                                               valign = "center"), rows = table_row1, cols = 1:2)
  data_rows <- (table_row1 + 1):(table_row1 + nrow(guide))
  addStyle(wb, sheet_name, style = createStyle(wrapText = TRUE, valign = "top", textDecoration = "bold"),
          rows = data_rows, cols = 1)
  addStyle(wb, sheet_name, style = createStyle(wrapText = TRUE, valign = "top"), rows = data_rows, cols = 2)

  # hauteur de ligne calculee (colonne "Quoi faire" de ~95 caracteres de large, ~16 unites de hauteur
  # par ligne de texte enveloppee) plutot que laissee a la hauteur par defaut : sans cela, un paragraphe
  # qui s'enveloppe sur plusieurs lignes apparait tronque a l'ouverture (Excel ne recalcule pas toujours
  # la hauteur des lignes generees par openxlsx) tant que l'utilisateur ne double-clique pas la bordure
  # de la ligne pour l'ajuster manuellement.
  for (i in seq_len(nrow(guide))) {
    n_wrapped_lines <- ceiling(nchar(guide[["Quoi faire"]][i]) / 95)
    setRowHeights(wb, sheet_name, rows = table_row1 + i, heights = max(20, n_wrapped_lines * 16))
  }

  setColWidths(wb, sheet_name, cols = 1, widths = 26)
  setColWidths(wb, sheet_name, cols = 2, widths = 110)
  freezePane(wb, sheet_name, firstActiveRow = table_row1 + 1)

  invisible(wb)
}

# Integre le journal de nettoyage additionnel DIRECTEMENT dans le classeur "suivi a verifier" (2
# onglets : "Journal_additionnel" pour des corrections champ par champ non signalees par la revue
# automatique, "Journal_suppression" pour supprimer des enquetes entieres) - remplace l'ancien modele
# telecharge separement (voir prepare_additional_cleaning_log()/prepare_additional_deletion_log(), qui
# lisent desormais ces onglets PAR NOM directement dans le classeur rempli, plutot qu'un fichier a part).
# Pas de feuille "raw_data" dupliquee ici (contrairement a l'ancien modele) : la feuille "raw_data" du
# classeur "suivi a verifier" lui-meme (voir 2_nettoyage_final.R) sert deja de reference, avec les
# corrections automatiques (RAS, reponses sans information utile...) deja appliquees.
ADDITIONAL_LOG_HEADERS <- c("uuid", "variable", "issue", "feedback", "change_type", "old_value", "new_value")
ADDITIONAL_LOG_CHANGE_TYPES <- c("change_response", "blank_response", "no_action", "remove_survey")
DELETION_LOG_HEADERS <- c("uuid", "reason_deletion")

# previous_additional/previous_deletion : lignes DEJA TAPEES dans le classeur "suivi a verifier"
# precedent (voir previous_review_path/carry_forward_decisions() dans 2_nettoyage_final.R), a
# recopier telles quelles dans le nouveau classeur. Contrairement a "A_verifier"/"Verification_prix"/
# "Reponses_autres" (regenerees automatiquement a chaque execution, avec un row_id pour faire
# correspondre une ligne fraiche a une decision precedente), ces 2 feuilles sont a saisie 100% manuelle
# et ne sont jamais regenerees - "reprendre" une ligne precedente n'a donc pas de sens a faire
# correspondre par identifiant, juste a recopier ce qui a deja ete tape, pour eviter de tout retaper
# quand la collecte se poursuit et que le classeur est regenere sur des donnees agrandies. Data frame
# avec les colonnes de ADDITIONAL_LOG_HEADERS/DELETION_LOG_HEADERS (dans n'importe quel ordre - reordonne
# ci-dessous), ou NULL si aucune ligne a reprendre.
add_additional_log_sheets <- function(wb, sheet_name = "6_Journal_additionnel", sheet_name2 = "7_Journal_suppression",
                                      previous_additional = NULL, previous_deletion = NULL) {
  sheet <- sheet_name
  addWorksheet(wb, sheet)
  writeData(wb, sheet, as.list(ADDITIONAL_LOG_HEADERS), startRow = 1, colNames = FALSE)
  if (!is.null(previous_additional) && nrow(previous_additional) > 0) {
    writeData(wb, sheet, previous_additional[, ADDITIONAL_LOG_HEADERS, drop = FALSE], startRow = 2, colNames = FALSE)
  }
  addStyle(wb, sheet, style = createStyle(textDecoration = "bold", fgFill = GREY_HEADER),
          rows = 1, cols = seq_along(ADDITIONAL_LOG_HEADERS))
  freezePane(wb, sheet, firstActiveRow = 2)
  addFilter(wb, sheet, rows = 1, cols = seq_along(ADDITIONAL_LOG_HEADERS))
  setColWidths(wb, sheet, cols = seq_along(ADDITIONAL_LOG_HEADERS), widths = c(24, 22, 30, 30, 16, 20, 20))

  # zone de saisie deverrouillee genereuse (1000 lignes) - au-dela, les cellules resteraient
  # verrouillees par defaut (comportement standard Excel sur une feuille protegee)
  addStyle(wb, sheet, style = createStyle(locked = FALSE), rows = 2:1000,
          cols = seq_along(ADDITIONAL_LOG_HEADERS), gridExpand = TRUE)
  change_type_col <- which(ADDITIONAL_LOG_HEADERS == "change_type")
  # liste litterale en ligne (pas de sqref sur plage multi-cellules), donc pas concernee par la
  # contrainte "sqref = xsd:list separe par des espaces, pas des virgules" de queue_list_validation() -
  # voir la note associee dans ce fichier.
  queue_list_validation(sheet, col = change_type_col, rows = 2:1000,
                        list_range = paste0('"', paste(ADDITIONAL_LOG_CHANGE_TYPES, collapse = ","), '"'))

  protectWorksheet(wb, sheet, protect = TRUE)

  sheet2 <- sheet_name2
  addWorksheet(wb, sheet2)
  writeData(wb, sheet2, as.list(DELETION_LOG_HEADERS), startRow = 1, colNames = FALSE)
  if (!is.null(previous_deletion) && nrow(previous_deletion) > 0) {
    writeData(wb, sheet2, previous_deletion[, DELETION_LOG_HEADERS, drop = FALSE], startRow = 2, colNames = FALSE)
  }
  addStyle(wb, sheet2, style = createStyle(textDecoration = "bold", fgFill = GREY_HEADER),
          rows = 1, cols = seq_along(DELETION_LOG_HEADERS))
  freezePane(wb, sheet2, firstActiveRow = 2)
  addFilter(wb, sheet2, rows = 1, cols = seq_along(DELETION_LOG_HEADERS))
  setColWidths(wb, sheet2, cols = seq_along(DELETION_LOG_HEADERS), widths = c(24, 40))
  addStyle(wb, sheet2, style = createStyle(locked = FALSE), rows = 2:1000,
          cols = seq_along(DELETION_LOG_HEADERS), gridExpand = TRUE)
  protectWorksheet(wb, sheet2, protect = TRUE)

  invisible(wb)
}

# Ecrit la feuille de verification consolidee dans un classeur existant (meme logique que
# add_other_responses_sheet() : l'appelant cree/charge `wb` et sauvegarde une fois toutes les feuilles
# ajoutees). `df` doit avoir les colonnes techniques de REVIEW_HEADERS_FR ; "action" peut etre vide (a
# remplir par l'equipe) ou deja une valeur de ACTION_LABELS_FR.
write_review_sheet <- function(wb, df, sheet_name = "3_A_verifier") {
  # bloc de variables liees (prix + contexte d'un meme article - voir item_blocks/expand_price_item_row
  # dans 2_nettoyage_final.R) - extrait AVANT le select() ci-dessous, puisque block_id n'est pas une des
  # colonnes de REVIEW_HEADERS_FR (donc jamais affichee comme colonne), seulement utilisee ici pour
  # appliquer un bandeau de couleur alterne aux lignes d'un meme bloc (voir plus bas).
  block_id <- if ("block_id" %in% colnames(df)) df$block_id else rep(NA_character_, nrow(df))
  # is_calculated_field (Verification_prix uniquement - voir expand_price_item_row() dans
  # 2_nettoyage_final.R) : certaines variables du bloc (ex. "bucket_20l_price_unit_item") sont elles-
  # memes un champ CALCULATE de l'outil kobo, recalcule automatiquement par recompute_price_items() a
  # partir des VRAIS champs saisis - la cellule "Que faire ?" reste EDITABLE comme les autres (demande
  # utilisateur : verrouiller ces cellules rendait la revue trop penible), mais recoit une note Excel
  # (pas un verrou) expliquant que toute correction saisie ici sera de toute facon recalculee - et que ce
  # recalcul est desormais lui-meme journalise (voir price_recalc_log dans 2_nettoyage_final.R), donc
  # rien n'est jamais perdu silencieusement.
  is_calculated_field <- if ("is_calculated_field" %in% colnames(df)) {
    df$is_calculated_field
  } else {
    rep(FALSE, nrow(df))
  }
  is_calculated_field[is.na(is_calculated_field)] <- FALSE

  df_display <- df %>% select(names(REVIEW_HEADERS_FR))
  colnames(df_display) <- unname(REVIEW_HEADERS_FR[colnames(df_display)])
  df_display <- sanitize_df_text(df_display)

  addWorksheet(wb, sheet_name)
  writeData(wb, sheet_name, df_display, colNames = TRUE)

  n <- nrow(df_display) + 1
  # 2:n vaut c(2, 1) (sequence DESCENDANTE, pas une plage vide) quand n == 1, c'est-a-dire quand
  # df_display a 0 ligne - cas reel (pas seulement theorique) d'un jeu de donnees sans aucun probleme
  # detecte. Sans ce garde-fou, les styles/validations "lignes de donnees" ci-dessous retoucheraient
  # aussi la ligne d'en-tete (ligne 1), la deverrouillant notamment via le style locked = FALSE plus bas.
  data_rows <- if (n >= 2) 2:n else integer(0)
  col_of_display <- function(key) match(REVIEW_HEADERS_FR[[key]], colnames(df_display))
  action_col <- col_of_display("action")
  nouvelle_reponse_col <- col_of_display("nouvelle_reponse")
  statut_col <- col_of_display("statut")
  # colonnes A a L (toutes les colonnes d'information/contexte, de "Identifiant enquete" a "Statut") :
  # protegees en bloc - l'equipe de suivi n'a besoin de modifier que "Que faire ?"/"Nouvelle reponse"/
  # "Commentaire" (colonnes M a O). "statut" est la derniere colonne de REVIEW_HEADERS_FR avant "action" -
  # 1:statut_col couvre donc exactement A:L tant que l'ordre des colonnes de REVIEW_HEADERS_FR ne change pas.
  technical_id_cols <- 1:statut_col

  # en-tete : fond rouge, texte blanc, gras, figee, avec filtres
  addStyle(wb, sheet_name, style = createStyle(fgFill = RED_HEADER, fontColour = "#FFFFFF",
                                               textDecoration = "bold", wrapText = TRUE, valign = "center",
                                               halign = "center"),
           rows = 1, cols = 1:ncol(df_display))
  freezePane(wb, sheet_name, firstActiveRow = 2, firstActiveCol = 2)
  addFilter(wb, sheet_name, rows = 1, cols = 1:ncol(df_display))

  # texte avec retour a la ligne pour toutes les colonnes de contenu - deverrouillees (locked = FALSE)
  # puisque ce sont les cellules que l'equipe de suivi doit remplir (voir protectWorksheet() ci-dessous)
  addStyle(wb, sheet_name, style = createStyle(wrapText = TRUE, valign = "top", locked = FALSE),
           rows = data_rows, cols = 1:ncol(df_display), gridExpand = TRUE)

  # colonnes techniques (a ne pas modifier) : fond gris pour bien les distinguer des colonnes a remplir,
  # et verrouillees (locked = TRUE, explicite car addStyle() remplace entierement le style precedent sur
  # ces memes cellules - sans cela elles resteraient deverrouillees par le style ci-dessus)
  addStyle(wb, sheet_name, style = createStyle(fgFill = GREY_HEADER, fontColour = "#595959", wrapText = TRUE,
                                               valign = "top", locked = TRUE),
           rows = data_rows, cols = technical_id_cols, gridExpand = TRUE)

  # alternance de fond gris TRES clair / gris clair PAR BLOC + bordure MOYENNE autour du perimetre de
  # chaque bloc de lignes liees (prix + std/other d'un meme article - voir expand_price_item_row() dans
  # 2_nettoyage_final.R) : les deux se combinent (le fond fait percevoir le bloc comme une bande
  # continue, la bordure en marque nettement la limite - demande explicite de l'utilisateur de garder
  # les deux, pas de remplacer l'un par l'autre). Chaque appel definit un style COMPLET (fond +
  # police/verrouillage + bordure eventuelle) en une seule fois plutot que d'accumuler plusieurs
  # bordures avec stack = TRUE sur les memes cellules. Le fond est d'abord pose sur tout le rectangle du
  # bloc (sans bordure), puis les 4 appels perimetre (gauche/droite/haut/bas, MEME fond) l'ecrasent
  # uniquement sur leur cote - les 2 derniers ("haut"/"bas", pleine largeur) ecrasent volontairement les
  # coins deja touches par "gauche"/"droite" : le cadre n'est donc pas ferme au pixel pres aux 4 coins,
  # mais reste tres lisible (lignes horizontales continues sur toute la largeur du bloc).
  BLOCK_SHADE_LIGHT <- "#F2F2F2"
  BLOCK_SHADE_MED <- "#E3E3E3"
  if (any(!is.na(block_id))) {
    blocks_rle <- rle(block_id)
    block_end_idx <- cumsum(blocks_rle$lengths)
    block_start_idx <- block_end_idx - blocks_rle$lengths + 1
    shade_idx <- 0
    apply_block_style <- function(rows_range, cols_range, shade, border_side = NULL) {
      grey_cols <- intersect(cols_range, technical_id_cols)
      content_cols <- setdiff(cols_range, grey_cols)
      border_args <- if (is.null(border_side)) list() else list(border = border_side, borderStyle = "medium")
      if (length(grey_cols) > 0) {
        addStyle(wb, sheet_name,
                style = do.call(createStyle, c(list(fgFill = shade, fontColour = "#595959", wrapText = TRUE,
                                                    valign = "top", locked = TRUE), border_args)),
                rows = rows_range, cols = grey_cols, gridExpand = TRUE)
      }
      if (length(content_cols) > 0) {
        addStyle(wb, sheet_name,
                style = do.call(createStyle, c(list(fgFill = shade, wrapText = TRUE, valign = "top",
                                                    locked = FALSE), border_args)),
                rows = rows_range, cols = content_cols, gridExpand = TRUE)
      }
    }
    for (b in seq_along(blocks_rle$values)) {
      if (is.na(blocks_rle$values[b])) next
      shade_idx <- shade_idx + 1
      shade <- if (shade_idx %% 2 == 1) BLOCK_SHADE_LIGHT else BLOCK_SHADE_MED
      r1 <- block_start_idx[b] + 1  # +1 pour la ligne d'en-tete
      r2 <- block_end_idx[b] + 1
      all_cols <- 1:ncol(df_display)
      apply_block_style(r1:r2, all_cols, shade)
      apply_block_style(r1:r2, 1, shade, "left")
      apply_block_style(r1:r2, ncol(df_display), shade, "right")
      apply_block_style(r1, all_cols, shade, "top")
      apply_block_style(r2, all_cols, shade, "bottom")
    }
  }

  # liste deroulante pour la colonne "Que faire ?" - liste en ligne (pas besoin d'une feuille annexe,
  # les 4 options tiennent largement dans la limite de 255 caracteres d'Excel)
  action_choices <- paste0('"', paste(unname(ACTION_LABELS_FR), collapse = ","), '"')
  queue_list_validation(sheet_name, col = action_col, rows = data_rows, list_range = action_choices)

  # note Excel (pas un verrou, voir commentaire en tete de fonction) sur la cellule "Que faire ?" des
  # lignes is_calculated_field - visible en survolant le petit triangle rouge ajoute par Excel dans le
  # coin de la cellule. Mise en attente (queue_comment()) plutot qu'ecrite directement ici avec
  # openxlsx::writeComment() - voir le commentaire au-dessus de PENDING_COMMENTS/apply_classic_dropdowns()
  # pour la raison (openxlsx2, appele juste apres pour corriger les listes deroulantes, faisait
  # disparaitre les commentaires ecrits par openxlsx).
  calculated_rows <- which(is_calculated_field) + 1  # +1 pour la ligne d'en-tete
  if (length(calculated_rows) > 0) {
    recalcul_note <- paste0("Cette valeur est un champ calcule automatiquement a partir des autres variables ",
                            "du bloc (prix/unite saisis ci-dessus). Vous pouvez taper une correction ici si ",
                            "besoin, mais elle sera de toute facon recalculee automatiquement si l'une des ",
                            "autres variables change - le recalcul final sera lui-meme visible dans le journal ",
                            "de nettoyage.")
    for (r in calculated_rows) queue_comment(sheet_name, col = action_col, row = r, text = recalcul_note)
  }

  # liste deroulante pour "Nouvelle reponse" quand la variable signalee est une question a choix
  # (select_one/select_multiple) : chaque ligne obtient la liste des choix VALIDES DE SA PROPRE
  # question, pour eviter que l'equipe tape un libelle qui ne correspondra a aucun code une fois
  # retraduit (voir read_review_sheet()). Les libelles sont stockes dans une feuille cachee (une
  # colonne par liste de choix kobo referencee, partagee entre les questions qui utilisent la meme
  # liste) - meme principe que la feuille "Listes_..." de add_other_responses_sheet(). Pour les
  # select_multiple, si plusieurs choix sont necessaires, l'equipe peut les separer par ";" (voir
  # read_review_sheet()) - la liste deroulante n'insere qu'un seul choix a la fois (limitation Excel).
  choice_lists_sheet <- substr(paste0("Listes_", sheet_name), 1, 31)
  addWorksheet(wb, choice_lists_sheet, visible = FALSE)

  if (length(data_rows) > 0) {
    variable_lists <- vapply(df$variable, function(v) {
      info <- get_variable_choice_list(v)
      if (is.null(info)) NA_character_ else info$list_name
    }, character(1))

    unique_lists <- unique(variable_lists[!is.na(variable_lists)])
    for (li in seq_along(unique_lists)) {
      list_name <- unique_lists[li]
      choice_labels <- kobo_choices$label[kobo_choices$list_name == list_name]
      choice_labels <- choice_labels[!is.na(choice_labels) & nzchar(choice_labels)]
      if (length(choice_labels) == 0) next

      writeData(wb, choice_lists_sheet, choice_labels, startCol = li)
      col_letter <- openxlsx::int2col(li)
      list_values <- paste0("'", choice_lists_sheet, "'!$", col_letter, "$1:$", col_letter,
                            "$", length(choice_labels))

      target_rows <- which(variable_lists == list_name) + 1  # +1 pour la ligne d'en-tete
      queue_list_validation(sheet_name, col = nouvelle_reponse_col, rows = target_rows,
                            list_range = list_values)
    }
  }

  # colonnes elargies (notamment "Nouvelle reponse", au libelle tres long) pour que le texte de l'entete
  # tienne sur moins de lignes une fois enveloppe - permet en retour une hauteur de ligne d'en-tete plus
  # raisonnable juste en dessous (demande utilisateur : entete auparavant beaucoup trop haute). "Question"
  # (8eme colonne, le libelle complet de la question kobo) elargie a son tour (demande utilisateur : les
  # lignes de donnees etaient trop hautes, ce texte s'enveloppant sur de nombreuses lignes courtes).
  setColWidths(wb, sheet_name, cols = 1:ncol(df_display),
              widths = c(22, 30, 16, 18, 16, 16, 20, 28, 55, 22, 34, 16, 28, 38, 44))
  setRowHeights(wb, sheet_name, rows = 1, heights = 26)

  # protection de la feuille : empeche de modifier l'en-tete (ligne 1, verrouillee par defaut) et les
  # colonnes A a L (verrouillees ci-dessus), tout en laissant M a O ("Que faire ?"/"Nouvelle reponse"/
  # "Commentaire", deverrouillees ci-dessus) editables pour toutes les lignes, et le tri/filtre
  # disponibles. Aucun mot de passe defini : protection uniquement contre les modifications
  # accidentelles (un clic sur "Oter la protection" dans Excel suffit a la retirer), pas une mesure de
  # securite.
  protectWorksheet(wb, sheet_name, protect = TRUE)

  invisible(wb)
}

# Relit une feuille de verification consolidee remplie par l'equipe de suivi, et retraduit les
# en-tetes/la liste deroulante "Que faire ?" vers les noms/valeurs techniques attendus par
# cleaningtools (create_clean_data(), review_cleaning_log(), ...) - meme logique de "traduction a la
# frontiere" que pour la feuille reponses_autres : l'equipe ne voit jamais ces noms techniques, seul ce
# point de lecture les reintroduit.

# Retraduit le contenu de "Nouvelle reponse" (voir la liste deroulante ajoutee par write_review_sheet())
# du libelle kobo choisi par l'equipe vers le code technique attendu par cleaningtools -
# correspondance exacte en priorite (cas le plus courant : un seul libelle choisi dans la liste
# deroulante) ; pour les select_multiple, repli sur un decoupage par ";" (et non par espace : les
# libelles eux-memes contiennent des espaces) si l'equipe a combine plusieurs choix dans la meme
# cellule. Toute portion qui ne correspond a aucun libelle connu est conservee telle quelle (deja un
# code valide, ou texte libre) plutot que rejetee.
choice_translate <- function(variable, value) {
  if (length(value) == 0) return(NA_character_)
  if (is.na(value) || !nzchar(value)) return(value)
  info <- get_variable_choice_list(variable)
  if (is.null(info)) return(value)

  choices <- kobo_choices[kobo_choices$list_name == info$list_name, ]
  exact <- choices$name[choices$label == value]
  if (length(exact) > 0) return(exact[1])

  if (info$q_type == "select_multiple" && grepl(";", value, fixed = TRUE)) {
    parts <- trimws(str_split(value, ";")[[1]])
    parts <- parts[nzchar(parts)]
    codes <- vapply(parts, function(p) {
      m <- choices$name[choices$label == p]
      if (length(m) > 0) m[1] else p
    }, character(1))
    return(paste(codes, collapse = " "))
  }

  value
}

# translate_choices = FALSE : ne retraduit PAS "nouvelle_reponse" (libelle choisi -> code kobo) - utilise
# par carry_forward_decisions() (voir plus bas) quand on relit un classeur PRECEDENT pour en reprendre
# les decisions dans un classeur fraichement regenere : on veut alors reinjecter exactement le libelle
# francais tel que l'equipe l'avait choisi (ce que write_review_sheet() reecrira tel quel dans la
# cellule), pas le code technique cleaningtools - qui, lui, resterait correct pour create_clean_data()
# mais s'afficherait tel quel (illisible) si jamais reaffiche dans une cellule "Nouvelle reponse".
read_review_sheet <- function(path, sheet_name = "3_A_verifier", translate_choices = TRUE) {
  raw <- read_excel(path, sheet = sheet_name, col_types = "text")
  header_rev <- setNames(names(REVIEW_HEADERS_FR), unname(REVIEW_HEADERS_FR))
  match_idx <- match(colnames(raw), names(header_rev))
  colnames(raw)[!is.na(match_idx)] <- header_rev[match_idx[!is.na(match_idx)]]

  action_rev <- setNames(names(ACTION_LABELS_FR), unname(ACTION_LABELS_FR))
  raw <- raw %>% mutate(change_type = unname(action_rev[action]))

  # une valeur de "Que faire ?" non reconnue (faute de frappe, texte tape au lieu de choisi dans la
  # liste deroulante, ...) donnerait un change_type NA silencieux sans ce garde-fou - la ligne serait
  # ensuite ignoree plus loin dans le pipeline sans que personne ne s'en apercoive.
  invalid_action <- raw %>% filter(!is.na(action) & nzchar(action) & is.na(change_type))
  if (nrow(invalid_action) > 0) {
    detail <- paste0("uuid=", invalid_action$uuid, ", variable=", invalid_action$variable,
                     " : valeur '", invalid_action$action, "' non reconnue pour \"Que faire ?\"")
    stop("Valeur(s) invalide(s) dans la colonne \"Que faire ?\" du classeur suivi a verifier (feuille '",
        sheet_name, "', fichier ", basename(path), ") - utilisez uniquement les options de la liste ",
        "deroulante ('",
        paste(unname(ACTION_LABELS_FR), collapse = "', '"), "') :\n- ", paste(detail, collapse = "\n- "))
  }

  if (translate_choices) {
    raw <- raw %>%
      rowwise() %>%
      mutate(nouvelle_reponse = choice_translate(variable, nouvelle_reponse)) %>%
      ungroup()
  }

  raw %>%
    rename(question = variable, old_value = reponse_actuelle, new_value = nouvelle_reponse,
          issue = probleme, comment = commentaire)
}

# Relit une feuille "Reponses_autres" remplie et la retraduit vers les noms techniques (memes noms que
# df_other/OTHER_HEADERS_FR - pas de renommage supplementaire type read_review_sheet(), puisque
# OTHER_HEADERS_FR utilise deja directement les noms internes comme cles). Utilisee a la fois par la
# PASSE 2 (2_nettoyage_final.R) et par carry_forward_decisions() (voir plus bas) - un seul point de
# traduction "francais affiche -> nom technique" pour cette feuille, au lieu de dupliquer la logique.
read_other_responses_sheet <- function(path, sheet_name = "5_Reponses_autres") {
  raw <- read_excel(path, sheet = sheet_name, col_types = "text")
  header_en <- setNames(names(OTHER_HEADERS_FR), unname(OTHER_HEADERS_FR))
  match_idx <- match(colnames(raw), names(header_en))
  colnames(raw)[!is.na(match_idx)] <- header_en[match_idx[!is.na(match_idx)]]
  raw
}

# Reprend les decisions (action + valeur(s) de remplacement) d'un classeur "suivi a verifier" PRECEDENT
# dans les lignes fraichement regenerees pour l'execution en cours, ligne par ligne, par identifiant
# stable (voir row_id, cree dans 2_nettoyage_final.R avant l'appel a write_review_sheet()/
# add_other_responses_sheet()) - pour relancer le pipeline sur des donnees qui se sont AGRANDIES
# (nouvelles soumissions Kobo) sans perdre le travail de revue deja fait sur les enquetes deja
# presentes avant, ni devoir copier-coller a la main entre deux classeurs. Fonction generique (aucune
# connaissance Excel/nom de feuille) - reutilisee pour "A_verifier", "Verification_prix" et
# "Reponses_autres" depuis 2_nettoyage_final.R, avec des noms de colonnes differents a chaque fois.
#
# new_df        : lignes fraichement generees pour CETTE execution (row_id/statut deja presents,
#                 colonnes de decision encore vides).
# previous_df   : lignes du classeur precedent, DEJA retraduites en noms techniques identiques a ceux
#                 de new_df (voir les appels dans 2_nettoyage_final.R) - peut contenir des lignes non
#                 remplies (action encore vide) : ces lignes ne reprennent aucune decision (rien a
#                 reprendre), mais passent quand meme au statut "Repris (deja traite)" (voir plus bas) -
#                 elles ont deja ete PRESENTEES a l'equipe de suivi lors d'un cycle precedent, meme si
#                 aucune decision n'a encore ete prise dessus, ce n'est donc pas une ligne "Nouveau".
# id_col        : colonne identifiant stable commune aux deux data frames ("row_id").
# action_col    : colonne "Que faire ?" - la decision n'est reprise que si elle a une valeur non vide
#                 ici (une ligne du classeur precedent restee vide n'a rien a transmettre a ce niveau,
#                 mais compte quand meme comme "deja vue" - voir statut_col plus bas).
# carry_cols    : autres colonnes de decision a reprendre quand une correspondance valide ET une
#                 decision non vide sont trouvees (ex. "nouvelle_reponse"/"commentaire", ou
#                 "new_value_text"/"new_value_choice").
# old_value_col : colonne de contexte a comparer avant de reprendre une decision ET avant de marquer
#                 "Repris" (garde-fou : si la valeur d'origine a change depuis le classeur precedent -
#                 ne devrait jamais arriver si les anciennes donnees sont vraiment figees, mais peu
#                 couteux a verifier - la ligne reste "Nouveau" pour une revue fraiche plutot que
#                 d'afficher a tort une ligne perimee comme deja traitee). NULL pour desactiver cette
#                 verification.
# statut_col    : colonne de statut ("Nouveau"/"Repris (deja traite)") - passe a "Repris (deja
#                 traite)" pour TOUTE ligne dont le row_id existait deja dans le classeur precedent,
#                 QU'ELLE AIT OU NON UNE DECISION REMPLIE (demande utilisateur explicite : une ligne
#                 deja presentee au cycle precedent n'est plus "Nouveau" meme si elle n'a pas encore
#                 ete traitee - ce qui, en pratique, marque aussi automatiquement TOUT un bloc de
#                 "Verification_prix" comme "Repris" des qu'une seule de ses lignes liees a recu une
#                 decision, puisque toutes les lignes d'un meme bloc existent deja ensemble dans le
#                 classeur precedent).
carry_forward_decisions <- function(new_df, previous_df, id_col = "row_id", action_col = "action",
                                    carry_cols = character(0), old_value_col = NULL,
                                    statut_col = "statut") {
  if (is.null(previous_df) || nrow(previous_df) == 0 || !id_col %in% colnames(previous_df)) {
    return(new_df)
  }

  # un row_id ne devrait apparaitre qu'une fois dans un classeur "suivi a verifier" normal - s'il y en a
  # quand meme (ex. classeur precedent lui-meme issu d'une fusion de plusieurs fichiers), garder la
  # PREMIERE ligne rencontree pour chacun plutot que de faire planter toute la generation du classeur
  # pour une anomalie mineure portant sur d'anciennes donnees.
  dup_ids <- previous_df[[id_col]][duplicated(previous_df[[id_col]])]
  if (length(dup_ids) > 0) {
    warning("Classeur precedent : ", length(unique(dup_ids)), " identifiant(s) en double - seule la ",
           "premiere ligne rencontree est reprise pour chacun : ", paste(unique(dup_ids), collapse = ", "))
    previous_df <- previous_df[!duplicated(previous_df[[id_col]]), , drop = FALSE]
  }

  match_idx <- match(new_df[[id_col]], previous_df[[id_col]])
  matched <- !is.na(match_idx)

  if (!is.null(old_value_col) && old_value_col %in% colnames(new_df) &&
      old_value_col %in% colnames(previous_df)) {
    # NA et "" traites comme EQUIVALENTS ("valeur absente") avant comparaison - sans cela, un champ
    # legitimement vide des les deux executions (ex. "_other_unit_item" quand l'unite standard a ete
    # utilisee, jamais rempli) casse silencieusement la reprise : construit fraichement, il vaut ""
    # (voir expand_price_item_row() dans 2_nettoyage_final.R, qui convertit NA en "" explicitement) ;
    # relu depuis le classeur precedent via read_review_sheet()/read_other_responses_sheet(), la MEME
    # cellule vide revient en NA (openxlsx ecrit "" comme une cellule Excel vide, readxl relit une
    # cellule vide en NA) - "" != NA (resultat NA en R, converti en FALSE juste apres) bloquait donc a
    # tort la reprise, precisement pour les champs les plus souvent vides.
    norm_empty <- function(x) ifelse(is.na(x) | !nzchar(x), "", x)
    old_matches <- norm_empty(new_df[[old_value_col]]) == norm_empty(previous_df[[old_value_col]][match_idx])
    old_matches[is.na(old_matches)] <- FALSE
    matched <- matched & old_matches
  }

  if (!any(matched)) return(new_df)

  # statut : TOUTE ligne matchee passe a "Repris (deja traite)", decision presente ou non (voir
  # commentaire en tete de fonction).
  if (statut_col %in% colnames(new_df)) {
    new_df[[statut_col]][matched] <- "Repris (deja traite)"
  }

  # decision (action + carry_cols) : reprise uniquement pour le sous-ensemble des lignes matchees qui
  # ont effectivement une valeur non vide dans action_col du classeur precedent - une ligne matchee mais
  # restee vide n'a rien a transmettre a ce niveau (elle a deja obtenu son statut "Repris" ci-dessus).
  src_idx <- match_idx[matched]
  has_decision <- !is.na(previous_df[[action_col]][src_idx]) & nzchar(previous_df[[action_col]][src_idx])
  if (any(has_decision)) {
    can_carry <- matched
    can_carry[matched] <- has_decision
    src_idx_decided <- src_idx[has_decision]
    new_df[[action_col]][can_carry] <- previous_df[[action_col]][src_idx_decided]
    for (col in carry_cols) {
      if (col %in% colnames(previous_df)) {
        new_df[[col]][can_carry] <- previous_df[[col]][src_idx_decided]
      }
    }
  }

  new_df
}

#############################################################################################################
# JOURNAL DE NETTOYAGE ADDITIONNEL (optionnel - complement manuel au classeur "A_verifier")
#############################################################################################################
# ADDITIONAL_LOG_HEADERS/ADDITIONAL_LOG_CHANGE_TYPES definis plus haut, avec add_additional_log_sheets()
# qui genere les onglets "Journal_additionnel"/"Journal_suppression" du classeur "suivi a verifier" -
# memes noms de colonnes que la feuille cleaning_log du classeur final (voir build_final_workbook()).

# Verifie chaque ligne du journal additionnel importe et construit un message d'erreur DETAILLE (numero
# de ligne Excel, uuid, variable) pour CHAQUE probleme trouve, plutot que de laisser create_clean_data()
# (cleaningtools) echouer plus loin avec un message generique qui ne dit pas quelle ligne du fichier
# importe en est la cause. S'arrete (stop()) avec la liste complete des problemes si au moins un est
# trouve, pour que l'utilisateur puisse tout corriger en une seule fois plutot que de re-importer a
# repetition pour decouvrir les erreurs une par une.
validate_additional_cleaning_log <- function(df, raw_data_all) {
  missing_cols <- setdiff(ADDITIONAL_LOG_HEADERS, colnames(df))
  if (length(missing_cols) > 0) {
    stop("Le journal de nettoyage additionnel importe n'a pas les colonnes attendues (colonnes manquantes : ",
        paste(missing_cols, collapse = ", "), ") - utilisez le modele telecharge depuis le tableau de bord, ",
        "sans modifier l'en-tete.")
  }

  problems <- character(0)
  for (i in seq_len(nrow(df))) {
    excel_row <- i + 1  # +1 pour la ligne d'en-tete
    uuid <- df$uuid[i]
    variable <- df$variable[i]
    change_type <- df$change_type[i]
    new_value <- df$new_value[i]

    ref <- paste0("Ligne ", excel_row, " (uuid=", ifelse(is.na(uuid) || !nzchar(uuid), "manquant", uuid),
                 ", variable=", ifelse(is.na(variable) || !nzchar(variable), "manquant", variable), ")")

    if (is.na(uuid) || !nzchar(uuid)) {
      problems <- c(problems, paste0(ref, " : identifiant d'enquete (uuid) manquant."))
      next
    }
    if (!uuid %in% raw_data_all$uuid) {
      problems <- c(problems, paste0(ref, " : cet identifiant d'enquete n'existe pas dans les donnees brutes importees."))
      next
    }
    if (is.na(change_type) || !change_type %in% ADDITIONAL_LOG_CHANGE_TYPES) {
      problems <- c(problems, paste0(ref, " : 'change_type' doit etre l'une de ces valeurs : ",
                                     paste(ADDITIONAL_LOG_CHANGE_TYPES, collapse = ", "), " (valeur trouvee : '",
                                     ifelse(is.na(change_type), "vide", change_type), "')."))
      next
    }
    if (change_type != "remove_survey") {
      if (is.na(variable) || !nzchar(variable)) {
        problems <- c(problems, paste0(ref, " : 'variable' est obligatoire sauf quand change_type = 'remove_survey'."))
        next
      }
      if (!variable %in% colnames(raw_data_all)) {
        problems <- c(problems, paste0(ref, " : la variable '", variable, "' n'existe pas dans les donnees brutes."))
        next
      }
    }
    if (change_type == "change_response" && (is.na(new_value) || !nzchar(new_value))) {
      problems <- c(problems, paste0(ref, " : 'new_value' est obligatoire quand change_type = 'change_response'."))
    }
  }

  if (length(problems) > 0) {
    stop("Le journal de nettoyage additionnel contient des erreurs - a corriger avant de continuer :\n- ",
        paste(problems, collapse = "\n- "))
  }

  invisible(TRUE)
}

# Lit, valide (voir validate_additional_cleaning_log()) et prepare le journal de nettoyage additionnel
# pour etre concatene avec le classeur "A_verifier" rempli (voir 2_nettoyage_final.R, passe 2) - memes
# noms de colonnes internes (question/comment) que read_review_sheet(), pour que tout le reste du
# pipeline (filled_field_level, deletion_log, combined_cleaning_log, ...) le traite exactement comme des
# lignes de plus du classeur rempli, sans logique dupliquee. Lu par NOM DE FEUILLE (pas par position :
# "Journal_additionnel" est desormais un onglet DE PLUS du classeur "suivi a verifier" lui-meme - voir
# add_additional_log_sheets() - plus un fichier separe ou il aurait ete la 1ere feuille). Renvoie NULL si
# l'onglet est absent (classeur genere par une version anterieure du pipeline) ou vide (present mais non
# rempli) - optionnel dans les deux cas, ce n'est pas une erreur.
prepare_additional_cleaning_log <- function(path, raw_data_all) {
  df <- tryCatch(read_excel(path, sheet = "6_Journal_additionnel", col_types = "text"),
                error = function(e) NULL)
  if (is.null(df)) return(NULL)
  df <- df[rowSums(!is.na(df)) > 0, , drop = FALSE]
  if (nrow(df) == 0) return(NULL)

  validate_additional_cleaning_log(df, raw_data_all)

  df %>%
    rename(question = variable, comment = feedback) %>%
    select(uuid, question, change_type, old_value, new_value, issue, comment)
}

#############################################################################################################
# JOURNAL DE SUPPRESSION ADDITIONNEL (2eme onglet du meme classeur que le journal de nettoyage
# additionnel - DELETION_LOG_HEADERS definie plus haut, voir add_additional_log_sheets()) - permet de
# signaler des enquetes entieres a supprimer sans devoir remplir toutes les colonnes du journal de
# nettoyage general (change_type=remove_survey, variable vide, etc.). Pas de colonne enum_id :
# l'identifiant de l'enqueteur est de toute facon re-derive automatiquement dans le classeur final via
# l'enum_lookup base sur les donnees brutes (voir deletion_log_final dans 2_nettoyage_final.R).
#############################################################################################################

# Meme logique que validate_additional_cleaning_log() : message d'erreur DETAILLE (numero de ligne
# Excel, uuid) pour chaque probleme trouve, tous rapportes en une seule fois.
validate_additional_deletion_log <- function(df, raw_data_all) {
  missing_cols <- setdiff(DELETION_LOG_HEADERS, colnames(df))
  if (length(missing_cols) > 0) {
    stop("L'onglet 'Journal_suppression' du journal de nettoyage additionnel n'a pas les colonnes attendues ",
        "(colonnes manquantes : ", paste(missing_cols, collapse = ", "),
        ") - utilisez le modele telecharge depuis le tableau de bord, sans modifier l'en-tete.")
  }

  problems <- character(0)
  for (i in seq_len(nrow(df))) {
    excel_row <- i + 1  # +1 pour la ligne d'en-tete
    uuid <- df$uuid[i]
    reason <- df$reason_deletion[i]

    ref <- paste0("Ligne ", excel_row, " (uuid=", ifelse(is.na(uuid) || !nzchar(uuid), "manquant", uuid), ")")

    if (is.na(uuid) || !nzchar(uuid)) {
      problems <- c(problems, paste0(ref, " : identifiant d'enquete (uuid) manquant."))
      next
    }
    if (!uuid %in% raw_data_all$uuid) {
      problems <- c(problems, paste0(ref, " : cet identifiant d'enquete n'existe pas dans les donnees brutes importees."))
      next
    }
    if (is.na(reason) || !nzchar(reason)) {
      problems <- c(problems, paste0(ref, " : 'reason_deletion' (raison de la suppression) est obligatoire."))
    }
  }

  if (length(problems) > 0) {
    stop("L'onglet 'Journal_suppression' du journal de nettoyage additionnel contient des erreurs - a corriger avant de continuer :\n- ",
        paste(problems, collapse = "\n- "))
  }

  invisible(TRUE)
}

# Lit, valide et prepare l'onglet "Journal_suppression" pour etre concatene avec le classeur "A_verifier"
# rempli exactement comme prepare_additional_cleaning_log() - memes colonnes internes (question/comment),
# change_type fixe a "remove_survey" (question=NA, comme les lignes "toute l'enquete" issues de
# A_verifier). Renvoie NULL si l'onglet est absent (classeur genere par une version anterieure du modele)
# ou vide (present mais non rempli) - optionnel, ce n'est pas une erreur dans les deux cas.
prepare_additional_deletion_log <- function(path, raw_data_all) {
  df <- tryCatch(read_excel(path, sheet = "7_Journal_suppression", col_types = "text"),
                error = function(e) NULL)
  if (is.null(df)) return(NULL)
  df <- df[rowSums(!is.na(df)) > 0, , drop = FALSE]
  if (nrow(df) == 0) return(NULL)

  validate_additional_deletion_log(df, raw_data_all)

  df %>%
    transmute(uuid = uuid, question = NA_character_, change_type = "remove_survey",
             old_value = NA_character_, new_value = NA_character_, issue = NA_character_,
             comment = reason_deletion)
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
# d'etiquette de la feuille. Complete aussi les colonnes USD calculees (voir "RECALCUL DES PRIX EN
# DOLLARS US" dans 2_nettoyage_final.R) qui n'existent pas dans l'outil kobo et n'auraient donc sinon
# aucune etiquette - price_conversion_labels est defini dans ce meme script (meme environnement
# d'execution, voir scripts/0_configuration.R) ; absent (raw_data, 1_suivi_quotidien.R) -> ignore.
apply_var_labels <- function(df){
  lbls <- kobo_survey$`label::french`
  names(lbls) <- kobo_survey$name
  lbls <- lbls[intersect(names(df), names(lbls))]
  if (exists("price_conversion_labels", inherits = TRUE)) {
    lbls <- c(lbls, price_conversion_labels[intersect(names(df), names(price_conversion_labels))])
  }
  labelled::var_label(df) <- as.list(lbls)
  df
}

GREY_HEADER <- "#D1D3D4"
RED_HEADER <- "#EE5859"
DARK_GREY_BANNER <- "#818183"

# Une seule ligne d'en-tete figee, fond gris clair - utilisee pour cleaning_log/deletion_log/variable_tracker/survey/choices
write_simple_sheet <- function(wb, sheet_name, df, auto_width = TRUE){
  df <- sanitize_df_text(df)
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
  label_row <- sanitize_xml_text(label_row)
  df <- sanitize_df_text(df)

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
fill_checklist_data_cleaning <- function(wb){
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
  set_check(39, "y")   # Verifications logiques
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

# estime la hauteur de ligne necessaire pour un texte enveloppe (wrapText) dans une colonne de largeur
# donnee (en "unites de caractere" Excel, meme echelle que setColWidths()/wb_set_col_widths()) - openxlsx
# et openxlsx2 n'ont pas de vrai autofit de hauteur de ligne (Excel ne le calcule qu'a l'ouverture
# manuelle, et seulement si on double-clique la bordure de ligne), d'ou cette approximation par nombre de
# caracteres : nombre de lignes ~= longueur du texte / caracteres par ligne (tient aussi compte des
# retours a la ligne explicites "\n" deja presents dans le texte), hauteur = nombre de lignes * hauteur
# d'une ligne + marge. Demande utilisateur : "make the rows width dependent on the content" (feuille
# READ_ME) - remplace les hauteurs fixes (45/30) utilisees jusque-la pour toutes les lignes sans
# distinction.
estimate_wrapped_row_height <- function(text, col_width_chars, line_height = 15, min_height = 15, padding = 6) {
  if (is.na(text) || !nzchar(text)) return(min_height)
  explicit_lines <- max(lengths(gregexpr("\n", text))) + 1
  chars_per_line <- max(1, floor(col_width_chars))
  wrapped_lines <- ceiling(nchar(text) / chars_per_line)
  n_lines <- max(explicit_lines, wrapped_lines)
  max(min_height, n_lines * line_height + padding)
}

# Construit la feuille READ_ME avec le meme langage visuel que la propre sortie d'analyse JMMI de
# REACH (onglet "Lisez_moi" de REACH_CAR_JMMI_Analyse_*.xlsx) : une banniere de titre gris fonce, une
# banniere de sous-titre gris clair, des libelles de section en gras en colonne A avec des descriptions
# avec retour a la ligne en colonne B, puis une barre d'en-tete rouge "Contenu" au-dessus d'un index
# des feuilles a deux colonnes. Hauteur de chaque ligne dependante de son contenu (voir
# estimate_wrapped_row_height() ci-dessus), pas une valeur fixe identique pour toutes les lignes.
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
    setRowHeights(wb, sheet, rows = r, heights = estimate_wrapped_row_height(s[2], col_width_chars = 100))
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
    setRowHeights(wb, sheet, rows = r, heights = estimate_wrapped_row_height(sheet_descriptions[[nm]], col_width_chars = 100))
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
                                 assessment_month_label, assessment_year){

  wb <- loadWorkbook(Sys.getenv("checklist_template"))
  for (sh in c("READ_ME","survey","choices","variable_tracker","raw_data","clean_data",
               "cleaning_log","deletion_log","sampling_info")) {
    if (sh %in% names(wb)) removeWorksheet(wb, sh)
  }
  freezePane(wb, "CHECKLIST_Data_Cleaning", firstActiveRow = 2)

  # deletion_log est toujours incluse, meme vide (0 suppression) - plus simple et plus previsible pour
  # le reviseur qu'une feuille qui apparait ou disparait d'un cycle a l'autre.
  has_deletions <- nrow(deletion_log_final) > 0
  fill_checklist_data_cleaning(wb)

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
  # reason_deletion contient un texte libre (repris du commentaire de l'equipe de suivi - voir
  # deletion_log_final dans 2_nettoyage_final.R) souvent plus long qu'une largeur "auto" ne l'anticipe -
  # elargie explicitement et avec retour a la ligne pour rester lisible
  reason_col <- which(colnames(deletion_log_final) == "reason_deletion")
  if (length(reason_col) == 1) {
    setColWidths(wb, "deletion_log", cols = reason_col, widths = 60)
    addStyle(wb, "deletion_log", style = createStyle(wrapText = TRUE, valign = "top"),
            rows = 2:(nrow(deletion_log_final) + 1), cols = reason_col, gridExpand = TRUE, stack = TRUE)
  }

  # loadWorkbook() conserve l'onglet actif du modele checklist_template.xlsx tel qu'il etait ouvert
  # dans Excel au moment ou ce fichier a ete enregistre (ex. "cleaning_log", 8eme onglet du modele
  # complet a 10 feuilles) - un index qui n'existe plus une fois les feuilles placeholder retirees
  # ci-dessus (7 feuilles restantes a ce stade). Sans cette reinitialisation explicite,
  # worksheetOrder<-() plus bas echoue ("N doesn't exist as sheet index") des que le modele a ete
  # resauvegarde avec un autre onglet que le premier actif - reinitialiser sur la 1ere feuille avant
  # tout reordonnancement rend ce comportement independant de l'etat du fichier modele.
  activeSheet(wb) <- 1

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
  # LOCALAPPDATA n'existe que sous Windows - vide (donc "") sur toute autre plateforme (ex. serveur
  # Linux type Posit Connect), ce qui ferait pointer file.path() vers la RACINE du systeme de fichiers
  # ("/jmmi_car_build") au lieu d'un dossier local. Repli sur tempdir() dans ce cas : le probleme de
  # verrouillage OneDrive que ce dossier de mise en attente contourne est de toute facon propre a
  # Windows (pas de OneDrive sur un serveur Linux), donc perdre la persistance entre sessions R n'a pas
  # d'impact la ou ce mecanisme n'a pas de raison d'etre sollicite.
  stage_dir <- if (nzchar(Sys.getenv("LOCALAPPDATA"))) {
    file.path(Sys.getenv("LOCALAPPDATA"), "jmmi_car_build")
  } else {
    file.path(tempdir(), "jmmi_car_build")
  }
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
