################################################################################
# JMMI RELABELLING - RE-ETIQUETAGE DES CHOIX D'UN CLASSEUR NETTOYE
################################################################################
# ENTREE      : Sys.getenv("clean_data_file") - le classeur final produit par le nettoyage JMMI (lit
#               ses feuilles "raw_data"/"clean_data", en-tete sur deux lignes) ; Sys.getenv("kobo_tool_file")
#               - l'outil kobo correspondant (feuilles "survey"/"choices")
# SORTIE      : Sys.getenv("output_dir")/<nom du fichier d'entree>_libelle.xlsx - une COPIE COMPLETE du
#               classeur d'entree (memes feuilles, meme ordre, meme mise en forme) ou, dans "raw_data" et
#               "clean_data" uniquement : (1) les CODES des reponses a choix (select_one/select_multiple)
#               sont remplaces par leur LIBELLE, (2) la ligne d'etiquette des colonnes binaires
#               select_multiple ("parent/code") est completee au format "libelle question/libelle choix".
#               Toutes les autres feuilles (survey, choices, cleaning_log, deletion_log, variable_tracker,
#               etc.) et les NOMS de colonnes techniques sont copies tels quels, sans modification.
# DESCRIPTION : Outil autonome et generique (voir scripts/fonctions_relabel.R) : ne relabellise QUE
#               les colonnes presentes a la fois dans les donnees et dans l'outil kobo, et ne modifie
#               jamais les noms de colonnes - seulement les valeurs des questions a choix et, pour les
#               sous-colonnes select_multiple, leur etiquette (ligne 1) uniquement.
# OBJECTIF    : Produire une version du classeur nettoye plus lisible pour l'analyse/le partage (des
#               libelles clairs plutot que des codes kobo), sans creer de nouvelle etape de nettoyage
#               ni toucher au classeur final officiel.
# CONTEXTE    : Copie autonome pour rca_jmmi_dashboard (voir en-tete de app.R) - source d'origine dans
#               rca_jmmi_relabelling/relabel_choices.R, a recopier ici si elle evolue. Prend en entree
#               la SORTIE du nettoyage (le classeur final), pas les donnees brutes.
################################################################################

library(readxl)
library(openxlsx)
library(dplyr)
library(stringr)

source("scripts/fonctions_relabel.R")

cat("--> Chargement de l'outil kobo...\n")
kobo_survey <- read_excel(Sys.getenv("kobo_tool_file"), sheet = "survey", col_types = "text") %>% rename_with(tolower)
kobo_choices <- read_excel(Sys.getenv("kobo_tool_file"), sheet = "choices", col_types = "text") %>% rename_with(tolower)

clean_data_file <- Sys.getenv("clean_data_file")

cat("--> Chargement du classeur nettoye (toutes feuilles)...\n")
wb <- loadWorkbook(clean_data_file)
sheets_to_relabel <- intersect(c("raw_data", "clean_data"), names(wb))
if (length(sheets_to_relabel) == 0) {
  stop("Aucune feuille 'raw_data' ou 'clean_data' trouvee dans le classeur nettoye - rien a re-etiqueter.")
}

for (sheet in sheets_to_relabel) {
  cat("--> Re-etiquetage de la feuille '", sheet, "'...\n", sep = "")
  sheet_content <- read_two_row_header_sheet(clean_data_file, sheet = sheet)
  relabelled_data <- relabel_choices(sheet_content$data, kobo_survey, kobo_choices)
  new_label_row <- relabel_multiple_choice_subcolumn_labels(
    colnames(sheet_content$data), sheet_content$label_row, kobo_survey, kobo_choices
  )
  overwrite_two_row_header_sheet(wb, sheet, relabelled_data, new_label_row)
}

output_dir <- Sys.getenv("output_dir")
if (!dir.exists(output_dir)) dir.create(output_dir, recursive = TRUE)
out_name <- paste0(tools::file_path_sans_ext(basename(clean_data_file)), "_libelle.xlsx")
out_path <- file.path(output_dir, out_name)

saveWorkbook(wb, out_path, overwrite = TRUE)

cat("\n--> Classeur re-etiquete ecrit dans :", out_path, "\n")
