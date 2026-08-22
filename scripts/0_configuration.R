################################################################################
# JMMI CAR - 0. CONFIGURATION COMMUNE (chargee par chaque script principal)
################################################################################
# ENTREE      : ressources/RCA_JMMI_2026-08_TOOL.xlsx (survey/choices kobo, via Sys.getenv("kobo")),
#               ressources/check_list_august.xlsx (via Sys.getenv("logical_checks"))
# SORTIE      : Objets globaux utilises par tous les scripts : kobo_survey, kobo_choices,
#               kobo_choices_check, cols_numeric, other_labels, other_db, logical_checks_list ; ainsi
#               que ressources/labels_questions_others.xlsx (ecrit en effet de bord)
# DESCRIPTION : Charge les packages necessaires, importe fonctions_verification_outil_kobo.R et
#               fonctions_nettoyage.R, charge et valide l'outil kobo, et charge la liste des
#               verifications logiques.
# OBJECTIF    : Un seul endroit pour garder la configuration identique entre 1_suivi_quotidien.R et
#               2_nettoyage_final.R, afin que la logique outil/etiquettes/liste de verifications ne
#               soit jamais dupliquee ni ne diverge.
# CONTEXTE    : Adapte de src/init.R de JMMI_ETH - remplace label::english par label::french et pointe
#               vers les chemins de l'outil kobo/checklist propres a la RCA.
################################################################################

#-------------------------------------------------------------------------------
# chargement des packages
#-------------------------------------------------------------------------------
library(tidyverse)
library(readxl)
library(writexl)
library(openxlsx)
# utilisee uniquement par build_final_workbook() (fonctions_nettoyage.R) pour cloner les feuilles
# survey/choices depuis l'outil kobo avec leur mise en forme d'origine - openxlsx ne sait pas copier une
# feuille (avec sa mise en forme) d'un classeur charge vers un autre classeur, openxlsx2 le permet.
library(openxlsx2)

# devtools::install_github("impact-initiatives/cleaningtools")
library(cleaningtools)

options(scipen = 999) # pour eviter la notation scientifique
options(dplyr.summarise.inform = FALSE)

#-------------------------------------------------------------------------------
# chargement des fonctions
#-------------------------------------------------------------------------------
source("scripts/fonctions_verification_outil_kobo.R", local = TRUE)
source("scripts/fonctions_nettoyage.R", local = TRUE)

#-------------------------------------------------------------------------------
# chargement de l'outil KOBO
#-------------------------------------------------------------------------------
cat("\n- CHARGEMENT de l'outil ...\n")
kobo_survey <- read_excel(Sys.getenv("kobo"), sheet="survey", col_types="text") %>%
  rename_with(.,tolower) %>%
  filter(!is.na(name)) %>%
  mutate(q_type=as.character(lapply(type, function(x) str_split(x, " ")[[1]][1])),
         list_name=as.character(lapply(type, function(x) str_split(x, " ")[[1]][2])))

# la feuille choices de la RCA a deja une colonne "label" simple (pas de suffixe "::french" comme la feuille survey)
kobo_choices <- read_excel(Sys.getenv("kobo"), sheet="choices", col_types="text") %>%
  rename_with(.,tolower) %>%
  filter(!is.na(list_name)) %>%
  select(list_name, name, label) %>%
  distinct()

kobo_choices_check <- kobo_choices %>%
  select(list_name, name, label)

# verifier la colonne "relevant" de l'outil
t <- check_constraints(kobo_survey, kobo_choices_check)
if (length(t)>0) warning(paste0("\n", "Problemes de conditions relevant : ", paste(t, collapse="; ")))

# verifier le kobo avec les fonctions de cleaningtools
cleaningtools::verify_valid_choices(kobo_choices)
cleaningtools::verify_valid_survey(kobo_survey)

cat("..OK\n")

#-------------------------------------------------------------------------------
# recuperer tous les noms de colonnes entiers/decimaux de l'outil kobo
cols_numeric <- filter(kobo_survey, type %in% c("integer", "decimal")) %>% pull(name)

#-------------------------------------------------------------------------------
# generer la base des other_labels
other_labels <- get_other_labels()

# generer la base des question_other
other_db <- get_other_db()

#-------------------------------------------------------------------------------
# charger la liste des verifications logiques
logical_checks_list <- read_excel(Sys.getenv("logical_checks")) %>%
  filter(id_verification != "skip")
