################################################################################
# JMMI CAR - SUIVI QUOTIDIEN (PENDANT LA COLLECTE)
################################################################################
# ENTREE      : export kobo brut/jeu de donnees.ressources/RCA_JMMI_2026-08_TOOL.xlsx (outil kobo)
# SORTIE      : output/1_suivi_couverture_et_prix/RCA_JMMI_<mois>_maj_quotidienne_<horodatage>.xlsx -
#               un seul classeur. Nombre_Prix et Prix_Median (commune en premier, avec les colonnes de
#               contexte prefecture/sous_prefecture) sont les feuilles prioritaires, placees en premier/
#               cote a cote, avec un en-tete sur deux lignes (libelle de la question puis nom de
#               colonne) et des bannieres de section entre communes/sous-prefectures/prefectures/
#               national ; Prix_Median a une echelle de couleur rouge (prix eleve) / vert (prix bas)
#               calculee separement pour chaque article. Suivent Resume, Detail et
#               Resume_Partenaire_Zone (style standard, inchange), puis raw_data (jeu de donnees brutes
#               importe, anonymise de la meme maniere que raw_data dans le classeur final - voir
#               2_nettoyage_final.R -, en-tete sur deux lignes), pour retrouver le detail d'une ligne
#               signalee sans devoir rouvrir l'export kobo brut.
# DESCRIPTION : Ne fait AUCUNE verification/nettoyage - uniquement un suivi de la couverture de la
#               collecte (nombre d'enquetes/de prix par commune/partenaire) et des prix medians par
#               article, pour un pilotage au jour le jour.
# OBJECTIF    : Donner a l'equipe une photo rapide de l'avancement de la collecte (ou manque-t-il des
#               prix ? les prix semblent-ils plausibles ?) sans melanger cela avec le nettoyage des
#               donnees, qui se fait entierement a la fin de la collecte (voir 2_nettoyage_final.R).
# CONTEXTE    : A executer autant de fois que necessaire pendant que la collecte est en cours (ex.
#               chaque jour). Remplace l'ancien 1_verification_pendant_collecte.R, qui melangeait ce
#               suivi avec des verifications intermediaires - le projet a ete simplifie pour que toutes
#               les verifications se fassent en une seule fois, a la fin de la collecte.
################################################################################

rm(list=ls())

# chargement des packages, des fonctions et de la liste de variables predefinies
source("scripts/0_configuration.R")

################################################################################
# CHARGEMENT DU JEU DE DONNEES BRUT
################################################################################

raw_data_all <- read_raw_data(Sys.getenv("raw_dataset")) %>%
  add_admin_org_labels()

#-------------------------------------------------------------------------------
# mois/cycle d'evaluation : derive de la plus ancienne valeur de la variable kobo "date" (le jour de
# la visite du marche), et non d'une valeur codee en dur dans .Renviron, pour que le nom de fichier
# reflete toujours ce qui a ete reellement collecte.
period <- get_assessment_period(raw_data_all, date_col = "date")
assessment_month <- period$assessment_month

#-------------------------------------------------------------------------------
# recuperer les noms des colonnes de prix (utilise par Nombre_Prix et Prix_Median)
items_list <- kobo_survey %>%
  filter(type == "calculate" & str_ends(name, "_price_item")) %>%
  pull(name)

################################################################################
# SUIVI DE LA COUVERTURE (nombre d'enquetes/de prix par commune et par partenaire)
################################################################################

# tableau detaille
detail_df <- raw_data_all %>%
  count(admin1_label, admin2_label, admin3_label, organisation_name_label) %>%
  pivot_wider(names_from = organisation_name_label, values_from = n, values_fill = 0) %>%
  rowwise() %>% mutate(Total = sum(c_across(where(is.numeric)))) %>% ungroup() %>%
  bind_rows(summarise(., across(where(is.numeric), sum, na.rm = TRUE), .groups = "drop") %>%
              mutate(admin1_label = "Total", admin2_label = NA, admin3_label = NA))

# tableau resume
summary_df <- raw_data_all %>%
  count(organisation_name_label, name = "nombre_enquetes") %>%
  arrange(desc(nombre_enquetes)) %>%
  bind_rows(tibble(organisation_name_label = "Total", nombre_enquetes = sum(.$nombre_enquetes)))

# nombre de prix par article (commune en premier - voir get_num_prices())
num_prices_df <- get_num_prices(raw_data_all)

# prix median par article (commune en premier - voir get_median_prices())
median_prices_df <- get_median_prices(raw_data_all)

# resume partenaire vs zone
partner_region_summary <- create_partner_region_summary(raw_data_all)

# donnees brutes a inclure dans le classeur, anonymisees de la meme maniere que raw_data dans le
# classeur final (voir 2_nettoyage_final.R) - permet de retrouver le detail d'une ligne signalee dans
# Nombre_Prix/Prix_Median sans devoir rouvrir l'export kobo brut.
remove_cols_raw <- c(
  intersect(colnames(raw_data_all), filter(kobo_survey, type %in% c("geopoint"))$name),
  "audit", "audit_URL", "deviceid", "enum_id", "shop_name", "organisation_name_label"
)
remove_cols_raw <- intersect(remove_cols_raw, colnames(raw_data_all))
raw_data_export <- select(raw_data_all, -all_of(remove_cols_raw))
if ("_index" %in% colnames(raw_data_export)) raw_data_export <- rename(raw_data_export, id = `_index`)

################################################################################
# ECRITURE DU CLASSEUR
################################################################################

if (!dir.exists("output/1_suivi_couverture_et_prix")) dir.create("output/1_suivi_couverture_et_prix", recursive = TRUE)
output_file <- paste0("./output/1_suivi_couverture_et_prix/RCA_JMMI_", assessment_month,
                      "_maj_quotidienne_", format(Sys.time(), "%Y-%m-%d_%Hh%M"), ".xlsx")

wb <- createWorkbook()

# Nombre_Prix et Prix_Median sont les deux feuilles les plus importantes du suivi quotidien (couverture
# et plausibilite des prix) - regroupees cote a cote, en premier dans le classeur. Elles utilisent
# write_priority_sheet() (en-tete standard sur deux lignes - libelle puis nom de colonne - avec des
# bannieres de section entre communes/sous-prefectures/prefectures/national) plutot que
# add_styled_sheet() ; les autres feuilles gardent le style simple habituel.
np_rows <- write_priority_sheet(wb, "Nombre_Prix", num_prices_df)
mp_rows <- write_priority_sheet(wb, "Prix_Median", median_prices_df)
add_styled_sheet(wb, "Resume", summary_df)
add_styled_sheet(wb, "Detail", detail_df)
add_styled_sheet(wb, "Resume_Partenaire_Zone", partner_region_summary)
write_labeled_data_sheet(wb, "raw_data", apply_var_labels(raw_data_export))

worksheetOrder(wb) <- match(
  c("Nombre_Prix", "Prix_Median", "Resume", "Detail", "Resume_Partenaire_Zone", "raw_data"),
  names(wb)
)
activeSheet(wb) <- "Nombre_Prix"

# Nombre_Prix : signaler en rouge les cases avec 2 prix ou moins (partout), et en jaune specifiquement
# au niveau commune avec 1 ou 2 prix (le niveau le plus fin, ou combler un manque est le plus urgent).
# cols est recalcule par nom (et non par position fixe) car num_prices_df a aussi des colonnes de
# contexte (zone, prefecture, sous_prefecture, niveau_admin) avant les colonnes d'articles ; rows
# utilise np_rows (renvoye par write_priority_sheet()) car les bannieres de section decalent les
# lignes reelles par rapport a l'ordre des lignes dans num_prices_df.
num_prix_cols <- which(colnames(num_prices_df) %in% items_list)
apply_conditional_formatting("Nombre_Prix", rows = np_rows, cols = num_prix_cols,
                             rule = "<=2",
                             style = "#EE5859")
apply_conditional_formatting("Nombre_Prix",
                             rows = np_rows[num_prices_df$niveau_admin == "commune"],
                             cols = num_prix_cols, rule = c("=1", "=2"),
                             style = "#f8dd4c")

# Prix_Median : echelle de couleur rouge (prix eleve) -> vert (prix bas), calculee separement pour
# chaque article (voir apply_price_color_scale()) pour un reperage visuel immediat sans avoir besoin
# de lire chaque chiffre
prix_cols <- which(colnames(median_prices_df) %in% items_list)
apply_price_color_scale(wb, "Prix_Median", rows = mp_rows, cols = prix_cols)

saveWorkbook(wb, output_file, overwrite = TRUE)

cat("\n--> Suivi quotidien ecrit dans :", output_file, "\n")
