################################################################################
# JMMI RCA - 4. ANALYSE MENSUELLE COMPLETE (descriptive + PMAS/MEB + Score de Fonctionnalite)
################################################################################
# ENTREE      : Sys.getenv("clean_data_file") - le classeur final produit par le nettoyage JMMI (memes
#               feuilles que build_final_workbook() : CHECKLIST_Data_Cleaning, survey, choices, raw_data,
#               clean_data, cleaning_log, deletion_log, variable_tracker, ...) ; Sys.getenv("kobo_tool_file")
#               - l'outil kobo correspondant (feuilles "survey"/"choices") ; Sys.getenv("eur_xrate")
#               (parite XAF/EUR). L'annee et le mois du cycle NE SONT PAS des parametres d'entree : ils
#               sont detectes automatiquement depuis la colonne "date" de clean_data (voir
#               get_assessment_period_from_data() dans scripts/fonctions_analyse.R). Optionnels :
#               Sys.getenv("analyse_previous_clean_data_file") + Sys.getenv("analyse_mois_precedent")
#               (pour la feuille "Evolution des paniers du PMAS" - laissee vide si absents).
# SORTIE      : Sys.getenv("output_dir")/<date>_CAR_JMMI_Dataset_Analyse_<mois>.xlsx - reprend toutes les
#               feuilles du classeur de nettoyage importe et y ajoute : Données nettoyées, Coût médian du
#               PMAS, Evolution des paniers du PMAS, Prix médian par article, Score de Fonctionnalité,
#               Analyse.
# DESCRIPTION : copie autonome pour rca_jmmi_dashboard (voir en-tete de app.R) de rca_jmmi_analysis/
#               1_analyse_complete.R - meme logique, adaptee aux variables d'environnement du tableau de
#               bord (au lieu des chemins fixes "input/"/"output/" du projet autonome). La liste
#               d'analyse (LOA) utilisee pour la feuille "Analyse" est generee automatiquement depuis
#               l'outil kobo importe (scripts/fonctions_analyse.R::build_loa()), pas depuis un fichier
#               DAP maintenu a la main.
# CONTEXTE    : source d'origine dans rca_jmmi_analysis/1_analyse_complete.R, a recopier ici si elle
#               evolue (voir note d'en-tete de app.R - ce dossier est une copie autonome).
################################################################################

library(tidyverse)
library(readxl)
library(writexl)
library(openxlsx)
library(openxlsx2)
library(analysistools)
library(presentresults)
library(srvyr)

options(scipen = 999)
options(dplyr.summarise.inform = FALSE)

source("scripts/fonctions_analyse.R", local = TRUE)
source("scripts/mfs.R", local = TRUE)

################################################################################
# PARAMETRES (variables d'environnement - voir run_analyse_pipeline() dans app.R)
################################################################################
INPUT_PATH <- Sys.getenv("clean_data_file")
KOBO_TOOL_FILE <- Sys.getenv("kobo_tool_file")
PAYS <- "RCA"
EUR_XRATE <- as.numeric(Sys.getenv("eur_xrate"))
PREVIOUS_INPUT_PATH <- Sys.getenv("analyse_previous_clean_data_file")   # "" si non fourni
PREVIOUS_MOIS <- Sys.getenv("analyse_mois_precedent")

if (!nzchar(INPUT_PATH) || !file.exists(INPUT_PATH)) {
  stop("Classeur d'entree introuvable : le classeur final nettoye n'a pas ete correctement importe.")
}
if (!nzchar(KOBO_TOOL_FILE) || !file.exists(KOBO_TOOL_FILE)) {
  stop("Outil kobo introuvable : l'outil kobo n'a pas ete correctement importe.")
}
if (is.na(EUR_XRATE)) stop("Taux de change XAF/EUR invalide.")

################################################################################
# CHARGEMENT DE L'OUTIL KOBO
################################################################################
cat("\n- CHARGEMENT de l'outil kobo ...\n")
kobo_survey <- read_excel(KOBO_TOOL_FILE, sheet = "survey", col_types = "text") %>%
  rename_with(tolower) %>%
  filter(!is.na(name)) %>%
  mutate(q_type = as.character(lapply(type, function(x) str_split(x, " ")[[1]][1])),
        list_name = as.character(lapply(type, function(x) str_split(x, " ")[[1]][2])))

kobo_choices <- read_excel(KOBO_TOOL_FILE, sheet = "choices", col_types = "text") %>%
  rename_with(tolower) %>%
  filter(!is.na(list_name)) %>%
  select(list_name, name, label) %>%
  distinct()
cat("..OK (", nrow(kobo_survey), " questions, ", nrow(kobo_choices), " choix)\n", sep = "")

################################################################################
# LECTURE DU CLASSEUR DE NETTOYAGE (base du classeur final - conserve sa mise en forme d'origine)
################################################################################
cat("\n- LECTURE du classeur de nettoyage :", INPUT_PATH, "...\n")
wb_final <- wb_load(INPUT_PATH)
cleaning_sheet_names <- wb_get_sheet_names(wb_final)
cat("  feuilles trouvees :", paste(cleaning_sheet_names, collapse = ", "), "\n")

raw_data_all <- read_excel(INPUT_PATH, sheet = "raw_data", col_types = "text")
clean_data <- read_excel(INPUT_PATH, sheet = "clean_data", skip = 1, col_types = "text")
cat("  raw_data :", nrow(raw_data_all), "lignes | clean_data :", nrow(clean_data), "lignes\n")

# annee/mois du cycle detectes automatiquement depuis les donnees (voir get_assessment_period_from_data()
# dans scripts/fonctions_analyse.R) plutot que demandes a l'utilisateur - cette information est deja
# presente dans le classeur importe (nom de fichier ET feuille README).
period <- get_assessment_period_from_data(clean_data, date_col = "date")
MOIS <- period$assessment_month_label
ANNEE <- period$assessment_year
cat("  periode detectee automatiquement :", MOIS, ANNEE, "\n")

################################################################################
# "DONNEES NETTOYEES" (jeu de donnees etiquete) -- DESACTIVE TEMPORAIREMENT (demande utilisateur : ne
# produire que "Coût médian du PMAS" + "Prix médian par article" pour l'instant). Decommenter le bloc
# if (FALSE) ci-dessous pour la reactiver.
################################################################################
if (FALSE) {
cat("\n- CONSTRUCTION du jeu de donnees etiquete (Donnees nettoyées) ...\n")
donnees_nettoyees <- build_labeled_dataset(clean_data, kobo_survey, kobo_choices)
wb_final <- wb_add_worksheet(wb_final, "Données nettoyées")
wb_final <- wb_add_data(wb_final, "Données nettoyées", donnees_nettoyees)
wb_final <- wb_add_font(wb_final, "Données nettoyées", dims = "A1", bold = TRUE)
wb_final <- wb_freeze_pane(wb_final, "Données nettoyées", firstActiveRow = 2, firstActiveCol = 2)
cat("  ..OK (", nrow(donnees_nettoyees), " lignes, ", ncol(donnees_nettoyees), " colonnes)\n", sep = "")
}

################################################################################
# PMAS / MEB : "Coût médian du PMAS", "Evolution des paniers du PMAS", "Prix médian par article"
################################################################################
cat("\n- CALCUL du PMAS (panier minimum) ...\n")

admin_lookup <- clean_data %>%
  select(admin1_label, admin2_label, admin3_label) %>%
  distinct(admin3_label, .keep_all = TRUE)
admin2_lookup <- clean_data %>%
  select(admin1_label, admin2_label) %>%
  distinct(admin2_label, .keep_all = TRUE)

# meb_food/meb_nfi/meb_wash/meb_total (calculate_meb(), non filtree par la methodologie de suffisance)
# restent utilises plus bas pour la feuille "Evolution des paniers du PMAS".
meb_food <- calculate_meb(clean_data, MEB_FOOD_WEIGHTS) %>% rename(meb_food = MEB)
meb_nfi <- calculate_meb(clean_data, MEB_NFI_WEIGHTS) %>% rename(meb_nfi = MEB)
meb_wash <- calculate_meb(clean_data, MEB_WASH_WEIGHTS) %>% rename(meb_wash = MEB)
meb_total <- calculate_meb(clean_data, MEB_WEIGHTS_ALL) %>% rename(meb_total = MEB)

# --- "Coût médian du PMAS" : niveaux Commune / Sous-Préfecture / Préfecture / National empilés dans une
# seule feuille (14 colonnes de couts + 1 colonne "Note", un bloc par niveau separe par une ligne vide),
# meme disposition que reference/REACH_CAR_JMMI_Dataset_Analyse_Juin_Juillet_2026 (1).xlsx. Chaque niveau
# applique la methodologie de suffisance des donnees (voir calculate_pmas_by_level()/
# calculate_pmas_national() dans scripts/fonctions_analyse.R) : panier total calcule seulement si >=12/17
# articles ont assez de prix locaux a ce niveau, sinon chaque panier (alimentaire/non alimentaire/
# hygiene) est calcule independamment s'il atteint son propre seuil, avec une note explicative.
PMAS_COST_COLS <- function(p) {
  list(
    `Coût du panier alimentaire` = p$meb_food,
    `Coût du panier non alimentaire` = p$meb_nfi,
    `Coût du panier des produits d'hygiène` = p$meb_wash,
    `Coût total des paniers` = p$meb_total,
    `Coût du panier alimentaire (en euro)` = p$meb_food / EUR_XRATE,
    `Coût du panier non alimentaire (en euro)` = p$meb_nfi / EUR_XRATE,
    `Coût du panier des produits d'hygiène (en euro)` = p$meb_wash / EUR_XRATE,
    `Coût total des paniers (en euro)` = p$meb_total / EUR_XRATE,
    Note = p$note
  )
}

# niveau Commune (admin3)
p_commune <- calculate_pmas_by_level(clean_data, "admin3_label") %>%
  filter(admin3_label != "National") %>%
  left_join(admin_lookup, by = "admin3_label")
pmas_commune <- p_commune %>%
  transmute(
    `Année` = ANNEE, Mois = MOIS, Pays = PAYS,
    `Préfecture` = admin1_label, `Sous-Préfecture` = admin2_label, Commune = admin3_label,
    !!!PMAS_COST_COLS(p_commune)
  ) %>%
  arrange(`Préfecture`, `Sous-Préfecture`, Commune)

# niveau Sous-Préfecture (admin2) - agrege directement depuis les prix individuels groupes par admin2
p_sp <- calculate_pmas_by_level(clean_data, "admin2_label") %>%
  left_join(admin2_lookup, by = "admin2_label")
pmas_sous_prefecture <- p_sp %>%
  transmute(
    `Année` = ANNEE, Mois = MOIS, Pays = PAYS,
    `Préfecture` = admin1_label, `Sous-Préfecture` = admin2_label, Commune = NA_character_,
    !!!PMAS_COST_COLS(p_sp)
  ) %>%
  arrange(`Préfecture`, `Sous-Préfecture`)

# niveau Préfecture (admin1) - idem, groupe directement par admin1
p_pref <- calculate_pmas_by_level(clean_data, "admin1_label")
pmas_prefecture <- p_pref %>%
  transmute(
    `Année` = ANNEE, Mois = MOIS, Pays = PAYS,
    `Préfecture` = admin1_label, `Sous-Préfecture` = NA_character_, Commune = NA_character_,
    !!!PMAS_COST_COLS(p_pref)
  ) %>%
  arrange(`Préfecture`)

# niveau National - une seule ligne, "Ensemble des marchés évalués" (meme libelle que le classeur de
# reference) dans la colonne Préfecture
p_nat <- calculate_pmas_national(clean_data)
pmas_national <- tibble(
  `Année` = ANNEE, Mois = MOIS, Pays = PAYS,
  `Préfecture` = "Ensemble des marchés évalués", `Sous-Préfecture` = NA_character_, Commune = NA_character_,
  !!!PMAS_COST_COLS(p_nat)
)

# ecriture : en-tete une seule fois, puis les 4 blocs empiles avec une ligne vide entre chacun - na=""
# pour que les cellules NA (Sous-Prefecture/Commune absents aux niveaux agreges, paniers/Note non
# calcules) restent VRAIMENT vides plutot que d'ecrire une erreur Excel #N/A (comportement par defaut
# de wb_add_data())
wb_final <- wb_add_worksheet(wb_final, "Coût médian du PMAS")
wb_final <- wb_add_data(wb_final, "Coût médian du PMAS", t(colnames(pmas_commune)), start_row = 1, col_names = FALSE, na = "")

row_cursor <- 2
wb_final <- wb_add_data(wb_final, "Coût médian du PMAS", pmas_commune, start_row = row_cursor, col_names = FALSE, na = "")
sep_row_1 <- row_cursor + nrow(pmas_commune)
row_cursor <- sep_row_1 + 1
wb_final <- wb_add_data(wb_final, "Coût médian du PMAS", pmas_sous_prefecture, start_row = row_cursor, col_names = FALSE, na = "")
sep_row_2 <- row_cursor + nrow(pmas_sous_prefecture)
row_cursor <- sep_row_2 + 1
wb_final <- wb_add_data(wb_final, "Coût médian du PMAS", pmas_prefecture, start_row = row_cursor, col_names = FALSE, na = "")
sep_row_3 <- row_cursor + nrow(pmas_prefecture)
row_cursor <- sep_row_3 + 1
wb_final <- wb_add_data(wb_final, "Coût médian du PMAS", pmas_national, start_row = row_cursor, col_names = FALSE, na = "")
# fusionne Préfecture/Sous-Préfecture/Commune sur la ligne nationale (seule colonne D remplie -
# "Ensemble des marchés évalués" - E et F vides) pour que le libelle s'affiche sur les 3 colonnes
wb_final <- wb_merge_cells(wb_final, "Coût médian du PMAS", dims = wb_dims(rows = row_cursor, cols = 4:6))

# ligne separatrice GRISE (fine) entre chaque bloc d'agregation (Commune / Sous-Préfecture / Préfecture
# / National) pour delimiter clairement les sections - la ligne vide entre chaque bloc est remplie en
# gris et sa hauteur reduite pour ressembler a une fine ligne de separation plutot qu'a une ligne de
# donnees vide
sep_rows <- c(sep_row_1, sep_row_2, sep_row_3)
wb_final <- wb_add_fill(wb_final, "Coût médian du PMAS", dims = wb_dims(rows = sep_rows, cols = 1:15), color = wb_color(hex = "FFBFBFBF"))
wb_final <- wb_set_row_heights(wb_final, "Coût médian du PMAS", rows = sep_rows, heights = 6)

# mise en forme reprise A L'IDENTIQUE (police, couleurs, bordures, formats numeriques) du classeur de
# reference REACH_CAR_JMMI_Dataset_Analyse_Juin_Juillet_2026, styles extraits directement du XML OOXML :
# en-tete = fond rouge corail FFEE5958 + police Roboto Condensed 14 gras, texte BLANC, ligne haute (46)
# pour les libelles enveloppes ; corps = police Segoe UI 12 gris fonce (couleur texte du theme, tintee
# +0.15 -> #262626), GRAS pour Commune et les 2 colonnes "Total" (XAF/EUR) - bordures fines noires
# partout ; ligne nationale nettement plus marquee (police Roboto Condensed 18 gras, bordure EPAISSE
# noire en haut ET en bas, format comptable XAF). Colonne "Note" (15e, O) ajoutee par rapport au
# classeur de reference, avec un style texte assorti. Lignes separatrices grises : convention
# deliberement differente du classeur de reference (qui utilise un vert pale) - demande explicite de
# l'utilisateur, conservee telle quelle.
last_row <- row_cursor + nrow(pmas_national) - 1
national_row_r <- row_cursor
full_dims <- wb_dims(rows = 1:last_row, cols = 1:15)
BODY_GREY <- "FF262626"

# wb_add_border() sur une PLAGE RECTANGULAIRE multi-cellules ne dessine par defaut qu'un cadre EXTERIEUR
# autour de la plage entiere (comme le ferait un simple "contour"), PAS une grille sur chaque cellule -
# les parametres inner_hgrid/inner_vgrid (constate empiriquement, teste isolement) sont necessaires pour
# obtenir une bordure sur CHAQUE cellule interieure, sans quoi la plupart des cellules du tableau
# restent sans bordure du tout.
wb_final <- wb_add_border(wb_final, "Coût médian du PMAS", dims = full_dims, inner_hgrid = "thin", inner_vgrid = "thin")
wb_final <- wb_add_border(wb_final, "Coût médian du PMAS", dims = "A1:O1",
                          left_border = "thin", left_color = wb_color(hex = "FFFFFFFF"),
                          right_border = "thin", right_color = wb_color(hex = "FFFFFFFF"),
                          top_border = "thin", top_color = wb_color(hex = "FFFFFFFF"),
                          bottom_border = "thin", bottom_color = wb_color(hex = "FF000000"))
wb_final <- wb_add_fill(wb_final, "Coût médian du PMAS", dims = "A1:O1", color = wb_color(hex = "FFEE5958"))
wb_final <- wb_add_font(wb_final, "Coût médian du PMAS", dims = "A1:O1", name = "Roboto Condensed", size = 14, bold = TRUE, color = wb_color(hex = "FFFFFFFF"))
wb_final <- wb_add_cell_style(wb_final, "Coût médian du PMAS", dims = "A1:O1", horizontal = "center", vertical = "center", wrap_text = "1")
wb_final <- wb_set_row_heights(wb_final, "Coût médian du PMAS", rows = 1, heights = 46)

# corps (communes/sous-prefectures/prefectures - PAS la ligne nationale, stylee separement plus bas)
body_rows <- setdiff(2:last_row, national_row_r)
wb_final <- wb_add_font(wb_final, "Coût médian du PMAS", dims = wb_dims(rows = body_rows, cols = 1:15), name = "Segoe UI", size = 12, color = wb_color(hex = BODY_GREY))
wb_final <- wb_add_font(wb_final, "Coût médian du PMAS", dims = wb_dims(rows = body_rows, cols = c(6, 10, 14)), name = "Segoe UI", size = 12, bold = TRUE, color = wb_color(hex = BODY_GREY))
wb_final <- wb_add_cell_style(wb_final, "Coût médian du PMAS", dims = wb_dims(rows = body_rows, cols = c(1,2,3,7:14)), horizontal = "center", vertical = "center")
wb_final <- wb_add_cell_style(wb_final, "Coût médian du PMAS", dims = wb_dims(rows = body_rows, cols = 4:6), horizontal = "left", vertical = "center", indent = 1)
wb_final <- wb_add_cell_style(wb_final, "Coût médian du PMAS", dims = wb_dims(rows = body_rows, cols = 15), horizontal = "left", vertical = "center", wrap_text = "1")
wb_final <- wb_add_numfmt(wb_final, "Coût médian du PMAS", dims = wb_dims(rows = body_rows, cols = 7:10),
                          numfmt = "#,##0\\ [$XAF];[Red]#,##0\\ [$XAF]")
wb_final <- wb_add_numfmt(wb_final, "Coût médian du PMAS", dims = wb_dims(rows = body_rows, cols = 11:14),
                          numfmt = "#,##0\\ \"€\";[Red]#,##0\\ \"€\"")

# ligne nationale : nettement plus marquee que les autres (police Roboto Condensed 18 gras, bordure
# EPAISSE noire en haut ET en bas, format comptable pour les couts XAF), meme convention que le classeur
# de reference - le libelle fusionne (Préfecture:Commune) y est CENTRE, contrairement aux lignes de
# donnees ou ces colonnes sont alignees a gauche.
wb_final <- wb_add_font(wb_final, "Coût médian du PMAS", dims = wb_dims(rows = national_row_r, cols = 1:15), name = "Roboto Condensed", size = 18, bold = TRUE, color = wb_color(hex = BODY_GREY))
wb_final <- wb_add_cell_style(wb_final, "Coût médian du PMAS", dims = wb_dims(rows = national_row_r, cols = c(1,2,3,7:14)), horizontal = "center", vertical = "center")
wb_final <- wb_add_cell_style(wb_final, "Coût médian du PMAS", dims = wb_dims(rows = national_row_r, cols = 4:6), horizontal = "center", vertical = "center")
wb_final <- wb_add_cell_style(wb_final, "Coût médian du PMAS", dims = wb_dims(rows = national_row_r, cols = 15), horizontal = "left", vertical = "center", wrap_text = "1")
wb_final <- wb_add_numfmt(wb_final, "Coût médian du PMAS", dims = wb_dims(rows = national_row_r, cols = 7:10),
                          numfmt = "_-* #,##0\\ [$XAF]_-;-* #,##0\\ [$XAF]_-;_-* \"-\"??\\ [$XAF]_-;_-@_-")
wb_final <- wb_add_numfmt(wb_final, "Coût médian du PMAS", dims = wb_dims(rows = national_row_r, cols = 11:14),
                          numfmt = "#,##0\\ \"€\";[Red]#,##0\\ \"€\"")
wb_final <- wb_add_border(wb_final, "Coût médian du PMAS", dims = wb_dims(rows = national_row_r, cols = 1:15),
                          top_border = "thick", bottom_border = "thick", left_border = "none", right_border = "none")

# colonne "Note" (15e, O) large pour que le texte y tienne sans faire deborder la hauteur de ligne ;
# hauteur de ligne NORMALE (uniforme) sur toutes les lignes de donnees (SAUF la ligne nationale, plus
# haute - voir ci-dessus) - sans cela, Excel recalcule une hauteur de ligne variable (parfois tres
# grande) selon la longueur du texte enveloppe dans la colonne Note, ce qui rend le tableau incoherent
# visuellement. Seules la ligne d'en-tete (1) et les lignes separatrices grises gardent leur propre
# hauteur (respectivement plus grande et plus petite).
wb_final <- wb_set_col_widths(wb_final, "Coût médian du PMAS", cols = 1:15,
                              widths = c(5, 11, 14, 10, 21, 20, 27, 23, 22, 23, 22, 26, 26, 30, 90))
data_rows <- setdiff(body_rows, sep_rows)
wb_final <- wb_set_row_heights(wb_final, "Coût médian du PMAS", rows = data_rows, heights = 20)
wb_final <- wb_set_row_heights(wb_final, "Coût médian du PMAS", rows = national_row_r, heights = 24)
wb_final <- wb_freeze_pane(wb_final, "Coût médian du PMAS", firstActiveRow = 2)
cat("  ..OK (Coût médian du PMAS : ", nrow(pmas_commune), " communes, ", nrow(pmas_sous_prefecture),
   " sous-prefectures, ", nrow(pmas_prefecture), " prefectures, 1 ligne nationale)\n", sep = "")

# --- Explication de la methodologie de suffisance des donnees (memes 4 regles, meme texte que
# l'info-box du tableau de bord - voir app.R, Étape 3 analyse) - affichee directement DANS le classeur
# pour que la methodologie soit lisible sans devoir rouvrir le tableau de bord. Constantes
# correspondantes : MEB_MIN_OBS_DEFAULT/MEB_MIN_OBS_OVERRIDE, MEB_TOTAL_MIN_ITEMS,
# MEB_BASKET_MIN_STANDALONE, MEB_BASKET_MIN_CAUTION (voir scripts/fonctions_analyse.R).
PMAS_EXPLANATION <- c(
  "1. Prix médian par article : à chaque niveau (Commune, Sous-Préfecture, Préfecture, National), le prix médian d'un article n'est calculé à partir des prix collectés localement que s'il y a au moins 3 prix disponibles pour cet article à ce niveau - sauf pour la viande de bœuf, où 1 seul prix suffit.",
  "2. PMAS total : calculé seulement si au moins 12 des 17 articles ont un prix médian local suffisant (règle 1). Les articles restants (jusqu'à 5) sont alors complétés par la médiane nationale de cet article, pour pouvoir quand même calculer un total complet.",
  "3. Moins de 12 articles disponibles localement : le PMAS total n'est PAS calculé, mais chaque panier est calculé séparément dès que son propre seuil est atteint - panier alimentaire si au moins 6 des 9 produits alimentaires sont disponibles, panier non alimentaire si au moins 4 des 6 produits sont disponibles, panier d'hygiène si les 2 produits sont disponibles (les articles manquants dans un panier calculé sont complétés par la médiane nationale). Une note l'indique dans la colonne \"Note\".",
  "4. Entre 12 et 14 articles disponibles : le PMAS total est calculé, mais si un panier a été calculé avec peu de prix locaux (moins de 6/9 pour l'alimentaire, 4/6 pour le non alimentaire, ou 1/2 pour l'hygiène), une note invite à interpréter ce panier avec prudence."
)
explain_title_row <- last_row + 3
explain_rows <- explain_title_row + seq_along(PMAS_EXPLANATION)

wb_final <- wb_add_data(wb_final, "Coût médian du PMAS", data.frame("Comment le PMAS est calculé"),
                        start_row = explain_title_row, start_col = 1, col_names = FALSE)
wb_final <- wb_add_font(wb_final, "Coût médian du PMAS", dims = wb_dims(rows = explain_title_row, cols = 1),
                        name = "Roboto Condensed", size = 13, bold = TRUE)
for (k in seq_along(PMAS_EXPLANATION)) {
  r <- explain_rows[k]
  wb_final <- wb_add_data(wb_final, "Coût médian du PMAS", data.frame(PMAS_EXPLANATION[k]),
                          start_row = r, start_col = 1, col_names = FALSE)
  wb_final <- wb_merge_cells(wb_final, "Coût médian du PMAS", dims = wb_dims(rows = r, cols = 1:15))
}
wb_final <- wb_add_font(wb_final, "Coût médian du PMAS", dims = wb_dims(rows = explain_rows, cols = 1), name = "Segoe UI", size = 11)
wb_final <- wb_add_cell_style(wb_final, "Coût médian du PMAS", dims = wb_dims(rows = explain_rows, cols = 1),
                              horizontal = "left", vertical = "center", wrap_text = "1")
wb_final <- wb_set_row_heights(wb_final, "Coût médian du PMAS", rows = explain_rows, heights = 30)
cat("  ..OK (Explication de la méthodologie de suffisance ajoutée)\n")

# --- "Composition du PMAS" -- DESACTIVE TEMPORAIREMENT (demande utilisateur : ne produire que
# "Coût médian du PMAS" + "Prix médian par article" pour l'instant). Decommenter pour la reactiver.
if (FALSE) {
cat("\n- CONSTRUCTION du tableau de reference 'Composition du PMAS' ...\n")
composition_pmas <- build_meb_composition_table()

wb_final <- wb_add_worksheet(wb_final, "Composition du PMAS")
wb_final <- wb_add_data(wb_final, "Composition du PMAS", composition_pmas)
wb_final <- wb_add_font(wb_final, "Composition du PMAS", dims = "A1:D1", bold = TRUE)
wb_final <- wb_add_fill(wb_final, "Composition du PMAS", dims = "A1:D1", color = wb_color(hex = "FFD9D9D9"))
wb_final <- wb_set_col_widths(wb_final, "Composition du PMAS", cols = 1:4, widths = c(26, 26, 34, 28))
wb_final <- wb_freeze_pane(wb_final, "Composition du PMAS", firstActiveRow = 2)

# fond de couleur legere par categorie (panier), pour les distinguer visuellement d'un coup d'oeil -
# chaque categorie forme un bloc de lignes contigu (voir build_meb_composition_table(), bind_rows() dans
# le meme ordre que les 3 categories ci-dessous)
cat_colors <- c("Panier alimentaire" = "FFFCE9D6", "Panier non alimentaire" = "FFDCE9F5",
               "Panier des produits d'hygiène" = "FFE1F0DC")
for (cat_name in names(cat_colors)) {
  rows <- which(composition_pmas$Catégorie == cat_name) + 1
  if (length(rows) > 0) {
    wb_final <- wb_add_fill(wb_final, "Composition du PMAS", dims = wb_dims(rows = rows, cols = 1:4),
                            color = wb_color(hex = cat_colors[[cat_name]]))
  }
}
cat("  ..OK (", nrow(composition_pmas), " articles)\n", sep = "")
}

# --- Evolution des paniers du PMAS -- DESACTIVE TEMPORAIREMENT (idem ci-dessus). Decommenter pour la
# reactiver (necessite aussi meb_food/meb_nfi/meb_wash/meb_total via calculate_meb(), ci-dessous).
if (FALSE) {
meb_food <- calculate_meb(clean_data, MEB_FOOD_WEIGHTS) %>% rename(meb_food = MEB)
meb_nfi <- calculate_meb(clean_data, MEB_NFI_WEIGHTS) %>% rename(meb_nfi = MEB)
meb_wash <- calculate_meb(clean_data, MEB_WASH_WEIGHTS) %>% rename(meb_wash = MEB)
meb_total <- calculate_meb(clean_data, MEB_WEIGHTS_ALL) %>% rename(meb_total = MEB)

if (nzchar(PREVIOUS_INPUT_PATH) && file.exists(PREVIOUS_INPUT_PATH)) {
  clean_data_prev <- read_excel(PREVIOUS_INPUT_PATH, sheet = "clean_data", skip = 1, col_types = "text")
  meb_food_prev <- calculate_meb(clean_data_prev, MEB_FOOD_WEIGHTS) %>% rename(prev_food = MEB)
  meb_nfi_prev <- calculate_meb(clean_data_prev, MEB_NFI_WEIGHTS) %>% rename(prev_nfi = MEB)
  meb_wash_prev <- calculate_meb(clean_data_prev, MEB_WASH_WEIGHTS) %>% rename(prev_wash = MEB)
  meb_total_prev <- calculate_meb(clean_data_prev, MEB_WEIGHTS_ALL) %>% rename(prev_total = MEB)

  pct_evol <- function(prev, cur) if_else(!is.na(prev) & !is.na(cur) & prev != 0, (cur - prev) / prev, NA_real_)

  evolution_base <- meb_food %>%
    full_join(meb_nfi, by = "admin3_label") %>%
    full_join(meb_wash, by = "admin3_label") %>%
    full_join(meb_total, by = "admin3_label") %>%
    filter(admin3_label != "National") %>%
    full_join(meb_food_prev, by = "admin3_label") %>%
    full_join(meb_nfi_prev, by = "admin3_label") %>%
    full_join(meb_wash_prev, by = "admin3_label") %>%
    full_join(meb_total_prev, by = "admin3_label") %>%
    left_join(admin_lookup, by = "admin3_label")

  evolution_pmas <- evolution_base %>%
    transmute(
      Mois = MOIS, Pays = PAYS,
      `Préfecture` = admin1_label, `Sous-Préfecture` = admin2_label, Commune = admin3_label,
      prev_food, meb_food, evol_food = pct_evol(prev_food, meb_food),
      prev_nfi, meb_nfi, evol_nfi = pct_evol(prev_nfi, meb_nfi),
      prev_wash, meb_wash, evol_wash = pct_evol(prev_wash, meb_wash),
      prev_total, meb_total, evol_total = pct_evol(prev_total, meb_total)
    ) %>%
    arrange(`Préfecture`, `Sous-Préfecture`, Commune)

  paniers <- c("Coût du panier alimentaire", "Coût du panier non-alimentaire",
              "Coût du panier des produits d'hygiène", "Coût total des paniers")
  header1 <- c("Mois", "Pays", "Préfecture", "Sous-Préfecture", "Commune", rep(paniers, each = 3))
  header2 <- c("", "", "", "", "", rep(c(if (nzchar(PREVIOUS_MOIS)) PREVIOUS_MOIS else "Mois précédent", MOIS, "Evolution"), times = 4))
  evol_display <- evolution_pmas %>%
    transmute(Mois, Pays, `Préfecture`, `Sous-Préfecture`, Commune,
             across(c(prev_food, meb_food, evol_food, prev_nfi, meb_nfi, evol_nfi,
                      prev_wash, meb_wash, evol_wash, prev_total, meb_total, evol_total),
                    ~if_else(is.na(.), "Non évalué", as.character(round(., 4)))))
  colnames(evol_display) <- paste0("c", seq_len(ncol(evol_display)))

  wb_final <- wb_add_worksheet(wb_final, "Evolution des paniers du PMAS")
  wb_final <- wb_add_data(wb_final, "Evolution des paniers du PMAS", t(header1), start_row = 1, col_names = FALSE)
  wb_final <- wb_add_data(wb_final, "Evolution des paniers du PMAS", t(header2), start_row = 2, col_names = FALSE)
  wb_final <- wb_add_data(wb_final, "Evolution des paniers du PMAS", evol_display, start_row = 3, col_names = FALSE)
  for (grp in 0:3) {
    start_col <- 6 + grp * 3
    wb_final <- wb_merge_cells(wb_final, "Evolution des paniers du PMAS",
                               dims = wb_dims(rows = 1, cols = start_col:(start_col + 1)))
  }
  wb_final <- wb_add_font(wb_final, "Evolution des paniers du PMAS", dims = "A1:R2", bold = TRUE)
  wb_final <- wb_freeze_pane(wb_final, "Evolution des paniers du PMAS", firstActiveRow = 3)
  cat("  ..OK (Evolution des paniers du PMAS :", nrow(evolution_pmas), "communes)\n")
} else {
  cat("  (pas de classeur du mois precedent fourni - feuille 'Evolution des paniers du PMAS' laissee vide)\n")
  wb_final <- wb_add_worksheet(wb_final, "Evolution des paniers du PMAS")
}
}

# --- Prix median par article ---
# libelles alignes sur le classeur de reference REACH_CAR_JMMI_Dataset_Analyse_Juin_Juillet_2026 (ex.
# "Seau en plastique" et non "Seau") - "Tomates" present dans ce classeur de reference n'existe plus
# comme variable dans l'outil kobo actuel (RCA_JMMI_2026-08_TOOL.xlsx), donc pas repris ici : voir le
# filtre item_var %in% colnames(clean_data) ci-dessous, qui retire deja automatiquement tout article
# absent de l'outil du mois courant.
PRICE_ITEM_CATALOG <- tribble(
  ~item_var, ~label_fr, ~categorie,
  "maize_price_item", "Maïs", "Produits alimentaires",
  "manioc_price_item", "Manioc", "Produits alimentaires",
  "rice_price_item", "Riz", "Produits alimentaires",
  "bean_price_item", "Haricot", "Produits alimentaires",
  "peanut_price_item", "Arachide", "Produits alimentaires",
  "beef_meat_price_item", "Viande de bœuf", "Produits alimentaires",
  "vegetable_oil_price_item", "Huile végétale", "Produits alimentaires",
  "sugar_price_item", "Sucre", "Produits alimentaires",
  "salt_price_item", "Sel", "Produits alimentaires",
  "mosquito_net_price_item", "Moustiquaire", "Produits non alimentaires",
  "jerrycan_price_item", "Bidon", "Produits non alimentaires",
  "bed_sheets_price_item", "Drap", "Produits non alimentaires",
  "carpet_price_item", "Natte", "Produits non alimentaires",
  "tarpaulin_price_item", "Bâche", "Produits non alimentaires",
  "cooking_pot_price_item", "Marmite", "Produits non alimentaires",
  "soap_price_item", "Savon", "Produits non alimentaires",
  "bucket_price_item", "Seau en plastique", "Produits non alimentaires",
  "cloth_price_item", "Pagne", "Produits supplémentaires",
  "bowl_price_item", "Cuvette métallique", "Produits supplémentaires",
  "kettle_price_item", "Théière / Bouta", "Produits supplémentaires",
  "firewood_price_item", "Bois de chauffage", "Produits supplémentaires",
  "petrol_price_item", "Essence", "Produits supplémentaires",
  "water_price_item", "Eau remplie dans un bidon", "Produits supplémentaires",
  "egg_price_item", "Œufs", "Produits supplémentaires"
)
# pour chaque article, prefere la variable "_price_unit_item" (prix ramene a une unite/quantite
# standardisee, ex. maize_350g_price_unit_item - la meme utilisee pour le calcul du PMAS, voir
# MEB_FOOD_WEIGHTS etc. dans scripts/fonctions_analyse.R) quand elle existe dans l'outil kobo ; sinon
# retombe sur la variable "_price_item" brute (prix pour la quantite telle que rapportee, non
# standardisee).
resolve_price_unit_item <- function(plain_var, cols) {
  base <- sub("_price_item$", "", plain_var)
  candidates <- grep(paste0("^", base, "_.*_price_unit_item$"), cols, value = TRUE)
  if (length(candidates) >= 1) candidates[1] else plain_var
}
PRICE_ITEM_CATALOG$item_var <- vapply(PRICE_ITEM_CATALOG$item_var, resolve_price_unit_item,
                                      character(1), cols = colnames(clean_data))
# ne garde que les articles reellement presents dans ce round (l'outil kobo peut ajouter/retirer des
# articles d'un mois a l'autre)
PRICE_ITEM_CATALOG <- PRICE_ITEM_CATALOG %>% filter(item_var %in% colnames(clean_data))
cat("  (", nrow(PRICE_ITEM_CATALOG), " articles de prix disponibles ce mois-ci)\n", sep = "")

# suffisance des donnees par article/commune (meme convention que le classeur de reference, voir sa
# legende "Moins de 3 cotations rapportées" / "Aucune cotation rapportée") : le prix median LOCAL n'est
# affiche que si au moins PRIX_ITEM_MIN_OBS prix ont ete collectes pour cet article dans cette commune ;
# sinon la mediane NATIONALE (tous marches confondus) est affichee a la place, avec un marquage visuel
# applique plus bas (fond rose = 1 ou 2 prix locaux, police rouge sans fond = aucun prix local).
PRIX_ITEM_MIN_OBS <- 3
price_items <- as.character(PRICE_ITEM_CATALOG$item_var)
n_items <- length(price_items)

item_data <- clean_data %>%
  select(admin1_label, admin2_label, admin3_label, all_of(price_items)) %>%
  mutate(across(all_of(price_items), as.numeric))

national_median_item <- item_data %>% summarise(across(all_of(price_items), ~median(., na.rm = TRUE)))
local_median_item <- item_data %>% group_by(admin1_label, admin2_label, admin3_label) %>%
  summarise(across(all_of(price_items), ~median(., na.rm = TRUE)), .groups = "drop") %>%
  arrange(admin1_label, admin2_label, admin3_label)
local_count_item <- item_data %>% group_by(admin1_label, admin2_label, admin3_label) %>%
  summarise(across(all_of(price_items), ~sum(!is.na(.))), .groups = "drop") %>%
  arrange(admin1_label, admin2_label, admin3_label)

prix_median <- local_median_item
for (it in price_items) {
  prix_median[[it]] <- if_else(local_count_item[[it]] >= PRIX_ITEM_MIN_OBS,
                               local_median_item[[it]], national_median_item[[it]][1])
}
colnames(prix_median)[-(1:3)] <- PRICE_ITEM_CATALOG$label_fr[match(colnames(prix_median)[-(1:3)], PRICE_ITEM_CATALOG$item_var)]
colnames(prix_median)[1:3] <- c("Préfecture", "Sous-Préfecture", "Commune")
n_communes <- nrow(prix_median)

# categorie ecrite UNE SEULE FOIS par bloc (premiere colonne du bloc), le reste vide, puis les cellules
# du bloc sont FUSIONNEES (voir plus bas) - delimitee en plus par une bordure MEDIUM a gauche du debut
# de chaque nouveau bloc (sauf le tout premier)
first_of_block <- !duplicated(PRICE_ITEM_CATALOG$categorie)
cat_header_items <- ifelse(first_of_block, PRICE_ITEM_CATALOG$categorie, "")
cat_header <- c("", "", "", cat_header_items)
item_header <- c("Préfecture", "Sous-Préfecture", "Commune", PRICE_ITEM_CATALOG$label_fr)

wb_final <- wb_add_worksheet(wb_final, "Prix médian par article")
wb_final <- wb_add_data(wb_final, "Prix médian par article", t(cat_header), start_row = 2, col_names = FALSE, na = "")
wb_final <- wb_add_data(wb_final, "Prix médian par article", t(item_header), start_row = 3, col_names = FALSE, na = "")
wb_final <- wb_add_data(wb_final, "Prix médian par article", prix_median, start_row = 4, col_names = FALSE, na = "")

# limites de chaque bloc de categorie contigu (ligne 2) - reutilisees plus bas pour la bordure MEDIUM de
# separation entre blocs (a gauche du premier bloc de chaque categorie, sauf la toute premiere qui
# commence au bord de la feuille ; a droite de la derniere colonne de chaque bloc, reprise sur l'entete
# ET toutes les lignes de donnees, meme convention que le classeur de reference)
cat_blocks <- rle(PRICE_ITEM_CATALOG$categorie)
block_end <- cumsum(cat_blocks$lengths)
block_start <- block_end - cat_blocks$lengths + 1
block_start_cols <- 3 + block_start
block_end_cols <- 3 + block_end

# fusionne chaque bloc de categorie (Produits alimentaires / non alimentaires / supplementaires) en une
# seule cellule (demande explicite de l'utilisateur - le classeur de reference original ne fusionne pas
# ces cellules, mais le rendu fusionne est prefere ici)
for (b in seq_along(cat_blocks$values)) {
  if (block_end_cols[b] > block_start_cols[b]) {
    wb_final <- wb_merge_cells(wb_final, "Prix médian par article", dims = wb_dims(rows = 2, cols = block_start_cols[b]:block_end_cols[b]))
  }
}

# ligne "Ensemble des marchés évalués" (national, calculee sur l'ensemble des prix collectes - jamais
# marquee rose/rouge puisqu'elle EST la reference nationale utilisee pour les substitutions ci-dessus)
national_row_r <- 4 + n_communes
# ecrit le libelle (texte) et les valeurs (numeriques) separement - les combiner dans un seul c() les
# forcerait TOUS en texte (coercion de type habituelle de c()), ce qui empecherait le format numerique
# XAF applique plus bas de s'afficher (un format numerique n'a aucun effet sur une cellule texte)
wb_final <- wb_add_data(wb_final, "Prix médian par article", data.frame("Ensemble des marchés évalués"),
                        start_row = national_row_r, start_col = 1, col_names = FALSE)
wb_final <- wb_add_data(wb_final, "Prix médian par article", national_median_item,
                        start_row = national_row_r, start_col = 4, col_names = FALSE, na = "")
wb_final <- wb_merge_cells(wb_final, "Prix médian par article", dims = wb_dims(rows = national_row_r, cols = 1:3))

# legende (identique au classeur de reference, meme texte et memes styles - voir plus bas)
legend_r1 <- national_row_r + 4
legend_r2 <- legend_r1 + 1
wb_final <- wb_add_data(wb_final, "Prix médian par article", data.frame("*"), start_row = legend_r1, start_col = 1, col_names = FALSE)
wb_final <- wb_add_data(wb_final, "Prix médian par article",
                        data.frame("Moins de 3 cotations rapportées -  la médiane utilisée est celle de l'ensemble des marchés évalués"),
                        start_row = legend_r1, start_col = 2, col_names = FALSE)
wb_final <- wb_add_data(wb_final, "Prix médian par article", data.frame("*"), start_row = legend_r2, start_col = 1, col_names = FALSE)
wb_final <- wb_add_data(wb_final, "Prix médian par article",
                        data.frame("Aucune cotation rapportée -  la médiane utilisée est celle de l'ensemble des marchés évalués"),
                        start_row = legend_r2, start_col = 2, col_names = FALSE)

################ mise en forme (reprise a l'identique du classeur de reference, styles extraits du XML OOXML) ################
last_col <- 3 + n_items
# gris fonce (texte du corps/national) et bleu pale (fond des en-tetes d'articles) : couleurs extraites
# du theme du classeur de reference (theme "1" tinte +0.35 -> #595959 ; theme "8"/accent5 #4BACC6 tinte
# +0.8 -> #DBEEF4), pas de simple approximation a l'oeil
GREY_TEXT <- "FF595959"
LIGHT_BLUE_FILL <- "FFDBEEF4"

# categorie (ligne 2) : police Roboto Condensed 16 gras rouge (EF5958) sur la cellule de libelle
# (cellules du bloc fusionnees ci-dessus), alignee a GAUCHE avec un retrait, bordure MEDIUM a gauche du
# debut de chaque nouveau bloc (sauf le tout premier)
wb_final <- wb_add_font(wb_final, "Prix médian par article", dims = wb_dims(rows = 2, cols = 4:last_col),
                        name = "Roboto Condensed", size = 16, bold = TRUE, color = wb_color(hex = "FFEF5958"))
wb_final <- wb_add_cell_style(wb_final, "Prix médian par article", dims = wb_dims(rows = 2, cols = 4:last_col),
                              horizontal = "left", vertical = "center", indent = 1)
# wb_add_border() applique a un dims MULTI-CELLULE (colonnes non contigues separees par des virgules)
# perd silencieusement les cotes demandes sur la TOUTE PREMIERE cellule de la liste (bug constate
# empiriquement, teste isolement) - boucle donc sur chaque colonne separement (1 appel par colonne, dims
# alors toujours une plage a une seule colonne, contigue si plusieurs lignes) plutot qu'un seul appel
# multi-cellule, ici et pour tous les separateurs de bloc ci-dessous.
if (length(block_start_cols) > 1) {
  for (col in block_start_cols[-1]) {
    wb_final <- wb_add_border(wb_final, "Prix médian par article", dims = wb_dims(rows = 2, cols = col),
                              left_border = "medium", top_border = "none", right_border = "none", bottom_border = "none")
  }
}
wb_final <- wb_set_row_heights(wb_final, "Prix médian par article", rows = 2, heights = 21)

# en-tete colonnes (ligne 3) : DEUX tons, comme le classeur de reference - colonnes admin (Préfecture/
# Sous-Préfecture/Commune) en fond corail + texte blanc ; colonnes d'articles en fond bleu pale + texte
# gris fonce
wb_final <- wb_add_fill(wb_final, "Prix médian par article", dims = wb_dims(rows = 3, cols = 1:3), color = wb_color(hex = "FFEE5958"))
wb_final <- wb_add_font(wb_final, "Prix médian par article", dims = wb_dims(rows = 3, cols = 1:3), name = "Roboto Condensed", size = 13, bold = TRUE, color = wb_color(hex = "FFFFFFFF"))
wb_final <- wb_add_fill(wb_final, "Prix médian par article", dims = wb_dims(rows = 3, cols = 4:last_col), color = wb_color(hex = LIGHT_BLUE_FILL))
wb_final <- wb_add_font(wb_final, "Prix médian par article", dims = wb_dims(rows = 3, cols = 4:last_col), name = "Roboto Condensed", size = 13, bold = TRUE, color = wb_color(hex = GREY_TEXT))
wb_final <- wb_add_cell_style(wb_final, "Prix médian par article", dims = wb_dims(rows = 3, cols = 1:last_col), horizontal = "center", vertical = "center", wrap_text = "1")
for (col in block_end_cols) {
  wb_final <- wb_add_border(wb_final, "Prix médian par article", dims = wb_dims(rows = 3, cols = col),
                            right_border = "medium", left_border = "none", top_border = "none", bottom_border = "none")
}
wb_final <- wb_set_row_heights(wb_final, "Prix médian par article", rows = 3, heights = 38)

# corps des donnees (lignes 4 a la ligne nationale) : police Segoe UI 12 gris fonce, format numerique
# avec separateur de milliers, centre pour les prix / gauche pour les libelles admin, meme separateur
# vertical MEDIUM entre blocs de categorie que l'entete ci-dessus
wb_final <- wb_add_font(wb_final, "Prix médian par article", dims = wb_dims(rows = 4:national_row_r, cols = 1:last_col), name = "Segoe UI", size = 12, color = wb_color(hex = GREY_TEXT))
wb_final <- wb_add_cell_style(wb_final, "Prix médian par article", dims = wb_dims(rows = 4:(national_row_r - 1), cols = 1:3), horizontal = "left", vertical = "center")
wb_final <- wb_add_cell_style(wb_final, "Prix médian par article", dims = wb_dims(rows = 4:(national_row_r - 1), cols = 4:last_col), horizontal = "center", vertical = "center")
wb_final <- wb_add_numfmt(wb_final, "Prix médian par article", dims = wb_dims(rows = 4:(national_row_r - 1), cols = 4:last_col), numfmt = "#,##0;[Red]#,##0")
for (col in block_end_cols) {
  wb_final <- wb_add_border(wb_final, "Prix médian par article", dims = wb_dims(rows = 4:(national_row_r - 1), cols = col),
                            right_border = "medium", left_border = "none", top_border = "none", bottom_border = "none")
}
wb_final <- wb_set_row_heights(wb_final, "Prix médian par article", rows = 4:(national_row_r - 1), heights = 22)

# ligne nationale : police Roboto Condensed 14 gras (comme le classeur de reference, pas Segoe UI),
# format XAF, libelle centre, bordure MEDIUM en haut ET en bas pour la delimiter clairement des communes
# individuelles (meme convention que les separateurs entre blocs de "Coût médian du PMAS")
wb_final <- wb_add_font(wb_final, "Prix médian par article", dims = wb_dims(rows = national_row_r, cols = 1:last_col), name = "Roboto Condensed", size = 14, bold = TRUE, color = wb_color(hex = GREY_TEXT))
wb_final <- wb_add_cell_style(wb_final, "Prix médian par article", dims = wb_dims(rows = national_row_r, cols = 1:3), horizontal = "center", vertical = "center")
wb_final <- wb_add_cell_style(wb_final, "Prix médian par article", dims = wb_dims(rows = national_row_r, cols = 4:last_col), horizontal = "center", vertical = "center")
wb_final <- wb_add_numfmt(wb_final, "Prix médian par article", dims = wb_dims(rows = national_row_r, cols = 4:last_col), numfmt = "#,##0\\ [$XAF];[Red]#,##0\\ [$XAF]")
wb_final <- wb_add_border(wb_final, "Prix médian par article", dims = wb_dims(rows = national_row_r, cols = 1:last_col),
                          top_border = "medium", bottom_border = "medium", left_border = "none", right_border = "none")
for (col in block_end_cols) {
  wb_final <- wb_add_border(wb_final, "Prix médian par article", dims = wb_dims(rows = national_row_r, cols = col),
                            top_border = "medium", bottom_border = "medium", right_border = "medium", left_border = "none")
}
wb_final <- wb_set_row_heights(wb_final, "Prix médian par article", rows = national_row_r, heights = 29.5)

# marquage rose ("moins de 3 cotations") / rouge ("aucune cotation") des cellules substituees par la
# mediane nationale - regroupe les cellules a marquer par style pour appliquer wb_add_fill()/
# wb_add_font() en un minimum d'appels (dims accepte une liste de cellules separees par des virgules)
pink_cells <- character(0)
red_cells <- character(0)
for (j in seq_along(price_items)) {
  it <- price_items[j]
  col_num <- 3 + j
  counts <- local_count_item[[it]]
  for (i in seq_len(n_communes)) {
    if (counts[i] == 0) {
      red_cells <- c(red_cells, wb_dims(rows = 3 + i, cols = col_num))
    } else if (counts[i] < PRIX_ITEM_MIN_OBS) {
      pink_cells <- c(pink_cells, wb_dims(rows = 3 + i, cols = col_num))
    }
  }
}
if (length(pink_cells) > 0) {
  pink_dims <- paste(pink_cells, collapse = ",")
  wb_final <- wb_add_fill(wb_final, "Prix médian par article", dims = pink_dims, color = wb_color(hex = "FFFFC7CE"))
  wb_final <- wb_add_font(wb_final, "Prix médian par article", dims = pink_dims, name = "Segoe UI", size = 12, color = wb_color(hex = "FFFF0000"))
}
if (length(red_cells) > 0) {
  red_dims <- paste(red_cells, collapse = ",")
  wb_final <- wb_add_font(wb_final, "Prix médian par article", dims = red_dims, name = "Segoe UI", size = 12, color = wb_color(hex = "FFFF0000"))
}
cat("  (marquage donnees insuffisantes : ", length(pink_cells), " cellules < 3 cotations, ",
   length(red_cells), " cellules sans aucune cotation)\n", sep = "")

# legende : memes styles que les cellules marquees ci-dessus (le "*" reprend exactement le style d'une
# cellule reelle marquee), texte explicatif en italique
wb_final <- wb_add_fill(wb_final, "Prix médian par article", dims = wb_dims(rows = legend_r1, cols = 1), color = wb_color(hex = "FFFFC7CE"))
wb_final <- wb_add_font(wb_final, "Prix médian par article", dims = wb_dims(rows = legend_r1, cols = 1), name = "Segoe UI", size = 12, color = wb_color(hex = "FFFF0000"))
wb_final <- wb_add_cell_style(wb_final, "Prix médian par article", dims = wb_dims(rows = legend_r1, cols = 1), horizontal = "center", vertical = "center")
wb_final <- wb_add_font(wb_final, "Prix médian par article", dims = wb_dims(rows = legend_r2, cols = 1), name = "Segoe UI", size = 12, color = wb_color(hex = "FFFF0000"))
wb_final <- wb_add_cell_style(wb_final, "Prix médian par article", dims = wb_dims(rows = legend_r2, cols = 1), horizontal = "center", vertical = "center")
wb_final <- wb_add_font(wb_final, "Prix médian par article", dims = wb_dims(rows = c(legend_r1, legend_r2), cols = 2), name = "Segoe UI", size = 12, italic = TRUE)

wb_final <- wb_set_col_widths(wb_final, "Prix médian par article", cols = 1:3, widths = c(19.5, 19.5, 27.7))
wb_final <- wb_set_col_widths(wb_final, "Prix médian par article", cols = 4:last_col, widths = rep(13, n_items))
wb_final <- wb_freeze_pane(wb_final, "Prix médian par article", firstActiveRow = 4, firstActiveCol = 4)
cat("  ..OK (Prix médian par article :", n_communes, "communes)\n")

################################################################################
# "SCORE DE FONCTIONNALITE" (MFS - voir scripts/mfs.R)
################################################################################
cat("\n- CALCUL du Score de Fonctionnalité des Marchés (SFM/MFS) ...\n")
mfs_raw <- calculate_mfs(clean_data)

score_fonctionnalite <- mfs_raw %>%
  ungroup() %>%  # calculate_mfs() renvoie un tibble encore groupe (summarise() ne retire qu'un niveau de
                 # regroupement a la fois) - sans ungroup(), transmute() garderait admin1_label/admin2_label
                 # en plus des colonnes explicitement listees ci-dessous
  transmute(
    `Année` = ANNEE, Mois = MOIS, Pays = PAYS,
    `Préfecture` = admin1_label, `Sous-préfecture` = admin2_label, Commune = admin3_label,
    `Accessibilité des marchés` = round(mfs_accessibility_pillar_score / 100, 2),
    `Disponibilité des produits` = round(mfs_availability_pillar_score / 100, 2),
    `Abordabilité des prix des produits` = round(mfs_affordability_pillar_score / 100, 2),
    `Résilience des circuits d'approvisionnement` = round(mfs_resilience_pillar_score / 100, 2),
    `Niveau des infrastructures du marché` = round(mfs_infrastructure_pillar_score / 100, 2),
    `Total Score SFM` = round(mfs_score / 100, 2),
    !!paste0("Classification SFM ", MOIS, " ", ANNEE) := functionality
  ) %>%
  arrange(`Préfecture`, `Sous-préfecture`, Commune)

wb_final <- wb_add_worksheet(wb_final, "Score de Fonctionnalité")
wb_final <- wb_add_data(wb_final, "Score de Fonctionnalité", score_fonctionnalite, na = "")

# mise en forme reprise du classeur de reference REACH_CAR_JMMI_Dataset_Analyse_Juin_Juillet_2026 (memes
# constantes que "Coût médian du PMAS" ci-dessus : en-tete fond rouge corail FFEE5958 + police Roboto
# Condensed 14 gras texte blanc ; corps police Segoe UI 12 BODY_GREY, colonnes de score au format
# pourcentage (numFmt Excel integre id 9 - "0%", puisque mfs_*_pillar_score/100 est deja une fraction
# 0-1) ; "Total Score SFM" en gras 14pt (comme les colonnes "Total" de PMAS). Pas de ligne/colonne
# d'espacement au debut (contrairement au classeur de reference) : convention deja etablie par "Coût
# médian du PMAS" dans ce classeur, gardee identique ici pour la coherence entre les deux feuilles. Pas
# non plus des 3 niveaux d'agregation Sous-Prefecture/Prefecture/National du classeur de reference
# (calculate_mfs() ne groupe qu'au niveau Commune) - demande utilisateur portant sur la mise en forme,
# pas sur le calcul ; a envisager separement si ces niveaux sont souhaites.
n_mfs <- nrow(score_fonctionnalite)
last_row_mfs <- n_mfs + 1
classif_col_name <- paste0("Classification SFM ", MOIS, " ", ANNEE)
classif_col <- which(colnames(score_fonctionnalite) == classif_col_name)
full_dims_mfs <- wb_dims(rows = 1:last_row_mfs, cols = 1:13)

wb_final <- wb_add_border(wb_final, "Score de Fonctionnalité", dims = full_dims_mfs, inner_hgrid = "thin", inner_vgrid = "thin")
wb_final <- wb_add_border(wb_final, "Score de Fonctionnalité", dims = "A1:M1",
                          left_border = "thin", left_color = wb_color(hex = "FFFFFFFF"),
                          right_border = "thin", right_color = wb_color(hex = "FFFFFFFF"),
                          top_border = "thin", top_color = wb_color(hex = "FFFFFFFF"),
                          bottom_border = "thin", bottom_color = wb_color(hex = "FF000000"))
wb_final <- wb_add_fill(wb_final, "Score de Fonctionnalité", dims = "A1:M1", color = wb_color(hex = "FFEE5958"))
wb_final <- wb_add_font(wb_final, "Score de Fonctionnalité", dims = "A1:M1", name = "Roboto Condensed", size = 14, bold = TRUE, color = wb_color(hex = "FFFFFFFF"))
wb_final <- wb_add_cell_style(wb_final, "Score de Fonctionnalité", dims = "A1:M1", horizontal = "center", vertical = "center", wrap_text = "1")
wb_final <- wb_set_row_heights(wb_final, "Score de Fonctionnalité", rows = 1, heights = 62)

body_rows_mfs <- 2:last_row_mfs
wb_final <- wb_add_font(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = body_rows_mfs, cols = 1:13), name = "Segoe UI", size = 12, color = wb_color(hex = BODY_GREY))
wb_final <- wb_add_font(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = body_rows_mfs, cols = 12), name = "Segoe UI", size = 14, bold = TRUE, color = wb_color(hex = BODY_GREY))
wb_final <- wb_add_font(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = body_rows_mfs, cols = classif_col), name = "Segoe UI", size = 14, bold = TRUE, color = wb_color(hex = BODY_GREY))
wb_final <- wb_add_cell_style(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = body_rows_mfs, cols = c(1,2,3,7:12)), horizontal = "center", vertical = "center")
wb_final <- wb_add_cell_style(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = body_rows_mfs, cols = 4:6), horizontal = "left", vertical = "center", indent = 1)
wb_final <- wb_add_cell_style(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = body_rows_mfs, cols = classif_col), horizontal = "left", vertical = "center", wrap_text = "1")
wb_final <- wb_add_numfmt(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = body_rows_mfs, cols = 7:12), numfmt = "0%")

# code couleur de la classification (fond plein par categorie, meme palette que le classeur de reference
# - "Donnees insuffisantes" n'a pas de fond, seulement un texte rouge, contrairement aux 4 autres) -
# applique APRES le style de corps general ci-dessus pour le remplacer sur cette seule colonne/ces
# seules lignes.
MFS_CATEGORY_STYLES <- list(
  "Problèmes graves" = list(fill = "FFFF0000"),
  "Mauvaise fonctionnalité" = list(fill = "FFF79646"),
  "Fonctionnalité limitée" = list(fill = "FFFFFFBF"),
  "Fonctionnalité complète" = list(fill = "FF249EA0"),
  "Données insuffisantes" = list(font_color = "FFFF0000")
)
classif_values <- score_fonctionnalite[[classif_col_name]]
for (cat_name in names(MFS_CATEGORY_STYLES)) {
  cat_rows <- which(classif_values == cat_name) + 1  # +1 pour la ligne d'en-tete
  if (length(cat_rows) == 0) next
  cat_style <- MFS_CATEGORY_STYLES[[cat_name]]
  if (!is.null(cat_style$fill)) {
    wb_final <- wb_add_fill(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = cat_rows, cols = classif_col), color = wb_color(hex = cat_style$fill))
  }
  if (!is.null(cat_style$font_color)) {
    wb_final <- wb_add_font(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = cat_rows, cols = classif_col), name = "Segoe UI", size = 14, bold = TRUE, color = wb_color(hex = cat_style$font_color))
  }
}

wb_final <- wb_set_col_widths(wb_final, "Score de Fonctionnalité", cols = 1:13,
                              widths = c(10, 15, 10, 19, 20, 25, 19, 18, 21, 26, 25, 18, 36))
wb_final <- wb_set_row_heights(wb_final, "Score de Fonctionnalité", rows = body_rows_mfs, heights = 20)
wb_final <- wb_freeze_pane(wb_final, "Score de Fonctionnalité", firstActiveRow = 2)
cat("  ..OK (", nrow(score_fonctionnalite), " communes, ",
   sum(is.na(score_fonctionnalite$`Total Score SFM`)), " avec donnees insuffisantes)\n", sep = "")

################################################################################
# "ANALYSE" (LOA descriptive) -- DESACTIVEE TEMPORAIREMENT (demande utilisateur : cette section est
# lente a executer). Decommenter le bloc if (FALSE) ci-dessous pour la reactiver.
################################################################################
if (FALSE) {
cat("\n- ANALYSE DESCRIPTIVE (feuille 'Analyse') ...\n")
my_loa <- build_loa(kobo_survey, group_var = "admin3_label")
missing_vars <- setdiff(my_loa$analysis_var, colnames(clean_data))
if (length(missing_vars) > 0) {
  cat("  (", length(missing_vars), " variable(s) de l'outil absente(s) de clean_data, ignorees)\n", sep = "")
  my_loa <- my_loa %>% filter(analysis_var %in% colnames(clean_data))
}
cat("  LOA generee :", nrow(my_loa), "lignes (", n_distinct(my_loa$analysis_var), "variables )\n")

my_design <- srvyr::as_survey_design(clean_data)
my_results <- create_analysis(my_design, loa = my_loa, sm_separator = "/")
gxv <- my_results$results_table %>% create_table_group_x_variable(value_columns = c("stat"))

analyse_tmp_path <- tempfile(fileext = ".xlsx")
create_xlsx_group_x_variable(gxv, table_sheet = "Analyse", file_path = analyse_tmp_path,
                             write_file = TRUE, overwrite = TRUE)
wb_analyse <- wb_load(analyse_tmp_path)
wb_final <- wb_clone_worksheet(wb_final, old = "Analyse", new = "Analyse", from = wb_analyse)
unlink(analyse_tmp_path)
cat("  ..OK (", ncol(gxv) - 1, " variables x ", nrow(gxv), " communes)\n", sep = "")
}

################################################################################
# ORDRE FINAL DES FEUILLES ET SAUVEGARDE
################################################################################
desired_order <- c("READ_ME", "Lisez_moi", "CHECKLIST_Data_Cleaning", "Données nettoyées", "Coût médian du PMAS",
                   "Composition du PMAS", "Evolution des paniers du PMAS", "Prix médian par article",
                   "Score de Fonctionnalité", "Analyse", "survey", "choices", "raw_data", "clean_data",
                   "cleaning_log", "deletion_log", "variable_tracker")
current_sheets <- wb_get_sheet_names(wb_final)
final_order <- intersect(desired_order, current_sheets)
final_order <- c(final_order, setdiff(current_sheets, final_order))  # garde-fou : n'oublie aucune feuille
wb_final <- wb_set_order(wb_final, match(final_order, current_sheets))

output_dir <- Sys.getenv("output_dir")
if (!dir.exists(output_dir)) dir.create(output_dir, recursive = TRUE)
mois_slug <- if (nzchar(MOIS)) MOIS else "mois"
out_path <- file.path(output_dir, paste0(Sys.Date(), "_CAR_JMMI_Dataset_Analyse_", mois_slug, ".xlsx"))
wb_save(wb_final, out_path, overwrite = TRUE)
cat("\n--> Classeur d'analyse complet exporte :", out_path, "\n")
