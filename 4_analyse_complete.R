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
#               ATTENTION : malgre son nom (conserve pour compatibilite avec les variables d'env
#               existantes), "analyse_previous_clean_data_file" attend un classeur d'ANALYSE JMMI du
#               mois precedent (sortie de ce meme script - lu via sa feuille "Coût médian du PMAS"),
#               PAS un classeur "Finale" de nettoyage/clean_data - voir read_prev_pmas() plus bas.
# SORTIE      : Sys.getenv("output_dir")/REACH_CAR_Analyse_JMMI_<Mois>-<Annee>.xlsx - reprend toutes les
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
source("scripts/format_results.R", local = TRUE)
# relabel_choices()/relabel_multiple_choice_subcolumn_labels() - re-etiquetage automatique des codes de
# choix dans "Données nettoyées" (voir plus bas) - demande utilisateur explicite : "the etiquetage de
# choix does not deserve a step, it needs to be automatically included in the analysis calculations"
# (etait auparavant une etape separee du tableau de bord - voir app.R, mode retire).
source("scripts/fonctions_relabel.R", local = TRUE)

################################################################################
# PARAMETRES (variables d'environnement - voir run_analyse_pipeline() dans app.R)
################################################################################
INPUT_PATH <- Sys.getenv("clean_data_file")
KOBO_TOOL_FILE <- Sys.getenv("kobo_tool_file")
PAYS <- "RCA"
PAYS_ISO3 <- "CAR"  # demande utilisateur explicite (pas le code ISO 3166-1 alpha-3 "CAF" - convention deja
                    # utilisee partout ailleurs dans ce projet, ex. noms de fichiers "REACH_CAR_...") -
                    # feuilles "meb_analysis"/"mfs_analysis" (colonne "country"), modele externe REACH.
EUR_XRATE <- as.numeric(Sys.getenv("eur_xrate"))
USD_XRATE <- as.numeric(Sys.getenv("usd_xrate"))
PREVIOUS_INPUT_PATH <- Sys.getenv("analyse_previous_clean_data_file")   # "" si non fourni
PREVIOUS_MOIS <- Sys.getenv("analyse_mois_precedent")

if (!nzchar(INPUT_PATH) || !file.exists(INPUT_PATH)) {
  stop("Classeur d'entree introuvable : le classeur final nettoye n'a pas ete correctement importe.")
}
if (!nzchar(KOBO_TOOL_FILE) || !file.exists(KOBO_TOOL_FILE)) {
  stop("Outil kobo introuvable : l'outil kobo n'a pas ete correctement importe.")
}
if (is.na(EUR_XRATE)) stop("Taux de change XAF/EUR invalide.")
if (is.na(USD_XRATE)) stop("Taux de change XAF/USD invalide.")

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
  standardise_choices_label() %>%  # "label" ou "label::french" (voir scripts/fonctions_analyse.R)
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

# garde-fou : le classeur importe doit etre la sortie "Finale" du NETTOYAGE, pas un classeur d'ANALYSE
# deja produit (par erreur de reimport, ex. l'utilisateur reimporte le fichier telecharge la fois
# precedente). Sans ce controle, wb_add_worksheet() ci-dessous ne detecte PAS le nom de feuille deja
# present (aucune erreur a ce stade) et les wb_merge_cells() suivants finissent par entrer en collision
# avec les fusions DEJA presentes dans l'ancienne feuille de meme nom, avec un message cryptique
# ("Merge intersects with existing merged cells") qui ne dit rien sur la vraie cause - repere/reproduit
# en importatant un classeur d'analyse deja produit comme classeur "Finale".
ANALYSIS_SHEET_NAMES <- c("Coût médian du PMAS", "Composition du PMAS", "Evolution des paniers du PMAS",
                          "Prix médian par article", "Score de Fonctionnalité", "Checklist_Data_Analysis",
                          "meb_composition", "Données nettoyées", "meb_analysis", "mfs_analysis")
already_analysed_sheets <- intersect(ANALYSIS_SHEET_NAMES, cleaning_sheet_names)
if (length(already_analysed_sheets) > 0) {
  stop("Le classeur importé contient déjà des feuilles d'analyse (", paste(already_analysed_sheets, collapse = ", "),
      ") : il s'agit probablement d'un classeur d'ANALYSE déjà produit, pas du classeur \"Finale\" issu du ",
      "nettoyage. Réimportez le classeur produit par l'étape \"Finale\" du nettoyage (pas un fichier déjà ",
      "téléchargé depuis cette étape \"Analyse\").")
}

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
# "DONNEES NETTOYEES" -- copie de la feuille "clean_data" du classeur de nettoyage importe (meme mise en
# forme, memes colonnes, meme ordre - demande utilisateur : "just copy the clean_data sheet in all its
# formatting, dont change it"), avec (1) les colonnes de prix individuels/taux de change masquees (deja
# affichees ailleurs - "Prix médian par article" pour les prix, taux de change non pertinents pour une
# lecture generale des donnees) et (2) les CODES des questions a choix (select_one/select_multiple)
# remplaces par leur LIBELLE lisible - demande utilisateur explicite : "the etiquetage de choix does not
# deserve a step, it needs to be automatically included in the analysis calculations ... include this
# step in the analysis script" (etait auparavant une etape manuelle separee du tableau de bord -
# "Étiquetage des choix" / 3_relabel_choices.R -, desormais retiree, voir app.R). Seules les VALEURS sont
# modifiees (jamais les noms de colonnes/variables, ni clean_data/raw_data qui restent au format
# technique brut necessaire aux calculs de ce script et a la tracabilite) - voir
# scripts/fonctions_relabel.R (fonctions generiques, deja utilisees et testees pour ce meme usage).
################################################################################
cat("\n- COPIE de la feuille \"clean_data\" (Données nettoyées) ...\n")
wb_final <- wb_clone_worksheet(wb_final, old = "clean_data", new = "Données nettoyées", from = wb_final)

cat("  ..re-etiquetage automatique des codes de choix ...\n")
label_row_clean_data <- suppressMessages(
  read_excel(INPUT_PATH, sheet = "clean_data", range = readxl::cell_rows(1), col_names = FALSE, col_types = "text")
)
label_row_clean_data <- as.character(unlist(label_row_clean_data[1, ], use.names = FALSE))
relabelled_nettoyees <- relabel_choices(clean_data, kobo_survey, kobo_choices, verbose = FALSE)
new_label_row_nettoyees <- relabel_multiple_choice_subcolumn_labels(
  colnames(clean_data), label_row_clean_data, kobo_survey, kobo_choices
)
# na = "" : de nombreuses cellules de clean_data sont vides par conception (questions non applicables/
# sautees selon la logique de l'outil kobo - "relevant") - sans ce parametre, wb_add_data() les affiche
# comme une erreur Excel "#N/A" litterale plutot que des cellules vides (comportement par defaut, constate
# empiriquement plus tot dans cette session), ce qui rend "Données nettoyées" illisible - demande
# utilisateur explicite suite a ce constat.
wb_final <- wb_add_data(wb_final, "Données nettoyées", t(new_label_row_nettoyees), start_row = 1, col_names = FALSE, na = "")
wb_final <- wb_add_data(wb_final, "Données nettoyées", relabelled_nettoyees, start_row = 2, col_names = TRUE, na = "")

# colonnes masquees (PAS supprimees - openxlsx2 ne permet pas de decaler des colonnes sans risquer de
# corrompre fusions/styles/formules de la feuille clonee) : prix individuels par article
# (*_price_item), taux de change EUR (*_eur_xrate_official) et USD (*_usd_xrate_official). Les noms de
# colonnes technique proviennent de la ligne 2 du classeur ("clean_data" a ete lu avec skip = 1 - voir
# plus haut), qui correspond donc 1:1 aux colonnes reelles de la feuille (meme ordre, aucune colonne
# ajoutee/retiree par read_excel()).
hide_cols_nettoyees <- which(grepl("_price_item|_eur_xrate|usd_xrate", colnames(clean_data)))
if (length(hide_cols_nettoyees) > 0) {
  wb_final <- wb_set_col_widths(wb_final, "Données nettoyées", cols = hide_cols_nettoyees, hidden = TRUE)
}
cat("  ..OK (", nrow(clean_data), " lignes, ", ncol(clean_data), " colonnes, ", length(hide_cols_nettoyees),
   " colonnes masquees)\n", sep = "")

################################################################################
# PMAS / MEB : "Coût médian du PMAS", "Evolution des paniers du PMAS", "Prix médian par article"
################################################################################
cat("\n- CALCUL du PMAS (panier minimum) ...\n")

admin_lookup <- clean_data %>%
  select(admin1_label, admin2_label, admin3_label) %>%
  distinct(admin3_label, .keep_all = TRUE)

# --- "Coût médian du PMAS" : niveaux Commune / National uniquement (demande utilisateur - Sous-Préfecture
# et Préfecture retires, meme structure a 2 niveaux que "Score de Fonctionnalité"), 14 colonnes de couts +
# 1 colonne "Note", un bloc par niveau separe par une ligne vide. Chaque niveau applique la methodologie
# de suffisance des donnees (voir calculate_pmas_by_level()/calculate_pmas_national() dans
# scripts/fonctions_analyse.R) : panier total calcule seulement si >=12/17 articles ont assez de prix
# locaux a ce niveau, sinon chaque panier (alimentaire/non alimentaire/hygiene) est calcule
# independamment s'il atteint son propre seuil, avec une note explicative.
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

# articles PMAS/MEB attendus mais absents des donnees ce cycle-ci (variable renommee/retiree de l'outil
# kobo) - calcule ICI (avant le premier appel a calculate_pmas_by_level(), qui l'affiche aussi en
# cat() console) pour etre reutilise a la fois par la regle 8 de PMAS_EXPLANATION ci-dessous ET par la
# feuille "meb_composition" plus loin, sans dupliquer le calcul.
missing_meb_items <- check_meb_items_missing(clean_data)

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

# niveau National - une seule ligne, "Ensemble des marchés évalués" (meme libelle que le classeur de
# reference) dans la colonne Préfecture
p_nat <- calculate_pmas_national(clean_data)
pmas_national <- tibble(
  `Année` = ANNEE, Mois = MOIS, Pays = PAYS,
  `Préfecture` = "Ensemble des marchés évalués", `Sous-Préfecture` = NA_character_, Commune = NA_character_,
  !!!PMAS_COST_COLS(p_nat)
)

# ecriture : en-tete une seule fois, puis les 2 blocs (Commune / National) empiles avec une ligne vide
# entre les deux - na="" pour que les cellules NA (Sous-Prefecture/Commune absents au niveau national,
# paniers/Note non calcules) restent VRAIMENT vides plutot que d'ecrire une erreur Excel #N/A
# (comportement par defaut de wb_add_data())
wb_final <- wb_add_worksheet(wb_final, "Coût médian du PMAS")
wb_final <- wb_add_data(wb_final, "Coût médian du PMAS", t(colnames(pmas_commune)), start_row = 1, col_names = FALSE, na = "")

row_cursor <- 2
wb_final <- wb_add_data(wb_final, "Coût médian du PMAS", pmas_commune, start_row = row_cursor, col_names = FALSE, na = "")
sep_row_1 <- row_cursor + nrow(pmas_commune)
row_cursor <- sep_row_1 + 1
wb_final <- wb_add_data(wb_final, "Coût médian du PMAS", pmas_national, start_row = row_cursor, col_names = FALSE, na = "")
# fusionne Préfecture/Sous-Préfecture/Commune sur la ligne nationale (seule colonne D remplie -
# "Ensemble des marchés évalués" - E et F vides) pour que le libelle s'affiche sur les 3 colonnes
wb_final <- wb_merge_cells(wb_final, "Coût médian du PMAS", dims = wb_dims(rows = row_cursor, cols = 4:6))

# ligne separatrice GRISE (fine) entre les 2 blocs (Commune / National) pour delimiter clairement les
# sections - la ligne vide entre les deux blocs est remplie en gris et sa hauteur reduite pour ressembler
# a une fine ligne de separation plutot qu'a une ligne de donnees vide
sep_rows <- c(sep_row_1)
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
cat("  ..OK (Coût médian du PMAS : ", nrow(pmas_commune), " communes, 1 ligne nationale)\n", sep = "")

# --- Explication de la methodologie (7 regles - regle 2 reecrite pour decrire la mediane des PANIERS
# communaux, PAS une mediane de prix par article recomposee au niveau national, demande utilisateur
# explicite ; regle 7 ajoutee sur le nombre de communes contribuant a la mediane nationale - meme texte
# que l'info-box du tableau de bord - voir app.R, Étape 3 analyse) - affichee directement DANS le classeur
# pour que la methodologie soit lisible sans devoir rouvrir le tableau de bord. TOUS les nombres cites
# (17 articles, 12 requis, seuils par panier, ...) sont calcules ICI a partir des constantes de
# scripts/fonctions_analyse.R (MEB_MIN_OBS_DEFAULT/MEB_MIN_OBS_OVERRIDE, MEB_TOTAL_MIN_ITEMS,
# MEB_TOTAL_CAUTION_MAX_ITEMS, MEB_BASKET_MIN_STANDALONE, MEB_BASKET_MIN_CAUTION, MEB_*_WEIGHTS) plutot
# que ecrits en dur dans le texte - demande utilisateur, suite a l'audit de robustesse : ces nombres
# etaient auparavant recopies a la main a plusieurs endroits (ici, app.R, et les commentaires de
# fonctions_analyse.R), sans aucune garantie de rester synchronises si un article ou un seuil change.
n_meb_total <- length(MEB_WEIGHTS_ALL)
n_food_items <- length(MEB_FOOD_WEIGHTS)
n_nfi_items <- length(MEB_NFI_WEIGHTS)
n_wash_items <- length(MEB_WASH_WEIGHTS)
n_extra_allowed <- n_meb_total - MEB_TOTAL_MIN_ITEMS
override_labels <- unname(MEB_ITEM_LABELS[names(MEB_MIN_OBS_OVERRIDE)])
override_text <- if (length(MEB_MIN_OBS_OVERRIDE) > 0) {
  paste0(" - sauf pour ", paste(paste0(override_labels, " (seuil réduit à ", MEB_MIN_OBS_OVERRIDE, ")"), collapse = ", "), ".")
} else ""

PMAS_EXPLANATION <- c(
  "1. Marché = Commune : chaque commune évaluée correspond à un marché (plusieurs vendeurs y sont enquêtés). Le prix médian de chaque article y est calculé directement à partir des prix individuels collectés auprès de ses vendeurs (voir la feuille \"Prix médian par article\") ; chaque commune compose ensuite SON PROPRE panier (alimentaire/non alimentaire/hygiène/total) à partir de ces prix médians et des poids du PMAS (voir la feuille \"meb_composition\").",
  "2. National = médiane des paniers communaux déjà calculés : le panier national N'EST PAS recomposé à partir de prix par article agrégés au niveau national, mais correspond à la médiane des paniers (alimentaire, non alimentaire, hygiène, total) déjà calculés pour chaque commune (règle 1) - chaque commune compte pour une seule observation, et chaque panier est médianné indépendamment des trois autres.",
  paste0("3. Suffisance des données (au niveau d'une commune) : un article est considéré \"mesuré localement\" dans un marché s'il y a au moins ", MEB_MIN_OBS_DEFAULT, " prix locaux disponibles pour cet article dans ce marché", override_text, " Sinon, la médiane nationale des médianes de marché (voir feuille \"Prix médian par article\") est utilisée comme valeur de secours pour cet article dans ce marché, afin que la commune puisse quand même composer son propre panier."),
  paste0("4. PMAS total (d'une commune) : calculé seulement si au moins ", MEB_TOTAL_MIN_ITEMS, " des ", n_meb_total, " articles y sont mesurés localement (règle 3). Les articles restants (jusqu'à ", n_extra_allowed, ") sont alors complétés par leur valeur de secours, pour pouvoir quand même calculer un total complet."),
  paste0("5. Moins de ", MEB_TOTAL_MIN_ITEMS, " articles disponibles localement (dans une commune) : son PMAS total n'est PAS calculé, mais chaque panier est calculé séparément dès que son propre seuil est atteint - panier alimentaire si au moins ", MEB_BASKET_MIN_STANDALONE[["food"]], " des ", n_food_items, " produits alimentaires sont disponibles, panier non alimentaire si au moins ", MEB_BASKET_MIN_STANDALONE[["nfi"]], " des ", n_nfi_items, " produits sont disponibles, panier d'hygiène si au moins ", MEB_BASKET_MIN_STANDALONE[["wash"]], " des ", n_wash_items, " produits sont disponibles (les articles manquants dans un panier calculé sont complétés par leur valeur de secours). Une note l'indique dans la colonne \"Note\"."),
  paste0("6. Entre ", MEB_TOTAL_MIN_ITEMS, " et ", MEB_TOTAL_CAUTION_MAX_ITEMS, " articles disponibles (dans une commune) : son PMAS total est calculé, mais si un panier a été calculé avec peu d'articles mesurés localement (moins de ", MEB_BASKET_MIN_CAUTION[["food"]], "/", n_food_items, " pour l'alimentaire, ", MEB_BASKET_MIN_CAUTION[["nfi"]], "/", n_nfi_items, " pour le non alimentaire, ou ", MEB_BASKET_MIN_CAUTION[["wash"]], "/", n_wash_items, " pour l'hygiène), une note invite à interpréter ce panier avec prudence."),
  "7. Ligne nationale calculée sur moins de communes : si un panier n'a pas pu être calculé pour certaines communes (règle 5), la médiane nationale de ce panier (règle 2) est calculée uniquement sur les communes où il a pu l'être - une note l'indique dans la colonne \"Note\" de la ligne \"Ensemble des marchés évalués\"."
)
# regle 8 : n'apparait QUE si des articles PMAS/MEB attendus sont absents des donnees ce cycle-ci
# (missing_meb_items, calcule plus haut avec check_meb_items_missing()) - typiquement une variable
# renommee/retiree de l'outil kobo. AVANT cette regle, ce cas passait totalement inapercu sur cette
# feuille : le panier se recalculait avec un article en moins, sans aucune indication ici que quelque
# chose manquait (demande utilisateur explicite, suite a l'audit de robustesse - voir aussi la feuille
# "meb_composition", ou l'article manquant est marque en rouge).
has_missing_meb_items <- length(missing_meb_items) > 0
if (has_missing_meb_items) {
  PMAS_EXPLANATION <- c(PMAS_EXPLANATION, paste0(
    "8. ATTENTION - article(s) manquant(s) ce cycle-ci : ", paste(missing_meb_items, collapse = ", "),
    " ne correspond(ent) à aucune colonne des données importées (variable probablement renommée ou retirée de l'outil kobo depuis le dernier cycle) et sont donc totalement EXCLU(S) du calcul de tous les paniers ci-dessus, à tous les niveaux - voir la feuille \"meb_composition\" pour le détail (article marqué en rouge)."
  ))
}
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
if (has_missing_meb_items) {
  # regle 8 (la derniere) mise en evidence - fond rose/police rouge gras, memes couleurs que le
  # marquage "aucune cotation" de "Prix médian par article"/"meb_composition"
  warn_row <- explain_rows[length(explain_rows)]
  wb_final <- wb_add_fill(wb_final, "Coût médian du PMAS", dims = wb_dims(rows = warn_row, cols = 1:15), color = wb_color(hex = "FFFFC7CE"))
  wb_final <- wb_add_font(wb_final, "Coût médian du PMAS", dims = wb_dims(rows = warn_row, cols = 1), name = "Segoe UI", size = 11, bold = TRUE, color = wb_color(hex = "FFFF0000"))
  wb_final <- wb_set_row_heights(wb_final, "Coût médian du PMAS", rows = warn_row, heights = 45)
}
cat("  ..OK (Explication de la méthodologie de suffisance ajoutée)\n")

# --- Evolution des paniers du PMAS : niveaux Commune + National uniquement (demande utilisateur
# explicite : "only keep the commune and national level. This is all that will [be] available from now
# on" - meme simplification a 2 niveaux que "Coût médian du PMAS"/"Score de Fonctionnalité"). Reactivee
# (etait desactivee plus tot dans la session) et reconstruite pour utiliser
# calculate_pmas_by_level()/calculate_pmas_national() - MEME methodologie mediane-des-medianes que "Coût
# médian du PMAS" (y compris le garde-fou "article manquant", check_meb_items_missing()) - plutot que
# l'ancien calculate_meb()/calculate_meb_by_level() (seuil de suffisance different, pas de mediane des
# medianes, pas d'alerte article manquant) utilise avant sa desactivation : reprend p_commune/p_nat DEJA
# calcules plus haut pour le cycle COURANT (memes valeurs qu'affichees sur "Coût médian du PMAS"), et
# applique exactement le meme calcul au cycle PRECEDENT si fourni. Mise en forme reprise A L'IDENTIQUE de
# la version d'origine (demande utilisateur : "exactly as is") - 2 lignes d'en-tete, 3 colonnes par
# panier (mois precedent / mois courant / evolution), pas de couleur - avec juste la ligne nationale en
# plus, au meme format que les lignes communales (pas de mise en evidence particuliere, coherent avec le
# style minimal d'origine de cette feuille).
# lit les couts PMAS DEJA CALCULES depuis la feuille "Coût médian du PMAS" d'un classeur d'ANALYSE
# precedent (PAS un classeur "Finale" de nettoyage/clean_data) - demande utilisateur explicite : "you
# shouldn't need those columns [admin1_label/admin2_label/admin3_label de clean_data], just the data in
# the evolution des paniers du pmas for the [cycle] ... aka june 2026". Plus robuste dans le temps que de
# relire des donnees BRUTES et TOUT recalculer : la feuille "Coût médian du PMAS" a un format de sortie
# STABLE (le notre, invariable d'un cycle a l'autre), independant de la structure de clean_data qui a
# elle deja change plusieurs fois par le passe (ex. le classeur de reference de juin 2026 utilise encore
# l'ancien "admin1"/"admin2"/"admin3" sur une seule ligne d'en-tete, incompatible avec le format actuel
# "admin1_label"/etc. sur 2 lignes) - lire le RESULTAT deja produit par ce meme pipeline evite cette
# fragilite entierement, plutot que de la contourner cas par cas. Ne garde que les lignes ou "Commune"
# est renseignee (exclut d'eventuelles lignes Sous-Prefecture/Prefecture d'un ancien classeur a 4
# niveaux, sans effet sur un classeur au format actuel a 2 niveaux qui n'en a de toute facon pas) ; la
# ligne nationale est identifiee par son libelle "Ensemble des marchés évalués" dans la colonne
# "Préfecture" (secours sur la mediane des communes lues si elle est introuvable).
read_prev_pmas <- function(path) {
  peek <- read_excel(path, sheet = "Coût médian du PMAS", col_names = FALSE, n_max = 5, col_types = "text")
  hdr_row <- 1
  for (r in seq_len(nrow(peek))) {
    if ("Préfecture" %in% as.character(unlist(peek[r, ]))) { hdr_row <- r; break }
  }
  raw <- read_excel(path, sheet = "Coût médian du PMAS", skip = hdr_row - 1, col_types = "text")
  # normalise les en-tetes (retours a la ligne/espaces multiples a l'interieur d'un nom de colonne) -
  # constate empiriquement sur le classeur de reference de juin 2026, dont l'en-tete "Coût total des
  # paniers" contient un retour a la ligne litteral ("Coût total des\npaniers", pour un retour visuel a
  # l'affichage a l'epoque) qui faisait echouer la correspondance exacte de nom de colonne ci-dessous.
  colnames(raw) <- trimws(gsub("\\s+", " ", colnames(raw)))
  cost_cols <- c("Coût du panier alimentaire", "Coût du panier non alimentaire",
                "Coût du panier des produits d'hygiène", "Coût total des paniers")
  missing_cols <- setdiff(c("Préfecture", "Commune", cost_cols), colnames(raw))
  if (length(missing_cols) > 0) {
    stop("la feuille 'Coût médian du PMAS' de ce classeur ne contient pas les colonnes attendues (",
        paste(missing_cols, collapse = ", "), ") - ce n'est probablement pas un classeur d'ANALYSE ",
        "JMMI (sortie de 4_analyse_complete.R).")
  }
  raw <- raw %>% mutate(across(all_of(cost_cols), as.numeric))

  # distinct(admin3_label) : un ancien classeur a 4 niveaux (Commune/Sous-Prefecture/Prefecture/
  # National, avant leur retrait de ce pipeline) peut avoir sa ligne "Sous-Prefecture" repetant a tort
  # le meme nom de Commune que la ligne Commune correspondante (constate empiriquement sur le classeur
  # de reference de juin 2026 : sa colonne "Commune" n'est blanchie qu'au niveau Prefecture, pas au
  # niveau Sous-Prefecture) - garder seulement la 1ere occurrence de chaque commune evite de dupliquer
  # ces lignes dans la comparaison (la 1ere occurrence est toujours la vraie ligne Commune, puisque ce
  # niveau est toujours ecrit en premier dans le classeur).
  commune <- raw %>% filter(!is.na(Commune)) %>%
    transmute(admin3_label = Commune, meb_food = `Coût du panier alimentaire`,
             meb_nfi = `Coût du panier non alimentaire`,
             meb_wash = `Coût du panier des produits d'hygiène`,
             meb_total = `Coût total des paniers`) %>%
    distinct(admin3_label, .keep_all = TRUE)
  national <- raw %>% filter(`Préfecture` == "Ensemble des marchés évalués") %>%
    transmute(meb_food = `Coût du panier alimentaire`, meb_nfi = `Coût du panier non alimentaire`,
             meb_wash = `Coût du panier des produits d'hygiène`, meb_total = `Coût total des paniers`)
  if (nrow(national) == 0) {
    national <- commune %>%
      summarise(meb_food = median(meb_food, na.rm = TRUE), meb_nfi = median(meb_nfi, na.rm = TRUE),
               meb_wash = median(meb_wash, na.rm = TRUE), meb_total = median(meb_total, na.rm = TRUE))
  }
  list(commune = commune, national = national[1, ])
}

if (nzchar(PREVIOUS_INPUT_PATH) && file.exists(PREVIOUS_INPUT_PATH)) {
  # tout le calcul du cycle precedent est protege par tryCatch() : un classeur precedent illisible ou
  # structurellement incompatible ne doit PAS faire echouer tout le pipeline apres ~9 minutes de calcul
  # par ailleurs reussi sur le cycle courant (constate empiriquement avec un ancien classeur "Finale" -
  # crash cryptique "Column admin3_label doesn't exist" au lieu d'un message clair) - en cas d'echec, la
  # feuille est simplement laissee vide (avec un message explicatif), comme si aucun classeur precedent
  # n'avait ete fourni.
  evolution_result <- tryCatch({
    prev_pmas <- read_prev_pmas(PREVIOUS_INPUT_PATH)
    p_commune_prev <- prev_pmas$commune
    p_nat_prev <- prev_pmas$national

    pct_evol <- function(prev, cur) if_else(!is.na(prev) & !is.na(cur) & prev != 0, (cur - prev) / prev, NA_real_)

    # niveau Commune : jointure a GAUCHE depuis p_commune (cycle COURANT) - demande utilisateur explicite :
    # "only do the comparison and show for the markets for which we collected data in the latest round, do
    # not show all other markets" - ne garde donc que les communes evaluees CE cycle-ci ; une commune non
    # evaluee au cycle precedent (pas de correspondance a droite) affiche simplement des cellules
    # prev_*/evol_* vides plutot que d'etre exclue. p_commune porte deja admin1_label/admin2_label (joints
    # plus haut), pas besoin de les rejoindre ici.
    evolution_commune <- p_commune %>%
      left_join(p_commune_prev %>% select(admin3_label, prev_food = meb_food, prev_nfi = meb_nfi,
                                          prev_wash = meb_wash, prev_total = meb_total),
               by = "admin3_label") %>%
      transmute(
        Mois = MOIS, Pays = PAYS,
        `Préfecture` = admin1_label, `Sous-Préfecture` = admin2_label, Commune = admin3_label,
        prev_food, meb_food, evol_food = pct_evol(prev_food, meb_food),
        prev_nfi, meb_nfi, evol_nfi = pct_evol(prev_nfi, meb_nfi),
        prev_wash, meb_wash, evol_wash = pct_evol(prev_wash, meb_wash),
        prev_total, meb_total, evol_total = pct_evol(prev_total, meb_total)
      ) %>%
      arrange(`Préfecture`, `Sous-Préfecture`, Commune)

    # niveau National : 1 seule ligne - mediane des paniers communaux deja calcules (p_nat/p_nat_prev),
    # meme convention et meme libelle que "Coût médian du PMAS".
    evolution_national <- tibble(
      Mois = MOIS, Pays = PAYS,
      `Préfecture` = "Ensemble des marchés évalués", `Sous-Préfecture` = NA_character_, Commune = NA_character_,
      prev_food = p_nat_prev$meb_food, meb_food = p_nat$meb_food,
      evol_food = pct_evol(p_nat_prev$meb_food, p_nat$meb_food),
      prev_nfi = p_nat_prev$meb_nfi, meb_nfi = p_nat$meb_nfi,
      evol_nfi = pct_evol(p_nat_prev$meb_nfi, p_nat$meb_nfi),
      prev_wash = p_nat_prev$meb_wash, meb_wash = p_nat$meb_wash,
      evol_wash = pct_evol(p_nat_prev$meb_wash, p_nat$meb_wash),
      prev_total = p_nat_prev$meb_total, meb_total = p_nat$meb_total,
      evol_total = pct_evol(p_nat_prev$meb_total, p_nat$meb_total)
    )

    evolution_pmas <- bind_rows(evolution_commune, evolution_national)

    # colonnes GAP (NA) inserees entre chaque groupe de panier pour reproduire l'espacement du classeur
    # de reference REACH_CAR_JMMI_Dataset_Analyse_Juin_Juillet_2026 (colonnes I/M/Q vides entre les
    # groupes alimentaire/non-alimentaire/hygiene/total) - demande utilisateur explicite : "formatted
    # exactly like the april july analysis (...) same colors, cell formatting and spaces". Valeurs
    # numeriques REELLES (pas de texte "Non évalué" comme avant) pour que les formats numeriques Excel
    # (devise XAF, %) s'appliquent correctement - une commune sans donnee au cycle precedent affiche
    # simplement des cellules vides (na = "" plus bas), comme dans le classeur de reference.
    evol_display <- evolution_pmas %>%
      transmute(
        c1 = Mois, c2 = Pays, c3 = `Préfecture`, c4 = `Sous-Préfecture`, c5 = Commune,
        c6 = prev_food, c7 = meb_food, c8 = evol_food, c9 = NA_real_,
        c10 = prev_nfi, c11 = meb_nfi, c12 = evol_nfi, c13 = NA_real_,
        c14 = prev_wash, c15 = meb_wash, c16 = evol_wash, c17 = NA_real_,
        c18 = prev_total, c19 = meb_total, c20 = evol_total
      )

    prev_label <- if (nzchar(PREVIOUS_MOIS)) PREVIOUS_MOIS else "Mois précédent"
    paniers <- c("Coût du panier alimentaire", "Coût du panier non-alimentaire",
                "Coût du panier des produits d'hygiène", "Coût total des paniers")
    header1 <- rep("", 20)
    header1[1:5] <- c("Mois", "Pays", "Préfecture", "Sous-Préfecture", "Commune")
    header2 <- rep("", 20)
    for (g in seq_along(paniers)) {
      start_col <- 6 + (g - 1) * 4
      header1[start_col] <- paniers[g]
      header2[start_col:(start_col + 2)] <- c(prev_label, MOIS, "Evolution")
    }

    list(ok = TRUE, header1 = header1, header2 = header2, evol_display = evol_display,
        n_communes = nrow(evolution_commune))
  }, error = function(e) list(ok = FALSE, message = conditionMessage(e)))

  wb_final <- wb_add_worksheet(wb_final, "Evolution des paniers du PMAS")
  if (evolution_result$ok) {
    # mise en forme reprise A L'IDENTIQUE du classeur de reference
    # REACH_CAR_JMMI_Dataset_Analyse_Juin_Juillet_2026 (styles extraits directement du XML OOXML), et
    # alignee sur les memes constantes que "Coût médian du PMAS" plus haut (FFEE5958/Roboto
    # Condensed/Segoe UI/BODY_GREY) - demande utilisateur explicite : "formatted exactly like the april
    # july analysis, with the same colors, cell formatting and spaces". Ligne 1 vide (marge visuelle,
    # comme le classeur de reference), ligne 2 = en-tete de groupe (fond rouge corail, fusionne sur 3
    # colonnes par panier), ligne 3 = sous-en-tete de cycle (mois precedent / mois courant / evolution),
    # donnees a partir de la ligne 4, colonnes I/M/Q vides entre les groupes (memes largeurs que le
    # classeur de reference).
    n_data <- nrow(evolution_result$evol_display)
    last_row <- 3 + n_data
    national_row_r <- last_row
    id_block <- 1:5
    val_blocks <- list(6:8, 10:12, 14:16, 18:20)
    all_blocks <- c(list(id_block), val_blocks)

    wb_final <- wb_add_data(wb_final, "Evolution des paniers du PMAS", t(evolution_result$header1), start_row = 2, col_names = FALSE)
    wb_final <- wb_add_data(wb_final, "Evolution des paniers du PMAS", t(evolution_result$header2), start_row = 3, col_names = FALSE)
    # na = "" : cellules NA (comparaisons impossibles - commune non evaluee au cycle precedent,
    # Sous-Préfecture/Commune absents sur la ligne nationale) laissees VRAIMENT vides plutot qu'une
    # erreur Excel "#N/A" litterale (comportement par defaut de wb_add_data(), constate empiriquement).
    wb_final <- wb_add_data(wb_final, "Evolution des paniers du PMAS", evolution_result$evol_display, start_row = 4, col_names = FALSE, na = "")

    for (blk in val_blocks) {
      wb_final <- wb_merge_cells(wb_final, "Evolution des paniers du PMAS", dims = wb_dims(rows = 2, cols = blk))
    }

    # ligne 2 (en-tete de groupe) : fond rouge corail + police blanche (ID = taille 12, paniers = taille
    # 14, comme le classeur de reference)
    wb_final <- wb_add_fill(wb_final, "Evolution des paniers du PMAS", dims = wb_dims(rows = 2, cols = id_block), color = wb_color(hex = "FFEE5958"))
    wb_final <- wb_add_font(wb_final, "Evolution des paniers du PMAS", dims = wb_dims(rows = 2, cols = id_block), name = "Roboto Condensed", size = 12, bold = TRUE, color = wb_color(hex = "FFFFFFFF"))
    for (blk in val_blocks) {
      wb_final <- wb_add_fill(wb_final, "Evolution des paniers du PMAS", dims = wb_dims(rows = 2, cols = blk), color = wb_color(hex = "FFEE5958"))
      wb_final <- wb_add_font(wb_final, "Evolution des paniers du PMAS", dims = wb_dims(rows = 2, cols = blk), name = "Roboto Condensed", size = 14, bold = TRUE, color = wb_color(hex = "FFFFFFFF"))
    }
    wb_final <- wb_add_cell_style(wb_final, "Evolution des paniers du PMAS", dims = wb_dims(rows = 2, cols = unlist(all_blocks)), horizontal = "center", vertical = "center", wrap_text = "1")
    for (blk in all_blocks) {
      wb_final <- wb_add_border(wb_final, "Evolution des paniers du PMAS", dims = wb_dims(rows = 2, cols = blk),
                                top_border = "thin", top_color = wb_color(hex = "FFFFFFFF"),
                                left_border = "thin", left_color = wb_color(hex = "FFFFFFFF"),
                                right_border = "thin", right_color = wb_color(hex = "FFFFFFFF"),
                                bottom_border = "thin", bottom_color = wb_color(hex = "FF000000"))
    }

    # ligne 3 (sous-en-tete de cycle) + corps (lignes 4:last_row) : grille fine noire par bloc (colonnes
    # I/M/Q laissees entierement vides - ni bordure ni fond, comme le classeur de reference)
    for (blk in all_blocks) {
      wb_final <- wb_add_border(wb_final, "Evolution des paniers du PMAS", dims = wb_dims(rows = 3:last_row, cols = blk), inner_hgrid = "thin", inner_vgrid = "thin")
    }
    wb_final <- wb_add_font(wb_final, "Evolution des paniers du PMAS", dims = wb_dims(rows = 3, cols = unlist(val_blocks)), name = "Segoe UI", size = 12, bold = TRUE, color = wb_color(hex = BODY_GREY))
    wb_final <- wb_add_cell_style(wb_final, "Evolution des paniers du PMAS", dims = wb_dims(rows = 3, cols = unlist(all_blocks)), horizontal = "center", vertical = "center")

    body_rows <- setdiff(4:last_row, national_row_r)
    wb_final <- wb_add_font(wb_final, "Evolution des paniers du PMAS", dims = wb_dims(rows = body_rows, cols = c(id_block, unlist(val_blocks))), name = "Segoe UI", size = 12, color = wb_color(hex = BODY_GREY))
    wb_final <- wb_add_font(wb_final, "Evolution des paniers du PMAS", dims = wb_dims(rows = body_rows, cols = c(5, 18, 19, 20)), name = "Segoe UI", size = 12, bold = TRUE, color = wb_color(hex = BODY_GREY))
    wb_final <- wb_add_cell_style(wb_final, "Evolution des paniers du PMAS", dims = wb_dims(rows = body_rows, cols = c(1, 2, unlist(val_blocks))), horizontal = "center", vertical = "center")
    wb_final <- wb_add_cell_style(wb_final, "Evolution des paniers du PMAS", dims = wb_dims(rows = body_rows, cols = 3:5), horizontal = "left", vertical = "center", indent = 1)
    wb_final <- wb_add_numfmt(wb_final, "Evolution des paniers du PMAS", dims = wb_dims(rows = body_rows, cols = c(6, 7, 10, 11, 14, 15, 18, 19)), numfmt = "#,##0\\ [$XAF];[Red]#,##0\\ [$XAF]")
    wb_final <- wb_add_numfmt(wb_final, "Evolution des paniers du PMAS", dims = wb_dims(rows = body_rows, cols = c(8, 12, 16, 20)), numfmt = "0%")

    # ligne nationale : reprend EXACTEMENT le style de la ligne nationale de "Coût médian du PMAS"
    # (police Roboto Condensed 18 gras, bordure EPAISSE noire en haut ET en bas sur toute la largeur, y
    # compris les colonnes vides I/M/Q pour un trait continu - update = TRUE pour FUSIONNER cette
    # bordure epaisse avec la grille fine deja posee plus haut, sans quoi wb_add_border() REMPLACE tout
    # le style de bordure existant sur les cellules touchees, y compris les bords gauche/droit deja
    # corrects, comme constate empiriquement plus tot dans cette session) - le libelle fusionne
    # (Préfecture:Commune) est CENTRE, contrairement aux lignes de donnees ou ces colonnes sont alignees
    # a gauche.
    wb_final <- wb_merge_cells(wb_final, "Evolution des paniers du PMAS", dims = wb_dims(rows = national_row_r, cols = 3:5))
    wb_final <- wb_add_font(wb_final, "Evolution des paniers du PMAS", dims = wb_dims(rows = national_row_r, cols = 1:20), name = "Roboto Condensed", size = 18, bold = TRUE, color = wb_color(hex = BODY_GREY))
    wb_final <- wb_add_cell_style(wb_final, "Evolution des paniers du PMAS", dims = wb_dims(rows = national_row_r, cols = c(1, 2, unlist(val_blocks))), horizontal = "center", vertical = "center")
    wb_final <- wb_add_cell_style(wb_final, "Evolution des paniers du PMAS", dims = wb_dims(rows = national_row_r, cols = 3:5), horizontal = "center", vertical = "center")
    wb_final <- wb_add_numfmt(wb_final, "Evolution des paniers du PMAS", dims = wb_dims(rows = national_row_r, cols = c(6, 7, 10, 11, 14, 15, 18, 19)),
                              numfmt = "_-* #,##0\\ [$XAF]_-;-* #,##0\\ [$XAF]_-;_-* \"-\"??\\ [$XAF]_-;_-@_-")
    wb_final <- wb_add_numfmt(wb_final, "Evolution des paniers du PMAS", dims = wb_dims(rows = national_row_r, cols = c(8, 12, 16, 20)), numfmt = "0%")
    wb_final <- wb_add_border(wb_final, "Evolution des paniers du PMAS", dims = wb_dims(rows = national_row_r, cols = 1:20),
                              top_border = "thick", bottom_border = "thick", left_border = "none", right_border = "none", update = TRUE)

    # cellules "non evaluable" (comparaison impossible pour cette commune/ce panier) : reprises A
    # L'IDENTIQUE du classeur de reference REACH_CAR_JMMI_Dataset_Analyse_Juin_Juillet_2026 (styles
    # extraits directement du XML OOXML) - demande utilisateur explicite : "put a cell like this where
    # the data is not available and N/A for the evolution column when we cannot compare the 2 months,
    # essentially the same formatting as in the previous juin juillet file". Colonne mois precedent/mois
    # courant NA -> texte "Non évalué" (fond gris FFD9D9D9, police italique) ; colonne Evolution NA ->
    # texte "N/A" (police grasse, aucun fond) - bordures/alignement/format numerique deja poses par les
    # blocs ci-dessus restent inchanges (Excel les ignore simplement pour du texte). Applique EN DERNIER
    # pour remplacer, cellule par cellule, le style "normal" (corps ou ligne nationale) deja pose plus
    # haut - l'ordre des appels determine quel style "gagne" sur une meme cellule.
    for (blk in val_blocks) {
      prev_col <- blk[1]; cur_col <- blk[2]; evol_col <- blk[3]
      prev_na_rows <- 3 + which(is.na(evolution_result$evol_display[[paste0("c", prev_col)]]))
      cur_na_rows  <- 3 + which(is.na(evolution_result$evol_display[[paste0("c", cur_col)]]))
      evol_na_rows <- 3 + which(is.na(evolution_result$evol_display[[paste0("c", evol_col)]]))
      for (r in prev_na_rows) {
        wb_final <- wb_add_data(wb_final, "Evolution des paniers du PMAS", "Non évalué", start_row = r, start_col = prev_col, col_names = FALSE)
        wb_final <- wb_add_fill(wb_final, "Evolution des paniers du PMAS", dims = wb_dims(rows = r, cols = prev_col), color = wb_color(hex = "FFD9D9D9"))
        wb_final <- wb_add_font(wb_final, "Evolution des paniers du PMAS", dims = wb_dims(rows = r, cols = prev_col), name = "Segoe UI", size = 12, italic = TRUE, color = wb_color(hex = BODY_GREY))
      }
      for (r in cur_na_rows) {
        wb_final <- wb_add_data(wb_final, "Evolution des paniers du PMAS", "Non évalué", start_row = r, start_col = cur_col, col_names = FALSE)
        wb_final <- wb_add_fill(wb_final, "Evolution des paniers du PMAS", dims = wb_dims(rows = r, cols = cur_col), color = wb_color(hex = "FFD9D9D9"))
        wb_final <- wb_add_font(wb_final, "Evolution des paniers du PMAS", dims = wb_dims(rows = r, cols = cur_col), name = "Segoe UI", size = 12, italic = TRUE, color = wb_color(hex = BODY_GREY))
      }
      for (r in evol_na_rows) {
        wb_final <- wb_add_data(wb_final, "Evolution des paniers du PMAS", "N/A", start_row = r, start_col = evol_col, col_names = FALSE)
        wb_final <- wb_add_font(wb_final, "Evolution des paniers du PMAS", dims = wb_dims(rows = r, cols = evol_col), name = "Segoe UI", size = 12, bold = TRUE, color = wb_color(hex = BODY_GREY))
      }
    }

    wb_final <- wb_set_col_widths(wb_final, "Evolution des paniers du PMAS", cols = 1:20,
                                  widths = c(11, 9, 21, 21, 28, 20, 20, 14, 4, 18, 18, 15, 4, 19, 18, 15, 4, 18, 18, 15))
    wb_final <- wb_set_row_heights(wb_final, "Evolution des paniers du PMAS", rows = 2, heights = 40)
    wb_final <- wb_set_row_heights(wb_final, "Evolution des paniers du PMAS", rows = 3, heights = 22)
    wb_final <- wb_set_row_heights(wb_final, "Evolution des paniers du PMAS", rows = body_rows, heights = 20)
    wb_final <- wb_set_row_heights(wb_final, "Evolution des paniers du PMAS", rows = national_row_r, heights = 26)
    wb_final <- wb_freeze_pane(wb_final, "Evolution des paniers du PMAS", firstActiveRow = 4, firstActiveCol = 6)
    cat("  ..OK (Evolution des paniers du PMAS :", evolution_result$n_communes, "communes, 1 ligne nationale)\n")
  } else {
    wb_final <- wb_add_data(wb_final, "Evolution des paniers du PMAS",
                            data.frame(paste0("Classeur du mois précédent illisible ou incompatible - feuille laissée vide. Détail : ", evolution_result$message)),
                            start_row = 1, col_names = FALSE)
    cat("  ATTENTION : classeur du mois precedent illisible/incompatible (", evolution_result$message,
       ") - feuille 'Evolution des paniers du PMAS' laissee (quasi) vide.\n", sep = "")
  }
} else {
  cat("  (pas de classeur du mois precedent fourni - feuille 'Evolution des paniers du PMAS' laissee vide)\n")
  wb_final <- wb_add_worksheet(wb_final, "Evolution des paniers du PMAS")
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
# sinon la mediane NATIONALE BRUTE (tous prix individuels confondus) est affichee a la place, avec un
# marquage visuel applique plus bas (fond rose = 1 ou 2 prix locaux, police rouge sans fond = aucun prix
# local). Cette mediane BRUTE (national_median_item) sert UNIQUEMENT de secours pour les communes trop
# peu couvertes - PAS la reference nationale affichee en bas de tableau, voir national_median_of_markets
# plus bas (demande utilisateur : mediane des medianes de marche, pas mediane recalculee sur tous les
# prix individuels).
PRIX_ITEM_MIN_OBS <- 3
price_items <- as.character(PRICE_ITEM_CATALOG$item_var)
n_items <- length(price_items)

item_data <- clean_data %>%
  select(admin1_label, admin2_label, admin3_label, all_of(price_items)) %>%
  mutate(across(all_of(price_items), as.numeric))

local_median_item <- item_data %>% group_by(admin1_label, admin2_label, admin3_label) %>%
  summarise(across(all_of(price_items), ~median(., na.rm = TRUE)), .groups = "drop") %>%
  arrange(admin1_label, admin2_label, admin3_label)
local_count_item <- item_data %>% group_by(admin1_label, admin2_label, admin3_label) %>%
  summarise(across(all_of(price_items), ~sum(!is.na(.))), .groups = "drop") %>%
  arrange(admin1_label, admin2_label, admin3_label)
# reference nationale (ligne "Ensemble des marchés évalués" plus bas ET secours pour les communes avec
# trop peu de cotations) : MEDIANE DES MEDIANES DE MARCHE (chaque commune = un marche, voir
# calculate_market_medians() dans fonctions_analyse.R pour la meme methodologie appliquee au PMAS) - PAS
# une mediane recalculee sur l'ensemble des prix individuels regroupes, pour qu'un marche avec beaucoup
# de vendeurs enquetes ne pese pas plus qu'un marche avec peu de vendeurs. Une seule reference nationale
# utilisee pour les deux usages (demande utilisateur explicite, corrige une premiere version qui
# affichait cette mediane des medianes mais secourait quand meme avec la mediane brute sur tous les prix
# individuels confondus).
national_median_of_markets <- local_median_item %>% summarise(across(all_of(price_items), ~median(., na.rm = TRUE)))

prix_median <- local_median_item
for (it in price_items) {
  prix_median[[it]] <- if_else(local_count_item[[it]] >= PRIX_ITEM_MIN_OBS,
                               local_median_item[[it]], national_median_of_markets[[it]][1])
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
wb_final <- wb_add_data(wb_final, "Prix médian par article", national_median_of_markets,
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
    `Accessibilité des marchés (score maximum : 25)` = round(mfs_accessibility_pillar_score, 0),
    `Disponibilité des produits (score maximum : 30)` = round(mfs_availability_pillar_score, 0),
    `Abordabilité des prix des produits (score maximum : 15)` = round(mfs_affordability_pillar_score, 0),
    `Résilience des circuits d'approvisionnement (score maximum : 20)` = round(mfs_resilience_pillar_score, 0),
    `Niveau des infrastructures du marché (score maximum : 10)` = round(mfs_infrastructure_pillar_score, 0),
    `Total Score SFM` = round(mfs_score / 100, 2),
    !!paste0("Classification SFM ", MOIS, " ", ANNEE) := functionality
  ) %>%
  arrange(`Préfecture`, `Sous-préfecture`, Commune)

classif_col_name <- paste0("Classification SFM ", MOIS, " ", ANNEE)

# ligne nationale : MEDIANE DES SCORES PAR MARCHE (chaque commune/marche compte pour une seule
# observation) pour chaque pilier et pour le score total - meme principe que la mediane des medianes
# utilisee pour le PMAS ci-dessus (demande utilisateur explicite : "median of the scores by market"). La
# classification de cette ligne nationale est recalculee en reprenant EXACTEMENT la meme logique que
# calculate_mfs() (scripts/mfs.R lignes ~462-491) appliquee aux valeurs medianes : un pilier est "bas"
# s'il est sous 50% de son maximum (12.5/15/7.5/10/5), puis Fonctionnalité complete/limitee/Mauvaise/
# Problemes graves selon le score total median et le nombre de piliers bas.
mfs_national_pillars <- mfs_raw %>%
  ungroup() %>%
  summarise(across(c(mfs_accessibility_pillar_score, mfs_availability_pillar_score,
                     mfs_affordability_pillar_score, mfs_resilience_pillar_score,
                     mfs_infrastructure_pillar_score, mfs_score), ~median(., na.rm = TRUE)))

nat_low_dims <- sum(
  mfs_national_pillars$mfs_accessibility_pillar_score < 12.5,
  mfs_national_pillars$mfs_availability_pillar_score < 15,
  mfs_national_pillars$mfs_affordability_pillar_score < 7.5,
  mfs_national_pillars$mfs_resilience_pillar_score < 10,
  mfs_national_pillars$mfs_infrastructure_pillar_score < 5
)
nat_score <- mfs_national_pillars$mfs_score
# MFS_LVL_* (scripts/mfs.R, source()e plus haut) : memes 5 chaines EXACTES que calculate_mfs(), plus
# jamais retapees a la main ici (voir commentaire dans mfs.R).
nat_functionality <- case_when(
  is.na(nat_score) ~ MFS_LVL_INSUFFICIENT,
  nat_score < 25 | nat_low_dims >= 3 ~ MFS_LVL_SEVERE,
  nat_score < 50 | nat_low_dims >= 2 ~ MFS_LVL_POOR,
  nat_score > 80 & nat_low_dims == 0 ~ MFS_LVL_FULL,
  TRUE ~ MFS_LVL_LIMITED
)

score_fonctionnalite_national <- tibble(
  `Année` = ANNEE, Mois = MOIS, Pays = PAYS,
  `Préfecture` = "Ensemble des marchés évalués", `Sous-préfecture` = NA_character_, Commune = NA_character_,
  `Accessibilité des marchés (score maximum : 25)` = round(mfs_national_pillars$mfs_accessibility_pillar_score, 0),
  `Disponibilité des produits (score maximum : 30)` = round(mfs_national_pillars$mfs_availability_pillar_score, 0),
  `Abordabilité des prix des produits (score maximum : 15)` = round(mfs_national_pillars$mfs_affordability_pillar_score, 0),
  `Résilience des circuits d'approvisionnement (score maximum : 20)` = round(mfs_national_pillars$mfs_resilience_pillar_score, 0),
  `Niveau des infrastructures du marché (score maximum : 10)` = round(mfs_national_pillars$mfs_infrastructure_pillar_score, 0),
  `Total Score SFM` = round(nat_score / 100, 2),
  !!classif_col_name := nat_functionality
)

wb_final <- wb_add_worksheet(wb_final, "Score de Fonctionnalité")
wb_final <- wb_add_data(wb_final, "Score de Fonctionnalité", score_fonctionnalite, na = "")

# ligne separatrice grise (fine) entre les communes et la ligne nationale, meme convention que "Coût
# médian du PMAS" ci-dessus
n_mfs_communes <- nrow(score_fonctionnalite)
sep_row_mfs <- n_mfs_communes + 2
national_row_mfs <- sep_row_mfs + 1
wb_final <- wb_add_fill(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = sep_row_mfs, cols = 1:13), color = wb_color(hex = "FFBFBFBF"))
wb_final <- wb_set_row_heights(wb_final, "Score de Fonctionnalité", rows = sep_row_mfs, heights = 6)

wb_final <- wb_add_data(wb_final, "Score de Fonctionnalité", score_fonctionnalite_national,
                        start_row = national_row_mfs, col_names = FALSE, na = "")
# fusionne Préfecture/Sous-préfecture/Commune sur la ligne nationale (seule colonne D remplie), meme
# convention que "Coût médian du PMAS"
wb_final <- wb_merge_cells(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = national_row_mfs, cols = 4:6))

# mise en forme reprise du classeur de reference REACH_CAR_JMMI_Dataset_Analyse_Juin_Juillet_2026 (memes
# constantes que "Coût médian du PMAS" ci-dessus : en-tete fond rouge corail FFEE5958 + police Roboto
# Condensed 14 gras texte blanc ; corps police Segoe UI 12 BODY_GREY, colonnes de pilier en entiers (le
# score maximum de chaque pilier - 25/30/15/20/10 - est indique directement dans l'en-tete de colonne,
# demande utilisateur, plutot qu'un affichage en pourcentage qui masquerait ces poids differents) ;
# "Total Score SFM" reste en pourcentage (c'est bien un pourcentage du score total sur 100, contrairement
# aux piliers qui ont chacun un maximum different) et en gras 14pt (comme les colonnes "Total" de PMAS).
# Pas de ligne/colonne
# d'espacement au debut (contrairement au classeur de reference) : convention deja etablie par "Coût
# médian du PMAS" dans ce classeur, gardee identique ici pour la coherence entre les deux feuilles. Pas
# non plus des 3 niveaux d'agregation Sous-Prefecture/Prefecture/National du classeur de reference
# (calculate_mfs() ne groupe qu'au niveau Commune) - demande utilisateur portant sur la mise en forme,
# pas sur le calcul ; a envisager separement si ces niveaux sont souhaites.
n_mfs <- n_mfs_communes
last_row_mfs <- national_row_mfs
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

# corps (communes uniquement - PAS la ligne separatrice ni la ligne nationale, stylees separement plus bas)
body_rows_mfs <- 2:(n_mfs + 1)
wb_final <- wb_add_font(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = body_rows_mfs, cols = 1:13), name = "Segoe UI", size = 12, color = wb_color(hex = BODY_GREY))
wb_final <- wb_add_font(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = body_rows_mfs, cols = 12), name = "Segoe UI", size = 14, bold = TRUE, color = wb_color(hex = BODY_GREY))
wb_final <- wb_add_font(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = body_rows_mfs, cols = classif_col), name = "Segoe UI", size = 14, bold = TRUE, color = wb_color(hex = BODY_GREY))
wb_final <- wb_add_cell_style(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = body_rows_mfs, cols = c(1,2,3,7:12)), horizontal = "center", vertical = "center")
wb_final <- wb_add_cell_style(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = body_rows_mfs, cols = 4:6), horizontal = "left", vertical = "center", indent = 1)
wb_final <- wb_add_cell_style(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = body_rows_mfs, cols = classif_col), horizontal = "left", vertical = "center", wrap_text = "1")
wb_final <- wb_add_numfmt(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = body_rows_mfs, cols = 7:11), numfmt = "0")
wb_final <- wb_add_numfmt(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = body_rows_mfs, cols = 12), numfmt = "0%")

# ligne nationale : meme mise en forme marquee que la ligne nationale de "Coût médian du PMAS" (police
# Roboto Condensed 18 gras, bordure EPAISSE noire en haut ET en bas) - le libelle fusionne
# (Préfecture:Commune) y est CENTRE, contrairement aux lignes de donnees ou ces colonnes sont indentees a
# gauche.
wb_final <- wb_add_font(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = national_row_mfs, cols = 1:13), name = "Roboto Condensed", size = 18, bold = TRUE, color = wb_color(hex = BODY_GREY))
wb_final <- wb_add_cell_style(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = national_row_mfs, cols = c(1,2,3,7:12)), horizontal = "center", vertical = "center")
wb_final <- wb_add_cell_style(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = national_row_mfs, cols = 4:6), horizontal = "center", vertical = "center")
wb_final <- wb_add_cell_style(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = national_row_mfs, cols = classif_col), horizontal = "left", vertical = "center", wrap_text = "1")
wb_final <- wb_add_numfmt(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = national_row_mfs, cols = 7:11), numfmt = "0")
wb_final <- wb_add_numfmt(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = national_row_mfs, cols = 12), numfmt = "0%")
wb_final <- wb_add_border(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = national_row_mfs, cols = 1:13),
                          top_border = "thick", bottom_border = "thick", left_border = "none", right_border = "none")

# code couleur de la classification (fond plein par categorie, meme palette que le classeur de reference
# - "Donnees insuffisantes" n'a pas de fond, seulement un texte rouge, contrairement aux 4 autres) -
# applique APRES le style de corps general ci-dessus pour le remplacer sur cette seule colonne/ces
# seules lignes. Cle = MFS_LVL_* (scripts/mfs.R), jamais la chaine retapee a la main (voir commentaire
# dans mfs.R).
MFS_CATEGORY_STYLES <- list()
MFS_CATEGORY_STYLES[[MFS_LVL_SEVERE]] <- list(fill = "FFFF0000")
MFS_CATEGORY_STYLES[[MFS_LVL_POOR]] <- list(fill = "FFF79646")
MFS_CATEGORY_STYLES[[MFS_LVL_LIMITED]] <- list(fill = "FFFFFFBF")
MFS_CATEGORY_STYLES[[MFS_LVL_FULL]] <- list(fill = "FF249EA0")
MFS_CATEGORY_STYLES[[MFS_LVL_INSUFFICIENT]] <- list(font_color = "FFFF0000")
classif_values <- score_fonctionnalite[[classif_col_name]]
for (cat_name in names(MFS_CATEGORY_STYLES)) {
  cat_rows <- which(classif_values == cat_name) + 1  # +1 pour la ligne d'en-tete
  if (isTRUE(nat_functionality == cat_name)) cat_rows <- c(cat_rows, national_row_mfs)  # ligne nationale
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
wb_final <- wb_set_row_heights(wb_final, "Score de Fonctionnalité", rows = national_row_mfs, heights = 24)
wb_final <- wb_freeze_pane(wb_final, "Score de Fonctionnalité", firstActiveRow = 2)
cat("  ..OK (", nrow(score_fonctionnalite), " communes, ",
   sum(is.na(score_fonctionnalite$`Total Score SFM`)), " avec donnees insuffisantes, 1 ligne nationale [",
   nat_functionality, "])\n", sep = "")

# legende visuelle sous le tableau (memes couleurs que MFS_CATEGORY_STYLES/la colonne "Classification
# SFM" ci-dessus, seuils de score simplifies fournis par l'utilisateur pour la legende - le calcul reel
# de calculate_mfs() combine score total ET nombre de dimensions basses, voir Methodologie_MFS.md, mais
# la legende reste une lecture simplifiee par seuil de score, comme dans le classeur de reference)
legend_title_row <- last_row_mfs + 3
legend_header_row <- legend_title_row + 2
legend_data_rows <- (legend_header_row + 1):(legend_header_row + 4)
legend_cols <- 1:4  # A:D
legend_range_cols <- 1:2   # A:B - "Plage de score"
legend_class_cols <- 3:4  # C:D - "Classification"

wb_final <- wb_add_data(wb_final, "Score de Fonctionnalité",
                        data.frame("Classification des marchés en fonction du score de fonctionnalité"),
                        start_row = legend_title_row, start_col = 1, col_names = FALSE)
wb_final <- wb_merge_cells(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = legend_title_row, cols = legend_cols))
wb_final <- wb_add_font(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = legend_title_row, cols = legend_cols[1]),
                        name = "Segoe UI", size = 14, bold = TRUE, color = wb_color(hex = BODY_GREY))
wb_final <- wb_add_cell_style(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = legend_title_row, cols = legend_cols[1]), vertical = "center")

# en-tetes des 2 colonnes (demande utilisateur : la legende passe d'un texte fusionne par ligne a un
# tableau a 2 colonnes - plage de score / classification - chacune toujours coloree par categorie)
wb_final <- wb_add_data(wb_final, "Score de Fonctionnalité", data.frame("Plage de score"), start_row = legend_header_row, start_col = legend_range_cols[1], col_names = FALSE)
wb_final <- wb_merge_cells(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = legend_header_row, cols = legend_range_cols))
wb_final <- wb_add_data(wb_final, "Score de Fonctionnalité", data.frame("Classification"), start_row = legend_header_row, start_col = legend_class_cols[1], col_names = FALSE)
wb_final <- wb_merge_cells(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = legend_header_row, cols = legend_class_cols))
wb_final <- wb_add_font(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = legend_header_row, cols = c(legend_range_cols[1], legend_class_cols[1])),
                        name = "Segoe UI", size = 12, bold = TRUE, color = wb_color(hex = BODY_GREY))
wb_final <- wb_add_cell_style(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = legend_header_row, cols = legend_cols), horizontal = "left", vertical = "center", indent = 1)

MFS_LEGEND_ITEMS <- list(
  list(range = "Score < 25 %", classification = MFS_LVL_SEVERE, fill = MFS_CATEGORY_STYLES[[MFS_LVL_SEVERE]]$fill),
  list(range = "Score entre 26 % et 49 %", classification = MFS_LVL_POOR, fill = MFS_CATEGORY_STYLES[[MFS_LVL_POOR]]$fill),
  list(range = "Score entre 50 % et 79 %", classification = MFS_LVL_LIMITED, fill = MFS_CATEGORY_STYLES[[MFS_LVL_LIMITED]]$fill),
  list(range = "Score ≥ 80 %", classification = MFS_LVL_FULL, fill = MFS_CATEGORY_STYLES[[MFS_LVL_FULL]]$fill)
)
for (i in seq_along(MFS_LEGEND_ITEMS)) {
  item <- MFS_LEGEND_ITEMS[[i]]
  r <- legend_data_rows[i]
  wb_final <- wb_add_data(wb_final, "Score de Fonctionnalité", data.frame(item$range), start_row = r, start_col = legend_range_cols[1], col_names = FALSE)
  wb_final <- wb_merge_cells(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = r, cols = legend_range_cols))
  wb_final <- wb_add_data(wb_final, "Score de Fonctionnalité", data.frame(item$classification), start_row = r, start_col = legend_class_cols[1], col_names = FALSE)
  wb_final <- wb_merge_cells(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = r, cols = legend_class_cols))
  wb_final <- wb_add_fill(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = r, cols = legend_cols), color = wb_color(hex = item$fill))
  wb_final <- wb_add_font(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = r, cols = c(legend_range_cols[1], legend_class_cols[1])),
                          name = "Segoe UI", size = 12, bold = TRUE, color = wb_color(hex = BODY_GREY))
  wb_final <- wb_add_cell_style(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = r, cols = legend_cols), horizontal = "left", vertical = "center", indent = 1)
}
wb_final <- wb_add_border(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = c(legend_header_row, legend_data_rows), cols = legend_cols),
                          inner_hgrid = "thin", inner_vgrid = "thin", top_border = "thin", bottom_border = "thin", left_border = "thin", right_border = "thin")
wb_final <- wb_set_row_heights(wb_final, "Score de Fonctionnalité", rows = c(legend_title_row, legend_header_row, legend_data_rows), heights = 22)

# tableau des poids ET definitions des dimensions du SFM (demande utilisateur - les poids sont deja
# indiques dans les en-tetes de colonne ci-dessus, repris ici sous forme de tableau ; la colonne
# "Définition" reprend la traduction francaise complete (pas un resume) des questions-guides de la
# methodologie MFS de REACH pour chaque dimension, fournies par l'utilisateur en anglais - voir
# Methodologie_MFS.md pour le document source.
weights_title_row <- legend_data_rows[length(legend_data_rows)] + 3
weights_header_row <- weights_title_row + 2
weights_data_rows <- (weights_header_row + 1):(weights_header_row + 5)
name_cols <- 1:3
max_col <- 4
def_cols <- 5:13

wb_final <- wb_add_data(wb_final, "Score de Fonctionnalité", data.frame("SFM : poids et définitions des dimensions"),
                        start_row = weights_title_row, start_col = 1, col_names = FALSE)
wb_final <- wb_merge_cells(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = weights_title_row, cols = 1:13))
wb_final <- wb_add_font(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = weights_title_row, cols = 1),
                        name = "Segoe UI", size = 14, bold = TRUE, color = wb_color(hex = BODY_GREY))
wb_final <- wb_add_cell_style(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = weights_title_row, cols = 1), vertical = "center")

wb_final <- wb_add_data(wb_final, "Score de Fonctionnalité", data.frame("Dimension"), start_row = weights_header_row, start_col = name_cols[1], col_names = FALSE)
wb_final <- wb_merge_cells(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = weights_header_row, cols = name_cols))
wb_final <- wb_add_data(wb_final, "Score de Fonctionnalité", data.frame("Score maximum"), start_row = weights_header_row, start_col = max_col, col_names = FALSE)
wb_final <- wb_add_data(wb_final, "Score de Fonctionnalité", data.frame("Définition"), start_row = weights_header_row, start_col = def_cols[1], col_names = FALSE)
wb_final <- wb_merge_cells(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = weights_header_row, cols = def_cols))
wb_final <- wb_add_font(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = weights_header_row, cols = c(name_cols[1], max_col, def_cols[1])),
                        name = "Segoe UI", size = 12, bold = TRUE, color = wb_color(hex = BODY_GREY))
wb_final <- wb_add_cell_style(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = weights_header_row, cols = 1:13), horizontal = "left", vertical = "center", indent = 1)

MFS_DIMENSION_DEFS <- list(
  list(name = "Accessibilité des marchés", max = 25,
       def = paste(
         "Tous les acteurs du marché, y compris les clients, ont-ils un accès physique à ce marché (c'est-à-dire que la plupart des personnes peuvent se rendre sur les marchés et/ou dans les commerces et y entrer depuis leur lieu d'hébergement sans effort physique ou coût majeur, et que les clients sont physiquement capables de ramener les biens achetés chez eux en quantité suffisante) ?",
         "Tous les acteurs du marché, y compris les clients, ont-ils un accès social à ce marché (c'est-à-dire qu'ils ne sont pas empêchés d'accéder aux marchés et/ou aux commerces, ou d'obtenir des biens, en raison de leur genre, de leur appartenance ethnique, de leur affiliation, ou d'autres caractéristiques physiques ou sociales) ?",
         "Ce marché et les routes qui y mènent sont-ils sûrs et sécurisés (c'est-à-dire que les clients et les autres acteurs du marché peuvent se rendre sur les marchés et/ou dans les commerces sans mettre leur sécurité en danger) ?"
       )),
  list(name = "Disponibilité des produits", max = 30,
       def = "Les vendeurs de ce marché peuvent-ils fournir de manière fiable tous les articles essentiels dont les ménages locaux ont besoin d'acheter régulièrement ?"),
  list(name = "Abordabilité des prix des produits", max = 15,
       def = paste(
         "Les clients ont-ils un accès financier à ce marché (c'est-à-dire que les articles essentiels y sont vendus de manière constante à des prix qu'un ménage local moyen peut se permettre) ?",
         "Les prix des articles essentiels sont-ils stables sur ce marché (c'est-à-dire qu'ils évoluent assez lentement pour permettre aux vendeurs et aux clients de planifier leurs dépenses futures) ?"
       )),
  list(name = "Résilience des circuits d'approvisionnement", max = 20,
       def = paste(
         "Les chaînes d'approvisionnement des articles essentiels sur ce marché fonctionnent-elles de manière fiable ?",
         "Les vendeurs de ce marché sont-ils systématiquement en mesure de réapprovisionner les articles essentiels qu'ils vendent avant d'en manquer ?",
         "Les acteurs de ce marché s'approvisionnent-ils en biens auprès de villes et/ou d'itinéraires d'approvisionnement variés, ou la plupart des biens arrivent-ils sur ce marché via un itinéraire d'approvisionnement unique susceptible d'être perturbé ?"
       )),
  list(name = "Niveau des infrastructures du marché", max = 10,
       def = paste(
         "Les infrastructures physiques dans et autour de ce marché (bâtiments, routes, etc.) sont-elles en assez bon état pour soutenir les activités normales de subsistance et de commerce ?",
         "Les vendeurs de ce marché ont-ils accès à des installations de stockage sécurisées et verrouillables pour conserver leur stock ?",
         "Une infrastructure financière existe-t-elle sur ce marché pour soutenir des modalités de paiement alternatives, au-delà de l'argent liquide physique et du crédit informel ?"
       ))
)
for (i in seq_along(MFS_DIMENSION_DEFS)) {
  d <- MFS_DIMENSION_DEFS[[i]]
  r <- weights_data_rows[i]
  wb_final <- wb_add_data(wb_final, "Score de Fonctionnalité", data.frame(d$name), start_row = r, start_col = name_cols[1], col_names = FALSE)
  wb_final <- wb_merge_cells(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = r, cols = name_cols))
  wb_final <- wb_add_data(wb_final, "Score de Fonctionnalité", data.frame(paste0(d$max, " / 100")), start_row = r, start_col = max_col, col_names = FALSE)
  wb_final <- wb_add_data(wb_final, "Score de Fonctionnalité", data.frame(d$def), start_row = r, start_col = def_cols[1], col_names = FALSE)
  wb_final <- wb_merge_cells(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = r, cols = def_cols))
  wb_final <- wb_add_font(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = r, cols = c(name_cols[1], max_col, def_cols[1])),
                          name = "Segoe UI", size = 11, color = wb_color(hex = BODY_GREY))
  wb_final <- wb_add_font(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = r, cols = name_cols[1]),
                          name = "Segoe UI", size = 11, bold = TRUE, color = wb_color(hex = BODY_GREY))
  wb_final <- wb_add_cell_style(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = r, cols = 1:13), horizontal = "left", vertical = "top", indent = 1, wrap_text = "1")
}
wb_final <- wb_add_cell_style(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = weights_data_rows, cols = max_col), horizontal = "center")
wb_final <- wb_add_border(wb_final, "Score de Fonctionnalité", dims = wb_dims(rows = c(weights_header_row, weights_data_rows), cols = 1:13),
                          inner_hgrid = "thin", inner_vgrid = "thin", top_border = "thin", bottom_border = "thin", left_border = "thin", right_border = "thin")
# hauteur de ligne proportionnelle a la longueur du texte de definition (traduction complete, pas un
# resume - la largeur fusionnee E:M fait environ 210 caracteres, ~15px par ligne de texte a 11pt)
def_chars_per_line <- 210
def_row_heights <- pmax(30, ceiling(nchar(sapply(MFS_DIMENSION_DEFS, function(d) d$def)) / def_chars_per_line) * 15 + 10)
wb_final <- wb_set_row_heights(wb_final, "Score de Fonctionnalité", rows = weights_data_rows, heights = def_row_heights)
wb_final <- wb_set_row_heights(wb_final, "Score de Fonctionnalité", rows = c(weights_title_row, weights_header_row), heights = 22)

################################################################################
# "ANALYSE" (LOA descriptive - reinstauree, demande utilisateur : mise en forme via
# scripts/format_results.R::format_analysis_results() (fournie par l'utilisateur, deplacee ici depuis la
# racine du projet, couleur d'accent changee de #A04040 a #D2CBB8 sur demande), LOA inspiree de
# RCA_JMMI_loa.xlsx, chaque variable desagregee au niveau national (colonne "Total") ET par
# sous-prefecture (une colonne par sous-prefecture) - format_analysis_results() distingue nativement les
# lignes group_var = NA (colonne "Total"/nationale) des lignes group_var = "admin2_label" (une colonne
# par groupe), sans jamais fusionner deux variables techniques differentes entre elles - contrairement a
# l'ancienne approche par pivot_wider() sur les libelles (abandonnee) qui fusionnait a tort les variables
# techniques partageant un meme libelle generique reutilise par article dans l'outil kobo (ex. "Précisez
# la préfecture de provenance:" repete pour 23 articles differents) et faisait planter l'export xlsx.
################################################################################
cat("\n- ANALYSE DESCRIPTIVE (feuille 'Analyse') ...\n")

# la loa generee melange, pour chaque variable, une passe nationale (group_var = NA, placee en premier)
# et une passe par sous-prefecture (group_var = "admin2_label") - voir build_loa() dans
# scripts/fonctions_analyse.R. price_vars reprend le catalogue deja resolu par la feuille "Prix median
# par article" ci-dessus (memes variables techniques "<item>_price_item", deja filtrees aux articles
# reellement presents ce mois-ci) plutot que de re-deviner les noms de colonnes de prix depuis l'outil
# kobo - evite toute divergence entre les deux feuilles.
my_loa <- build_loa(kobo_survey, group_var = "admin2_label", price_vars = as.character(PRICE_ITEM_CATALOG$item_var))
missing_vars <- setdiff(my_loa$analysis_var, colnames(clean_data))
if (length(missing_vars) > 0) {
  cat("  (", length(missing_vars), " variable(s) de l'outil absente(s) de clean_data, ignorees)\n", sep = "")
  my_loa <- my_loa %>% filter(analysis_var %in% colnames(clean_data))
}
cat("  LOA generee :", nrow(my_loa), "lignes (", n_distinct(my_loa$analysis_var), "variables, national + admin2)\n")

# clean_data est lu entierement en texte (voir LECTURE DU CLASSEUR ci-dessus) - les variables de prix
# (analysis_type = "median") doivent etre numeriques pour srvyr::survey_median(), sans quoi
# create_analysis() echoue ("Character vectors not allowed in survey functions") - constate a l'usage.
# Copie de clean_data (pas de modification en place : les autres feuilles de ce script utilisent
# clean_data en texte, ex. l'etiquetage "Donnees nettoyees").
median_vars_present <- intersect(unique(my_loa$analysis_var[my_loa$analysis_type == "median"]), colnames(clean_data))
analyse_data <- clean_data %>% mutate(across(all_of(median_vars_present), as.numeric))

my_design <- srvyr::as_survey_design(analyse_data)
my_results <- create_analysis(my_design, loa = my_loa, sm_separator = "/")

# format_analysis_results() lit les libelles DIRECTEMENT depuis survey/choices (pas besoin de
# create_label_dictionary()/add_label_columns_to_results_table() ni de pivot_wider() sur les libelles -
# elle construit les lignes en iterant par analysis_var technique, donc jamais de collision entre deux
# variables qui partageraient le meme libelle). kobo_choices n'a qu'une colonne "label" generique (pas
# de suffixe de langue), kobo_survey n'a que "label::french" (pas de version anglaise dans cet outil) :
# harmonisees ici sous le nom "label" (copie locale de kobo_survey, ne touche pas l'original utilise
# ailleurs dans ce script) pour que label_col="label" designe la meme colonne dans les deux feuilles,
# comme l'exige format_analysis_results() (voir scripts/format_results.R).
kobo_survey_label <- kobo_survey %>% select(-q_type, -list_name) %>% rename(label = `label::french`)

analyse_tmp_path <- tempfile(fileext = ".xlsx")
format_analysis_results(
  results_table = my_results$results_table,
  survey        = kobo_survey_label,
  choices       = kobo_choices,
  output_path   = analyse_tmp_path,
  label_col     = "label",
  sheet_name    = "Analyse"
)
wb_analyse <- wb_load(analyse_tmp_path)
wb_final <- wb_clone_worksheet(wb_final, old = "Analyse", new = "Analyse", from = wb_analyse)
unlink(analyse_tmp_path)
n_admin2 <- n_distinct(my_loa$group_var[!is.na(my_loa$group_var)])
cat("  ..OK (", n_distinct(my_loa$analysis_var), " variables, national + ", n_admin2, " niveau(x) de desagregation)\n", sep = "")

################################################################################
# "Checklist_Data_Analysis" - feuille "CHECKLIST_Data_Cleaning" du MODELE de checklist (fichier
# ressources/Checklist-Template_..._FRENCH.xlsx), clonee telle quelle (vierge, pas l'auto-evaluation
# deja remplie par le script de nettoyage) sous un nouveau nom - demande utilisateur : "Checklist_Data_
# Analysis" plutot que le nom complet du fichier modele (limite Excel de 31 caracteres pour un nom
# d'onglet, largement depassee par le nom du fichier).
CHECKLIST_TEMPLATE_FILE <- Sys.getenv("checklist_template")
if (nzchar(CHECKLIST_TEMPLATE_FILE) && file.exists(CHECKLIST_TEMPLATE_FILE)) {
  cat("\n- AJOUT de la feuille \"Checklist_Data_Analysis\" (auto-evaluation pre-remplie) ...\n")
  wb_checklist_template <- wb_load(CHECKLIST_TEMPLATE_FILE)
  wb_final <- wb_clone_worksheet(wb_final, old = "CHECKLIST_Data_Cleaning", new = "Checklist_Data_Analysis",
                                 from = wb_checklist_template)

  # pre-remplissage de la colonne A ("Checks (y/n/na)") - demande utilisateur : "precompilee comme
  # CHECKLIST_Data_Cleaning". Les memes numeros de ligne/libelles que CHECKLIST_Data_Cleaning (feuille
  # clonee a l'identique du modele ci-dessus) recoivent les MEMES valeurs y/n/na que
  # fill_checklist_data_cleaning() (voir rca_jmmi_cleaning/scripts/fonctions_nettoyage.R, ~lignes
  # 1888-1945) : ces normes minimales restent valables telles quelles a ce stade car le classeur d'analyse
  # final reprend integralement les feuilles deja verifiees au nettoyage (Données nettoyées, cleaning_log,
  # deletion_log, ...), inchangees. Traduction directe en openxlsx2 (wb_add_data) de la version openxlsx
  # (writeData) du script de nettoyage - meme fonction, wb_final etant ici un objet openxlsx2.
  set_check_analysis <- function(row, value) {
    wb_final <<- wb_add_data(wb_final, "Checklist_Data_Analysis", x = value,
                             start_row = row, start_col = 1, col_names = FALSE)
  }
  # Documentation, Transparence et Reproductibilite
  for (r in 7:14) set_check_analysis(r, "y")
  set_check_analysis(15, "na"); set_check_analysis(16, "na")
  set_check_analysis(18, "y"); set_check_analysis(19, "y")
  set_check_analysis(21, "na"); set_check_analysis(22, "y")
  # Qualite et Exactitude des Donnees
  for (r in c(27:34, 36:42)) set_check_analysis(r, "y")
  set_check_analysis(50, "na"); set_check_analysis(51, "na")
  # Verification du Nettoyage des Donnees
  for (r in 46:49) set_check_analysis(r, "y")
  # Points supplementaires
  set_check_analysis(55, "y"); set_check_analysis(56, "na")
  for (r in 57:59) set_check_analysis(r, "y")
  set_check_analysis(61, "y"); set_check_analysis(62, "y"); set_check_analysis(63, "y")
  set_check_analysis(64, "na"); set_check_analysis(65, "y")
  cat("  ..OK\n")
}

################################################################################
# "meb_composition" - tableau de reference documentant la composition du PMAS/MEB (categorie, article,
# variable technique, poids applique au prix median de l'article) - demande utilisateur, ajoutee a la
# toute fin du classeur (voir "meb_composition" en dernier dans desired_order plus bas). Reprend
# build_meb_composition_table() (scripts/fonctions_analyse.R), deja utilisee par l'ancienne feuille
# "Composition du PMAS" (desactivee plus tot dans la session) - meme tableau, nom d'onglet different.
# Colonne "Disponible ce mois-ci" + marquage rouge (demande utilisateur, suite a l'audit de robustesse) :
# AVANT cet ajout, un article dont la variable avait disparu de l'outil kobo (renommee/retiree)
# disparaissait SILENCIEUSEMENT du panier calcule (voir garde-fou check_meb_items_missing() dans
# calculate_pmas_by_level(), scripts/fonctions_analyse.R) tout en restant liste ici comme si de rien
# n'etait - confirme empiriquement (mediane des medianes ~19% plus basse sur un panier alimentaire sans
# aucune note nulle part). La feuille reflete maintenant explicitement ce qui manque.
################################################################################
cat("\n- CONSTRUCTION du tableau de reference \"meb_composition\" ...\n")
composition_pmas <- build_meb_composition_table(missing_meb_items)  # missing_meb_items calcule plus haut

wb_final <- wb_add_worksheet(wb_final, "meb_composition")
wb_final <- wb_add_data(wb_final, "meb_composition", composition_pmas)
wb_final <- wb_add_font(wb_final, "meb_composition", dims = "A1:E1", bold = TRUE)
wb_final <- wb_add_fill(wb_final, "meb_composition", dims = "A1:E1", color = wb_color(hex = "FFD9D9D9"))
wb_final <- wb_set_col_widths(wb_final, "meb_composition", cols = 1:5, widths = c(26, 26, 34, 28, 46))
wb_final <- wb_freeze_pane(wb_final, "meb_composition", firstActiveRow = 2)

# fond de couleur legere par categorie (panier), pour les distinguer visuellement d'un coup d'oeil -
# chaque categorie forme un bloc de lignes contigu (voir build_meb_composition_table(), bind_rows() dans
# le meme ordre que les 3 categories ci-dessous)
cat_colors <- c("Panier alimentaire" = "FFFCE9D6", "Panier non alimentaire" = "FFDCE9F5",
               "Panier des produits d'hygiène" = "FFE1F0DC")
for (cat_name in names(cat_colors)) {
  rows <- which(composition_pmas$Catégorie == cat_name) + 1
  if (length(rows) > 0) {
    wb_final <- wb_add_fill(wb_final, "meb_composition", dims = wb_dims(rows = rows, cols = 1:5),
                            color = wb_color(hex = cat_colors[[cat_name]]))
  }
}

# marquage rouge des articles manquants (memes couleurs que le marquage "aucune cotation" de "Prix
# médian par article") - applique APRES le fond par categorie ci-dessus pour le remplacer sur ces lignes
missing_rows <- which(composition_pmas$`Disponible ce mois-ci` != "Oui") + 1
if (length(missing_rows) > 0) {
  wb_final <- wb_add_fill(wb_final, "meb_composition", dims = wb_dims(rows = missing_rows, cols = 1:5), color = wb_color(hex = "FFFFC7CE"))
  wb_final <- wb_add_font(wb_final, "meb_composition", dims = wb_dims(rows = missing_rows, cols = 1:5), color = wb_color(hex = "FFFF0000"))
}
wb_final <- wb_add_cell_style(wb_final, "meb_composition", dims = wb_dims(rows = 2:(nrow(composition_pmas) + 1), cols = 5), wrap_text = "1", vertical = "center")

if (length(missing_meb_items) > 0) {
  cat("  ATTENTION : ", length(missing_meb_items), " article(s) PMAS/MEB absent(s) des donnees ce cycle-ci (",
     paste(missing_meb_items, collapse = ", "), ") - exclu(s) du panier, marque(s) en rouge dans meb_composition\n", sep = "")
}
cat("  ..OK (", nrow(composition_pmas), " articles, ", length(missing_meb_items), " absent(s) ce mois-ci)\n", sep = "")

################################################################################
# "meb_analysis" / "mfs_analysis" - REPLIQUE DU MODELE EXTERNE REACH (voir 07_sample_output_file.xlsx,
# feuilles "meb_analysis"/"mfs_analysis" - demande utilisateur explicite : "look at the meb_analysis and
# mfs_analysis and replicate it exactly"), ajoutees a la toute fin du classeur. 4 niveaux d'agregation
# empiles (admin3/admin2/admin1/admin0) - PAS les libelles Prefecture/Sous-Prefecture/Commune utilises
# ailleurs dans ce classeur, mais les codes/libelles PCode + le mot-cle de niveau attendu par le modele.
# Methodologie d'agregation (demande utilisateur explicite) : "always take the market medians and take
# their median or report the same value if we only have one value" - CHAQUE niveau (admin2/admin1/admin0)
# est calcule INDEPENDAMMENT comme la mediane des valeurs de marche (admin3) qui lui appartiennent,
# JAMAIS en cascade via un niveau intermediaire deja agrege (evite qu'une mediane de mediane de mediane
# ne s'eloigne de la vraie mediane des marches) - voir aggregate_admin_levels() ci-dessous.
################################################################################
cat("\n- CONSTRUCTION des feuilles \"meb_analysis\"/\"mfs_analysis\" (modele externe REACH) ...\n")

admin_codes <- clean_data %>%
  distinct(admin1_code, admin1_label, admin2_code, admin2_label, admin3_code, admin3_label)
admin2_codes <- admin_codes %>% distinct(admin1_code, admin1_label, admin2_code, admin2_label)
admin1_codes <- admin_codes %>% distinct(admin1_code, admin1_label)

# empile les 4 niveaux d'agregation (admin3 = marche, tel quel ; admin2/admin1/admin0 = mediane des
# valeurs de marche qui appartiennent a ce groupe, jamais recalculee a partir d'un niveau deja agrege)
# pour un ensemble de colonnes numeriques donne - reutilisee pour le MEB et le MFS. Chaque bloc est trie
# par libelle admin (Préfecture/Sous-Préfecture/Commune) - meme convention que "Coût médian du PMAS"/
# "Prix médian par article" (arrange(Préfecture, Sous-Préfecture, Commune)) - demande utilisateur
# explicite : coherence du tri entre les feuilles d'analyse.
aggregate_admin_levels <- function(commune_df, value_cols) {
  d3 <- commune_df %>% left_join(admin_codes, by = "admin3_label")

  r3 <- d3 %>%
    mutate(admin_level_aggregation = "admin3") %>%
    select(admin_level_aggregation, admin1_code, admin1_label, admin2_code, admin2_label,
          admin3_code, admin3_label, all_of(value_cols)) %>%
    arrange(admin1_label, admin2_label, admin3_label)

  r2 <- d3 %>% group_by(admin2_code) %>%
    summarise(across(all_of(value_cols), ~median(., na.rm = TRUE)), .groups = "drop") %>%
    left_join(admin2_codes, by = "admin2_code") %>%
    mutate(admin_level_aggregation = "admin2", admin3_code = NA_character_, admin3_label = NA_character_) %>%
    select(admin_level_aggregation, admin1_code, admin1_label, admin2_code, admin2_label,
          admin3_code, admin3_label, all_of(value_cols)) %>%
    arrange(admin1_label, admin2_label)

  r1 <- d3 %>% group_by(admin1_code) %>%
    summarise(across(all_of(value_cols), ~median(., na.rm = TRUE)), .groups = "drop") %>%
    left_join(admin1_codes, by = "admin1_code") %>%
    mutate(admin_level_aggregation = "admin1", admin2_code = NA_character_, admin2_label = NA_character_,
          admin3_code = NA_character_, admin3_label = NA_character_) %>%
    select(admin_level_aggregation, admin1_code, admin1_label, admin2_code, admin2_label,
          admin3_code, admin3_label, all_of(value_cols)) %>%
    arrange(admin1_label)

  r0 <- d3 %>%
    summarise(across(all_of(value_cols), ~median(., na.rm = TRUE))) %>%
    mutate(admin_level_aggregation = "admin0", admin1_code = NA_character_, admin1_label = NA_character_,
          admin2_code = NA_character_, admin2_label = NA_character_,
          admin3_code = NA_character_, admin3_label = NA_character_) %>%
    select(admin_level_aggregation, admin1_code, admin1_label, admin2_code, admin2_label,
          admin3_code, admin3_label, all_of(value_cols))

  bind_rows(r3, r2, r1, r0)
}

# "year"/"month" au format impose par le modele (annee en entier, mois en anglais complet) - MOIS/ANNEE
# (utilises partout ailleurs dans ce classeur) sont en francais/texte, traduits ici uniquement pour ces 2
# feuilles.
MOIS_EN_LOOKUP <- c(Janvier = "January", Fevrier = "February", Mars = "March", Avril = "April", Mai = "May",
                    Juin = "June", Juillet = "July", Aout = "August", Septembre = "September",
                    Octobre = "October", Novembre = "November", Decembre = "December")
ANALYSIS_YEAR_INT <- as.integer(ANNEE)
ANALYSIS_MONTH_EN <- unname(MOIS_EN_LOOKUP[MOIS])

# mise en forme commune aux 2 feuilles (demande utilisateur : "format the table nicer with the red first
# row and generally good coloring or bordering ... to distinguish the different level of analysis, do not
# add empty rows") - PAS de ligne vide entre les 4 blocs (admin3/admin2/admin1/admin0, toujours dans cet
# ordre - voir aggregate_admin_levels()) : chaque niveau recoit son propre fond (de plus en plus fonce a
# mesure que le niveau d'agregation monte), et une bordure MOYENNE marque chaque changement de niveau a la
# place d'une ligne separatrice.
style_admin_level_sheet <- function(wb, sheet, df, value_groups = NULL) {
  n_col <- ncol(df)
  last_row <- nrow(df) + 1
  level_col <- df$admin_level_aggregation

  wb <- wb_add_fill(wb, sheet, dims = wb_dims(rows = 1, cols = 1:n_col), color = wb_color(hex = "FFEE5958"))
  wb <- wb_add_font(wb, sheet, dims = wb_dims(rows = 1, cols = 1:n_col), name = "Roboto Condensed", size = 12, bold = TRUE, color = wb_color(hex = "FFFFFFFF"))
  wb <- wb_add_cell_style(wb, sheet, dims = wb_dims(rows = 1, cols = 1:n_col), horizontal = "center", vertical = "center", wrap_text = "1")
  wb <- wb_set_row_heights(wb, sheet, rows = 1, heights = 32)

  wb <- wb_add_border(wb, sheet, dims = wb_dims(rows = 1:last_row, cols = 1:n_col), inner_hgrid = "thin", inner_vgrid = "thin")
  wb <- wb_add_font(wb, sheet, dims = wb_dims(rows = 2:last_row, cols = 1:n_col), name = "Segoe UI", size = 11, color = wb_color(hex = "FF262626"))
  wb <- wb_add_cell_style(wb, sheet, dims = wb_dims(rows = 1:last_row, cols = 1:n_col), horizontal = "center", vertical = "center")

  level_fills <- c(admin3 = "FFFFFFFF", admin2 = "FFF2F2F2", admin1 = "FFD9D9D9", admin0 = "FFBFBFBF")
  for (lvl in names(level_fills)) {
    rows <- which(level_col == lvl) + 1
    if (length(rows) == 0) next
    wb <- wb_add_fill(wb, sheet, dims = wb_dims(rows = rows, cols = 1:n_col), color = wb_color(hex = level_fills[[lvl]]))
  }
  # ligne nationale (admin0) en gras, meme convention que les autres lignes nationales de ce classeur
  admin0_rows <- which(level_col == "admin0") + 1
  if (length(admin0_rows) > 0) {
    wb <- wb_add_font(wb, sheet, dims = wb_dims(rows = admin0_rows, cols = 1:n_col), name = "Segoe UI", size = 11, bold = TRUE, color = wb_color(hex = "FF262626"))
  }

  # bordure MOYENNE a la 1ere ligne de chaque nouveau bloc de niveau (admin3 -> admin2 -> admin1 -> admin0)
  # - update = TRUE est ESSENTIEL ici (constate empiriquement, teste isolement) : par defaut,
  # wb_add_border() REMPLACE integralement la bordure des cellules touchees par ce qui est explicitement
  # passe dans CET appel (les cotes non passes deviennent vides, quel que soit ce qui avait ete pose par
  # un appel precedent - ex. l'inner_hgrid/inner_vgrid du tableau entier plus haut) ; update = TRUE fusionne
  # au contraire avec la bordure deja presente, ne touchant que les cotes explicitement donnes.
  n3 <- sum(level_col == "admin3"); n2 <- sum(level_col == "admin2")
  n1 <- sum(level_col == "admin1"); n0 <- sum(level_col == "admin0")
  block_starts <- c(2 + n3, 2 + n3 + n2, 2 + n3 + n2 + n1)
  block_starts <- block_starts[c(n2, n1, n0) > 0 & block_starts <= last_row]
  for (br in block_starts) {
    wb <- wb_add_border(wb, sheet, dims = wb_dims(rows = br, cols = 1:n_col), top_border = "medium", update = TRUE)
  }

  # demande utilisateur : bordures plus epaisses autour de chaque groupe de colonnes "variables" (meb :
  # local_currency / eur_xrate_official / usd_xrate_official ; mfs : les scores) - encadre chaque groupe
  # d'une bordure MOYENNE (gauche de sa 1ere colonne, droite de sa derniere), update = TRUE (voir
  # commentaire ci-dessus) pour fusionner avec la grille/bordures deja posees plutot que les remplacer.
  if (!is.null(value_groups)) {
    col_names <- colnames(df)
    for (grp in value_groups) {
      idx <- match(grp, col_names)
      idx <- idx[!is.na(idx)]
      if (length(idx) == 0) next
      wb <- wb_add_border(wb, sheet, dims = wb_dims(rows = 2:last_row, cols = min(idx)), left_border = "medium", update = TRUE)
      wb <- wb_add_border(wb, sheet, dims = wb_dims(rows = 2:last_row, cols = max(idx)), right_border = "medium", update = TRUE)
    }
  }

  wb <- wb_freeze_pane(wb, sheet, firstActiveRow = 2)
  wb <- wb_set_col_widths(wb, sheet, cols = 1:n_col, widths = 18)
  wb
}

# --- meb_analysis --------------------------------------------------------------------------------------
p_commune_meb <- p_commune %>%
  transmute(
    admin3_label,
    meb_food_local_currency = meb_food,
    meb_nfi_local_currency = meb_nfi,
    meb_personal_hygiene_local_currency = meb_wash,
    meb_total_local_currency = meb_total,
    meb_food_eur_xrate_official = meb_food / EUR_XRATE,
    meb_nfi_eur_xrate_official = meb_nfi / EUR_XRATE,
    meb_personal_hygiene_eur_xrate_official = meb_wash / EUR_XRATE,
    meb_total_eur_xrate_official = meb_total / EUR_XRATE,
    meb_food_usd_xrate_official = meb_food / USD_XRATE,
    meb_nfi_usd_xrate_official = meb_nfi / USD_XRATE,
    meb_personal_hygiene_usd_xrate_official = meb_wash / USD_XRATE,
    meb_total_usd_xrate_official = meb_total / USD_XRATE
  )
meb_value_cols <- setdiff(colnames(p_commune_meb), "admin3_label")

meb_analysis <- aggregate_admin_levels(p_commune_meb, meb_value_cols) %>%
  mutate(across(all_of(meb_value_cols), ~round(., 0)),  # demande utilisateur : arrondi a l'entier
        year = ANALYSIS_YEAR_INT, month = ANALYSIS_MONTH_EN, country = PAYS_ISO3, currency_code = "XAF") %>%
  select(admin_level_aggregation, year, month, country, admin1_code, admin1_label, admin2_code, admin2_label,
        admin3_code, admin3_label, currency_code, all_of(meb_value_cols))

wb_final <- wb_add_worksheet(wb_final, "meb_analysis")
wb_final <- wb_add_data(wb_final, "meb_analysis", meb_analysis, na = "")
meb_value_groups <- list(
  local_currency = grep("_local_currency$", meb_value_cols, value = TRUE),
  eur_xrate_official = grep("_eur_xrate_official$", meb_value_cols, value = TRUE),
  usd_xrate_official = grep("_usd_xrate_official$", meb_value_cols, value = TRUE)
)
wb_final <- style_admin_level_sheet(wb_final, "meb_analysis", meb_analysis, value_groups = meb_value_groups)
cat("  ..OK (meb_analysis :", nrow(meb_analysis), "lignes,", ncol(meb_analysis), "colonnes)\n")

# --- mfs_analysis ---------------------------------------------------------------------------------------
p_commune_mfs <- mfs_raw %>%
  ungroup() %>%
  transmute(
    admin3_label,
    mfs_accessibility_score = mfs_accessibility_pillar_score,
    mfs_availability_score = mfs_availability_pillar_score,
    mfs_affordability_score = mfs_affordability_pillar_score,
    mfs_resilience_score = mfs_resilience_pillar_score,
    mfs_infrastructure_score = mfs_infrastructure_pillar_score,
    mfs_total_score = mfs_score
  )
mfs_value_cols <- setdiff(colnames(p_commune_mfs), "admin3_label")

# classification (mfs_functionality_classification) : uniquement au niveau marche (admin3) - "optional,
# also there is no need to aggregate it at admin1 admin2 and admin0" (modele externe) ; libelles traduits
# en anglais pour correspondre exactement au modele (voir la logique originale/en anglais dans
# calculate_mfs(), scripts/mfs.R : "Categorising outcomes").
MFS_CLASSIF_EN <- setNames(
  c("Full_functionality", "Limited_functionality", "Poor_functionality", "Severe_issues", "Insufficient_data"),
  c(MFS_LVL_FULL, MFS_LVL_LIMITED, MFS_LVL_POOR, MFS_LVL_SEVERE, MFS_LVL_INSUFFICIENT)
)
classif_by_admin3 <- mfs_raw %>% ungroup() %>%
  transmute(admin3_label, mfs_functionality_classification = unname(MFS_CLASSIF_EN[functionality]))

mfs_analysis <- aggregate_admin_levels(p_commune_mfs, mfs_value_cols) %>%
  mutate(across(all_of(mfs_value_cols), ~round(., 0)),  # demande utilisateur : arrondi a l'entier
        year = ANALYSIS_YEAR_INT, month = ANALYSIS_MONTH_EN, country = PAYS_ISO3) %>%
  left_join(classif_by_admin3, by = "admin3_label") %>%
  select(admin_level_aggregation, year, month, country, admin1_code, admin1_label, admin2_code, admin2_label,
        admin3_code, admin3_label, all_of(mfs_value_cols), mfs_functionality_classification)

wb_final <- wb_add_worksheet(wb_final, "mfs_analysis")
wb_final <- wb_add_data(wb_final, "mfs_analysis", mfs_analysis, na = "")
mfs_value_groups <- list(scores = mfs_value_cols)
wb_final <- style_admin_level_sheet(wb_final, "mfs_analysis", mfs_analysis, value_groups = mfs_value_groups)
cat("  ..OK (mfs_analysis :", nrow(mfs_analysis), "lignes,", ncol(mfs_analysis), "colonnes)\n")

################################################################################
# MISE A JOUR DE LA FEUILLE READ_ME : DESCRIPTION DES FEUILLES D'ANALYSE
################################################################################
# demande utilisateur : "when the analysis is done add a description of all the sheets in the file, in
# the dedicated part of the read me" - complete le tableau "Contenu" deja rempli par le NETTOYAGE
# (write_read_me_sheet() dans fonctions_nettoyage.R) avec les feuilles ajoutees ici par l'ANALYSE, dans
# le meme style (police/gras colonne A, retour a la ligne colonne B). La derniere ligne DEJA remplie est
# detectee dynamiquement (pas de numero de ligne code en dur : le nombre de lignes du nettoyage varie
# d'un cycle a l'autre - deletion_log present/absent, etc.). Applique aussi retroactivement
# estimate_wrapped_row_height() (scripts/fonctions_analyse.R - meme logique que cote nettoyage) aux
# lignes du tableau "Contenu" deja remplies par le nettoyage : le classeur importe a pu etre nettoye
# AVANT cette fonctionnalite et porter encore des hauteurs de ligne fixes (demande utilisateur : "make
# the rows width dependent on the content").
readme_sheet_name <- intersect(c("READ_ME", "Lisez_moi"), wb_get_sheet_names(wb_final))[1]
if (!is.na(readme_sheet_name)) {
  cat("\n- MISE A JOUR de la feuille \"", readme_sheet_name, "\" (index des feuilles) ...\n", sep = "")
  readme_df <- wb_to_df(wb_final, sheet = readme_sheet_name, col_names = FALSE)
  readme_last_row <- max(which(!is.na(readme_df[[1]]) | !is.na(readme_df[[2]])))
  contenu_banner_row <- which(readme_df[[1]] == "Contenu")

  if (length(contenu_banner_row) > 0) {
    index_first_row <- contenu_banner_row[1] + 2  # +1 = ligne d'en-tete "Onglets"/"Description", +1 = 1ere ligne de donnees

    # hauteur dependante du contenu pour les lignes d'INTRODUCTION du haut (Contexte du projet, Objectif
    # general, ...) egalement deja remplies par le nettoyage - lignes 1 et 2 (bannieres titre/sous-titre,
    # texte court sur une seule ligne, hauteur fixe volontaire) et la banniere "Contenu" elle-meme
    # exclues.
    intro_last_row <- contenu_banner_row[1] - 1
    if (3 <= intro_last_row) {
      for (rr in 3:intro_last_row) {
        wb_final <- wb_set_row_heights(wb_final, readme_sheet_name, rows = rr,
                                       heights = estimate_wrapped_row_height(readme_df[[2]][rr], col_width_chars = 100))
      }
    }

    # hauteur dependante du contenu pour les lignes DEJA remplies par le nettoyage dans le tableau
    # "Contenu"
    if (index_first_row <= readme_last_row) {
      for (rr in index_first_row:readme_last_row) {
        txt <- readme_df[[2]][rr]
        wb_final <- wb_set_row_heights(wb_final, readme_sheet_name, rows = rr,
                                       heights = estimate_wrapped_row_height(txt, col_width_chars = 100))
      }
    }

    ANALYSIS_SHEET_DESCRIPTIONS <- list(
      "Checklist_Data_Analysis" = "Auto-évaluation des normes minimales, reprise à l'étape d'analyse - remplie automatiquement pour les vérifications que le script peut certifier, le reste à compléter par le réviseur.",
      "Données nettoyées" = "Copie de clean_data (mêmes données et mise en forme), avec les colonnes de prix individuels/taux de change masquées (déjà détaillées dans les feuilles de prix/PMAS ci-dessous).",
      "Coût médian du PMAS" = "Coût du Panier Minimum d'Articles de Survie (PMAS/MEB), par commune et au niveau national (médiane des paniers communaux déjà calculés) - voir l'explication détaillée dans la feuille elle-même.",
      "Prix médian par article" = "Prix médian de chaque article du PMAS, par commune et au niveau national (médiane des médianes de marché).",
      "Score de Fonctionnalité" = "Score de Fonctionnalité des Marchés (SFM/MFS), par commune et au niveau national (médiane des scores communaux déjà calculés), avec classification et définitions des dimensions.",
      "Analyse" = "Analyse descriptive complète de l'ensemble du questionnaire (national et par sous-préfecture).",
      "meb_composition" = "Tableau de référence : composition du PMAS/MEB (articles, variables techniques, poids appliqués à chaque prix médian).",
      "meb_analysis" = "PMAS/MEB au format standard REACH (mêmes valeurs que \"Coût médian du PMAS\"), empilé par niveau d'agrégation (admin3/admin2/admin1/admin0), en devise locale (XAF), en euro et en dollar.",
      "mfs_analysis" = "Score de Fonctionnalité des Marchés au format standard REACH (mêmes valeurs que \"Score de Fonctionnalité\"), empilé par niveau d'agrégation (admin3/admin2/admin1/admin0)."
    )

    r <- readme_last_row + 1
    for (nm in names(ANALYSIS_SHEET_DESCRIPTIONS)) {
      desc <- ANALYSIS_SHEET_DESCRIPTIONS[[nm]]
      wb_final <- wb_add_data(wb_final, readme_sheet_name, nm, start_row = r, start_col = 1, col_names = FALSE)
      wb_final <- wb_add_data(wb_final, readme_sheet_name, desc, start_row = r, start_col = 2, col_names = FALSE)
      wb_final <- wb_add_font(wb_final, readme_sheet_name, dims = wb_dims(rows = r, cols = 1), bold = TRUE)
      wb_final <- wb_add_cell_style(wb_final, readme_sheet_name, dims = wb_dims(rows = r, cols = 1), vertical = "top")
      wb_final <- wb_add_cell_style(wb_final, readme_sheet_name, dims = wb_dims(rows = r, cols = 2), vertical = "top", wrap_text = "1")
      wb_final <- wb_set_row_heights(wb_final, readme_sheet_name, rows = r,
                                     heights = estimate_wrapped_row_height(desc, col_width_chars = 100))
      r <- r + 1
    }
    cat("  ..OK (", length(ANALYSIS_SHEET_DESCRIPTIONS), " feuilles d'analyse ajoutées à l'index)\n", sep = "")
  } else {
    cat("  .. tableau \"Contenu\" introuvable dans cette feuille, mise a jour ignoree\n")
  }
}

################################################################################
# ORDRE FINAL DES FEUILLES ET SAUVEGARDE
################################################################################
# couleur des onglets (demande utilisateur) : rouge = feuille de presentation (READ_ME/Lisez_moi, selon la
# langue du classeur importe) ; vert = feuilles d'ANALYSE ajoutees par ce script (donnees nettoyees + les 5
# feuilles de resultats "principales" - Coût médian du PMAS, Evolution des paniers du PMAS (demande
# utilisateur explicite, ajoutee ici - etait grise auparavant), Prix médian par article, Score de
# Fonctionnalité, Analyse - "Composition du PMAS" reste grise, non citee par l'utilisateur) ; gris = tout
# le reste (checklists, feuilles techniques du nettoyage : survey, choices, raw_data, clean_data,
# cleaning_log, deletion_log, variable_tracker). Rouge/vert reprennent des couleurs deja utilisees ailleurs
# dans ce classeur (FFEE5958 = rouge corail des en-tetes de tableau, FFBFBFBF = gris des lignes
# separatrices) pour rester coherent visuellement ; vert = vert standard Office (aucun vert n'est deja
# utilise ailleurs dans ce classeur).
# IMPORTANT : ceci doit s'executer AVANT wb_set_order() ci-dessous - constate empiriquement, wb_set_order()
# change l'ordre renvoye par wb_get_sheet_names() mais PAS l'ordre physique de la liste wb$worksheets ;
# apres un reordonnancement, match(nom, wb_get_sheet_names(wb)) pointe donc vers le mauvais element de
# wb$worksheets (les couleurs se retrouvent sur les mauvais onglets). Avant tout reordonnancement, les deux
# ordres coincident encore (chaque feuille vient d'etre ajoutee dans l'ordre), donc l'association nom <->
# indice physique est fiable ici.
TAB_COLOR_RED   <- "FFEE5958"
TAB_COLOR_GREEN <- "FF00B050"
TAB_COLOR_GREY  <- "FFBFBFBF"
tab_sheets_red   <- c("READ_ME", "Lisez_moi")
tab_sheets_green <- c("Coût médian du PMAS", "Evolution des paniers du PMAS", "Prix médian par article",
                      "Score de Fonctionnalité", "Analyse", "Données nettoyées")

set_tab_color <- function(wb, sheet_name, hex) {
  idx <- match(sheet_name, wb_get_sheet_names(wb))
  if (!is.na(idx)) wb$worksheets[[idx]]$sheetPr <- sprintf('<sheetPr><tabColor rgb="%s"/></sheetPr>', hex)
  wb
}
for (sn in wb_get_sheet_names(wb_final)) {
  hex <- if (sn %in% tab_sheets_red) TAB_COLOR_RED else if (sn %in% tab_sheets_green) TAB_COLOR_GREEN else TAB_COLOR_GREY
  wb_final <- set_tab_color(wb_final, sn, hex)
}

# demande utilisateur : les 2 checklists sont deplacees APRES "Analyse" (etaient juste apres READ_ME)
desired_order <- c("READ_ME", "Lisez_moi",
                   "Données nettoyées", "Coût médian du PMAS",
                   "Composition du PMAS", "Evolution des paniers du PMAS", "Prix médian par article",
                   "Score de Fonctionnalité", "Analyse",
                   "CHECKLIST_Data_Cleaning", "Checklist_Data_Analysis",
                   "survey", "choices", "raw_data", "clean_data",
                   "cleaning_log", "deletion_log", "variable_tracker",
                   "meb_composition", "meb_analysis", "mfs_analysis")  # demande utilisateur : toujours en tout dernier
current_sheets <- wb_get_sheet_names(wb_final)
final_order <- intersect(desired_order, current_sheets)
final_order <- c(final_order, setdiff(current_sheets, final_order))  # garde-fou : n'oublie aucune feuille
wb_final <- wb_set_order(wb_final, match(final_order, current_sheets))

output_dir <- Sys.getenv("output_dir")
if (!dir.exists(output_dir)) dir.create(output_dir, recursive = TRUE)
mois_slug <- if (nzchar(MOIS)) MOIS else "mois"
annee_slug <- if (nzchar(ANNEE)) ANNEE else format(Sys.Date(), "%Y")
# nom de fichier demande par l'utilisateur : "REACH_CAR_Analyse_JMMI_<Mois>-<Annee>.xlsx" (mois/annee
# detectes automatiquement depuis les donnees, voir get_assessment_period_from_data() plus haut - pas
# de date du jour dans le nom, contrairement a l'ancien format).
out_path <- file.path(output_dir, paste0("REACH_CAR_Analyse_JMMI_", mois_slug, "-", annee_slug, ".xlsx"))
wb_save(wb_final, out_path, overwrite = TRUE)
cat("\n--> Classeur d'analyse complet exporte :", out_path, "\n")
