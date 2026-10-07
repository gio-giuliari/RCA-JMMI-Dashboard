################################################################################
# JMMI RCA - TABLEAU DE BORD (habillage Shiny du pipeline de nettoyage)
################################################################################
# Ce dossier est AUTONOME (self-contained) : 1_suivi_quotidien.R, 2_nettoyage_final.R, 4_analyse_complete.R,
# scripts/ et ressources/ sont des COPIES du projet rca_jmmi_cleaning (source d'origine, a cote), placees
# ici pour que ce dossier puisse etre deploye/publie seul (ex. Posit Connect/shinyapps.io), sans dependre
# des dossiers voisins a l'execution. scripts/fonctions_relabel.R est une copie de jmmi_relabelling
# (re-etiquetage des codes de choix, desormais integree automatiquement a l'analyse - voir
# 4_analyse_complete.R - plutot qu'une etape separee du tableau de bord, demande utilisateur explicite).
# Si le pipeline de rca_jmmi_cleaning ou l'outil jmmi_relabelling evoluent, il faut recopier ces fichiers
# ici pour rester a jour (pas de synchronisation automatique).
# A chaque clic sur "Lancer", l'app copie ces scripts/ressources dans un dossier temporaire isole (propre
# a cette execution), y place les fichiers importes par l'utilisateur, puis EXECUTE le script demande
# DANS LA MEME SESSION R QUE L'APP SHINY (voir run_script_inprocess() plus bas) - pas de sous-processus
# Rscript.exe separe, pour rester compatible avec un deploiement Shiny standard (Posit Connect,
# shinyapps.io, Shiny Server) ou seule cette session R est disponible. L'isolation entre executions est
# obtenue via un environnement R frais (new.env()) a chaque appel plutot que via un processus separe.
################################################################################

library(shiny)
library(shinyjs)
library(shinythemes)
library(htmltools)
library(openxlsx)
# packages utilises par les scripts du pipeline (1_suivi_quotidien.R, 2_nettoyage_final.R,
# 4_analyse_complete.R, scripts/*.R), maintenant executes dans cette meme session R (voir
# run_script_inprocess()) - declares ici pour que la publication (ex. rsconnect::deployApp()) les
# detecte et les installe sur le serveur.
library(tidyverse)
library(readxl)
library(writexl)
library(openxlsx2)
library(cleaningtools)
library(labelled)
# utilises par 4_analyse_complete.R (analyse descriptive + PMAS/MEB + Score de Fonctionnalite) -
# analysistools/presentresults sont, comme cleaningtools, des paquets IMPACT installes depuis GitHub
# (pas le CRAN) : verifier qu'ils sont bien geles dans le meme manifeste de deploiement que cleaningtools
# (voir la note sur Posit Connect/shinyapps.io en tete de fichier).
library(analysistools)
library(presentresults)
library(srvyr)
# plyr n'est JAMAIS attache via library() ici expres - il n'est utilise que via plyr::rbind.fill()
# (namespace qualifie) dans 2_nettoyage_final.R, car plyr masque des fonctions dplyr (rename,
# summarise, ...) utilisees SANS prefixe ailleurs dans le pipeline. rsconnect detecte quand meme la
# dependance a partir de l'appel plyr:: lui-meme, pas besoin de library(plyr) pour ca.

## chemins fixes ---------------------------------------------------------------------------------
PIPELINE_ROOT <- normalizePath(".")
CHECKLIST_TEMPLATE_NAME <- "Checklist-Template_Nettoyage-Donnees-Quant_Normes-Min_v1_2_2026-03_FRENCH.xlsx"

## taux de change par defaut (source unique - demande utilisateur, suite a l'audit de robustesse : ces
## 2 valeurs etaient recopiees independamment a 2 endroits (env_vars du nettoyage ET de l'analyse), sans
## garantie de rester synchronisees si l'une des deux copies etait mise a jour sans l'autre). XAF/EUR est
## une parite fixe (ne change jamais) ; XAF/USD flotte sur le marche - DEFAULT_USD_XRATE doit etre mis a
## jour periodiquement.
DEFAULT_EUR_XRATE <- "655.95"
DEFAULT_USD_XRATE <- "570"

## constantes PMAS/MEB partagees avec 4_analyse_complete.R, et generateur du document de methodologie
## telechargeable (voir scripts/methodology.R - build_methodology_html()) qui a remplace l'ancienne
## boite d'info "Comment le PMAS est calculé" de l'interface (demande utilisateur explicite : "remove the
## ui with the meb explanation ... that can be downloaded directly from the dashboard"). source()es au
## niveau global de l'app (pas seulement copiees dans le dossier temporaire d'une execution, comme pour
## le reste du pipeline) : sans effet de bord au chargement (uniquement des definitions de constantes/
## fonctions), et les executions du pipeline restent isolees dans leur propre environnement (voir
## run_analyse_pipeline()/run_script_inprocess() plus bas), independantes de cette copie globale.
source(file.path(PIPELINE_ROOT, "scripts", "fonctions_analyse.R"), local = TRUE)
source(file.path(PIPELINE_ROOT, "scripts", "methodology.R"), local = TRUE)

#### UI DEFINITION ####
ui <- shiny::bootstrapPage(
  shinyjs::useShinyjs(),
  tags$head(tags$link(rel = "stylesheet", href = "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css")),
  tags$head(includeCSS("styles.css")),
  tags$head(tags$script(HTML("
    function highlightModeToggle(container) {
      $(container).find('.radio-inline').each(function() {
        var checked = $(this).find('input[type=radio]').is(':checked');
        $(this).toggleClass('active-choice', checked);
      });
    }
    $(document).on('shiny:connected shiny:value', function() {
      $('.mode-toggle').each(function() { highlightModeToggle(this); });
    });
    $(document).on('change', '.mode-toggle input[type=radio]', function() {
      highlightModeToggle($(this).closest('.mode-toggle'));
    });
  "))),

  navbarPage(
    id = "navbarPage",
    theme = shinythemes::shinytheme("flatly"),
    collapsible = TRUE,
    title = div(HTML('<a style="text-decoration:none;cursor:default;color:#FFFFFF;" class="active" href="#">JMMI RCA Dashboard </a>')),
    windowTitle = "JMMI RCA - Tableau de bord",
    position = "fixed-top",

    tabPanel("Nettoyage donn\u00e9es",
      div(class = "content-wrapper",
        fluidRow(
          div(class = "dashboard-box",
            br(),

            tags$h2(tags$b("Pipeline de nettoyage JMMI - RCA")),
            br(),

            div(class = "info-box",
              div(class = "info-title", "Comment \u00e7a marche"),
              div(class = "info-row", "1. Choisissez ce que vous voulez faire."),
              div(class = "blue-line"),
              div(class = "info-row", "2. Suivi quotidien / Nettoyage : importez le jeu de donn\u00e9es brutes et l'outil Kobo. Pour le nettoyage, choisissez \"Pour suivi\" la premi\u00e8re fois (cr\u00e9e le classeur \u00e0 faire remplir par l'\u00e9quipe de suivi), puis \"Finale\" une fois ce classeur rempli (\u00e0 importer en plus) pour produire le classeur final."),
              div(class = "blue-line"),
              div(class = "info-row", "3. Analyse mensuelle : importez le classeur nettoy\u00e9 (sortie de l'\u00e9tape \"Finale\") pour produire l'analyse compl\u00e8te (statistiques descriptives, PMAS/MEB, Score de Fonctionnalit\u00e9 des March\u00e9s). L'outil Kobo n'est pas \u00e0 r\u00e9importer : il est d\u00e9j\u00e0 inclus dans ce classeur (feuilles \"survey\"/\"choices\"). La feuille \"Donn\u00e9es nettoy\u00e9es\" du classeur produit a automatiquement ses codes de r\u00e9ponse remplac\u00e9s par leurs libell\u00e9s (plus besoin d'une \u00e9tape d'\u00e9tiquetage s\u00e9par\u00e9e).")
            ),

            br(),

            ## methodologie complete (PMAS/MEB + SFM/MFS), telechargeable independamment d'un lancement
            ## du pipeline - remplace l'ancienne boite d'info "Comment le PMAS est calcul\u00e9" (retiree,
            ## voir Etape 3 - Analyse ci-dessous) : demande utilisateur explicite, document exhaustif
            ## destine a un public interne ET externe (voir scripts/methodology.R).
            div(class = "info-box",
              div(class = "info-title", "M\u00e9thodologie compl\u00e8te (PMAS/MEB et Score de Fonctionnalit\u00e9 des March\u00e9s)"),
              div(class = "info-row", "Le d\u00e9tail exhaustif de la m\u00e9thodologie utilis\u00e9e par ce pipeline - composition et calcul du PMAS/MEB, calcul du Score de Fonctionnalit\u00e9 des March\u00e9s (SFM/MFS), et les hypoth\u00e8ses retenues pour l'analyse - est disponible en t\u00e9l\u00e9chargement ci-dessous, \u00e0 tout moment (pas besoin de lancer une analyse au pr\u00e9alable)."),
              downloadButton("download_methodology", "T\u00e9l\u00e9charger la m\u00e9thodologie compl\u00e8te (.html)", class = "run-button")
            ),

            br(),

            ## ETAPE 1 - mode -----------------------------------------------------------------------
            div(class = "step-box",
              div(class = "step-title", "Etape 1 - Que voulez-vous faire ?"),
              div(class = "mode-toggle",
                radioButtons("mode_select", label = NULL, inline = TRUE,
                  choiceNames = list(
                    HTML("Suivi quotidien<span class=\"pill-sub\">(collecte en cours)</span>"),
                    HTML("Nettoyage des donn\u00e9es<span class=\"pill-sub\">(fin de collecte)</span>"),
                    HTML("Analyse mensuelle<span class=\"pill-sub\">(post-nettoyage)</span>")
                  ),
                  choiceValues = list("suivi", "nettoyage", "analyse")
                )
              ),
              conditionalPanel(
                condition = "input.mode_select == 'nettoyage'",
                br(),
                div(class = "mode-toggle submode",
                  radioButtons("submode_select", label = "Etape du nettoyage :", inline = TRUE,
                    choiceNames = list("Pour suivi (pr\u00e9parer le classeur \u00e0 v\u00e9rifier)", "Finale (int\u00e9grer le classeur rempli)"),
                    choiceValues = list("pour_suivi", "finale")
                  )
                )
              )
            ),

            ## ETAPE 2 - fichiers -----------------------------------------------------------------
            div(class = "step-box",
              div(class = "step-title", "Etape 2 - Fichiers"),
              fluidRow(
                conditionalPanel(
                  condition = "input.mode_select != 'analyse'",
                  column(6, fileInput("file_raw", "Jeu de donn\u00e9es brutes (export Kobo, .xlsx)", accept = ".xlsx", width = "100%"))
                ),
                conditionalPanel(
                  condition = "input.mode_select == 'analyse'",
                  column(6, fileInput("file_clean_data", "Base des don\u00e9es nettoy\u00e9e (sortie de l'\u00e9tape \"Finale\", .xlsx)", accept = ".xlsx", width = "100%"))
                ),
                conditionalPanel(
                  condition = "input.mode_select != 'analyse'",
                  column(6, fileInput("file_kobo", "Outil Kobo (survey/choices, .xlsx)", accept = ".xlsx", width = "100%"))
                )
              ),
              conditionalPanel(
                condition = "input.mode_select == 'nettoyage' && input.submode_select == 'finale'",
                fluidRow(
                  column(4,
                    numericInput("usd_xrate_input", "Taux de change USD/XAF actuel",
                                 value = as.numeric(DEFAULT_USD_XRATE), min = 1, step = 0.01, width = "100%"),
                    div(class = "info-row", style = "margin-top:-8px;",
                      "Le XAF/EUR est une parite fixe (jamais saisie ici) ; le XAF/USD flotte sur le marche - verifiez le taux du jour avant de lancer.")
                  )
                )
              )
            ),

            ## ETAPE 3 - classeur precedent (uniquement nettoyage + pour_suivi) ----------------------
            ## Permet de relancer "pour_suivi" sur des donnees qui se sont agrandies (collecte encore
            ## en cours) sans perdre le travail de revue deja fait : les decisions deja prises pour les
            ## enquetes deja presentes avant sont reprises automatiquement dans le nouveau classeur
            ## (statut "Repris (deja traite)", voir carry_forward_decisions() dans
            ## scripts/fonctions_nettoyage.R), au lieu de devoir copier-coller a la main entre deux
            ## classeurs.
            conditionalPanel(
              condition = "input.mode_select == 'nettoyage' && input.submode_select == 'pour_suivi'",
              div(class = "step-box",
                div(class = "step-title", "Etape 3 - Nettoyage precedent (optionnel)"),
                checkboxInput("has_previous_review",
                  "Un nettoyage précédent existe-t-il pour ces données ? (collecte encore en cours, ce classeur est régénéré sur des données qui se sont agrandies)",
                  value = FALSE),
                conditionalPanel(
                  condition = "input.has_previous_review == true",
                  div(class = "info-row", style = "margin-bottom:10px;",
                    "Importez le classeur \"suivi à vérifier\" précédent (même format, rempli en partie ou en totalité) : les décisions déjà prises pour les enquêtes déjà présentes avant seront reprises automatiquement dans le nouveau classeur, seules les nouvelles lignes resteront à remplir. Les lignes déjà tapées dans Journal_additionnel/Journal_suppression sont également recopiées."),
                  fileInput("file_previous_review", "Classeur \"suivi à vérifier\" précédent (.xlsx)", accept = ".xlsx", width = "100%")
                )
              )
            ),

            ## ETAPE 3 - classeur rempli (uniquement nettoyage + finale) -----------------------------
            conditionalPanel(
              condition = "input.mode_select == 'nettoyage' && input.submode_select == 'finale'",
              div(class = "step-box",
                div(class = "step-title", "Etape 3 - Classeur rempli par l'\u00e9quipe de suivi"),
                div(class = "info-row", style = "margin-bottom:10px;",
                  "Le classeur \"suivi \u00e0 v\u00e9rifier\" (feuille \"Lisez_moi\") inclut d\u00e9sormais le journal de nettoyage additionnel (onglets Journal_additionnel/Journal_suppression) - pour une correction manuelle non signal\u00e9e par la revue automatique, remplissez-les directement dans ce m\u00eame classeur avant de le r\u00e9importer ici."),
                fileInput("file_filled", "Classeur \"suivi \u00e0 v\u00e9rifier\" rempli (.xlsx)", accept = ".xlsx", width = "100%")
              )
            ),

            ## ETAPE 3 - parametres de l'analyse (uniquement mode analyse) -----------------------------
            ## L'annee/le mois du cycle ne sont plus demandes ici : ils sont detectes automatiquement
            ## depuis la colonne "date" du classeur nettoye importe (voir get_assessment_period_from_data()
            ## dans scripts/fonctions_analyse.R) - cette information est deja presente dans le classeur
            ## (nom de fichier ET feuille README), inutile de la faire retaper.
            conditionalPanel(
              condition = "input.mode_select == 'analyse'",
              div(class = "step-box",
                div(class = "step-title", "Etape 3 - Comparaison optionnelle au mois précédent"),
                div(class = "info-row", style = "margin-bottom:10px;",
                  "Pour comparer le coût des paniers à celui du mois précédent (feuille \"Evolution des paniers du PMAS\") : importez le classeur d'ANALYSE JMMI de ce mois précédent (le fichier téléchargé à l'Étape 3 - Analyse le mois dernier, PAS le classeur nettoyé) et indiquez son nom. Laissez vide pour ignorer cette comparaison - la feuille sera alors laissée vide."),
                fluidRow(
                  column(3, textInput("analyse_mois_precedent", "Mois précédent (ex. Juillet)", width = "100%")),
                  column(6, fileInput("file_analyse_previous", "Classeur d'ANALYSE JMMI du mois précédent (.xlsx, optionnel)", accept = ".xlsx", width = "100%"))
                )
              )
            ),

            br(),
            fluidRow(
              column(3,
                tags$div(style = "position: relative;",
                  actionButton("run", "Lancer", class = "run-button"),
                  tags$div(id = "loading-spinner", style = "display:none; position:absolute; left:110%; top:50%; transform: translate(0,-50%);",
                    tags$img(src = "https://i.gifer.com/ZZ5H.gif", height = "24px", width = "24px")
                  )
                )
              ),
              column(3,
                uiOutput("download_ui")
              )
            ),

            br(),
            uiOutput("result_summary"),

            conditionalPanel(
              condition = "output.has_log",
              br(),
              div(class = "step-title", "Journal d'ex\u00e9cution"),
              tags$pre(class = "log-box", textOutput("log_text"))
            ),

            br(), br()
          )
        )
      )
    )
  )
)

## variables d'environnement lues (via Sys.getenv()) par les scripts du pipeline et leur configuration
## commune (scripts/0_configuration.R) - remises a "non definies" avant CHAQUE execution (voir
## run_script_inprocess()) pour qu'une valeur laissee par l'execution PRECEDENTE ne persiste pas par
## erreur (ex. additional_cleaning_log defini lors d'un "Lancer", absent au suivant).
ALL_PIPELINE_ENV_VARS <- c(
  "kobo", "raw_dataset", "logical_checks", "eur_xrate", "usd_xrate", "checklist_template",
  "final_workbook_dir", "suivi_final_a_remplir_dir", "suivi_final_remplis_path",
  "suivi_final_remplis_pattern", "previous_review_path",
  "kobo_tool_file", "clean_data_file", "output_dir",
  "analyse_mois_precedent", "analyse_previous_clean_data_file"
)

## execute un script du pipeline DANS LA SESSION R DE L'APP SHINY (pas de sous-processus Rscript.exe -
## voir la note d'en-tete) : isolation via un environnement frais (new.env()) plutot qu'un processus
## separe, variables de configuration passees via Sys.setenv() (equivalent du .Renviron ecrit
## auparavant pour un sous-processus), et capture de tout ce que le script imprime (cat()/print()/
## warning()) - via sink() vers un fichier temporaire - pour l'afficher dans le "Journal d'execution".
## Les erreurs stop() du script sont interceptees (tryCatch) plutot que de faire planter l'app entiere.
run_script_inprocess <- function(script_path, wd, env_vars = character(0)) {
  old_wd <- getwd()
  setwd(wd)
  on.exit(setwd(old_wd), add = TRUE)

  Sys.unsetenv(ALL_PIPELINE_ENV_VARS)
  if (length(env_vars) > 0) do.call(Sys.setenv, as.list(env_vars))
  on.exit(Sys.unsetenv(ALL_PIPELINE_ENV_VARS), add = TRUE)

  log_file <- tempfile(fileext = ".log")
  con <- file(log_file, open = "wt")
  sink(con, type = "output")
  sink(con, type = "message")

  env <- new.env()
  result <- tryCatch({
    source(script_path, local = env, echo = FALSE)
    list(ok = TRUE, error_message = NULL)
  }, error = function(e) {
    list(ok = FALSE, error_message = conditionMessage(e))
  }, finally = {
    sink(type = "message")
    sink(type = "output")
    close(con)
  })

  log_lines <- if (file.exists(log_file)) readLines(log_file, warn = FALSE) else character(0)
  unlink(log_file)
  if (!result$ok) log_lines <- c(log_lines, paste("Erreur :", result$error_message))

  # error_message renvoye SEPAREMENT du log complet (pas seulement noye dedans) pour que le message
  # d'erreur affiche a l'ecran (result_summary, voir diagnose_error()) puisse l'inclure directement,
  # plutot que de forcer l'utilisateur a deplier le journal d'execution pour voir le detail (uuid/
  # variable concernes, etc.) des erreurs de validation ligne par ligne (stop() avec message detaille).
  list(ok = result$ok, log = paste(log_lines, collapse = "\n"), error_message = result$error_message)
}

## construit et lance une execution isolee du pipeline dans un sous-dossier temporaire, EN PROCESSUS
## (voir run_script_inprocess()). Fonction de haut niveau (hors server()) pour pouvoir etre testee
## independamment de l'app.
run_pipeline <- function(kind, raw_data_path, kobo_path, filled_path = NULL, previous_review_path = NULL,
                         usd_xrate = DEFAULT_USD_XRATE) {

  script_basename <- if (kind == "suivi") "1_suivi_quotidien" else "2_nettoyage_final"

  tmp <- file.path(tempdir(), paste0("jmmi_run_", gsub("[^0-9]", "", format(Sys.time(), "%Y%m%d%H%M%OS3"))))
  dir.create(file.path(tmp, "input"), recursive = TRUE)
  dir.create(file.path(tmp, "ressources"), recursive = TRUE)
  dir.create(file.path(tmp, "scripts"), recursive = TRUE)
  dir.create(file.path(tmp, "output"), recursive = TRUE)

  # scripts et ressources fixes : copies ephemeres depuis les fichiers deja embarques dans ce dossier
  # (voir note d'en-tete - ce sont des copies de rca_jmmi_cleaning, autonomes)
  file.copy(list.files(file.path(PIPELINE_ROOT, "scripts"), full.names = TRUE),
            file.path(tmp, "scripts"), overwrite = TRUE)
  file.copy(file.path(PIPELINE_ROOT, "ressources", "check_list_august.xlsx"),
            file.path(tmp, "ressources", "check_list_august.xlsx"), overwrite = TRUE)
  file.copy(file.path(PIPELINE_ROOT, "ressources", CHECKLIST_TEMPLATE_NAME),
            file.path(tmp, "ressources", CHECKLIST_TEMPLATE_NAME), overwrite = TRUE)
  file.copy(file.path(PIPELINE_ROOT, paste0(script_basename, ".R")),
            file.path(tmp, paste0(script_basename, ".R")), overwrite = TRUE)

  # verification explicite que les ressources embarquees (pas les fichiers importes par l'utilisateur)
  # sont bien presentes ET non vides APRES la copie - sinon file.copy() echoue silencieusement
  # (renvoie FALSE sans erreur) si la source n'existe pas, ce qui produirait plus tard une erreur
  # generique et trompeuse ("fichier importe corrompu") au moment de lire ces fichiers, alors que le
  # vrai probleme est que ressources/ n'a pas ete inclus lors de la publication de l'app (ex. Posit
  # Connect) - message precis ici pour eviter de faire chercher le probleme du mauvais cote.
  required_resources <- c(
    file.path(tmp, "ressources", "check_list_august.xlsx"),
    file.path(tmp, "ressources", CHECKLIST_TEMPLATE_NAME),
    file.path(tmp, paste0(script_basename, ".R"))
  )
  missing_resources <- required_resources[!file.exists(required_resources) | file.size(required_resources) %in% c(0, NA)]
  if (length(missing_resources) > 0) {
    stop("Fichier(s) de configuration manquant(s) ou vide(s) dans le deploiement de l'application : ",
        paste(basename(missing_resources), collapse = ", "),
        " - le dossier ressources/ (et les scripts .R) n'ont probablement pas ete inclus lors de la publication.")
  }

  # fichiers importes par l'utilisateur (noms fixes, quel que soit le nom d'origine)
  file.copy(raw_data_path, file.path(tmp, "input", "raw_data.xlsx"), overwrite = TRUE)
  file.copy(kobo_path, file.path(tmp, "ressources", "kobo_tool.xlsx"), overwrite = TRUE)

  if (!is.null(filled_path)) {
    dir.create(file.path(tmp, "output", "3_classeurs_remplis"), recursive = TRUE)
    file.copy(filled_path, file.path(tmp, "output", "3_classeurs_remplis", "classeur_suivi_a_verifier.xlsx"), overwrite = TRUE)
  }

  # classeur "suivi a verifier" d'un tour PRECEDENT du pipeline (etape "pour_suivi" uniquement -
  # collecte encore en cours, ce nouveau classeur va etre regenere sur des donnees qui se sont
  # agrandies) : voir carry_forward_decisions() dans scripts/fonctions_nettoyage.R, appelee depuis
  # 2_nettoyage_final.R quand previous_review_path (ci-dessous) est renseigne.
  if (!is.null(previous_review_path)) {
    file.copy(previous_review_path, file.path(tmp, "input", "previous_review.xlsx"), overwrite = TRUE)
  }

  # variables de configuration lues par le script via Sys.getenv() (voir run_script_inprocess())
  env_vars <- c(
    kobo = file.path(tmp, "ressources", "kobo_tool.xlsx"),
    raw_dataset = file.path(tmp, "input", "raw_data.xlsx"),
    logical_checks = file.path(tmp, "ressources", "check_list_august.xlsx"),
    eur_xrate = DEFAULT_EUR_XRATE,
    usd_xrate = as.character(usd_xrate),
    checklist_template = file.path(tmp, "ressources", CHECKLIST_TEMPLATE_NAME),
    final_workbook_dir = paste0(file.path(tmp, "output", "4_final_donnees_nettoyees"), "/"),
    suivi_final_a_remplir_dir = paste0(file.path(tmp, "output", "2_a_verifier"), "/"),
    suivi_final_remplis_path = paste0(file.path(tmp, "output", "3_classeurs_remplis"), "/"),
    suivi_final_remplis_pattern = "classeur_suivi_a_verifier.xlsx$",
    # chaine vide (pas simplement omise) si non fourni : coherent avec Sys.getenv() qui renvoie "" pour
    # une variable non definie - previous_review_path est de toute facon deja remise a "non definie"
    # avant CHAQUE execution par Sys.unsetenv(ALL_PIPELINE_ENV_VARS) (voir run_script_inprocess()).
    previous_review_path = if (!is.null(previous_review_path)) file.path(tmp, "input", "previous_review.xlsx") else ""
  )

  res <- run_script_inprocess(file.path(tmp, paste0(script_basename, ".R")), wd = tmp, env_vars = env_vars)
  ok <- res$ok
  log_out <- res$log

  out_dir <- switch(kind,
    suivi = file.path(tmp, "output", "1_suivi_couverture_et_prix"),
    pour_suivi = file.path(tmp, "output", "2_a_verifier"),
    finale = file.path(tmp, "output", "4_final_donnees_nettoyees")
  )
  pattern <- switch(kind,
    suivi = "_maj_quotidienne_.*\\.xlsx$",
    pour_suivi = "suivi_a_verifier.*\\.xlsx$",
    finale = "^REACH_CAR_Dataset_JMMI_.*\\.xlsx$"
  )
  files <- if (dir.exists(out_dir)) list.files(out_dir, pattern = pattern, full.names = TRUE) else character()
  output_file <- if (length(files) > 0) files[1] else NULL

  if (ok && is.null(output_file)) {
    ok <- FALSE
    log_out <- c(log_out, "Le script s'est termin\u00e9 sans erreur mais aucun fichier de sortie attendu n'a \u00e9t\u00e9 trouv\u00e9.")
  }

  list(ok = ok, log = paste(log_out, collapse = "\n"), file = output_file, error_message = res$error_message)
}

## meme principe que run_pipeline() (dossier temporaire isole, execution EN PROCESSUS - voir
## run_script_inprocess()), pour l'analyse mensuelle complete (descriptive + PMAS/MEB + Score de
## Fonctionnalite) : prend en entree le
## classeur final deja produit par le nettoyage (pas les donnees brutes). L'annee/le mois du cycle sont
## detectes automatiquement depuis ce classeur (voir get_assessment_period_from_data() dans
## scripts/fonctions_analyse.R), pas passes ici. `previous_clean_path`/`previous_mois` sont optionnels
## (NULL) - voir 4_analyse_complete.R : la feuille "Evolution des paniers du PMAS" reste vide si absents.
## Pas de parametre kobo_path : l'outil kobo n'est PAS reimporte separement, ses feuilles "survey"/
## "choices" sont deja clonees telles quelles dans clean_data_path par build_final_workbook() (voir
## scripts/fonctions_nettoyage.R) - on les reextrait ci-dessous plutot que de redemander a l'utilisateur
## un fichier qu'il a deja fourni a l'etape de nettoyage.
run_analyse_pipeline <- function(clean_data_path,
                                 previous_clean_path = NULL, previous_mois = NULL,
                                 usd_xrate = DEFAULT_USD_XRATE) {

  tmp <- file.path(tempdir(), paste0("jmmi_analyse_", gsub("[^0-9]", "", format(Sys.time(), "%Y%m%d%H%M%OS3"))))
  dir.create(file.path(tmp, "input"), recursive = TRUE)
  dir.create(file.path(tmp, "scripts"), recursive = TRUE)
  dir.create(file.path(tmp, "output"), recursive = TRUE)

  dir.create(file.path(tmp, "ressources"), recursive = TRUE)

  file.copy(file.path(PIPELINE_ROOT, "scripts", "fonctions_analyse.R"),
            file.path(tmp, "scripts", "fonctions_analyse.R"), overwrite = TRUE)
  file.copy(file.path(PIPELINE_ROOT, "scripts", "mfs.R"),
            file.path(tmp, "scripts", "mfs.R"), overwrite = TRUE)
  # mise en forme de la feuille "Analyse" (voir format_analysis_results() dans 4_analyse_complete.R)
  file.copy(file.path(PIPELINE_ROOT, "scripts", "format_results.R"),
            file.path(tmp, "scripts", "format_results.R"), overwrite = TRUE)
  # relabel_choices()/relabel_multiple_choice_subcolumn_labels() - re-etiquetage automatique de "Données
  # nettoyées" (voir 4_analyse_complete.R) - demande utilisateur explicite : plus une etape separee.
  file.copy(file.path(PIPELINE_ROOT, "scripts", "fonctions_relabel.R"),
            file.path(tmp, "scripts", "fonctions_relabel.R"), overwrite = TRUE)
  file.copy(file.path(PIPELINE_ROOT, "4_analyse_complete.R"),
            file.path(tmp, "4_analyse_complete.R"), overwrite = TRUE)
  # modele de checklist (voir CHECKLIST_TEMPLATE_NAME) : sa feuille "CHECKLIST_Data_Cleaning" vierge est
  # clonee telle quelle dans le classeur d'analyse (onglet "Checklist_Data_Analysis", juste apres la
  # feuille "CHECKLIST_Data_Cleaning" deja auto-remplie issue du nettoyage) - demande utilisateur.
  file.copy(file.path(PIPELINE_ROOT, "ressources", CHECKLIST_TEMPLATE_NAME),
            file.path(tmp, "ressources", CHECKLIST_TEMPLATE_NAME), overwrite = TRUE)

  required_resources <- c(
    file.path(tmp, "scripts", "fonctions_analyse.R"),
    file.path(tmp, "scripts", "mfs.R"),
    file.path(tmp, "scripts", "format_results.R"),
    file.path(tmp, "scripts", "fonctions_relabel.R"),
    file.path(tmp, "4_analyse_complete.R"),
    file.path(tmp, "ressources", CHECKLIST_TEMPLATE_NAME)
  )
  missing_resources <- required_resources[!file.exists(required_resources) | file.size(required_resources) %in% c(0, NA)]
  if (length(missing_resources) > 0) {
    stop("Fichier(s) de configuration manquant(s) ou vide(s) dans le deploiement de l'application : ",
        paste(basename(missing_resources), collapse = ", "),
        " - le dossier scripts/ (et les fichiers .R) n'ont probablement pas ete inclus lors de la publication.")
  }

  file.copy(clean_data_path, file.path(tmp, "input", "clean_data.xlsx"), overwrite = TRUE)

  kobo_sheets <- tryCatch(readxl::excel_sheets(clean_data_path), error = function(e) character(0))
  if (!all(c("survey", "choices") %in% kobo_sheets)) {
    stop("Le classeur nettoyé importé ne contient pas les feuilles \"survey\"/\"choices\" attendues - importez le classeur produit par l'étape \"Finale\" du nettoyage (pas un autre fichier).")
  }
  kobo_survey_raw <- readxl::read_excel(clean_data_path, sheet = "survey", col_types = "text")
  kobo_choices_raw <- readxl::read_excel(clean_data_path, sheet = "choices", col_types = "text")
  writexl::write_xlsx(list(survey = kobo_survey_raw, choices = kobo_choices_raw),
                      file.path(tmp, "input", "kobo_tool.xlsx"))

  env_vars <- c(
    clean_data_file = file.path(tmp, "input", "clean_data.xlsx"),
    kobo_tool_file = file.path(tmp, "input", "kobo_tool.xlsx"),
    eur_xrate = DEFAULT_EUR_XRATE,
    usd_xrate = as.character(usd_xrate),
    output_dir = paste0(file.path(tmp, "output"), "/"),
    checklist_template = file.path(tmp, "ressources", CHECKLIST_TEMPLATE_NAME)
  )
  if (!is.null(previous_clean_path)) {
    file.copy(previous_clean_path, file.path(tmp, "input", "previous_clean_data.xlsx"), overwrite = TRUE)
    env_vars["analyse_previous_clean_data_file"] <- file.path(tmp, "input", "previous_clean_data.xlsx")
    env_vars["analyse_mois_precedent"] <- if (!is.null(previous_mois)) previous_mois else ""
  }

  res <- run_script_inprocess(file.path(tmp, "4_analyse_complete.R"), wd = tmp, env_vars = env_vars)
  ok <- res$ok
  log_out <- res$log

  out_dir <- file.path(tmp, "output")
  # nom de fichier "REACH_CAR_Analyse_JMMI_<Mois>-<Annee>.xlsx" (voir out_path dans 4_analyse_complete.R,
  # demande utilisateur - remplace l'ancien "<date>_CAR_JMMI_Dataset_Analyse_<mois>.xlsx")
  files <- if (dir.exists(out_dir)) list.files(out_dir, pattern = "^REACH_CAR_Analyse_JMMI_.*\\.xlsx$", full.names = TRUE) else character()
  output_file <- if (length(files) > 0) files[1] else NULL

  if (ok && is.null(output_file)) {
    ok <- FALSE
    log_out <- c(log_out, "Le script s'est terminé sans erreur mais aucun fichier de sortie attendu n'a été trouvé.")
  }

  list(ok = ok, log = paste(log_out, collapse = "\n"), file = output_file, error_message = res$error_message)
}

## traduit le journal d'exécution technique (texte brut du sous-processus Rscript.exe) en un diagnostic
## court, spécifique et actionnable en français, affiché AU-DESSUS du journal brut en cas d'échec (voir
## result_summary dans server()). Fonctionne par reconnaissance de motifs connus (messages d'erreur
## définis dans ce dossier, ou erreurs techniques R/Excel courantes) ; si aucun motif ne correspond,
## renvoie un message générique adapté à l'étape en cours plutôt que de laisser l'utilisateur seul face
## au texte technique brut.
ERROR_PATTERNS <- list(
  list(
    pattern = "Doublons detectes \\(identifiant d'enquete identique\\)",
    titre = "Doublons dans les données brutes",
    message = "Le fichier de données brutes contient plusieurs lignes avec le même identifiant d'enquête (uuid). C'est un problème technique d'export Kobo, pas quelque chose que l'équipe de suivi peut corriger dans un classeur.",
    action = "Ré-exportez les données depuis Kobo (ou contactez le support technique) pour obtenir un export sans doublons, puis relancez le traitement."
  ),
  list(
    pattern = "Certaines combinaisons enquete/variable sont signalees plus d'une fois",
    titre = "Le classeur rempli contient des lignes en double",
    message = "Dans le classeur \"suivi à vérifier\" renvoyé par l'équipe de suivi, une même combinaison enquête + variable apparaît plusieurs fois dans la feuille A_verifier.",
    action = "Ouvrez le classeur rempli, repérez les lignes dupliquées (même identifiant d'enquête et même variable dans les colonnes grisées) et supprimez ou harmonisez les doublons, puis réimportez le classeur corrigé."
  ),
  list(
    pattern = "Doublons detectes dans le journal de nettoyage des reponses autre",
    titre = "Doublons dans le recodage des réponses \"autre\"",
    message = "La feuille reponses_autres du classeur rempli contient plusieurs recodages différents pour la même réponse \"autre\" de la même enquête.",
    action = "Dans la feuille reponses_autres, vérifiez qu'une seule case (VRAI / EXISTANT / INVALIDE) a été remplie par ligne, puis réimportez le classeur corrigé."
  ),
  list(
    pattern = "Le journal de nettoyage additionnel contient des erreurs",
    kinds = "finale",
    titre = "Erreurs dans le journal de nettoyage additionnel",
    message = "L'onglet Journal_additionnel ou Journal_suppression du classeur \"suivi à vérifier\" rempli contient une ou plusieurs lignes invalides - le détail complet (numéro de ligne, identifiant d'enquête, variable concernée) est indiqué dans le journal d'exécution ci-dessous.",
    action = "Corrigez chaque ligne signalée dans le journal d'exécution ci-dessous, réenregistrez le classeur, puis réimportez-le."
  ),
  list(
    pattern = "n'a pas les colonnes attendues",
    kinds = "finale",
    titre = "Onglet Journal_additionnel/Journal_suppression modifié",
    message = "L'onglet Journal_additionnel ou Journal_suppression du classeur \"suivi à vérifier\" rempli n'a pas les colonnes attendues.",
    action = "N'ajoutez/ne renommez pas de colonnes dans ces onglets - repartez du classeur \"suivi à vérifier\" tel qu'il a été généré, puis réimportez-le."
  ),
  list(
    pattern = "Fichier\\(s\\) de configuration manquant\\(s\\) ou vide\\(s\\)",
    titre = "Fichiers de l'application manquants",
    message = "Un fichier interne à l'application (dans ressources/ ou un des scripts .R) est manquant ou vide sur le serveur - ce n'est pas un problème avec vos fichiers importés.",
    action = "Vérifiez que TOUT le contenu du dossier de l'application (y compris ressources/ et les fichiers .R) a bien été inclus lors de la publication, puis republiez l'application."
  ),
  list(
    pattern = "not found in workbook|Sheet .* not found|n'existe pas dans (ce|le) classeur",
    titre = "Onglet manquant dans le fichier importé",
    message = "Le fichier importé ne contient pas l'onglet (la feuille Excel) attendu - par exemple \"survey\" ou \"choices\" pour l'outil Kobo. C'est probablement le mauvais fichier, ou une version incomplète de l'outil Kobo.",
    action = "Vérifiez que chaque fichier est importé dans la bonne case (outil Kobo dans \"Outil Kobo\", données dans \"Jeu de données\"), puis réessayez."
  ),
  list(
    pattern = "not a zip file|invalid zip|corrupt workbook|Corrupt|zip end of central directory",
    titre = "Fichier illisible ou corrompu",
    message = "Le fichier importé n'a pas pu être ouvert comme un classeur Excel valide (.xlsx). Il est peut-être corrompu, encore au format .xls, ou le téléchargement a été interrompu.",
    action = "Ouvrez le fichier vous-même dans Excel pour vérifier qu'il s'ouvre correctement, ré-enregistrez-le au format .xlsx si besoin, puis réimportez-le."
  ),
  list(
    pattern = "object '[^']+' not found|could not find function",
    titre = "Colonne ou variable manquante",
    message = "Le traitement s'attend à trouver une colonne ou une variable qui est absente du fichier importé - le plus souvent parce que l'outil Kobo importé ne correspond pas exactement aux données (version différente de celle utilisée pour la collecte).",
    action = "Vérifiez que l'outil Kobo importé est bien la version utilisée pour CETTE collecte, et que le fichier de données provient bien du même outil, puis réessayez."
  ),
  list(
    pattern = "Le script s'est termin.{1,2} sans erreur mais aucun fichier de sortie",
    kinds = c("suivi", "pour_suivi"),
    titre = "Aucun résultat produit",
    message = "Le traitement s'est déroulé sans erreur technique, mais n'a produit aucun fichier de sortie - le plus souvent parce que les données importées sont vides, ou que la colonne \"date\" n'est pas correctement renseignée dans le fichier de données brutes.",
    action = "Vérifiez que le fichier de données brutes contient bien des enquêtes avec une date renseignée, puis réessayez."
  ),
  list(
    pattern = "Le script s'est termin.{1,2} sans erreur mais aucun fichier de sortie",
    kinds = "finale",
    titre = "Aucun classeur final produit",
    message = "Le traitement s'est déroulé sans erreur technique, mais n'a pas produit de classeur final - le plus souvent parce que le classeur \"suivi à vérifier\" importé ne correspond pas à celui généré à la première exécution (fichier renommé, ou provenant d'un autre jeu de données).",
    action = "Réimportez le classeur \"suivi à vérifier\" tel qu'il a été téléchargé et rempli, sans le renommer, en vous assurant qu'il correspond bien au même fichier de données brutes que la première exécution."
  ),
  list(
    pattern = "Classeur d'entree introuvable : le classeur final nettoye",
    kinds = "analyse",
    titre = "Mauvais fichier importé pour l'analyse",
    message = "Le fichier importé comme \"classeur final nettoyé\" n'a pas pu être utilisé. Cette étape attend en entrée le classeur FINAL produit par le nettoyage (\"Nettoyage des données\" > \"Finale\"), pas les données brutes ni un autre type de fichier.",
    action = "Importez le fichier REACH_CAR_Dataset_JMMI_<Mois>-<Année>.xlsx téléchargé à la fin de l'étape \"Finale\" du nettoyage, puis relancez l'analyse."
  ),
  list(
    pattern = "The following analysis variables are not present",
    kinds = "analyse",
    titre = "Outil kobo incompatible avec le classeur importé",
    message = "L'outil kobo importé décrit des questions qui n'existent pas dans le classeur final importé - l'outil kobo ne correspond probablement pas à celui utilisé pour produire ce classeur.",
    action = "Vérifiez que l'outil kobo importé est bien la version utilisée pour produire CE classeur final, puis réessayez."
  ),
  list(
    pattern = "Taux de change XAF/EUR invalide",
    kinds = "analyse",
    titre = "Taux de change invalide",
    message = "Le taux de change XAF/EUR configuré dans l'application n'est pas un nombre valide.",
    action = "Contactez la personne en charge du tableau de bord - il s'agit d'un problème de configuration de l'application, pas de vos fichiers importés."
  )
)

## `error_message` (optionnel) : le message exact du stop() qui a interrompu le script (voir
## run_script_inprocess()) - pour les erreurs de validation ligne par ligne (voir fonctions_nettoyage.R,
## ex. validate_additional_cleaning_log(), read_review_sheet()), ce message contient deja le detail
## precis (uuid, variable, valeur en cause) pour CHAQUE ligne fautive. Toujours affiche directement dans
## le message a l'ecran plutot que de forcer l'utilisateur a deplier le journal d'execution pour le voir.
diagnose_error <- function(kind, log_text, error_message = NULL) {
  diag <- NULL
  for (p in ERROR_PATTERNS) {
    if (!is.null(p$kinds) && !identical(kind, "") && !(kind %in% p$kinds)) next
    if (grepl(p$pattern, log_text, ignore.case = TRUE)) {
      diag <- list(titre = p$titre, message = p$message, action = p$action)
      break
    }
  }

  if (is.null(diag)) {
    generic_message <- switch(kind,
      suivi = "Causes fréquentes : le fichier de données brutes ne correspond pas à l'outil Kobo importé, ou la colonne \"date\" est manquante/mal renseignée.",
      pour_suivi = "Causes fréquentes : le fichier de données brutes ne correspond pas à l'outil Kobo importé, ou un fichier n'est pas au format .xlsx attendu.",
      finale = "Causes fréquentes : le classeur \"suivi à vérifier\" ne correspond pas au même fichier de données brutes que la première exécution, ou une ligne du classeur rempli est incohérente (voir le détail ci-dessous pour l'identifiant d'enquête concerné).",
      analyse = "Causes fréquentes : le classeur nettoyé importé ne provient pas de l'étape \"Finale\" du nettoyage (il doit contenir les feuilles \"survey\"/\"choices\"), ou le classeur du mois précédent (optionnel) n'est pas le bon fichier.",
      "Causes fréquentes : fichiers importés au mauvais endroit, fichier non conforme au format .xlsx, ou incohérence entre les données et l'outil Kobo."
    )

    diag <- list(
      titre = "Erreur inattendue",
      message = paste("Une erreur est survenue pendant le traitement.", generic_message),
      action = "Vérifiez vos fichiers importés puis réessayez. Si le problème persiste, transmettez le journal d'exécution ci-dessous à la personne en charge du tableau de bord."
    )
  }

  if (!is.null(error_message) && nzchar(error_message)) {
    diag$message <- paste0(diag$message, "\n\nDétail : ", error_message)
  }

  diag
}

#### SERVER ####
server <- function(input, output, session) {

  result <- reactiveValues(ok = NA, log = "", file = NULL, kind = NULL, error_message = NULL)

  observeEvent(input$run, {

    if (input$mode_select != "analyse" && is.null(input$file_kobo)) {
      showNotification("Veuillez importer l'outil Kobo.", type = "error")
      return(invisible())
    }

    if (input$mode_select == "nettoyage" && identical(input$submode_select, "finale") &&
       (is.null(input$usd_xrate_input) || is.na(input$usd_xrate_input) || input$usd_xrate_input <= 0)) {
      showNotification("Veuillez indiquer le taux de change USD/XAF actuel.", type = "error")
      return(invisible())
    }

    if (input$mode_select == "analyse") {
      if (is.null(input$file_clean_data)) {
        showNotification("Veuillez importer le classeur nettoy\u00e9.", type = "error")
        return(invisible())
      }
    } else if (is.null(input$file_raw)) {
      showNotification("Veuillez importer le jeu de donn\u00e9es brutes.", type = "error")
      return(invisible())
    }

    kind <- if (input$mode_select == "suivi") "suivi" else if (input$mode_select == "analyse") "analyse" else input$submode_select
    if (input$mode_select == "nettoyage" && is.null(input$submode_select)) {
      showNotification("Veuillez choisir 'Pour suivi' ou 'Finale'.", type = "error")
      return(invisible())
    }
    if (identical(kind, "finale") && is.null(input$file_filled)) {
      showNotification("Veuillez importer le classeur rempli par l'\u00e9quipe de suivi.", type = "error")
      return(invisible())
    }
    if (identical(kind, "pour_suivi") && isTRUE(input$has_previous_review) && is.null(input$file_previous_review)) {
      showNotification("Veuillez importer le classeur \"suivi \u00e0 v\u00e9rifier\" pr\u00e9c\u00e9dent, ou d\u00e9cochez la case.", type = "error")
      return(invisible())
    }
    shinyjs::runjs("$('#loading-spinner').show();")
    result$ok <- NA
    result$file <- NULL

    res <- tryCatch(
      if (identical(kind, "analyse")) {
        run_analyse_pipeline(
          clean_data_path = input$file_clean_data$datapath,
          previous_clean_path = if (!is.null(input$file_analyse_previous)) input$file_analyse_previous$datapath else NULL,
          previous_mois = if (nzchar(trimws(input$analyse_mois_precedent))) trimws(input$analyse_mois_precedent) else NULL,
          usd_xrate = DEFAULT_USD_XRATE
        )
      } else {
        run_pipeline(kind,
          raw_data_path = input$file_raw$datapath,
          kobo_path = input$file_kobo$datapath,
          filled_path = if (identical(kind, "finale")) input$file_filled$datapath else NULL,
          previous_review_path = if (identical(kind, "pour_suivi") && isTRUE(input$has_previous_review) &&
                                     !is.null(input$file_previous_review)) input$file_previous_review$datapath else NULL,
          usd_xrate = if (identical(kind, "finale")) input$usd_xrate_input else DEFAULT_USD_XRATE
        )
      },
      error = function(e) list(ok = FALSE, log = paste("Erreur :", conditionMessage(e)), file = NULL,
                               error_message = conditionMessage(e))
    )

    shinyjs::runjs("$('#loading-spinner').hide();")

    result$ok <- res$ok
    result$log <- res$log
    result$file <- res$file
    result$kind <- kind
    result$error_message <- res$error_message

    if (isTRUE(res$ok)) {
      showNotification("Termin\u00e9 - le fichier est pr\u00eat au t\u00e9l\u00e9chargement.", type = "message")
    } else {
      showNotification("Une erreur est survenue - consultez le journal d'ex\u00e9cution.", type = "error")
    }
  })

  output$has_log <- reactive({ nzchar(result$log) })
  outputOptions(output, "has_log", suspendWhenHidden = FALSE)

  output$log_text <- renderText({ result$log })

  output$result_summary <- renderUI({
    if (is.na(result$ok)) return(NULL)
    if (isTRUE(result$ok)) {
      div(class = "info-box",
        div(class = "info-title", "Succ\u00e8s"),
        div(class = "info-row", paste0("Fichier pr\u00eat : ", basename(result$file)))
      )
    } else {
      diag <- diagnose_error(result$kind, result$log, result$error_message)
      div(class = "red-box",
        div(class = "red-title", diag$titre),
        div(class = "red-row", diag$message),
        div(class = "red-line"),
        div(class = "red-row", tags$b("Que faire : "), diag$action)
      )
    }
  })

  output$download_ui <- renderUI({
    if (!isTRUE(result$ok) || is.null(result$file)) return(NULL)
    downloadButton("download", "T\u00e9l\u00e9charger le fichier", class = "run-button")
  })

  # Le nom du fichier telecharge est simplement celui produit par le pipeline (qui inclut deja le mois
  # et un horodatage complet pour le suivi quotidien) - pas de suffixe ajoute ici, pour eviter de
  # dupliquer l'heure dans le nom.
  output$download <- downloadHandler(
    filename = function() basename(result$file),
    content = function(file) file.copy(result$file, file, overwrite = TRUE)
  )

  # document de methodologie (PMAS/MEB + SFM/MFS) : disponible independamment d'un lancement du
  # pipeline (pas de dependance a `result`) - genere a la demande via build_methodology_html()
  # (scripts/methodology.R), pour que les chiffres PMAS/MEB affiches ne puissent jamais diverger des
  # constantes reellement utilisees par le calcul au moment du telechargement.
  output$download_methodology <- downloadHandler(
    filename = function() "Methodologie_JMMI_RCA_PMAS-MFS.html",
    content = function(file) writeLines(build_methodology_html(), file, useBytes = TRUE)
  )
}

#### APP RUNNING ####
shiny::shinyApp(ui = ui, server = server)
