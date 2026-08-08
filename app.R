################################################################################
# JMMI RCA - TABLEAU DE BORD (habillage Shiny du pipeline de nettoyage)
################################################################################
# Ce dossier est AUTONOME (self-contained) : 1_suivi_quotidien.R, 2_nettoyage_final.R, 3_relabel_choices.R,
# scripts/ et ressources/ sont des COPIES des projets rca_jmmi_cleaning et jmmi_relabelling (sources
# d'origine, a cote), placees ici pour que ce dossier puisse etre deploye/zippe/partage seul, sans
# dependre des dossiers voisins a l'execution.
# Si le pipeline de rca_jmmi_cleaning ou l'outil jmmi_relabelling evoluent, il faut recopier ces fichiers
# ici pour rester a jour (pas de synchronisation automatique).
# A chaque clic sur "Lancer", l'app copie ces scripts/ressources dans un dossier temporaire isole
# (propre a cette execution), y place les fichiers importes par l'utilisateur, ecrit un .Renviron
# pointant vers ces fichiers, puis lance le script demande dans un PROCESSUS R SEPARE (Rscript.exe) -
# exactement comme en ligne de commande. Cela evite toute pollution de l'environnement global de l'app
# Shiny et isole totalement les executions les unes des autres.
################################################################################

library(shiny)
library(shinyjs)
library(shinythemes)
library(htmltools)

## chemins fixes ---------------------------------------------------------------------------------
PIPELINE_ROOT <- normalizePath(".")
RSCRIPT_PATH <- "C:/Program Files/R/R-4.5.3/bin/Rscript.exe"
CHECKLIST_TEMPLATE_NAME <- "Checklist-Template_Nettoyage-Donnees-Quant_Normes-Min_v1_2_2026-03_FRENCH.xlsx"

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
              div(class = "info-row", "3. \u00c9tiquetage des choix : importez l'outil Kobo et le classeur nettoy\u00e9 (sortie de l'\u00e9tape \"Finale\" ci-dessus) pour remplacer les codes de r\u00e9ponse par leurs libell\u00e9s.")
            ),

            br(),

            ## ETAPE 1 - mode -----------------------------------------------------------------------
            div(class = "step-box",
              div(class = "step-title", "Etape 1 - Que voulez-vous faire ?"),
              div(class = "mode-toggle",
                radioButtons("mode_select", label = NULL, inline = TRUE,
                  choiceNames = list("Suivi quotidien (collecte en cours)", "Nettoyage des donn\u00e9es (fin de collecte)", "\u00c9tiquetage des choix (post-nettoyage)"),
                  choiceValues = list("suivi", "nettoyage", "etiquetage")
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
                  condition = "input.mode_select != 'etiquetage'",
                  column(6, fileInput("file_raw", "Jeu de donn\u00e9es brutes (export Kobo, .xlsx)", accept = ".xlsx", width = "100%"))
                ),
                conditionalPanel(
                  condition = "input.mode_select == 'etiquetage'",
                  column(6, fileInput("file_clean_data", "Base des don\u00e9es nettoy\u00e9e (sortie de l'\u00e9tape \"Finale\", .xlsx)", accept = ".xlsx", width = "100%"))
                ),
                column(6, fileInput("file_kobo", "Outil Kobo (survey/choices, .xlsx)", accept = ".xlsx", width = "100%"))
              )
            ),

            ## ETAPE 3 - classeur rempli (uniquement nettoyage + finale) -----------------------------
            conditionalPanel(
              condition = "input.mode_select == 'nettoyage' && input.submode_select == 'finale'",
              div(class = "step-box",
                div(class = "step-title", "Etape 3 - Classeur rempli par l'\u00e9quipe de suivi"),
                fileInput("file_filled", "Classeur \"suivi \u00e0 v\u00e9rifier\" rempli (.xlsx)", accept = ".xlsx", width = "100%")
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

## construit et lance une execution isolee du pipeline dans un sous-dossier temporaire + un processus
## Rscript.exe separe (jamais l'environnement de l'app Shiny elle-meme). Fonction de haut niveau
## (hors server()) pour pouvoir etre testee independamment de l'app.
run_pipeline <- function(kind, raw_data_path, kobo_path, filled_path = NULL) {

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

  # fichiers importes par l'utilisateur (noms fixes, quel que soit le nom d'origine)
  file.copy(raw_data_path, file.path(tmp, "input", "raw_data.xlsx"), overwrite = TRUE)
  file.copy(kobo_path, file.path(tmp, "ressources", "kobo_tool.xlsx"), overwrite = TRUE)

  if (!is.null(filled_path)) {
    dir.create(file.path(tmp, "output", "3_classeurs_remplis"), recursive = TRUE)
    file.copy(filled_path, file.path(tmp, "output", "3_classeurs_remplis", "classeur_suivi_a_verifier.xlsx"), overwrite = TRUE)
  }

  # .Renviron propre a cette execution - lu automatiquement par le sous-processus R au demarrage
  # (CWD = tmp), exactement comme en ligne de commande
  env_lines <- c(
    paste0('kobo = "', file.path(tmp, "ressources", "kobo_tool.xlsx"), '"'),
    paste0('raw_dataset = "', file.path(tmp, "input", "raw_data.xlsx"), '"'),
    paste0('logical_checks = "', file.path(tmp, "ressources", "check_list_august.xlsx"), '"'),
    'eur_xrate = 655.95',
    paste0('checklist_template = "', file.path(tmp, "ressources", CHECKLIST_TEMPLATE_NAME), '"'),
    paste0('final_workbook_dir = "', file.path(tmp, "output", "4_final_donnees_nettoyees"), '/"'),
    paste0('suivi_final_a_remplir_dir = "', file.path(tmp, "output", "2_a_verifier"), '/"'),
    paste0('suivi_final_remplis_path = "', file.path(tmp, "output", "3_classeurs_remplis"), '/"'),
    'suivi_final_remplis_pattern = "classeur_suivi_a_verifier.xlsx$"'
  )
  writeLines(env_lines, file.path(tmp, ".Renviron"))

  old_wd <- getwd()
  setwd(tmp)
  on.exit(setwd(old_wd), add = TRUE)

  log_out <- tryCatch(
    system2(RSCRIPT_PATH, args = shQuote(paste0(script_basename, ".R")), stdout = TRUE, stderr = TRUE),
    error = function(e) paste("ERREUR AU LANCEMENT DE R :", conditionMessage(e))
  )
  status <- attr(log_out, "status")
  ok <- is.null(status) || identical(status, 0L)

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

  list(ok = ok, log = paste(log_out, collapse = "\n"), file = output_file)
}

## nom de fichier sur - conserve le nom original importe par l'utilisateur (pour que le fichier de
## sortie du re-etiquetage suive la meme convention de nommage que le classeur d'entree, voir
## 3_relabel_choices.R), mais retire tout chemin et tout caractere invalide dans un nom de fichier
## Windows ; force l'extension .xlsx (seul format accepte par le fileInput).
safe_input_filename <- function(name) {
  base <- tools::file_path_sans_ext(basename(name))
  base <- gsub('[\\\\/:*?"<>|]', "_", base)
  if (!nzchar(base)) base <- "clean_data"
  paste0(base, ".xlsx")
}

## meme principe que run_pipeline() (dossier temporaire isole + processus Rscript.exe separe), mais
## pour l'outil de re-etiquetage des choix (independant du pipeline de nettoyage lui-meme : prend en
## entree le classeur nettoye deja produit, pas les donnees brutes).
run_relabel_pipeline <- function(clean_data_path, clean_data_name, kobo_path) {

  tmp <- file.path(tempdir(), paste0("jmmi_relabel_", gsub("[^0-9]", "", format(Sys.time(), "%Y%m%d%H%M%OS3"))))
  dir.create(file.path(tmp, "input"), recursive = TRUE)
  dir.create(file.path(tmp, "scripts"), recursive = TRUE)
  dir.create(file.path(tmp, "output"), recursive = TRUE)

  file.copy(file.path(PIPELINE_ROOT, "scripts", "fonctions_relabel.R"),
            file.path(tmp, "scripts", "fonctions_relabel.R"), overwrite = TRUE)
  file.copy(file.path(PIPELINE_ROOT, "3_relabel_choices.R"),
            file.path(tmp, "3_relabel_choices.R"), overwrite = TRUE)

  # nom d'entree preserve (voir safe_input_filename) - 3_relabel_choices.R nomme sa sortie
  # "<nom du fichier d'entree>_final.xlsx", donc conserver le nom original ici est ce qui garantit
  # que le fichier telecharge suit la convention de nommage standard au lieu d'un nom fixe generique.
  clean_data_filename <- safe_input_filename(clean_data_name)
  file.copy(clean_data_path, file.path(tmp, "input", clean_data_filename), overwrite = TRUE)
  file.copy(kobo_path, file.path(tmp, "input", "kobo_tool.xlsx"), overwrite = TRUE)

  env_lines <- c(
    paste0('kobo_tool_file = "', file.path(tmp, "input", "kobo_tool.xlsx"), '"'),
    paste0('clean_data_file = "', file.path(tmp, "input", clean_data_filename), '"'),
    paste0('output_dir = "', file.path(tmp, "output"), '/"')
  )
  writeLines(env_lines, file.path(tmp, ".Renviron"))

  old_wd <- getwd()
  setwd(tmp)
  on.exit(setwd(old_wd), add = TRUE)

  log_out <- tryCatch(
    system2(RSCRIPT_PATH, args = shQuote("3_relabel_choices.R"), stdout = TRUE, stderr = TRUE),
    error = function(e) paste("ERREUR AU LANCEMENT DE R :", conditionMessage(e))
  )
  status <- attr(log_out, "status")
  ok <- is.null(status) || identical(status, 0L)

  out_dir <- file.path(tmp, "output")
  files <- if (dir.exists(out_dir)) list.files(out_dir, pattern = "_final\\.xlsx$", full.names = TRUE) else character()
  output_file <- if (length(files) > 0) files[1] else NULL

  if (ok && is.null(output_file)) {
    ok <- FALSE
    log_out <- c(log_out, "Le script s'est terminé sans erreur mais aucun fichier de sortie attendu n'a été trouvé.")
  }

  list(ok = ok, log = paste(log_out, collapse = "\n"), file = output_file)
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
    pattern = "Aucune feuille 'raw_data' ou 'clean_data' trouvee",
    kinds = "etiquetage",
    titre = "Mauvais fichier importé pour l'étiquetage",
    message = "Le classeur importé ne contient ni feuille raw_data ni clean_data. Cette étape attend en entrée le classeur FINAL produit par le nettoyage (\"Nettoyage des données\" > \"Finale\"), pas les données brutes ni un autre type de fichier.",
    action = "Importez le fichier REACH_CAR_Dataset_JMMI_<Mois>-<Année>.xlsx téléchargé à la fin de l'étape \"Finale\" du nettoyage, puis relancez l'étiquetage."
  ),
  list(
    pattern = "ERREUR AU LANCEMENT DE R",
    titre = "Impossible de démarrer le moteur R",
    message = "L'application n'a pas réussi à lancer le processus R qui exécute le traitement (Rscript.exe).",
    action = "Ce n'est pas un problème avec vos fichiers - contactez la personne en charge de l'installation du tableau de bord (chemin de R probablement introuvable ou modifié)."
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
  )
)

diagnose_error <- function(kind, log_text) {
  for (p in ERROR_PATTERNS) {
    if (!is.null(p$kinds) && !identical(kind, "") && !(kind %in% p$kinds)) next
    if (grepl(p$pattern, log_text, ignore.case = TRUE)) {
      return(list(titre = p$titre, message = p$message, action = p$action))
    }
  }

  generic_message <- switch(kind,
    suivi = "Causes fréquentes : le fichier de données brutes ne correspond pas à l'outil Kobo importé, ou la colonne \"date\" est manquante/mal renseignée.",
    pour_suivi = "Causes fréquentes : le fichier de données brutes ne correspond pas à l'outil Kobo importé, ou un fichier n'est pas au format .xlsx attendu.",
    finale = "Causes fréquentes : le classeur \"suivi à vérifier\" ne correspond pas au même fichier de données brutes que la première exécution, ou une ligne du classeur rempli est incohérente (voir le journal ci-dessous pour l'identifiant d'enquête concerné).",
    etiquetage = "Causes fréquentes : le classeur nettoyé importé ne provient pas de l'étape \"Finale\" du nettoyage, ou l'outil Kobo importé ne correspond pas à ces données.",
    "Causes fréquentes : fichiers importés au mauvais endroit, fichier non conforme au format .xlsx, ou incohérence entre les données et l'outil Kobo."
  )

  list(
    titre = "Erreur inattendue",
    message = paste("Une erreur est survenue pendant le traitement.", generic_message),
    action = "Vérifiez vos fichiers importés puis réessayez. Si le problème persiste, transmettez le journal d'exécution ci-dessous à la personne en charge du tableau de bord."
  )
}

#### SERVER ####
server <- function(input, output, session) {

  result <- reactiveValues(ok = NA, log = "", file = NULL, kind = NULL)

  observeEvent(input$run, {

    if (is.null(input$file_kobo)) {
      showNotification("Veuillez importer l'outil Kobo.", type = "error")
      return(invisible())
    }

    if (identical(input$mode_select, "etiquetage")) {
      if (is.null(input$file_clean_data)) {
        showNotification("Veuillez importer le classeur nettoy\u00e9.", type = "error")
        return(invisible())
      }
    } else if (is.null(input$file_raw)) {
      showNotification("Veuillez importer le jeu de donn\u00e9es brutes.", type = "error")
      return(invisible())
    }

    kind <- if (input$mode_select == "suivi") "suivi" else if (input$mode_select == "etiquetage") "etiquetage" else input$submode_select
    if (input$mode_select == "nettoyage" && is.null(input$submode_select)) {
      showNotification("Veuillez choisir 'Pour suivi' ou 'Finale'.", type = "error")
      return(invisible())
    }
    if (identical(kind, "finale") && is.null(input$file_filled)) {
      showNotification("Veuillez importer le classeur rempli par l'\u00e9quipe de suivi.", type = "error")
      return(invisible())
    }

    shinyjs::runjs("$('#loading-spinner').show();")
    result$ok <- NA
    result$file <- NULL

    res <- tryCatch(
      if (identical(kind, "etiquetage")) {
        run_relabel_pipeline(
          clean_data_path = input$file_clean_data$datapath,
          clean_data_name = input$file_clean_data$name,
          kobo_path = input$file_kobo$datapath
        )
      } else {
        run_pipeline(kind,
          raw_data_path = input$file_raw$datapath,
          kobo_path = input$file_kobo$datapath,
          filled_path = if (identical(kind, "finale")) input$file_filled$datapath else NULL
        )
      },
      error = function(e) list(ok = FALSE, log = paste("Erreur :", conditionMessage(e)), file = NULL)
    )

    shinyjs::runjs("$('#loading-spinner').hide();")

    result$ok <- res$ok
    result$log <- res$log
    result$file <- res$file
    result$kind <- kind

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
      diag <- diagnose_error(result$kind, result$log)
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
}

#### APP RUNNING ####
shiny::shinyApp(ui = ui, server = server)
