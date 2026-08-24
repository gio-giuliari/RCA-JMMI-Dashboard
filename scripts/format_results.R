library(tidyverse)
library(openxlsx)

# ── private helpers ───────────────────────────────────────────────────────────

.strip_html <- function(x) gsub("<[^>]+>", "", x)

.build_section_map <- function(survey_df, label_col = "label") {
  result     <- list()
  sect_stack <- character(0)
  # "calculate" n'est PAS exclu (contrairement a la version d'origine) : les variables de prix/median
  # de ce rapport (ex. "peanut_price_item") sont de type "calculate" dans l'outil kobo - les exclure
  # les empechait d'heriter d'une section (elles tombaient toutes dans un groupe "NA" commun a tous
  # les articles, au lieu de rester rattachees a la section de leur PROPRE article), ce qui cassait a
  # la fois le regroupement par article et le tri "prix en tete de section" demandes par l'utilisateur.
  skip_types <- c("begin_repeat", "end_repeat", "note",
                  "hidden", "start", "end", "today", "deviceid", "audit")
  types <- trimws(sub(" .*", "", trimws(survey_df$type)))
  labels <- survey_df[[label_col]]
  names_ <- survey_df$name

  for (i in seq_len(nrow(survey_df))) {
    if (is.na(types[i])) next
    type_base <- types[i]
    name_val  <- names_[i]
    lbl_val   <- labels[i]

    if (type_base == "begin_group") {
      lbl <- if (!is.na(lbl_val) && nchar(trimws(lbl_val)) > 0) {
        .strip_html(lbl_val)
      } else if (!is.na(name_val)) name_val else "Section"
      sect_stack <- c(sect_stack, lbl)
    } else if (type_base == "end_group") {
      if (length(sect_stack) > 0) sect_stack <- sect_stack[-length(sect_stack)]
    } else if (!is.na(name_val) && !type_base %in% skip_types) {
      if (length(sect_stack) > 0)
        result[[name_val]] <- sect_stack[length(sect_stack)]
    }
  }
  result
}

# Repeated-per-item questions asked OUTSIDE the item's own begin_group/end_group block (ex.
# "<item>_availability_market_item" sits right before "begin_group <item>" in this kobo tool, so
# .build_section_map() assigns it to the OUTER enclosing group - "Produits alimentaires (Maïs,
# manioc, ...)" - instead of the item's own section "Arachide"/"Haricot"/etc. This makes the outer
# section header print again for every single item (once per availability question), fragmenting the
# report. Fixed generically (no item hard-coded) by reassigning any "<prefix>_availability_market_item"
# variable to whatever section its sibling "<prefix>_std_unit_item" resolves to, since that sibling is
# always the first question actually inside the item's own group.
.reassign_outside_group_sections <- function(section_map, survey_names) {
  avail_vars <- grep("_availability_market_item$", survey_names, value = TRUE)
  for (v in avail_vars) {
    anchor <- paste0(sub("_availability_market_item$", "", v), "_std_unit_item")
    if (!is.null(section_map[[anchor]])) section_map[[v]] <- section_map[[anchor]]
  }
  section_map
}

# Make a fast lookup key — NA-safe. Include analysis_type so mean/median
# rows with NA aval don't collide with each other or with proportion rows.
.key <- function(type, var, aval) paste(type, var, ifelse(is.na(aval), "__NA__", aval), sep = "|")

# Choix "non-reponse" exclus des lignes affichees (demande utilisateur) - communs a de nombreuses
# listes de choix de cet outil ("Je ne sais pas" / "Je (ne) souhaite/préfère (pas) répondre" / codes NA
# litteraux), verifie dans la feuille choices : ne portent aucune information substantielle pour un
# rapport d'analyse.
EXCLUDED_CHOICE_CODES <- c("ne_sait_pas", "prefere_pas_repondre", "na", "n/a", "not_applicable")

# ── main function ─────────────────────────────────────────────────────────────

format_analysis_results <- function(
    results_table,
    survey,
    choices,
    output_path,
    totals        = NULL,
    label_col     = "label",
    accent_colour = "#D2CBB8",
    header_bg     = "#EE5958",  # meme rouge que la ligne d'en-tete de l'onglet "Score de Fonctionnalité" (voir wb_add_fill(..., "A1:M1", color = "FFEE5958") dans 4_analyse_complete.R)
    sheet_name    = "Résultats"
) {

  # ── 0. Validate label_col ─────────────────────────────────────────────────
  if (!label_col %in% names(survey))
    stop(sprintf("Column '%s' not found in survey. Available: %s",
                 label_col, paste(names(survey), collapse = ", ")))
  if (!label_col %in% names(choices))
    stop(sprintf("Column '%s' not found in choices. Available: %s",
                 label_col, paste(names(choices), collapse = ", ")))

  # ── 1. Split grouped / ungrouped ─────────────────────────────────────────
  has_grp_col <- "group_var" %in% names(results_table)

  ungrouped <- if (!is.null(totals)) {
    totals %>% filter(is.na(group_var) | trimws(as.character(group_var)) == "")
  } else if (has_grp_col) {
    results_table %>% filter(is.na(group_var) | trimws(as.character(group_var)) == "")
  } else {
    results_table
  }

  grouped <- if (has_grp_col) {
    results_table %>% filter(!is.na(group_var) & trimws(as.character(group_var)) != "")
  } else {
    data.frame()
  }

  is_stratified <- nrow(grouped) > 0
  grp_values    <- if (is_stratified) sort(unique(as.character(grouped$group_var_value))) else character(0)
  has_total     <- nrow(ungrouped) > 0

  ref_tbl <- if (has_total) {
    ungrouped
  } else if (is_stratified) {
    grouped %>% distinct(analysis_type, analysis_var, analysis_var_value, .keep_all = TRUE)
  } else {
    results_table
  }

  # ── 2. Label lookups ──────────────────────────────────────────────────────
  var_label_map <- survey %>%
    filter(!is.na(name), !is.na(.data[[label_col]])) %>%
    transmute(name, var_label = .strip_html(.data[[label_col]])) %>%
    distinct(name, .keep_all = TRUE)

  var_list_map <- survey %>%
    filter(grepl("^select_(one|multiple)", type)) %>%
    transmute(name, list_name = gsub("^select_(one|multiple) ", "", type)) %>%
    distinct(name, .keep_all = TRUE)

  choice_label_map <- choices %>%
    select(list_name, name, choice_label = all_of(label_col)) %>%
    distinct(list_name, name, .keep_all = TRUE)

  grp_var_name <- if (is_stratified) unique(grouped$group_var)[1] else NA_character_
  grp_list_nm  <- if (!is.na(grp_var_name)) {
    var_list_map %>% filter(name == grp_var_name) %>% pull(list_name)
  } else character(0)

  grp_labels <- vapply(grp_values, function(gv) {
    if (length(grp_list_nm) > 0) {
      lbl <- choice_label_map %>%
        filter(list_name == grp_list_nm[1], name == gv) %>% pull(choice_label)
      if (length(lbl) > 0 && !is.na(lbl[1])) return(lbl[1])
    }
    tools::toTitleCase(gsub("_", " ", gv))
  }, character(1))

  section_map <- .build_section_map(survey, label_col = label_col) %>%
    .reassign_outside_group_sections(survey$name)

  # ── 2b. Ordre des lignes : sections dans leur ordre d'apparition REEL dans l'outil kobo (pas
  # l'ordre de la loa, qui place toutes les lignes de prix/median a la toute fin - voir build_loa()
  # dans scripts/fonctions_analyse.R), et a l'interieur d'une section, les lignes de prix en
  # PREMIER (demande utilisateur) puis le reste dans l'ordre kobo naturel. Sans ce tri, une section
  # repetee par article (ex. "Arachide") se retrouverait coupee en deux : ses questions
  # proportion/choix groupees pres du debut du tableau, sa question de prix isolee tout a la fin.
  kobo_pos <- setNames(seq_len(nrow(survey)), survey$name)

  # ── 3. Enrich tables ──────────────────────────────────────────────────────
  enrich <- function(tbl) {
    tbl %>%
      left_join(var_label_map, by = c("analysis_var" = "name")) %>%
      left_join(var_list_map,  by = c("analysis_var" = "name")) %>%
      mutate(choice_key = analysis_var_value) %>%
      left_join(choice_label_map, by = c("list_name", "choice_key" = "name")) %>%
      mutate(
        display_choice = case_when(
          analysis_type == "mean"   ~ "Moyenne",
          analysis_type == "median" ~ "Médiane",
          TRUE ~ coalesce(choice_label, analysis_var_value)
        ),
        is_pct    = analysis_type %in% c("prop_select_one", "prop_select_multiple"),
        var_label = coalesce(var_label, analysis_var)
      )
  }

  ref_labeled <- enrich(ref_tbl) %>%
    mutate(section = vapply(analysis_var, function(v) { s <- section_map[[v]]; if (is.null(s)) NA_character_ else s }, character(1)))

  grp_labeled <- if (is_stratified) enrich(grouped) else NULL
  ung_labeled <- if (has_total && !is.null(totals)) enrich(ungrouped) else ref_labeled

  # ── 3b. Pre-index lookup tables — O(1) per row in the loop ───────────────

  # Index ungrouped by (var, aval) → row index (first occurrence wins)
  if (has_total) {
    ung_key <- .key(ung_labeled$analysis_type, ung_labeled$analysis_var, ung_labeled$analysis_var_value)
    ung_idx <- setNames(seq_len(nrow(ung_labeled)), ung_key)
    ung_idx <- ung_idx[!duplicated(names(ung_idx))]
  }

  # Index grouped by (var, aval) → matrix of per-group stats (rows = keys, cols = groups).
  # Vectorised via match()/cbind() index assignment instead of a nested key x group loop -
  # at our scale (thousands of keys x tens of admin2 groups) the nested-loop version re-
  # subsets the whole grouped table per key, which gets slow fast. A single indexed
  # assignment does the same job in one pass. "First non-NA occurrence wins" per (key,
  # group) is preserved by dropping later duplicates before assigning (matrix indexing
  # would otherwise let the LAST matching row win, not the first).
  if (is_stratified) {
    grp_labeled$.key <- .key(grp_labeled$analysis_type, grp_labeled$analysis_var, grp_labeled$analysis_var_value)
    grp_keys <- unique(grp_labeled$.key)
    grp_mat  <- matrix(NA_real_, nrow = length(grp_keys), ncol = length(grp_values),
                       dimnames = list(grp_keys, grp_values))

    row_i <- match(grp_labeled$.key, grp_keys)
    col_i <- match(as.character(grp_labeled$group_var_value), grp_values)
    stat_v <- suppressWarnings(as.numeric(grp_labeled$stat))
    valid <- !is.na(row_i) & !is.na(col_i) & !is.na(stat_v)

    if (any(valid)) {
      keep <- !duplicated(paste(row_i[valid], col_i[valid]))
      ri <- row_i[valid][keep]; ci <- col_i[valid][keep]; sv <- stat_v[valid][keep]
      grp_mat[cbind(ri, ci)] <- sv
    }

    # nombre de reponses non-NA par (variable, groupe) - PAS par choix (demande utilisateur : le
    # nombre affiche sur la ligne question doit porter sur la question entiere, pas sur une modalite
    # particuliere). n_total (pas n) : n_total = somme des n par modalite = nombre de repondants
    # ayant fourni une reponse valide a la question dans ce groupe, un respondant = une modalite
    # (verifie dans analysistools:::create_analysis_prop_select_one - n subit une reecriture
    # surprenante en cas de valeurs manquantes, n_total reste la mesure fiable).
    if ("n_total" %in% names(grp_labeled)) {
      var_n_grp <- grp_labeled %>%
        distinct(analysis_var, group_var_value, n_total) %>%
        mutate(n_total = suppressWarnings(as.numeric(n_total)))
      n_mat <- matrix(NA_real_, nrow = length(unique(var_n_grp$analysis_var)), ncol = length(grp_values),
                      dimnames = list(unique(var_n_grp$analysis_var), grp_values))
      vr <- match(var_n_grp$analysis_var, rownames(n_mat))
      vc <- match(as.character(var_n_grp$group_var_value), grp_values)
      ok <- !is.na(vr) & !is.na(vc) & !is.na(var_n_grp$n_total)
      if (any(ok)) n_mat[cbind(vr[ok], vc[ok])] <- var_n_grp$n_total[ok]
    }
  }

  # ── 4. Build row list ─────────────────────────────────────────────────────
  # ordre des variables : d'abord par position de la SECTION dans l'outil kobo (premiere apparition
  # de n'importe laquelle de ses variables membres), puis a l'interieur d'une section les lignes de
  # prix/median en premier, puis le reste dans l'ordre naturel de l'outil kobo.
  all_vars_raw <- unique(ref_labeled$analysis_var)
  var_section  <- vapply(all_vars_raw, function(v) { s <- section_map[[v]]; if (is.null(s)) NA_character_ else s }, character(1))
  var_type1    <- ref_labeled$analysis_type[match(all_vars_raw, ref_labeled$analysis_var)]
  var_pos      <- kobo_pos[all_vars_raw]
  var_pos[is.na(var_pos)] <- Inf
  sect_pos     <- ave(var_pos, var_section, FUN = function(x) suppressWarnings(min(x, na.rm = TRUE)))
  is_price_row <- var_type1 %in% c("median", "mean")
  ord <- order(sect_pos, !is_price_row, var_pos)
  all_vars <- all_vars_raw[ord]

  prev_sect <- NA_character_
  rows      <- vector("list", length(all_vars) * 10L)   # pre-allocate generously
  ri        <- 0L

  # Split ref_labeled into a list of data frames by var — one filter total
  ref_split <- split(ref_labeled, ref_labeled$analysis_var)

  for (var in all_vars) {
    df   <- ref_split[[var]]
    sect <- df$section[1]

    if (!is.na(sect) && !isTRUE(sect == prev_sect)) {
      ri <- ri + 1L
      rows[[ri]] <- list(row_type = "section", label = sect)
      prev_sect  <- sect
    }

    n_q <- if ("n_total" %in% names(df)) as.integer(max(df$n_total, na.rm = TRUE)) else NA_integer_
    grp_n <- if (is_stratified && exists("n_mat") && var %in% rownames(n_mat)) n_mat[var, ] else rep(NA_real_, length(grp_values))

    ri <- ri + 1L
    rows[[ri]] <- list(
      row_type = "question",
      label    = df$var_label[1],
      var_name = var,
      n        = n_q,
      grp_n    = grp_n,
      is_pct   = df$is_pct[1]
    )

    data_df <- df[!is.na(df$display_choice) & nchar(trimws(df$display_choice)) > 0, ]
    # retire les choix "non-reponse" (Je ne sais pas / (ne) souhaite/préfère (pas) répondre / NA
    # litteral) des lignes affichees - demande utilisateur, ne s'applique qu'aux vraies modalites de
    # choix (analysis_var_value), jamais aux lignes moyenne/mediane qui n'en ont pas.
    data_df <- data_df[is.na(data_df$analysis_var_value) |
                       !tolower(data_df$analysis_var_value) %in% EXCLUDED_CHOICE_CODES, ]

    for (i in seq_len(nrow(data_df))) {
      aval <- data_df$analysis_var_value[i]
      k    <- .key(data_df$analysis_type[i], var, aval)

      if (has_total) {
        idx      <- ung_idx[k]
        if (!is.na(idx)) {
          stat_tot <- ung_labeled$stat[idx]
          n_val    <- if ("n" %in% names(ung_labeled)) as.integer(ung_labeled$n[idx]) else NA_integer_
        } else {
          stat_tot <- NA_real_; n_val <- NA_integer_
        }
      } else {
        stat_tot <- data_df$stat[i]
        n_val    <- if ("n" %in% names(data_df)) as.integer(data_df$n[i]) else NA_integer_
      }

      g_stats <- if (is_stratified) {
        row_k <- grp_mat[k, , drop = FALSE]
        if (nrow(row_k) > 0) as.numeric(row_k[1, ]) else rep(NA_real_, length(grp_values))
      } else numeric(0)

      ri <- ri + 1L
      rows[[ri]] <- list(
        row_type    = "data",
        label       = data_df$display_choice[i],
        var_name    = if (!is.na(aval)) as.character(aval) else "",
        n           = n_val,
        stat_total  = stat_tot,
        is_pct      = data_df$is_pct[i],
        group_stats = g_stats
      )
    }
  }
  rows <- rows[seq_len(ri)]   # trim pre-allocated slots

  # ── 5. Column layout ──────────────────────────────────────────────────────
  # colonne IC 95% retiree (demande utilisateur)
  COL_LABEL <- 1L; COL_VAR <- 2L; nc <- 2L

  COL_N     <- if (!is_stratified)              { nc <- nc + 1L; nc } else NA_integer_
  COL_TOTAL <- if (has_total || !is_stratified) { nc <- nc + 1L; nc } else NA_integer_
  COL_GRPS  <- if (is_stratified) seq(nc + 1L, nc + length(grp_values)) else integer(0)

  n_total_cols <- nc + length(grp_values)
  all_cols     <- seq_len(n_total_cols)

  col_hdrs <- c("Indicateur", "Variable",
                if (!is.na(COL_N))     "n",
                if (!is.na(COL_TOTAL)) "Total",
                grp_labels)

  # ── 6. Styles ─────────────────────────────────────────────────────────────
  GRAY_Q <- "#F0F0F0"; BORDER <- "#D0D0D0"; WHITE <- "#FFFFFF"
  DARK   <- "#1A1A1A"; GRAY_T <- "#888888"

  wb <- createWorkbook()
  addWorksheet(wb, sheet_name, gridLines = FALSE)

  st_col_hdr  <- createStyle(fgFill = header_bg, fontColour = WHITE,
                             textDecoration = "bold", fontSize = 11, halign = "center", valign = "center",
                             border = "Bottom", borderColour = accent_colour, borderStyle = "medium")
  st_section  <- createStyle(fgFill = accent_colour, fontColour = WHITE,
                             textDecoration = "bold", fontSize = 11, valign = "center", indent = 1)
  st_question <- createStyle(fgFill = GRAY_Q, fontColour = DARK,
                             textDecoration = "bold", fontSize = 10, valign = "center",
                             wrapText = TRUE, indent = 1,
                             border = "Left", borderColour = accent_colour, borderStyle = "medium")
  st_q_meta   <- createStyle(fgFill = GRAY_Q, fontColour = GRAY_T,
                             fontSize = 9, textDecoration = "italic", halign = "center", valign = "center",
                             border = "Left", borderColour = BORDER)
  st_q_count  <- createStyle(fgFill = GRAY_Q, fontColour = GRAY_T,
                             fontSize = 9, textDecoration = "italic", halign = "center", valign = "center",
                             border = "Left", borderColour = BORDER)
  st_data_lbl <- createStyle(fontColour = DARK, fontSize = 10, valign = "center",
                             wrapText = TRUE, indent = 2, border = "Bottom", borderColour = BORDER)
  st_data_var <- createStyle(fontColour = GRAY_T, fontSize = 8, textDecoration = "italic",
                             valign = "center", border = "Bottom", borderColour = BORDER)
  st_data_n   <- createStyle(fontColour = GRAY_T, fontSize = 9, halign = "center",
                             valign = "center", border = "Bottom", borderColour = BORDER)
  st_data_pct <- createStyle(fontColour = DARK, fontSize = 10, numFmt = "0%",
                             halign = "center", valign = "center", border = "Bottom", borderColour = BORDER)
  st_data_num <- createStyle(fontColour = DARK, fontSize = 10, numFmt = "0.00",
                             halign = "center", valign = "center", border = "Bottom", borderColour = BORDER)

  # Declinations (teintes/nuances) de la couleur rouge demandee #EE5859, du plus clair (peu de
  # reponses dans ce bucket, 0-20%) au plus fonce (80-100%) - remplace l'ancienne echelle
  # blanc -> #C0392B, memes bornes 20/40/60/80.
  pct_fills <- c("#FFFFFF", "#FCDEDE", "#F7B4B4", "#EE5859", "#CA4B4C")
  pct_fonts <- c(DARK,      DARK,      DARK,      WHITE,     WHITE)
  st_pct_heat <- lapply(seq_along(pct_fills), function(i) {
    createStyle(fgFill = pct_fills[i], fontColour = pct_fonts[i],
                fontSize = 10, numFmt = "0%", halign = "center", valign = "center",
                border = "Bottom", borderColour = BORDER)
  })

  # ── 7. Column headers ─────────────────────────────────────────────────────
  writeData(wb, sheet_name, as.data.frame(t(col_hdrs)), startRow = 1, startCol = 1,
            colNames = FALSE)
  addStyle(wb, sheet_name, st_col_hdr, rows = 1, cols = all_cols, gridExpand = TRUE)
  setRowHeights(wb, sheet_name, 1, 22)
  freezePane(wb, sheet_name, firstActiveRow = 2)

  # ── 8. Write rows — batched by type ───────────────────────────────────────
  n_rows      <- length(rows)
  row_types   <- vapply(rows, `[[`, character(1), "row_type")

  # Assign Excel row numbers
  excel_rows  <- seq_len(n_rows) + 1L   # offset by header row

  # -- 8a. Write all data values in one shot per column via writeData on a df
  #        Build a flat data frame of the scalar columns first
  labels_vec  <- vapply(rows, function(r) r$label %||% "",    character(1))
  varnam_vec  <- vapply(rows, function(r) r$var_name %||% "", character(1))

  writeData(wb, sheet_name, data.frame(labels_vec), startRow = 2, startCol = COL_LABEL, colNames = FALSE)
  writeData(wb, sheet_name, data.frame(varnam_vec), startRow = 2, startCol = COL_VAR,   colNames = FALSE)

  if (!is.na(COL_N)) {
    n_vec <- vapply(rows, function(r) if (!is.null(r$n) && !is.na(r$n)) r$n else NA_integer_, integer(1))
    writeData(wb, sheet_name, data.frame(n_vec), startRow = 2, startCol = COL_N, colNames = FALSE)
  }

  if (!is.na(COL_TOTAL)) {
    stat_vec <- vapply(rows, function(r) {
      v <- r$stat_total
      if (is.null(v) || is.na(v)) NA_real_ else as.numeric(v)
    }, numeric(1))
    writeData(wb, sheet_name, data.frame(stat_vec), startRow = 2, startCol = COL_TOTAL, colNames = FALSE)
  }

  for (j in seq_along(COL_GRPS)) {
    gv_col <- grp_values[j]
    g_vec  <- vapply(rows, function(r) {
      if (is.null(r$group_stats) || length(r$group_stats) < j) NA_real_
      else { v <- r$group_stats[j]; if (is.na(v)) NA_real_ else as.numeric(v) }
    }, numeric(1))
    writeData(wb, sheet_name, data.frame(g_vec), startRow = 2, startCol = COL_GRPS[j], colNames = FALSE)
  }

  # -- 8a-bis. "N=x" (Total) / "n=x" (par groupe) sur la ligne question - nombre de reponses non-NA
  # a la question dans ce groupe (demande utilisateur). Ecrit APRES les valeurs numeriques
  # ci-dessus : les cellules Total/groupe de la ligne question sont deja vides (NA) a ce stade
  # (r$stat_total/r$group_stats n'existent que pour les lignes "data"), donc ce texte les remplace
  # sans rien ecraser.
  # UN SEUL writeData() par ligne question (pas un par cellule) : COL_TOTAL et COL_GRPS sont des
  # colonnes CONTIGUES par construction (voir COL_TOTAL/COL_GRPS ci-dessus), donc toute la portion
  # Total+groupes d'une ligne question s'ecrit en un seul appel vectorise - a l'echelle de ce
  # classeur (~200 questions x ~25 groupes), ecrire cellule par cellule (~5000 appels) prenait
  # nettement plus longtemps qu'ecrire ligne par ligne (~200 appels), constate a l'usage.
  q_rows_idx <- which(row_types == "question")
  if (length(q_rows_idx) > 0 && (!is.na(COL_TOTAL) || length(COL_GRPS) > 0)) {
    write_start_col <- if (!is.na(COL_TOTAL)) COL_TOTAL else COL_GRPS[1]
    for (rr in q_rows_idx) {
      r <- rows[[rr]]
      tot_val <- if (!is.na(COL_TOTAL)) {
        if (!is.null(r$n) && !is.na(r$n)) paste0("N=", r$n) else ""
      } else NULL
      grp_vals <- vapply(seq_along(COL_GRPS), function(j) {
        v <- if (!is.null(r$grp_n) && length(r$grp_n) >= j) r$grp_n[j] else NA_real_
        if (!is.na(v)) paste0("n=", as.integer(v)) else ""
      }, character(1))
      row_vals <- c(tot_val, grp_vals)
      writeData(wb, sheet_name, matrix(row_vals, nrow = 1), startRow = excel_rows[rr],
               startCol = write_start_col, colNames = FALSE)
    }
  }

  # -- 8b. Merge section cells (must be per-row; no vectorised API)
  sect_rows  <- excel_rows[row_types == "section"]
  for (er in sect_rows) mergeCells(wb, sheet_name, cols = all_cols, rows = er)

  # -- 8c. Batch styles by row type
  data_rows <- excel_rows[row_types == "data"]
  q_rows    <- excel_rows[row_types == "question"]

  if (length(sect_rows) > 0) {
    addStyle(wb, sheet_name, st_section,  rows = sect_rows, cols = all_cols, gridExpand = TRUE)
    setRowHeights(wb, sheet_name, sect_rows, 26)
  }
  if (length(q_rows) > 0) {
    addStyle(wb, sheet_name, st_question, rows = q_rows, cols = COL_LABEL, gridExpand = TRUE)
    meta_cols <- c(COL_VAR, if (!is.na(COL_N)) COL_N)
    addStyle(wb, sheet_name, st_q_meta,   rows = q_rows, cols = meta_cols,  gridExpand = TRUE)
    count_cols <- c(if (!is.na(COL_TOTAL)) COL_TOTAL, COL_GRPS)
    if (length(count_cols) > 0)
      addStyle(wb, sheet_name, st_q_count, rows = q_rows, cols = count_cols, gridExpand = TRUE)
    setRowHeights(wb, sheet_name, q_rows, 34)
  }
  if (length(data_rows) > 0) {
    addStyle(wb, sheet_name, st_data_lbl, rows = data_rows, cols = COL_LABEL, gridExpand = TRUE)
    addStyle(wb, sheet_name, st_data_var, rows = data_rows, cols = COL_VAR,   gridExpand = TRUE)
    if (!is.na(COL_N))
      addStyle(wb, sheet_name, st_data_n, rows = data_rows, cols = COL_N, gridExpand = TRUE)
    setRowHeights(wb, sheet_name, data_rows, 18)

    # Pct vs numeric style split — batch each group
    is_pct_row <- vapply(rows[row_types == "data"], `[[`, logical(1), "is_pct")
    pct_excel  <- data_rows[is_pct_row]
    num_excel  <- data_rows[!is_pct_row]
    stat_cols  <- c(if (!is.na(COL_TOTAL)) COL_TOTAL, COL_GRPS)
    if (length(pct_excel) > 0 && length(stat_cols) > 0)
      addStyle(wb, sheet_name, st_data_pct, rows = pct_excel, cols = stat_cols, gridExpand = TRUE)
    if (length(num_excel) > 0 && length(stat_cols) > 0)
      addStyle(wb, sheet_name, st_data_num, rows = num_excel, cols = stat_cols, gridExpand = TRUE)

    # Per-bucket red heat map on pct rows only — explicit addStyle per bucket
    # so only the exact pct_excel rows are coloured (colourScale bleeds to all rows).
    # 5 buckets at 20% intervals: white → light red → #EE5859 → darker shade.
    if (length(pct_excel) > 0 && length(stat_cols) > 0) {
      # For each stat column, read the values for pct rows and assign bucket styles.
      # We re-extract from the rows list directly so it works with or without COL_TOTAL.
      all_data_rows_list <- rows[row_types == "data"]
      pct_rows_list      <- all_data_rows_list[is_pct_row]
      for (.col_idx in seq_along(stat_cols)) {
        .col <- stat_cols[.col_idx]
        pct_vals <- vapply(pct_rows_list, function(r) {
          if (.col_idx == 1L && !is.na(COL_TOTAL) && .col == COL_TOTAL) {
            v <- r$stat_total
          } else {
            g_idx <- which(COL_GRPS == .col)
            v <- if (length(g_idx) > 0 && length(r$group_stats) >= g_idx) r$group_stats[g_idx] else NA_real_
          }
          if (is.null(v) || is.na(v)) NA_real_ else as.numeric(v)
        }, numeric(1))
        breaks  <- c(-Inf, 0.2, 0.4, 0.6, 0.8, Inf)
        buckets <- as.integer(cut(pct_vals, breaks = breaks, labels = FALSE))
        buckets[is.na(buckets)] <- 1L
        for (b in seq_along(pct_fills)) {
          b_rows <- pct_excel[buckets == b]
          if (length(b_rows) == 0) next
          addStyle(wb, sheet_name, st_pct_heat[[b]], rows = b_rows,
                   cols = .col, gridExpand = TRUE, stack = TRUE)
        }
      }
    }
  }

  # ── 9. Column widths ──────────────────────────────────────────────────────
  setColWidths(wb, sheet_name, COL_LABEL, 48)
  setColWidths(wb, sheet_name, COL_VAR,   18)
  if (!is.na(COL_N))     setColWidths(wb, sheet_name, COL_N,     7)
  if (!is.na(COL_TOTAL)) setColWidths(wb, sheet_name, COL_TOTAL, 10)
  if (length(COL_GRPS) > 0) setColWidths(wb, sheet_name, COL_GRPS, 10)


  # ── 10. Print layout ──────────────────────────────────────────────────────
  orient <- if (is_stratified && length(grp_values) > 2) "landscape" else "portrait"
  pageSetup(wb, sheet_name, orientation = orient,
            fitToWidth = 1, fitToHeight = FALSE, paperSize = 9)

  saveWorkbook(wb, output_path, overwrite = TRUE)
  message(sprintf(
    "Saved: %s  (%d sections, %d questions, %d data rows%s)",
    output_path,
    sum(row_types == "section"),
    sum(row_types == "question"),
    sum(row_types == "data"),
    if (is_stratified) paste0(", groups: ", paste(grp_labels, collapse = " / ")) else ""
  ))
  invisible(rows)
}

# Null-coalescing operator (base R doesn't have one)
`%||%` <- function(a, b) if (!is.null(a)) a else b
