# JMMI RCA – Dashboard (`rca_jmmi_dashboard`)

Shiny dashboard that runs the full monthly **JMMI / ICSM (Initiative Conjointe de Suivi des Marchés)** data pipeline for REACH Central African Republic. The pipeline covers daily monitoring during collection, data cleaning, and the monthly analysis (PMAS/MEB, Market Functionality Score). Users do not need R. They upload Excel files in the browser, click **Lancer**, and download the resulting workbook.

- **Live app:** <https://impact-initiatives.shinyapps.io/rca_jmmi_dashboard/>
- **Interface language:** French (all scripts, workbooks and messages are in French)
- **Sample data:** August 2026 round, collected 10 Aug – 4 Sep 2026, 786 surveys (see [Sample files](#sample-files-and-how-to-use-them))

---

## 1. What the dashboard does

The app has three modes, chosen in **Etape 1 – Que voulez-vous faire ?**:

| Mode | When | Inputs | Output |
|---|---|---|---|
| **Suivi quotidien** | During collection, as often as needed | Raw Kobo export + Kobo tool | `RCA_JMMI_<mois>_maj_quotidienne_<date>_<heure>.xlsx` |
| **Nettoyage des données → Pour suivi** | Once collection is closed (pass 1) | Raw Kobo export + Kobo tool (+ optional previous "suivi à vérifier") | `RCA_JMMI_<mois>_suivi_a_verifier_<date>_<heure>.xlsx`, to be filled in by the field team |
| **Nettoyage des données → Finale** | After the field team has filled in the workbook (pass 2) | Raw Kobo export + Kobo tool + **filled** "suivi à vérifier" + current USD/XAF rate | `REACH_CAR_Dataset_JMMI_<Mois>-<Année>.xlsx` (final cleaned dataset) |
| **Analyse mensuelle** | After cleaning | Final cleaned dataset (+ optional previous month's analysis workbook) | `REACH_CAR_Analyse_JMMI_<Mois>-<Année>.xlsx` |

The page also has a button that downloads the full methodology document (PMAS/MEB composition and calculation, Market Functionality Score) as a standalone HTML file. No run is needed first.

### 1.1 Suivi quotidien (daily monitoring)
This mode only tracks coverage and prices. It makes **no** corrections. It produces one workbook with these sheets:
- `Nombre_Prix`: number of prices per item and commune. Cells with ≤ 2 prices are red, and communes with 1–2 prices are yellow.
- `Prix_Median`: median price per item, with a red (expensive) to green (cheap) colour scale per item.
- `Resume`, `Detail`, `Resume_Partenaire_Zone`: surveys per partner, per area, and partner × area.
- `raw_data`: the uploaded data, anonymised. It lets you look up a flagged row without reopening Kobo.

### 1.2 Nettoyage des données (cleaning, two passes)
Cleaning runs the same script (`2_nettoyage_final.R`) twice.

**Pass 1 – "Pour suivi"** runs all checks at once and writes a single plain-language workbook for the field team. The checks are:
1. **Exact duplicates** (same `uuid`). These stop the run, because they are an export problem, not a field issue.
2. **Automatic fixes, applied before review:**
   - "No information" free-text answers ("RAS", "rien à signaler", "merci", "pas de problème"…, typos included) are blanked.
   - "Autre" answers that violate the Kobo `relevant` logic are corrected.
3. **No consent** → the survey is flagged for deletion.
4. **Sentinel values** (99, 999, -999, 88, 888, -888).
5. **Logical checks** listed in `ressources/check_list_august.xlsx`. Stock and restock day values above 365 are also flagged.
6. **Price outliers**, using `cleaningtools::check_outliers`. These run only on normalised `<item>_price_item` columns.
7. **Near-duplicate surveys** and **abnormally high % of missing values**.
8. **"Autre" free-text answers** to recode.

If a previous "suivi à vérifier" is uploaded (**Etape 3 – Nettoyage précédent**), decisions already taken for surveys that existed before are carried over automatically. Those rows are marked *"Repris (déjà traité)"* and only *"Nouveau"* rows need review. Use this when you regenerate the workbook on a dataset that has grown.

**Pass 2 – "Finale"** reads the filled workbook and then:
- applies every decision, including the "autre" recoding, the additional cleaning log and the deletion log;
- re-checks the cleaning (`review_cleaning`);
- recalculates prices in EUR (fixed parity of 655.95 XAF) and USD (rate entered in the app, 570 by default);
- removes personal data (GPS, enumerator, shop name, device ID, audit) and drops questions that are entirely empty;
- builds the final workbook on the REACH *Minimum Standards* checklist template.

The final workbook contains `CHECKLIST_Data_Cleaning` (auto-filled), `READ_ME`, `survey`, `choices`, `variable_tracker`, `raw_data`, `clean_data`, `cleaning_log` and `deletion_log`.

### 1.3 Analyse mensuelle (monthly analysis)
This mode takes the **Finale** workbook. The Kobo tool does not need to be uploaded again, because the `survey`/`choices` sheets are already inside that workbook. The year and month are detected from the `date` column. The output keeps every cleaning sheet and adds the following:

| Sheet | Content |
|---|---|
| `Données nettoyées` | `clean_data` with choice codes replaced by their labels. Relabelling is automatic and is no longer a separate step. |
| `Coût médian du PMAS` | Cost of the minimum expenditure basket (food / NFI / WASH / total) by commune, plus a national value computed as the median of commune baskets. Items with too few local prices are imputed with the national median, and data-sufficiency thresholds decide whether a basket is computed or flagged "à interpréter avec prudence". |
| `Evolution des paniers du PMAS` | Change compared with the previous month. This sheet is only filled if last month's **analysis** workbook is uploaded. |
| `Prix médian par article` | Median price per item, by commune and national. |
| `Score de Fonctionnalité` | Market Functionality Score: 5 pillars, a score out of 100, and a classification per commune. See [`Methodologie_MFS.md`](Methodologie_MFS.md). |
| `Analyse` | Descriptive analysis of the whole questionnaire, using an analysis list (LOA) generated automatically from the Kobo tool. |
| `Checklist_Data_Analysis` | Minimum-standards checklist for the analysis step. |
| `meb_composition`, `meb_analysis`, `mfs_analysis` | REACH standard-format tables (admin3 → admin0, in XAF/EUR/USD). |

---

## 2. Folder contents

```
rca_jmmi_dashboard/
├── app.R                        # Shiny app (UI + server). Runs the scripts below in-process.
├── styles.css                   # Dashboard styling
├── 1_suivi_quotidien.R          # Mode "Suivi quotidien"
├── 2_nettoyage_final.R          # Mode "Nettoyage" (Pour suivi + Finale)
├── 4_analyse_complete.R         # Mode "Analyse mensuelle"
├── Methodologie_MFS.md          # Market Functionality Score methodology (French)
├── rca_jmmi_dashboard.Rproj     # RStudio project
├── scripts/
│   ├── 0_configuration.R                    # Packages, Kobo tool loading/validation, logical-check list
│   ├── fonctions_nettoyage.R                # Cleaning helpers, "suivi à vérifier" workbook, final workbook builder
│   ├── fonctions_verification_outil_kobo.R  # Checks that `relevant` expressions point to real choices
│   ├── fonctions_analyse.R                  # LOA generation, PMAS/MEB weights, thresholds and calculation
│   ├── mfs.R                                # Market Functionality Score (calculate_mfs)
│   ├── format_results.R                     # Formatting of the "Analyse" sheet
│   ├── fonctions_relabel.R                  # Choice code → label relabelling
│   └── methodology.R                        # Builds the downloadable methodology HTML
├── ressources/
│   ├── check_list_august.xlsx               # Logical checks (id_verification, condition, columns, description)
│   ├── Checklist-Template_Nettoyage-Donnees-Quant_Normes-Min_v1_2_2026-03_FRENCH.xlsx  # REACH minimum-standards template
│   └── labels_questions_others.xlsx         # "Autre" question labels (regenerated at each run)
├── rsconnect/                   # shinyapps.io deployment record (impact-initiatives account)
├── sample_input/                # Example inputs (August 2026 round)
└── sample_output/               # Example outputs (August 2026 round)
```

**This folder is self-contained.** The `.R` scripts and `ressources/` are copies of the standalone projects in this repository: `rca_jmmi_cleaning` for cleaning, `rca_jmmi_analysis` for analysis, and `rca_jmmi_relabelling` for relabelling. The copies exist so the dashboard can be deployed on its own. There is **no automatic sync**: if the source projects change, copy the files here again.

---

## 3. Sample files and how to use them

All samples come from the **August 2026 round** (data exported on 8 Sep 2026). You can use them to test every mode of the dashboard end-to-end.

> ⚠️ The sample inputs contain non-anonymised field data. The raw Kobo export has enumerator names, shop names and device IDs. The suivi workbook has enumerator names (its shop names are already replaced by anonymous IDs). The repository is private, but do not share these files outside the team. The final cleaned dataset and the analysis workbook have these columns removed.

### `sample_input/`

| File | What it is |
|---|---|
| `2026-07-28_RCA_JMMI_outil_aout.xlsx` | Kobo tool (XLSForm: `survey`, `choices`, `settings`) used for the August round. Upload it as **"Outil Kobo"**. |
| `REACH_RCA_JMMI_Aout2026_export_kobo_2026-09-08.xlsx` | Raw Kobo export: 786 surveys, all versions, XML headers. Upload it as **"Jeu de données brutes"**. The original Kobo filename was shortened because it exceeded the Windows path-length limit. |
| `RCA_JMMI_2026-08_suivi_a_verifier_2026-09-08_13h39.xlsx` | **"Suivi à vérifier" workbook, already filled in** by the field team. It is the output of *Nettoyage → Pour suivi* and the input of the next step (see below). |

#### About the "suivi à vérifier" workbook (both an output and an input)
The dashboard generates this workbook in **Nettoyage → Pour suivi**. The field team fills it in, and it is then uploaded back into the dashboard. The sample is the **filled version**, so you can skip the field step when testing.

Its tabs:

| Tab | Role | Filled in the sample |
|---|---|---|
| `1_Lisez_moi` | Instructions for the field team | – |
| `2_Donnees_brutes` | Anonymised raw data for reference (read only) | – |
| `3_A_verifier` | One row per flagged issue (outlier, logical inconsistency, near-duplicate…). Choose an action in **"Que faire ? (obligatoire)"**: *Ne rien changer / Corriger la réponse / Mettre une valeur manquante / Supprimer toute l'enquête* | 45 / 45 rows decided |
| `4_Verification_prix` | Same as `3_A_verifier`, for item prices. The rows for one item form one block. | 373 / 373 rows decided |
| `5_Reponses_autres` | "Autre" free-text answers. Choose *Réécrire / Réattribuer à un choix existant / Supprimer*. | 526 / 526 rows decided |
| `6_Journal_additionnel` | Manual corrections that were **not** flagged automatically (`uuid`, `variable`, `change_type`, `new_value`) | 298 rows |
| `7_Journal_suppression` | Whole surveys to delete that were not flagged (`uuid`, `reason_deletion`) | empty |
| `Listes_*` | Hidden lists that feed the dropdowns. Do not edit. | – |

**Rules for the field team:** do not rename the file or the tabs. Do not edit grey columns. Do not add columns to tabs 6 and 7.

You can use this sample in two ways:
1. **As the filled workbook (normal use):** *Nettoyage → Finale*, **Etape 3 – Classeur rempli**. This produces the final cleaned dataset.
2. **As a previous review (carry-forward test):** *Nettoyage → Pour suivi*, tick **"Un nettoyage précédent existe-t-il…"** and upload it. Each decision it contains is carried into the new workbook, and those rows are marked *"Repris (déjà traité)"*.

### `sample_output/`

| File | Produced by |
|---|---|
| `RCA_JMMI_2026-08_maj_quotidienne_2026-09-08_08h13.xlsx` | **Suivi quotidien**. Sheets: `Nombre_Prix`, `Prix_Median`, `Resume`, `Detail`, `Resume_Partenaire_Zone`, `raw_data`. |
| `REACH_CAR_Analyse_JMMI_Aout_Septembre_2026.xlsx` | **Analyse mensuelle**. It contains every cleaning sheet (`clean_data`, `cleaning_log`, …) plus the analysis sheets listed in §1.3. The dashboard names this file `REACH_CAR_Analyse_JMMI_<Mois>-<Année>.xlsx`; this sample was renamed afterwards to reflect the Aug–Sep collection window. |
| `REACH_RCA - ICSM - Fiche informative - Août_Septembre_26.pdf` / `.indd` | Final factsheet and its InDesign source. This is **not** generated by the dashboard; it is laid out manually from the analysis workbook. |

The intermediate `REACH_CAR_Dataset_JMMI_<Mois>-<Année>.xlsx` (output of *Finale*) is not included separately. Its sheets are all inside the analysis workbook, and you can regenerate it in step 3 of the walkthrough below.

### End-to-end walkthrough with the samples

1. **Suivi quotidien:** upload the raw export and the Kobo tool, then click **Lancer**. You get a `maj_quotidienne` workbook.
2. **Nettoyage → Pour suivi:** upload the same two files, then click **Lancer**. You get an empty `suivi_a_verifier` workbook, which in real use goes to the field team.
3. **Nettoyage → Finale:**
   - upload the raw export and the Kobo tool;
   - check the USD/XAF rate;
   - in **Etape 3**, upload `RCA_JMMI_2026-08_suivi_a_verifier_2026-09-08_13h39.xlsx`;
   - click **Lancer**.

   You get `REACH_CAR_Dataset_JMMI_Aout-2026.xlsx`.
4. **Analyse mensuelle:**
   - upload the file from step 3;
   - optionally, upload last month's **analysis** workbook (not the cleaned dataset) and type its month, e.g. `Juillet`;
   - click **Lancer**.

   You get `REACH_CAR_Analyse_JMMI_Aout-2026.xlsx`.

If a step fails, the dashboard shows a short French diagnosis ("Que faire : …") above the full execution log.

---

## 4. Running and deploying

### Requirements
R ≥ 4.3, with these packages:
- **CRAN:** `shiny`, `shinyjs`, `shinythemes`, `htmltools`, `tidyverse`, `readxl`, `writexl`, `openxlsx`, `openxlsx2`, `labelled`, `srvyr`, `stringi`, `plyr`
- **IMPACT packages on GitHub:**

```r
# install.packages("devtools")
devtools::install_github("impact-initiatives/cleaningtools")
devtools::install_github("impact-initiatives/analysistools")
devtools::install_github("impact-initiatives/presentresults")
```

### Run locally
Open `rca_jmmi_dashboard.Rproj` in RStudio, then:
```r
shiny::runApp()
```
Each run happens in an isolated temporary folder. Nothing is written in the project folder, except `ressources/labels_questions_others.xlsx`, which is regenerated.

### Deploy to shinyapps.io
```r
rsconnect::deployApp()   # uses rsconnect/shinyapps.io/impact-initiatives/rca_jmmi_dashboard.dcf
```
Make sure `scripts/` and `ressources/` are included in the bundle. If they are missing, the app shows the error *"Fichiers de l'application manquants"*.

---

## 5. Maintenance – every new round

- **Kobo tool changes:** choice codes and question names used by the Market Functionality Score are declared at the top of `scripts/mfs.R`, and PMAS items and weights are at the top of `scripts/fonctions_analyse.R` (`MEB_*_WEIGHTS`). A code that no longer exists does **not** raise an error; it silently scores 0. Check these against the new tool. Missing PMAS items are reported in the analysis.
- **Logical checks:** edit `ressources/check_list_august.xlsx`. Rows with `id_verification = "skip"` are ignored.
- **USD/XAF rate:** update `DEFAULT_USD_XRATE` in `app.R` periodically. It can also be overridden in the app at the *Finale* step. EUR/XAF is a fixed parity.
- **Source projects:** if `rca_jmmi_cleaning`, `rca_jmmi_analysis` or `rca_jmmi_relabelling` change, copy the updated scripts here.

---

## 6. What was done (history)

The dashboard grew out of the standalone cleaning and analysis scripts. The main changes were:

- **Pipeline simplified for a field team with limited technical training.** Interim checks during collection were removed, so collection only needs daily monitoring. The former scripts 2 and 3, which used separate workbooks for logical checks and outliers, were merged into **one cleaning script run twice**, with **one workbook to fill in**. That workbook is written in plain French and has dropdowns and grey "do not edit" columns.
- **Automatic pre-cleaning:**
  - "RAS"-type answers with no information are blanked;
  - "autre" answers that are inconsistent with the `relevant` logic are corrected;
  - stock and restock day values are flagged only above 365;
  - outlier detection is restricted to normalised prices.
- **Carry-forward of decisions** from a previous "suivi à vérifier", so a growing dataset can be reviewed incrementally.
- **Manual logs inside the same workbook:** the additional cleaning log and the deletion log are now tabs 6 and 7, validated row by row with precise error messages.
- **Final workbook aligned with REACH Minimum Standards.** It uses the checklist template, auto-fills the checklist where the script can certify it, and adds a `READ_ME`, a variable tracker, anonymisation, and EUR and USD prices.
- **Monthly analysis integrated into the dashboard:**
  - LOA generated automatically from the Kobo tool;
  - PMAS/MEB with local-data sufficiency rules and national-median imputation;
  - Market Functionality Score (IMPACT guidance v4, Nov. 2023);
  - month-on-month basket evolution;
  - REACH standard `meb_analysis` / `mfs_analysis` sheets;
  - automatic relabelling of choice codes (formerly a separate dashboard step).
- **Dashboard hardening:**
  - scripts run in-process (no `Rscript` subprocess), so the app works on shinyapps.io and Posit Connect;
  - each run has its own temporary folder and clean environment;
  - checks catch missing resources and wrong-file uploads, such as re-uploading an analysis workbook as a cleaned dataset;
  - French error diagnosis;
  - downloadable methodology document.
