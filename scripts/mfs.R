#############################################################################################################
### Market Functionality Score (MFS)
#############################################################################################################
# Verified against: "IMPACT Cash and Markets COP - Guidance on the Market Functionality Score" (v4, Nov 2023)
# and the "MFS indicator menu". Indicator definitions, thresholds, and dimension weights below follow that
# guidance unless a comment explains a deliberate country-level adaptation (the guidance explicitly allows
# dropping/renaming answer options to fit local context - see "Creating a list of indicators").
#
# IMPORTANT - kobo choice codes change most rounds (the underlying cause of most silent MFS bugs): a
# comparison against a choice code that no longer exists does NOT error, it just silently evaluates to
# FALSE/0 for every respondent, which can make an entire indicator always score "perfect" or always score
# "worst" without any visible error. All choice/question codes this script depends on are declared ONCE in
# the CHOICE CODES block below - when the kobo tool is updated, re-check that block against the new
# survey/choices sheets before re-running this script, rather than hunting for literal strings throughout.
#
# calculate_mfs(xml_data) takes the "clean_data" sheet (as a data frame, with admin1_label/admin2_label/
# admin3_label already present) and returns one row per commune (admin1_label x admin2_label x
# admin3_label) with the 5 pillar scores (0-100 scale, see max_scores below), the total mfs_score
# (0-100), and its functionality classification - called from 1_analyse_complete.R, which handles
# reading the input file and formatting the "Score de Fonctionnalité" sheet.
calculate_mfs <- function(xml_data) {

data <- xml_data
names(data) <- gsub("/", ".", names(data))
# les colonnes binaires des select_multiple (suffixe ".choix", ex-"/choix") et les colonnes de
# stock/reapprovisionnement doivent etre numeriques pour rowSums()/soustraction plus bas - explicite ici
# plutot que de compter sur le type detecte par read_excel(), puisque l'appelant peut avoir lu clean_data
# en col_types="text" (convention du pipeline de nettoyage, pour eviter que readxl ne mal-interprete les
# dates)
data <- data %>%
  mutate(across(contains(".") | ends_with(c("_stock_duration_days_item", "_restock_wait_days_item")),
                ~suppressWarnings(as.numeric(.))))

#############################################################################################################
# CHOICE CODES (update HERE, and only here, if the kobo tool renames a choice used below) - checked against
# 2026-08-09_RCA_JMMI_outil_aout.xlsx. Group names match the kobo "list_name" they come from.
#############################################################################################################

# acces_physique (mfs_ac1_physical_access_mkt_issues_sm)
CH_ACCES_PHYSIQUE <- list(
  no_issues = "aucun",
  hazardous_buildings = "batiments_dangereux",   # "Hazardous, damaged, or unsafe buildings" (guide's AC.1/IN.1 exclusion)
  hazardous_roads = "dangers_dommages",           # "Hazards or damage on roads leading to the marketplace" (guide's AC.1/AC.2)
  dont_know = "ne_sait_pas",
  prefer_not_to_answer = "prefere_pas_repondre"
)
# NB: the tool's "acces_physique" list has no separate "Inadequate facilities..." option (it was dropped
# from the local questionnaire, which the guidance explicitly permits) - only "batiments_dangereux" exists,
# and it is used for both the AC.1 exclusion and the IN.1 indicator, matching the guide's own re-use of this
# single answer option across AC.1/AC.2/IN.1.

# oui_aucun_groupe (mfs_ac3_physical_access_mkt_discrimination_so)
CH_DISCRIMINATION_YES <- "oui"

# oui_non, used both by mfs_ac4_..._yn (security incidents) and mfs_af3_..._ability_yn (price forecasting)
CH_OUI <- "oui"
CH_NON <- "non"

# difficultes_financieres (mfs_af2_financial_challenges_customers_sm)
CH_FINANCE_NO_ISSUES <- "aucun_probleme"
CH_FINANCE_DONT_KNOW <- "ne_sait_pas"
CH_FINANCE_PNTA <- "prefere_pas_repondre"

# reapprovisionnement (mfs_re4_shop_operational_difficulties_sm)
CH_RESTOCK_NO_DIFFICULTIES <- "aucune"
CH_RESTOCK_DONT_KNOW <- "ne_sait_pas"
CH_RESTOCK_PNTA <- "prefere_pas_repondre"

# stockage (mfs_in2_available_storage_facility_so) - the two "Yes" (good) options are excluded from IN.2's
# numerator per the guide; everything else (no storage elsewhere/at home/other) counts as an issue
CH_STOCKAGE_ISSUES <- c("stockage_autre_installation", "stockage_domicile", "autre")

# type_paiement (mfs_in3_accepted_payment_modality_sm) - excluded from IN.3's numerator: both cash options
# (only a local-currency cash option exists in this tool, no foreign-currency cash option), "prefer not to
# answer", and "don't know" (guide only lists cash/PNTA, but excluding "don't know" matches every other MFS
# indicator's treatment of non-answers). "aucune" (no payment type accepted at all) is also excluded here -
# it is not evidence of alternative-payment infrastructure, so it should not count toward the "good" signal.
CH_PAIEMENT_EXCLUDE <- c("especes_monnaie_locale", "ne_sait_pas", "prefere_pas_repondre", "aucune")

# disponibilite (<item>_availability_market_item)
CH_DISPONIBLE <- "disponible"
CH_LIMITE <- "limite"
CH_INDISPONIBLE <- "indisponible"

# nb_fournisseur (food_supplier_unique_ctgy / nfi_supplier_unique / wash_supplier_unique /
# fire_combustion_supplier_unique) - "oui" = relies mostly on a single supplier for that category
CH_SINGLE_SUPPLIER <- "oui"

#############################################################################################################
# functions
#############################################################################################################

# guide's RE.3 scoring table (3: <25%, 2: 25-50%, 1: 50-75%, 0: >75% of vendors relying on a single supplier)
assign_scores_resilience_3 <- function(x) {
  case_when(
    x < 25 ~ 3,
    x >= 25 & x < 50 ~ 2,
    x >= 50 & x < 75 ~ 1,
    x >= 75 ~ 0,
    TRUE ~ NA_real_
  )
}


#### ACCESSIBILITY #### -----------------------------------------

#AC.1
#calculate '% of vendors selecting an option other than "Hazardous, damaged, or unsafe buildings in the marketplace,"
#"Hazards or damage on roads leading to the marketplace," "No issues," "Don't know," or "Prefer not to answer"'

mfs_access_physical <- data %>%
  mutate(physical_access_true = if_else(rowSums(select(.,contains('mfs_ac1_physical_access_mkt_issues_sm.')
                                                       & -ends_with(CH_ACCES_PHYSIQUE$no_issues)
                                                       & -ends_with(CH_ACCES_PHYSIQUE$dont_know)
                                                       & -ends_with(CH_ACCES_PHYSIQUE$hazardous_buildings)
                                                       & -ends_with(CH_ACCES_PHYSIQUE$hazardous_roads)
                                                       & -ends_with(CH_ACCES_PHYSIQUE$prefer_not_to_answer))) > 0,1,0)) %>%
  group_by(admin1_label,admin2_label,admin3_label) %>%
  summarise(physical_access = round(sum(physical_access_true)/n()*100,1)) %>%
  mutate(physical_access_score = case_when(physical_access<5 ~ 8,
                                           physical_access>=5 & physical_access<10 ~ 6,
                                           physical_access>=10 & physical_access<25 ~ 4,
                                           physical_access>=25 & physical_access<50 ~ 2,
                                           physical_access>=50 ~ 0))

#AC.2
#calculate '% of vendors selecting "Hazards or damage on roads leading to the marketplace"'

mfs_access_roads <- data %>%
  group_by(admin1_label,admin2_label,admin3_label) %>%
  summarise(physical_roads_access = round(sum(.data[[paste0("mfs_ac1_physical_access_mkt_issues_sm.", CH_ACCES_PHYSIQUE$hazardous_roads)]])/n()*100,1)) %>%
  mutate(physical_roads_access_score = case_when(physical_roads_access<5 ~ 4,
                                                 physical_roads_access>=5 & physical_roads_access<10 ~ 3,
                                                 physical_roads_access>=10 & physical_roads_access<25 ~ 2,
                                                 physical_roads_access>=25 & physical_roads_access<50 ~ 1,
                                                 physical_roads_access>=50 ~ 0))
#AC.3
#if any vendor responds "Yes," the market is coded as a "Yes"

mfs_access_social <- data %>%
  mutate(accessibility_social_access = if_else(mfs_ac3_physical_access_mkt_discrimination_so == CH_DISCRIMINATION_YES,1,0)) %>%
  group_by(admin1_label,admin2_label,admin3_label) %>%
  summarise(social_access_score = if_else(sum(accessibility_social_access)>0,0,2))

#AC.4
#calculate % of vendors selecting an option other than "No issues" or "Prefer not to answer"
#
# NB: unlike AC.1/AC.3, the current tool does not offer AC.4 as a single select_multiple with a "no issues"
# option. It instead asks an unconditional yes/no gate (mfs_ac4_physical_access_security_issues_yn: "have
# you observed/heard about a security incident?") followed by an always-shown detail question listing
# specific incident types (mfs_ac4_physical_access_security_issues_sm) that has no "no issues" choice at
# all. The yes/no gate is therefore the correct source for "% of vendors reporting security concerns" -
# same pattern as AC.3 above.

mfs_access_safety <- data %>%
  mutate(safety_access_true = if_else(mfs_ac4_physical_access_security_issues_yn == CH_OUI,1,0)) %>%
  group_by(admin1_label,admin2_label,admin3_label) %>%
  summarise(safety_access = round(sum(safety_access_true, na.rm = TRUE)/n()*100,1)) %>%
  mutate(safety_access_score = case_when(safety_access<5 ~ 3,
                                         safety_access>=5 & safety_access<10 ~ 2,
                                         safety_access>=10 & safety_access<20 ~ 1,
                                         safety_access>=20 ~ 0))


#### AVAILABILITY #### -------------------------------------

#AV.1
#Mode of vendor responses for each item

# function to take most common answer (statistical mode), favouring available over limited over unavailable
# on ties - i.e. if one trader says 'available' and one says 'limited', the adm3 is given 'available'.
# Priority is expressed explicitly (rather than via alphabetical sort of the raw codes) so it keeps working
# regardless of how the kobo choice codes happen to be spelled/ordered in a future tool version.
mode_fun <- function(x) {
  x <- x[!is.na(x)]
  if (length(x) == 0) return(NA_character_)

  counts <- table(x)
  tied_at_max <- names(counts)[counts == max(counts)]

  priority <- c(CH_DISPONIBLE, CH_LIMITE, CH_INDISPONIBLE)
  best <- priority[priority %in% tied_at_max][1]
  if (!is.na(best)) return(best)

  # unexpected code (e.g. dont_know/prefer_not_to_answer) tied for the max - fall back to first found
  tied_at_max[1]
}

mfs_available <- data %>%
  select(admin1_label,admin2_label,admin3_label, ends_with('_availability_market_item')) %>% #NB excludes currencies
  group_by(admin1_label,admin2_label,admin3_label) %>%
  summarise(across(ends_with('_availability_market_item'), mode_fun)) %>% #use mode_fun to extract mode value
  mutate(across(ends_with('_availability_market_item'), ~ case_when( #convert to scores
    . == CH_DISPONIBLE ~ 3,
    . == CH_LIMITE ~ 2,
    . == CH_INDISPONIBLE ~ 0))) %>%
  ungroup() %>%
  mutate(availability_score = rowSums(across(ends_with('_availability_market_item')), na.rm = T)) %>% #sum rows to give overall score
  select(admin1_label,admin2_label,admin3_label, availability_score)


#### AFORDABILITY #### -------------------------------------

#AF.1
#Median of vendor responses for each item

# function to score prices based on national median price
price_score_fun <- function(item_median){
  item <- deparse(substitute(item_median))
  result = case_when(
    {{item_median}} > median_national[[item]] * 1.5 ~ -2,
    {{item_median}} > median_national[[item]] * 1.25 ~ -1.5,
    {{item_median}} > median_national[[item]] * 1.1 ~ -1,
    {{item_median}} > median_national[[item]] * 0.90 ~ 0,
    {{item_median}} > median_national[[item]] * 0.75 ~ 1,
    {{item_median}} > median_national[[item]] * 0.5 ~ 1.5,
    {{item_median}} <= median_national[[item]] * 0.5 ~ 2,
    TRUE ~ NA_real_
  )
  return (result)
}

# monitored items for the price comparison: every "<item>_price_item" column - the final, standardised
# (per fixed unit) price field that every monitored item has exactly one of, e.g. maize_price_item ("price
# per kg"), jerrycan_price_item ("price per litre of capacity"). Detected dynamically rather than
# hardcoded, because AF.1's thresholds only ever compare an item's median price to ITS OWN national median
# (a ratio) - the ratio is unaffected by which standardised unit is used, as long as it is used consistently
# for both the market-level and national medians, which it is here since both come from the same column.
# This also means a newly-added monitored item (e.g. egg, added to the tool this round) is picked up
# automatically the next time this script runs, without needing this list to be hand-maintained every month
# (the previous hardcoded 23-item list broke - i.e. referenced columns that no longer existed under those
# names - after 6 items were renamed this round; see the raw "*_price_unit_item"/"*_<size>_price_unit_item"
# variants in the tool if a different standardised unit is ever preferred instead).
items_list <- data %>% select(ends_with('_price_item')) %>% colnames()

data <- data %>% mutate(across(all_of(items_list), as.numeric))

#select concerned columns
selection_price <- data%>% select(admin1_label,admin2_label,admin3_label,all_of(items_list)) %>%
  mutate(across(all_of(items_list), as.numeric))
#remove those having no measurements
selection_price <- selection_price[, !sapply(selection_price, function(col) all(is.na(col)))]
items_list <- setdiff(colnames(selection_price), c("admin1_label","admin2_label","admin3_label"))

#create national median dataset
median_national <- selection_price %>% # NB does not include currency prices
  lapply(median, na.rm=T) %>% #ignore warning messages- these are from trying to calculate the median on the adm1, adm2 and adm3 columns
  as.data.frame() %>%
  mutate(across(c(admin1_label, admin2_label, admin3_label), as.character))

n_items <- length(items_list) #also used later in final pillar score calculations

afford_scale <- 12 # set the max score for the affordability price levels - the scores will be scaled based on this value

#calculate scores
mfs_afford_price <- data%>%
  select(admin1_label,admin2_label,admin3_label,all_of(items_list)) %>%
  group_by(admin1_label,admin2_label,admin3_label)%>%
  summarise(across(everything(), ~median(., na.rm = TRUE))) %>%
  ungroup() %>%
  mutate(across(all_of(items_list), price_score_fun), #apply function to score each item
         afford_price_sum = rowSums(across(all_of(items_list)), na.rm = T), # row sums to give total score - +/- 2x number of items)
         afford_price_score = ((afford_price_sum - (-2*n_items)) * (afford_scale-0)) / ((2*n_items) - (-2*n_items) + 0) # use scaling formula to scale values to 0-12 - optional
  ) %>%
  select(admin1_label,admin2_label,admin3_label,afford_price_sum, afford_price_score)

#AF.2
#% of vendors selecting an option other than "No issues," "Don't know," or "Prefer not to answer"

mfs_afford_finance <- data %>%
  mutate(afford_finance_true = if_else(rowSums(select(.,contains('mfs_af2_financial_challenges_customers_sm.')
                                                      & -ends_with(CH_FINANCE_NO_ISSUES)
                                                      & -ends_with(CH_FINANCE_DONT_KNOW)
                                                      & -ends_with(CH_FINANCE_PNTA))) > 0,1,0)) %>%
  group_by(admin1_label,admin2_label,admin3_label) %>%
  summarise(afford_finance = round(sum(afford_finance_true)/n()*100,1)) %>%
  mutate(afford_finance_score = case_when(afford_finance<10 ~ 9,
                                          afford_finance>=10 & afford_finance<25 ~ 6,
                                          afford_finance>=25 & afford_finance<50 ~ 3,
                                          afford_finance>=50 ~ 0))

#AF.3
#% of vendors selecting "no"

mfs_afford_price_vol <- data %>%
  mutate(affordability_price_volatility = if_else(mfs_af3_price_forecasting_ability_yn == CH_NON,1,0)) %>%
  group_by(admin1_label,admin2_label,admin3_label) %>%
  summarise(afford_price_vol = round(sum(affordability_price_volatility)/n()*100,1)) %>%
  mutate(afford_price_vol_score = case_when(afford_price_vol<10 ~ 6,
                                            afford_price_vol>=10 & afford_price_vol<25 ~ 4,
                                            afford_price_vol>=25 & afford_price_vol<50 ~ 2,
                                            afford_price_vol>=50 ~ 0))


#### RESILIENCE #### -------------------------------------

#RE.1
# For each vendor, subtract # restocking days from # days of remaining stock for each item or category; aggregate by taking the median of these vendor-level calculations

# get names of items with stock data
stock_items <- data %>%
  select(ends_with('_stock_duration_days_item')) %>% colnames() %>% str_replace_all('_stock_duration_days_item','')

#create empty dataset
mfs_resil_restock <- data %>% select(admin1_label,admin2_label,admin3_label)

#for loop for calculating restock days
for (item in stock_items) {
  #select the stock ('_stock_current) and restock duration ('_duration') columns for each item
  item_stock <- data %>% select(contains(paste0(item,'_')) & ends_with('_stock_duration_days_item')) %>% select(1) #edit these parameters to ensure only 1 item is selected for each iteration
  item_restock <- data %>% select(contains(paste0(item,'_')) & ends_with('_restock_wait_days_item')) %>% select(1) #edit these parameters to ensure only 1 item is selected for each iteration
  item_resilience <- item_stock - item_restock #calculate resilience days
  item_resilience <- ifelse(item_resilience > 3,3,
                            ifelse(item_resilience > 0,2,
                                   ifelse(item_resilience == 0,1,0))) #calculate scores

  colnames(item_resilience) <- paste0(item,'_restock_days')

  mfs_resil_restock <- cbind(mfs_resil_restock, item_resilience)
}

mfs_resil_restock <- mfs_resil_restock %>%
  group_by(admin1_label,admin2_label,admin3_label) %>%
  summarise(across(everything(), ~median(., na.rm = TRUE))) %>% #aggregate by adm3
  mutate(resil_restock_score = rowSums(across(ends_with('_restock_days')), na.rm = T)) %>%
  select(admin1_label,admin2_label,admin3_label,resil_restock_score)

#RE.2
#### see above

# RE.3	Resilience	Supplier diversity:
# % of interviewed vendors mostly relying on a single supplier to stock their businesses by category
#
# NB: the tool no longer asks supplier-uniqueness per PRICE ITEM (the "*_supplier_unique_item" columns the
# original version of this script relied on do not exist any more). It now asks it once per PRODUCT
# CATEGORY instead (food / non-food items / WASH items / fire & combustion items), via 4 unconditional
# select_one questions sharing the "nb_fournisseur" choice list (oui = relies mostly on a single supplier).
supplier_category_cols <- c(
  food = "food_supplier_unique_ctgy",
  nfi = "nfi_supplier_unique",
  wash = "wash_supplier_unique",
  fire_combustion = "fire_combustion_supplier_unique"
)
input_n_supplier_categories <- length(supplier_category_cols)

mfs_resil_supply_diverse <- data %>%
  select(admin1_label,admin2_label,admin3_label,all_of(unname(supplier_category_cols))) %>%
  mutate(across(all_of(unname(supplier_category_cols)), ~if_else(. == CH_SINGLE_SUPPLIER,1,0))) %>%
  group_by(admin1_label,admin2_label,admin3_label) %>%
  summarise(across(all_of(unname(supplier_category_cols)), ~round(sum(., na.rm=T)/sum(!is.na(.))*100,1))) %>%
  ungroup() %>%
  mutate(across(all_of(unname(supplier_category_cols)), assign_scores_resilience_3)) %>%
  mutate(resil_supply_diverse_score = rowSums(across(all_of(unname(supplier_category_cols))), na.rm = T)) %>%
  #sum rows to give final score - max score is 12 (3 x 4 categories)
  select(admin1_label,admin2_label,admin3_label,resil_supply_diverse_score)


#RE.4
#% of vendors selecting an option other than "No difficulties" or "Prefer not to answer"

mfs_resil_supply <- data %>%
  mutate(resilience_supply_true = if_else(rowSums(select(.,contains('mfs_re4_shop_operational_difficulties_sm.')
                                                         & -ends_with(CH_RESTOCK_NO_DIFFICULTIES)
                                                         & -ends_with(CH_RESTOCK_DONT_KNOW)
                                                         & -ends_with(CH_RESTOCK_PNTA))) > 0,1,0)) %>%
  group_by(admin1_label,admin2_label,admin3_label) %>%
  summarise(resil_supply = round(sum(resilience_supply_true)/n()*100,1)) %>%
  mutate(resil_supply_score = case_when(resil_supply<5 ~ 12,
                                        resil_supply>=5 & resil_supply<10 ~ 9,
                                        resil_supply>=10 & resil_supply<25 ~ 6,
                                        resil_supply>=25 & resil_supply<50 ~ 3,
                                        resil_supply>=50 ~ 0))


#### INFRASTRUCTURE #### -------------------------------------

#IN.1
#% of vendors selecting "Hazardous, damaged, or unsafe buildings in the marketplace"

mfs_infra_facilities <- data %>%
  group_by(admin1_label,admin2_label,admin3_label) %>%
  summarise(infra_facilities = round(sum(.data[[paste0("mfs_ac1_physical_access_mkt_issues_sm.", CH_ACCES_PHYSIQUE$hazardous_buildings)]])/n()*100,1)) %>%
  mutate(infra_facilities_score = case_when(infra_facilities<5 ~ 4,
                                            infra_facilities>=5 & infra_facilities<10 ~ 3,
                                            infra_facilities>=10 & infra_facilities<25 ~ 2,
                                            infra_facilities>=25 & infra_facilities<50 ~ 1,
                                            infra_facilities>=50 ~ 0))

#IN.2
#% of vendors selecting an option other than "Yes, within my own business facilities" or "Yes, elsewhere within the marketplace" (or "Prefer not to answer")

mfs_infra_storage <- data %>%
  mutate(infrastructure_storage = if_else(mfs_in2_available_storage_facility_so %in% CH_STOCKAGE_ISSUES, 1, 0)) %>%
  group_by(admin1_label,admin2_label,admin3_label) %>%
  summarise(infra_storage = round(sum(infrastructure_storage)/n()*100,1)) %>%
  mutate(infra_storage_score = case_when(infra_storage<10 ~ 3,
                                         infra_storage>=10 & infra_storage<25 ~ 2,
                                         infra_storage>=25 & infra_storage<50 ~ 1,
                                         infra_storage>=50 ~ 0))

#IN.3
#% of vendors selecting an option other than "Cash (local currency)", "Cash (foreign currencies)", or "Prefer not to answer"

mfs_infra_payment <- data %>%
  mutate(infra_payment_true = if_else(rowSums(select(.,ends_with('_payment_modality_sm.')
                                                     & -matches(paste0("(", paste(CH_PAIEMENT_EXCLUDE, collapse = "|"), ")$")))) > 0,1,0)) %>%
  group_by(admin1_label,admin2_label,admin3_label) %>%
  summarise(infra_payment = round(sum(infra_payment_true, na.rm = T)/n()*100,1)) %>%
  mutate(infra_payment_score = case_when(infra_payment>75 ~ 3,
                                         infra_payment>50 ~ 2,
                                         infra_payment>25 ~ 1,
                                         infra_payment<=25 ~ 0))


#======================================== CALCULATE MFS =============================================

mfs_data_list <- list(mfs_access_physical,mfs_access_roads,mfs_access_social,mfs_access_safety, #accessibility pillar - max 17
                      mfs_available,                                                            #availability pillar - max depends on number of items or categories
                      mfs_afford_price,mfs_afford_finance,mfs_afford_price_vol,                 #affordability pillar - max depends on number of items
                      mfs_resil_restock,mfs_resil_supply_diverse,mfs_resil_supply,              #resilience pillar - max depends on number of items or categories
                      mfs_infra_facilities,mfs_infra_storage,mfs_infra_payment)                 #infrastructure pillar - max 10



#create full dataset
mfs <- mfs_data_list %>% reduce(full_join, by=c('admin1_label','admin2_label','admin3_label'))




#calculate final mfs
mfs <- mfs %>%
  select(admin1_label,admin2_label,admin3_label, contains('_score'))%>%
  #calculate pillar scores, scale to 0-1 by dividing by the max score for that pillar, then apply weights
  # (dimension weights below are REACH's standard MFS weights - accessibility 25, availability 30,
  # affordability 15, resilience 20, infrastructure 10, out of 100 - see "MFS: Standard dimension weights"
  # in the guidance)
  mutate(mfs_accessibility_pillar_score = (rowSums(across(contains('access_score'))) /17) *25,
         mfs_availability_pillar_score = (availability_score /(n_items*3)) *30,
         mfs_affordability_pillar_score = (rowSums(across(contains('afford'))) /(afford_scale+15)) *15,
         mfs_resilience_pillar_score = (rowSums(across(c(resil_restock_score, resil_supply_diverse_score, resil_supply_score))) /
                                          (length(stock_items)*3 + input_n_supplier_categories*3 + 12)) * 20,

         mfs_infrastructure_pillar_score = (rowSums(across(contains('infra'))) /10) *10,

         #calculate final score
         mfs_score = rowSums(across(contains('_pillar_score')))) %>%
  select(admin1_label,admin2_label,admin3_label, contains(c('_pillar_score', 'mfs_score')))



# Define the maximum possible scores
max_scores <- list(
  mfs_accessibility_pillar_score = 25,
  mfs_availability_pillar_score = 30,
  mfs_affordability_pillar_score = 15,
  mfs_resilience_pillar_score = 20,
  mfs_infrastructure_pillar_score = 10,
  mfs_score = 100
)

mfs <- mfs %>%
  dplyr::mutate(
    low_score_accessibility= ifelse(mfs_accessibility_pillar_score < 12.5, 1, 0),
    low_score_availability= ifelse(mfs_availability_pillar_score < 15, 1, 0),
    low_score_affordability= ifelse(mfs_affordability_pillar_score < 7.5, 1, 0),
    low_score_resilience= ifelse(mfs_resilience_pillar_score < 10, 1, 0),
    low_score_infrastructure= ifelse(mfs_infrastructure_pillar_score < 5, 1, 0),
    sum_low_dimensions= rowSums(across(starts_with("low"))))

# Functionality categories per the guidance ("Categorising outcomes"):
#   Full functionality:    MFS > 80% of max AND no dimension < 50% of its max
#   Limited functionality: MFS > 50% of max OR  <=1 dimension  < 50% of its max
#   Poor functionality:    MFS < 50% of max OR  >=2 dimensions < 50% of their max
#   Severe issues:         MFS < 25% of max OR  >=3 dimensions < 50% of their max
#   Insufficient data:     one or more entire dimensions could not be calculated (mfs_score is NA)
# These "or" conditions overlap (e.g. a market scoring 60% with 2 low dimensions matches both "Limited"'s
# and "Poor"'s trigger) - the guidance's own worked table does not resolve this, so the checks below are
# evaluated worst-category-first and the first (most severe) match wins, which is the only ordering under
# which every category - including "Severe issues" - is actually reachable. Checking best-first instead
# (as this script previously did) makes "Poor" swallow every case that should be "Severe", since "Poor"'s
# own "sum_low_dimensions >= 2" branch is already true whenever "Limited"'s condition fails - "Severe" can
# then never be reached. The is.na() check must also come first, or a market missing an entire dimension
# (mfs_score = NA) silently falls into a normal category instead of being flagged.
mfs <- mfs %>%
  mutate(functionality = case_when(
    is.na(mfs_score) ~ MFS_LVL_INSUFFICIENT,
    mfs_score < 25 | sum_low_dimensions >= 3 ~ MFS_LVL_SEVERE,
    mfs_score < 50 | sum_low_dimensions >= 2 ~ MFS_LVL_POOR,
    mfs_score > 80 & sum_low_dimensions == 0 ~ MFS_LVL_FULL,
    TRUE ~ MFS_LVL_LIMITED))

return(mfs)

}

#############################################################################################################
# LES 5 CATEGORIES DE CLASSIFICATION MFS (memes chaines EXACTES que celles produites par
# calculate_mfs() ci-dessus) - definies ICI, au niveau du fichier (pas a l'interieur de la fonction),
# pour etre reutilisables telles quelles par 4_analyse_complete.R (qui source ce fichier) sans retaper
# ces chaines une seconde fois - demande utilisateur, suite a l'audit de robustesse : ces 5 chaines
# etaient auparavant retapees independamment a 4 endroits (le case_when() ci-dessus + 3 endroits dans
# 4_analyse_complete.R - code couleur, legende, traduction anglaise pour "mfs_analysis"), sans garantie
# de rester synchronisees si l'une d'elles etait corrigee sans les autres - un tel desaccord fait
# disparaitre silencieusement la couleur/la traduction de la categorie concernee, sans aucune erreur.
#############################################################################################################
MFS_LVL_INSUFFICIENT <- "Données insuffisantes"
MFS_LVL_SEVERE       <- "Problèmes graves"
MFS_LVL_POOR         <- "Mauvaise fonctionnalité"
MFS_LVL_LIMITED      <- "Fonctionnalité limitée"
MFS_LVL_FULL         <- "Fonctionnalité complète"
