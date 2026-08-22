# Méthodologie du Score de Fonctionnalité des Marchés (SFM / MFS)

Calculé par `calculate_mfs()` dans `scripts/mfs.R`, appelée depuis `1_analyse_complete.R`
(`rca_jmmi_dashboard/4_analyse_complete.R` pour la version intégrée au dashboard). Produit la feuille
"Score de Fonctionnalité" du classeur d'analyse mensuel, une ligne par commune.

Source : *IMPACT Cash and Markets COP - Guidance on the Market Functionality Score* (v4, nov. 2023) et le
"MFS indicator menu" associé. Les définitions d'indicateurs, seuils et poids ci-dessous suivent ce guide ;
les écarts (suppression/renommage d'options de réponse pour coller au contexte local) sont explicitement
permis par le guide et documentés au fil du texte.

**Entrée** : `clean_data` (une ligne par enquête/vendeur interrogé). **Sortie** : une ligne par commune
(regroupement `admin1_label` x `admin2_label` x `admin3_label`), avec les 5 scores de pilier, le score
total (0-100) et la classification de fonctionnalité.

## Vue d'ensemble : les 5 piliers et leur poids

| Pilier | Poids final (/100) |
|---|---|
| Accessibilité des marchés | 25 |
| Disponibilité des produits | 30 |
| Abordabilité des prix | 15 |
| Résilience des circuits d'approvisionnement | 20 |
| Infrastructures du marché | 10 |

Chaque pilier est d'abord calculé sur son **échelle brute** (somme des scores de ses indicateurs), puis
ramené à son poids final par une simple règle de trois : `score_pilier = (somme_brute / max_brut) * poids`.

## Pilier 1 - Accessibilité (poids 25, échelle brute /17)

**AC.1** (max 8) - % de vendeurs cochant une option *autre que* "aucun problème", "bâtiments dangereux",
"dangers/dommages sur les routes", "ne sait pas" ou "préfère ne pas répondre"
(`mfs_ac1_physical_access_mkt_issues_sm`, select_multiple).
Seuils : <5 % → 8 · 5-10 % → 6 · 10-25 % → 4 · 25-50 % → 2 · ≥50 % → 0.

**AC.2** (max 4) - % de vendeurs ayant coché spécifiquement "dangers/dommages sur les routes" (même
question que AC.1, choix `dangers_dommages`).
Seuils : <5 % → 4 · 5-10 % → 3 · 10-25 % → 2 · 25-50 % → 1 · ≥50 % → 0.

**AC.3** (max 2) - `mfs_ac3_physical_access_mkt_discrimination_so` (oui/non). Si **au moins un** vendeur
du marché signale de la discrimination ("oui"), la commune reçoit 0 ; sinon 2 (logique "un seul cas
suffit", pas un pourcentage).

**AC.4** (max 3) - % de vendeurs répondant "oui" à `mfs_ac4_physical_access_security_issues_yn`
("avez-vous observé/entendu parler d'un incident de sécurité ?"). L'outil local pose cette question comme
un simple oui/non suivi d'une liste de détails sans option "aucun problème" - c'est donc ce oui/non, et
non la liste de détails, qui sert de source pour ce pourcentage.
Seuils : <5 % → 3 · 5-10 % → 2 · 10-20 % → 1 · ≥20 % → 0.

`score_accessibilité = (AC.1 + AC.2 + AC.3 + AC.4) / 17 * 25`

## Pilier 2 - Disponibilité (poids 30, échelle brute /(3 x nb d'articles suivis))

**AV.1** - pour chaque article suivi (une colonne `<item>_availability_market_item` par article), on
prend la réponse **la plus fréquente** (mode statistique) par commune parmi "disponible" / "limité" /
"indisponible". En cas d'égalité, priorité disponible > limité > indisponible (pas un tri alphabétique
des codes kobo, pour rester correct même si les codes changent). Score par article : disponible = 3,
limité = 2, indisponible = 0. Le score du pilier est la somme sur tous les articles suivis.

`score_disponibilité = somme_articles / (3 * nb_articles) * 30`

## Pilier 3 - Abordabilité (poids 15, échelle brute /27 = 12 + 9 + 6)

**AF.1** (échelle 0-12, après mise à l'échelle) - pour chaque article suivi (une colonne `<item>_price_item`
par article - le prix déjà standardisé par unité fixe), calcule la **médiane du prix par commune**,
comparée à la **médiane nationale** de ce même article (ratio commune/national) :

| Ratio au prix national | Score |
|---|---|
| > 150 % | -2 |
| > 125 % | -1.5 |
| > 110 % | -1 |
| > 90 % | 0 |
| > 75 % | 1 |
| > 50 % | 1.5 |
| ≤ 50 % | 2 |

Somme sur tous les articles suivis, puis mise à l'échelle linéaire de la plage possible
`[-2 * nb_articles, +2 * nb_articles]` vers `[0, 12]`. Les articles nouvellement ajoutés à l'outil sont
détectés automatiquement (toute colonne `*_price_item`), pas besoin de maintenir une liste à la main.

**AF.2** (max 9) - % de vendeurs cochant une option *autre que* "aucun problème" / "ne sait pas" /
"préfère ne pas répondre" (`mfs_af2_financial_challenges_customers_sm`).
Seuils : <10 % → 9 · 10-25 % → 6 · 25-50 % → 3 · ≥50 % → 0.

**AF.3** (max 6) - % de vendeurs répondant "non" à `mfs_af3_price_forecasting_ability_yn` (capacité à
anticiper l'évolution des prix).
Seuils : <10 % → 6 · 10-25 % → 4 · 25-50 % → 2 · ≥50 % → 0.

`score_abordabilité = (AF.1 + AF.2 + AF.3) / 27 * 15`

## Pilier 4 - Résilience (poids 20, échelle brute /(3 x nb d'articles avec données de stock + 12 + 12))

**RE.1** (max 3 par article, agrégé par médiane) - pour chaque article ayant des données de stock
(colonnes `<item>_stock_duration_days_item` et `<item>_restock_wait_days_item`), calcule par vendeur
`jours de stock restant - jours d'attente de réapprovisionnement`, puis un score : >3 jours → 3, >0 → 2,
=0 → 1, <0 → 0 (négatif = rupture de stock avant le prochain réapprovisionnement). La commune reçoit la
**médiane** de ce score par article, sommée sur tous les articles concernés. (RE.2 du guide est fusionné
dans ce même calcul, pas un indicateur séparé dans cet outil.)

**RE.3** (max 12 = 3 x 4 catégories) - % de vendeurs dépendant majoritairement d'un seul fournisseur, par
catégorie de produits (alimentaire / non-alimentaire / WASH / combustible - 4 questions dédiées :
`food_supplier_unique_ctgy`, `nfi_supplier_unique`, `wash_supplier_unique`,
`fire_combustion_supplier_unique`). Score par catégorie : <25 % → 3 · 25-50 % → 2 · 50-75 % → 1 · ≥75 % → 0.
Somme des 4 catégories.

**RE.4** (max 12) - % de vendeurs cochant une option *autre que* "aucune difficulté" / "ne sait pas" /
"préfère ne pas répondre" (`mfs_re4_shop_operational_difficulties_sm`).
Seuils : <5 % → 12 · 5-10 % → 9 · 10-25 % → 6 · 25-50 % → 3 · ≥50 % → 0.

`score_résilience = (RE.1 + RE.3 + RE.4) / (3*nb_articles_stock + 12 + 12) * 20`

## Pilier 5 - Infrastructures (poids 10, échelle brute /10)

**IN.1** (max 4) - % de vendeurs ayant coché "bâtiments dangereux" (même question que AC.1/AC.2, choix
`batiments_dangereux` - réutilisée telle quelle, comme dans le guide).
Seuils : <5 % → 4 · 5-10 % → 3 · 10-25 % → 2 · 25-50 % → 1 · ≥50 % → 0.

**IN.2** (max 3) - % de vendeurs dont la réponse à `mfs_in2_available_storage_facility_so` n'est *pas*
une solution de stockage adéquate (stockage ailleurs sur le marché ou dans son propre local) - stockage à
domicile/autre/etc. comptent comme un problème.
Seuils : <10 % → 3 · 10-25 % → 2 · 25-50 % → 1 · ≥50 % → 0.

**IN.3** (max 3) - % de vendeurs acceptant au moins un mode de paiement alternatif (colonne
`*_payment_modality_sm`, en excluant "espèces (monnaie locale)", "ne sait pas", "préfère ne pas
répondre" et "aucun mode accepté").
Seuils : >75 % → 3 · >50 % → 2 · >25 % → 1 · ≤25 % → 0.

`score_infrastructure = (IN.1 + IN.2 + IN.3) / 10 * 10`

## Score total et classification

`score_total = score_accessibilité + score_disponibilité + score_abordabilité + score_résilience + score_infrastructure` (sur 100)

Une dimension est dite **"basse"** si son score est sous 50 % de son maximum (accessibilité <12.5,
disponibilité <15, abordabilité <7.5, résilience <10, infrastructure <5).

La classification suit les catégories du guide, évaluées **de la pire à la meilleure** (le premier critère
vérifié l'emporte - un marché à 60 % avec 2 dimensions basses vérifie à la fois le critère "limitée" et
"mauvaise", mais "mauvaise" doit gagner ; évaluer dans l'autre sens ferait disparaître la catégorie
"problèmes graves") :

1. **Données insuffisantes** - une dimension entière n'a pas pu être calculée (score total manquant).
2. **Problèmes graves** - score total < 25 % OU ≥ 3 dimensions basses.
3. **Mauvaise fonctionnalité** - score total < 50 % OU ≥ 2 dimensions basses.
4. **Fonctionnalité complète** - score total > 80 % ET 0 dimension basse.
5. **Fonctionnalité limitée** - tous les autres cas.

## Maintenance - à vérifier à chaque nouveau round

Les codes de réponse kobo (listes de choix) changent souvent d'un round à l'autre. Une comparaison à un
code qui n'existe plus **ne provoque pas d'erreur** - le résultat est silencieusement 0 (ou FALSE) pour
tout le monde, ce qui peut rendre un indicateur artificiellement "parfait" ou "au plus mauvais" sans aucun
signal visible. Tous les noms de question et codes de choix utilisés sont déclarés une seule fois, en haut
de `scripts/mfs.R` (section "CHOICE CODES") - à re-vérifier contre les feuilles `survey`/`choices` de
l'outil kobo du mois avant de faire confiance à un nouveau calcul, plutôt que de chercher les chaînes de
caractères dans tout le script.

Vérifié le 2026-08-21 contre `RCA_JMMI_2026-08_TOOL.xlsx` : toutes les questions et tous les codes de
choix listés ci-dessus sont bien présents dans l'outil actuel.

## Historique

Le calcul du SFM (avec la section "Analyse" descriptive) avait été temporairement désactivé dans
`1_analyse_complete.R`/`4_analyse_complete.R` (ces deux sections étant les plus lentes à exécuter, sur
demande utilisateur, pour ne produire que "Coût médian du PMAS" + "Prix médian par article" pendant une
phase de mise au point). Réactivé le 2026-08-21 - la section "Analyse" (LOA descriptive) reste désactivée
séparément (toujours dans un bloc `if (FALSE)`), le SFM tournant maintenant indépendamment à chaque
exécution.
