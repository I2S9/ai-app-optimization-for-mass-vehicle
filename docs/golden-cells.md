# Golden cells — tests calcul & sauvegarde

Remplir après edit dans l’app. Repères : **colonne Excel** + **ligne Excel** (`excelRow` dans la grille si affiché).

**Formules BD** : lues depuis `bd-sheet.json` (champ `f` à l’export). Vous n’avez pas à les saisir ici sauf **exception** (nouvelle règle web ≠ Excel) — dans ce cas, une ligne « règle » suffit, pas tout le classeur.

## Calcul BD → Synthesis

| # | Action | Cellule BD | Cellule Syn à vérifier | Excel attendu | App actuel | OK ? |
|---|--------|------------|------------------------|---------------|------------|------|
| 1 | Edit masse | ex. O47 = | ex. T14 = | | | |
| 2 | Edit filtre Syn | — | ex. ligne filtre col G | | | |
| 3 | | | | | | |

## Formules BD (HyperFormula)

| # | Cellule (col + row) | Excel | App après edit dépendance | OK ? |
|---|---------------------|-------|---------------------------|------|
| 4 | | | | |
| 5 | | | | |

## Sauvegarde (après API Databricks)

| # | Action | Résultat attendu | OK ? |
|---|--------|------------------|------|
| 6 | Edit + Enregistrer + F5 | Valeur identique | |
| 7 | Deux onglets même projet | Pas d’écrasement (revision) | |

## Synthesis affichage (hors moteur)

Colonnes A–J : valeurs OK sur beaucoup de lignes — noter ici seulement les **écarts restants** (K+, couleurs).

| Zone | Colonnes | Lignes | Problème | Priorité |
|------|----------|--------|----------|----------|
| ex. véhicules | K–P | 10–20 | couleur | P2 |
