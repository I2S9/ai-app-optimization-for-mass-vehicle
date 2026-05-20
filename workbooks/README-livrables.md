# Références livrables — `workbooks/livrables/`
Ce dossier regroupe les classeurs Excel de référence pour la **parité fonctionnelle et visuelle** de l’application web (WGHT Dashboard).  
Source de vérité calcul pour le cœur métier : **`base-de-donnees-complete-avec-liens.xlsm`** (toutes les feuilles liées entre elles).
Les autres fichiers servent surtout de **gabarits de sortie** (mise en page, couleurs, graphiques) ou de **modules indépendants** (Portfolio, Z-CdG).
---
## Environnement Excel (équipe)
| Élément | Valeur |
|--------|--------|
| Produit | Microsoft Excel pour **Microsoft 365** |
| Licence | Office 365 **E3** (version bureau récente) |
| Version relevée | **2404** — build **16.0.17531.20172** |
| Architecture | **32 bits** (Windows) |
Les formules et graphiques des références doivent rester compatibles avec ce niveau (fonctions classiques + éventuelles fonctions 365 récentes à inventorier dans le classeur principal).
**Macros :** non vérifiées à la main ; les fichiers `.xlsm` peuvent en contenir. L’app web vise les **formules cellule** et les **exports gabarit**, pas l’exécution VBA sauf décision ultérieure.
---
## Inventaire des fichiers (noms sur disque)
| Fichier | Taille (ordre de grandeur) | Rôle |
|---------|----------------------------|------|
| `base-de-donnees-complete-avec-liens.xlsm` | ~22 Mo | **Cœur lié** : BD, Synthesis, livrables intégrés, feuilles intermédiaires |
| `cdc.xlsm` | ~2 Mo | **Gabarit sortie CDC** uniquement (autre base collègue — ne pas fusionner les données) |
| `cdg-outil-evaluation.xlsx` | ~340 Ko | **Gabarit sortie CDG** — onglet cible `OUTPUT - Mass & CoG` |
| `generation-graphes-waterline.xlsm` | ~15 Mo | **Waterline** — données dispersées + génération graphique sur `Waterline tool` |
| `portfolio.xlsx` | ~1 Mo | **Portfolio** — onglet `Portefeuille SP2` (saisie / export indépendant de BD) |
| `split-vehicules.xlsx` | ~4,8 Mo | **Split** — comparaison masses / graphiques (menu déroulant véhicules) |
| `z-cdg.xlsm` | ~2 Mo | **Z-CdG** — saisie manuelle (tableau simple) |
Les exports JSON utilisent exclusivement les fichiers de ce dossier (pas de copies à la racine `workbooks/`).
---
## Graphe de dépendances métier (résumé)
```
                    ┌─────────────────────────────────────┐
                    │  BD (Database) — cœur données       │
                    └──────────────┬──────────────────────┘
                                   │
         ┌─────────────────────────┼─────────────────────────┐
         ▼                         ▼                         ▼
   SYNTHESIS              OUTPUT FOR CDC (intermédiaire)   Weight Tax, Graph PTF limits, …
         │                         │
         │                         ▼
         │                    CDC (dans xlsm principal)
         │                         │
         └────────────┬────────────┘
                      ▼
              Waterline (BD + CDC)
                      │
         Target status forecast ← SYNTHESIS
         Split ← BD + SYNTHESIS
   Portfolio, Z-CdG : indépendants (saisie / gabarit séparé)
```
- **Feuilles intermédiaires** : présentes dans Excel pour décomposer le travail ; **non exposées dans l’app** — le moteur peut les **calculer en interne** (plus rapide) ou les ignorer si équivalent direct BD → livrable.
- **Livrables visibles** : ce que l’équipe consulte ou exporte (BD, Synthesis, CDC format, CDG OUTPUT, Waterline + graphe, Split Cascading, etc.).
---
## Décisions produit (validées)
| Sujet | Décision |
|-------|----------|
| **Feuilles intermédiaires** (maître + outils) | **Masquées** — pas d’écran dédié. Calcul possible en arrière-plan sans reproduire le découpage Excel si on peut aller plus vite. |
| **Portfolio** | **Ignorer** l’onglet `Portfolio ` du maître. Source unique : **`portfolio.xlsx`** → onglet **`Portefeuille SP2`** (écran web + export). |
| **Z-CdG** | **`z-cdg.xlsm`** remplace l’onglet `Z-CdG` du maître (ne pas utiliser celui du classeur lié pour l’UI). |
| **CDC** | Gabarit **`CdC STLAM-HEV(EB2)`** (`cdc.xlsm`) — **format uniquement** ; données alimentées par **BD + OUTPUT FOR CDC** (calcul interne), **sans** données du classeur collègue. |
| **Split** | Écran + export = onglet **`Cascading`** (`split-vehicules.xlsx`). Autres onglets = intermédiaires (alimentation données). |
| **Waterline** | Écran = feuille **`Waterline`** du **maître** ; bouton **« Générer graphe »** reproduit en arrière-plan la logique de **`generation-graphes-waterline.xlsm`** (`Waterline tool` / Adjust axis) → visualisation **dans l’app** + **export image** (graphe complexe arc-en-ciel / jauge, pas un chart générique). |
| **Noms d’onglets avec espaces** | **Conserver à l’identique** pour l’export Excel (ex. `Weight Tax `) — au moins en v1. |
---
## 1. `base-de-donnees-complete-avec-liens.xlsm`
Classeur principal **multi-feuilles liées**. C’est la base pour le **moteur de calcul** (dépendances `=BD!…`, `=SYNTHESIS!…`, etc.) et pour une grande partie des écrans web.
### Onglets relevés automatiquement
| Onglet Excel | Module app / livrable | Notes |
|--------------|----------------------|--------|
| `BD` | **Database** | Cœur projet — données véhicule / masse |
| `SYNTHESIS` | **Synthesis** | Dépend de BD (ex. formules sur plages type `BD!$O$2:$O$3480`) |
| `Portfolio ` | **Hors scope UI** | Ne pas utiliser — Portfolio = `portfolio.xlsx` uniquement |
| `OUTPUT FOR CDC` | Intermédiaire (masqué) | Calcul interne vers livrable CDC |
| `Waterline` | **Waterline** (écran) | Données + sélection véhicules ; graphe via bouton + ref. `generation-graphes-waterline.xlsm` |
| `Weight Tax ` | Taxe masse | Espace final dans le nom — graphiques |
| `Graph PTF limits` | Limites de dimension | Menu app : « Limites de dimension » |
| `Target status forecast` | Target-Status-Forecast | Dépend de **Synthesis** + graphique |
| `Z-CdG` | **Hors scope UI** | Remplacé par `z-cdg.xlsm` |
| `CPVpourRepartDec5`, `RepartitionDec5` | Intermédiaires (masqués) | Peuvent servir au moteur / export JSON legacy |
| `Bhatch optim`, `SYNTHESIS CDC V19`, `MNS`, `Options SP2` | Intermédiaires (masqués) | Pas d’écran — accélérer par calcul direct si possible |
### Exemple de formule cross-feuille (Synthesis ← BD)
```excel
=*((BD!$O$2:$O$3480=T$14)+(BD!$O$2:$O$3480="TT"))
```
À valider en spike moteur (recalcul incrémental sur grande plage).
---
## 2. `cdc.xlsm` — gabarit CDC uniquement
**Attention :** classeur d’un **autre collègue / autre base** — ne pas importer ses données dans BD.  
Usage projet : **reproduire le format de sortie** (couleurs, organisation, structure) pour le livrable CDC de **votre** chaîne BD → OUTPUT FOR CDC → …
### Onglets présents
`HEV SYNTH`, `HEV EB2-EP6 - WEIGHTS`, `HEV EB2 - WALK`, `P64 PHEV LR`, `CdC-GAP_value`, `CdC-Donnees CDC`, `CdC-Inputs VP`, **`CdC STLAM-HEV(EB2)`**
### Livrable cible (format)
| Onglet | Rôle |
|--------|------|
| **`CdC STLAM-HEV(EB2)`** | **Référence visuelle et structure** (modèle véhicule STLAM-HEV EB2) — formatting identique, **données alimentées par notre pipeline** |
Les autres onglets = calculs intermédiaires de l’outil collègue (aperçu uniquement).
---
## 3. `cdg-outil-evaluation.xlsx` — gabarit CDG
Même logique : feuilles d’entrée / intermédiaires pour l’équipe ; **sortie produit** = un onglet.
| Onglet | Rôle |
|--------|------|
| `Updates`, `INPUT1 - Veh Geometry`, `INPUT2 - Veh Sub-Syst Mass`, `UNSPRUNG Mass estimation`, etc. | Intermédiaires / outil |
| **`OUTPUT - Mass & CoG`** | **Livrable** — même mise en page, **seules les données changent** |
---
## 4. Waterline — maître + `generation-graphes-waterline.xlsm`
### Écran produit (utilisateur)
| Élément | Détail |
|---------|--------|
| Feuille affichée | **`Waterline`** dans `base-de-donnees-complete-avec-liens.xlsm` |
| Action | Bouton **« Générer graphe »** |
| Résultat | Visualisation **dans l’app** du graphe complexe (arc-en-ciel / jauge, modèle vs concurrence) |
| Export | **Image** du graphe (en plus des exports données / Excel si besoin) |
### Fichier de référence technique (non affiché comme grille)
`generation-graphes-waterline.xlsm` sert à **comprendre et reproduire** la génération du graphe (logique aujourd’hui sur **`Waterline tool`**, action type **« Adjust axis »**).  
Données dispersées sur de nombreux onglets (interne, concurrence, segments, taxation, etc.) — **intermédiaires masqués** ; l’app doit refaire ce pipeline **en arrière-plan**, pas demander à l’utilisateur d’ouvrir ces feuilles.
Onglets notables (référence dev) : `Waterline tool`, `Official Waterlines`, `Waterline OD`, familles `ONE-X`, `HEV` / `PHEV` / `BEV`, `Database`, `Summary`, etc.
---
## 5. `portfolio.xlsx` — Portfolio (source canonique)
| Onglet | Rôle |
|--------|------|
| **`Portefeuille SP2`** | **Seul onglet** — saisie idées de projet, **format identique** à l’écran web et à l’**export** |
| Autres (`GLOBAL PORTFOLIO`, `export SP1 S26`, …) | Hors scope |
- **Ne pas** utiliser l’onglet `Portfolio ` du classeur maître.
- **Pas de dépendance BD.**
- Fichier **partagé** → persistance serveur / conflits à prévoir.
- Code actuel : `src/lib/portfolioExcel.ts` (feuille `Portefeuille SP2`).
---
## 6. `split-vehicules.xlsx` — Split
| Onglet | Rôle |
|--------|------|
| **`Cascading`** | **Écran app + export** — tableau interactif + graphiques (comparaison masse composants) |
| `Isodensity P3U`, `SUB-SYSTEM SPLIT`, `Isodensity J1X`, `Isodensity J2U`, `SUB-SYST SPLIT OD`, `Feuil1` | **Intermédiaires** — récupération / calcul de données (masqués) |
- Menu déroulant : sélection des véhicules (comme Excel).
- Dépendances données : **BD + Synthesis** (classeur maître).
---
## 7. `z-cdg.xlsm` — Z-CdG (remplace l’onglet du maître)
| Onglet | Rôle |
|--------|------|
| **`Z-CdG`** | Tableau simple — **saisie manuelle** (écran `/z-cdg` + export) |
| `CPVpourRepartDec5`, `RepartitionDec5` | Intermédiaires / legacy — **masqués** sauf besoin moteur |
- L’onglet `Z-CdG` du **maître** n’est **pas** utilisé pour le produit.
- Pas de dépendance BD (saisie utilisateur).
---
## Correspondance menu application ↔ références
| Route / menu app | Référence principale |
|----------------|---------------------|
| `/database` | `BD` dans `base-de-donnees-complete-avec-liens.xlsm` |
| `/synthesis` | `SYNTHESIS` (idem) |
| `/portfolio` | `portfolio.xlsx` → `Portefeuille SP2` |
| `/cdc` | Format : `cdc.xlsm` → `CdC STLAM-HEV(EB2)` ; données : chaîne BD / OUTPUT FOR CDC |
| `/cdg` | `cdg-outil-evaluation.xlsx` → `OUTPUT - Mass & CoG` |
| `/waterline` | Maître → `Waterline` + bouton graphe (ref. `generation-graphes-waterline.xlsm`) |
| `/split` | `split-vehicules.xlsx` → **`Cascading`** |
| `/weight-tax` | `Weight Tax ` (feuille du maître) |
| `/graph-ptf-limits` | `Graph PTF limits` |
| `/target-status-forecast` | `Target status forecast` |
| `/z-cdg` | `z-cdg.xlsm` → `Z-CdG` (pas l’onglet du maître) |
---
## Stack technique (actuel et cible)
### Application web (en place)
| Couche | Technologie | Usage |
|--------|-------------|--------|
| Framework | **Next.js 15** (App Router) | Pages, API (`/api/portfolio-state`, etc.) |
| Langage | **TypeScript** | Modèles, libs, composants |
| UI | **React 19** | Grilles, formulaires, navigation |
| Styles | **Tailwind CSS 3** | Fond blanc, accents bleu marine, pas d’emojis |
| Grilles performantes | **`SheetVirtualGrid`** + **`@tanstack/react-virtual`** | Database, Synthesis (scroll virtualisé) |
| I/O Excel (import/export léger) | **SheetJS (`xlsx`)** | Seeds JSON, exports données simples (sans styles complets) |
| Persistance interim | **`localStorage`** + fichier **`data/portfolio-state.json`** | Autosave BD/Synthesis ; Portfolio serveur |
| Données statiques build | **`public/data/*.json`** | Générées par `tools/Export-WorkbookData.ps1`, `export:synthesis-bundle`, etc. |
### Cible (à construire — chemin critique)
| Couche | Technologie prévue | Usage |
|--------|-------------------|--------|
| **Moteur de calcul** | **HyperFormula** (ou équivalent) dans un **Web Worker** | Formules Excel, refs inter-feuilles, recalcul incrémental |
| Modèle classeur | **`WorkbookSession`** (TypeScript) | Feuilles, cellules sparse, révision, cellules « dirty » |
| Persistance | **API Next** + JSON versionné ou **Postgres** | Un projet = une session ; plus de feuilles isolées |
| **Exports livrables** | **Python 3 + openpyxl** (service ou CLI) | Copie **gabarit** `.xlsx`/`.xlsm` + injection valeurs/formules + styles |
| Graphe Waterline | **Canvas / SVG / lib chart** côté React (ou rendu serveur) | Reproduire la logique de `generation-graphes-waterline.xlsm` ; export **PNG/SVG** |
| IA / optimisation (Phase 8) | **Python (FastAPI)** | Lit un snapshot structuré du moteur — pas le DOM |
**Hors scope runtime :** Excel ouvert pour chaque edit ; recalcul complet à chaque frappe ; n8n comme moteur (orchestration ops seulement si besoin).
Référence agents : fichier **`.cursorrules`** à la racine du dépôt (même plan, règles de dev).
---
## Moteur de calcul — principe et efficacité
### État actuel (sans moteur)
- Les JSON (`sheet-bd-values.json`, etc.) contiennent des **valeurs figées** exportées depuis Excel.
- Les **formules** sont stockées en texte (`formulas[][]`) mais **ne s’exécutent pas** : modifier une cellule ne met pas à jour les cellules dépendantes ni les autres feuilles (`=BD!…`).
- Conséquence : l’app est une **grille éditable**, pas un **classeur vivant** — d’où l’impression que « rien ne se met à jour ».
### Rôle du moteur
Le moteur joue le rôle d’**Excel en mémoire** :
1. Stocke valeur saisie **ou** formule par cellule.
2. Construit un **graphe de dépendances** (qui dépend de qui).
3. À chaque modification : **recalcul incrémental** uniquement sur les cellules impactées — pas tout le classeur.
4. Expose à l’UI : valeur affichée, formule, **antécédents** / **dépendants** (audit « envers du décor »).
### Pourquoi ce sera plus rapide qu’Excel (objectif produit)
| Excel (lent aujourd’hui) | App + moteur bien conçu |
|--------------------------|-------------------------|
| Fichier `.xlsm` ~22 Mo, 32 bits | Session allégée, données utiles seulement |
| Recalcul souvent très large | Recalcul **ciblé** (sous-graphe) |
| UI + graphiques + macros dans le même process | Calcul en **Worker** ; UI ne redraw que le viewport |
| Milliers de lignes materialisées | Stockage **sparse** + grilles **virtualisées** |
**Objectifs perf (à valider en Phase 0–2) :**
| Métrique | Cible |
|----------|--------|
| Ouverture session BD + Synthesis | &lt; 3–5 s (recalcul initial une fois) |
| Edit cellule typique (recalcul incrémental) | **&lt; 100 ms** perçu |
| Pire cas (grosse plage) | &lt; 1 s ou indicateur « Calcul en cours… » |
### Formules prioritaires (v1)
- `SOMME`, différences, `SOMME.PROD`, comparaisons
- Références inter-feuilles : `BD!…`, `SYNTHESIS!…`
- Plages longues type `=*((BD!$O$2:$O$3480=T$14)+…)` — évaluation **vectorielle** dans le moteur (spike Phase 2)
### Feuilles dans le modèle moteur
| Type | Exemples | UI |
|------|----------|-----|
| **Visibles** | `BD`, `SYNTHESIS`, `Waterline`, `Weight Tax `, `Graph PTF limits`, `Target status forecast` | Écrans app |
| **Internes (masquées)** | `OUTPUT FOR CDC`, `SYNTHESIS CDC V19`, `Bhatch optim`, onglets Split hors `Cascading`, onglets `generation-graphes-waterline.xlsm` | Calcul uniquement |
| **Hors maître** | `portfolio.xlsx`, `z-cdg.xlsm` | Modules séparés |
---
## Plan de réalisation — phases, tâches et features
### Phase 0 — Cadrage et baseline (≈ 1 semaine)
**Objectif :** mesurer et inventorier avant de coder le moteur.
| Tâche | Livrable |
|-------|----------|
| Lister onglets, plages utilisées, volume cellules non vides | Rapport inventaire xlsm |
| Extraire échantillon formules cross-feuilles (BD → Synthesis, etc.) | Liste + 20 cellules « golden tests » |
| Détecter macros VBA (présence ou non) | Note compatibilité |
| Mesurer perf Excel vs app actuelle (chargement, scroll) | Tableau baseline |
| Aligner scripts export sur `base-de-donnees-complete-avec-liens.xlsm` | Manifest / JSON à jour |
---
### Phase 1 — Session unique et persistance (≈ 2–3 semaines)
**Objectif :** une vérité partagée ; fin de la perte de données au rechargement.
| Tâche | Feature |
|-------|---------|
| Modèle `WorkbookSession` (id, revision, sheets) | API REST Next |
| Sauvegarde/chargement serveur | Statut « Enregistré » dans l’UI |
| Brancher Synthesis sur la **même session** que BD | Edit BD visible dans Synthesis (valeurs, avant moteur) |
| Conserver secours `localStorage` si défini | Résilience offline |
**Features utilisateur :** données cohérentes entre `/database` et `/synthesis` ; pas de retour au JSON figé après F5.
---
### Phase 2 — Spike moteur de calcul (≈ 1–2 semaines)
**Objectif :** prouver la faisabilité **&lt; 100 ms** sur formules réelles.
| Tâche | Livrable |
|-------|----------|
| POC HyperFormula en Web Worker sur extrait BD + Synthesis | Démo technique |
| Tester formule plage `BD!O2:O3480` | Temps p50 / p95 |
| Lister fonctions Excel non supportées | Plan de contournement |
| Décision go / no-go | Document court |
---
### Phase 3 — Moteur intégré + audit calcul (≈ 3–5 semaines)
**Objectif :** classeur vivant sur BD + Synthesis.
| Tâche | Feature |
|-------|---------|
| `getCell` / affichage = **valeur calculée** | Grilles BD & Synthesis à jour en temps réel |
| Recalcul incrémental sur edit | Fluidité cible |
| Barre de formule fiable | Texte formule réel |
| Panneau **antécédents / dépendants** | « Envers du décor » pour experts |
| Feuilles intermédiaires dans le modèle, **masquées** | OUTPUT FOR CDC calculé sans écran |
| Chargement : 1 recalcul global + écran chargement | Acceptable au démarrage |
---
### Phase 4 — Performance UI (≈ 2–3 semaines)
**Objectif :** scroll et frappe fluides sur gros volumes.
| Tâche | Feature |
|-------|---------|
| Virtualiser **Portfolio** (`PortfolioWorkbookClient`) | Même perf que Database |
| Mises à jour **dirty cells** seulement | Moins de lag React |
| Worker : aucun recalcul sur le thread principal | UI reste responsive |
| Optimiser chargement JSON / session | Temps d’ouverture réduit |
---
### Phase 5 — Exports livrables gabarit (≈ 3–5 semaines)
**Objectif :** fichiers **identiques** (couleurs, fusions, formules Excel) — pas grille vierge.
| Livrable | Gabarit | Données |
|----------|---------|---------|
| CDC | `cdc.xlsm` → `CdC STLAM-HEV(EB2)` | BD + OUTPUT FOR CDC (interne) |
| CDG | `cdg-outil-evaluation.xlsx` → `OUTPUT - Mass & CoG` | Pipeline CDG |
| Portfolio | `portfolio.xlsx` → `Portefeuille SP2` | Saisie app |
| Split | `split-vehicules.xlsx` → `Cascading` | BD + Synthesis |
| BD / Synthesis | Maître ou export dédié | Snapshot moteur |
| Tâche | Stack |
|-------|--------|
| Service Python openpyxl | Copie template + patch cellules |
| API Next `POST /api/export/...` | Déclenche export, retour fichier |
| Tests non-régression visuelle | Comparaison avec références du dossier `livrables/` |
---
### Phase 6 — Graphiques Excel et in-app (≈ 2–4 semaines)
| Module | Tâche | Feature |
|--------|-------|---------|
| Taxe masse | Plages données + template chart | Export xlsx avec graphiques |
| Target-Status-Forecast | Idem (dépend Synthesis) | Export + affichage |
| **Waterline (6b)** | Reverse-engineer `Waterline tool` / Adjust axis | Bouton **Générer graphe** |
| Waterline | Rendu arc-en-ciel / jauge dans l’app | Visualisation interactive |
| Waterline | Export **image** (PNG/SVG) | Livrable graphique |
---
### Phase 7 — Modules restants (continu)
| Route | Tâches |
|-------|--------|
| `/cdc` | Écran + export format STLAM ; pas de données collègue |
| `/cdg` | Écran OUTPUT Mass & CoG |
| `/waterline` | Feuille maître + graphe |
| `/split` | UI Cascading, sélecteur véhicules, graphiques |
| `/weight-tax`, `/graph-ptf-limits`, `/target-status-forecast` | Brancher session + exports |
| `/z-cdg` | Tableau manuel (déjà partiel) |
| Pages vides (CDC, Waterline, Split si placeholder) | Remplacer `EmptySectionPage` |
---
### Phase 8 — IA et optimisation masse (Python, après moteur stable)
| Tâche | Feature |
|-------|---------|
| Modèle métier (modules, masses, contraintes) depuis snapshot session | API Python |
| Scénarios d’optimisation réalistes | Recommandations dans l’UI |
| Bench / historique | Aide à la décision continue |
**Règle :** l’IA lit le **snapshot calculé**, jamais le DOM cellule par cellule.
---
## Carte des features (routes app)
| Route | Feature | Moteur | Export gabarit |
|-------|---------|--------|----------------|
| `/database` | Grille BD, edit, sections, undo | Phase 3 | Phase 5 |
| `/synthesis` | Grille Synthesis liée BD | Phase 3 | Phase 5 |
| `/portfolio` | Portefeuille SP2, partagé | Indépendant | Phase 5 |
| `/cdc` | Cas de charges | Phase 3 + interne CDC | Phase 5 |
| `/cdg` | Mass & CoG | Phase 3+ | Phase 5 |
| `/waterline` | Données + graphe | Phase 3 + 6b | Image + option xlsx |
| `/weight-tax` | Taxe masse + charts | Phase 3 | Phase 6 |
| `/graph-ptf-limits` | Limites dimension | Phase 3 | Phase 5 |
| `/target-status-forecast` | Forecast + chart | Phase 3 | Phase 6 |
| `/z-cdg` | Saisie manuelle | Indépendant | Phase 5 |
| `/split` | Cascading compare | Phase 3+ | Phase 5 |
---
## Roadmap — ordre résumé
1. **Phase 0** — Inventaire + baseline perf + golden tests  
2. **Phase 1** — `WorkbookSession` + persistance + lien BD → Synthesis  
3. **Phase 2** — Spike HyperFormula (go/no-go)  
4. **Phase 3** — Moteur intégré + audit calcul  
5. **Phase 4** — Perf UI (Portfolio virtualisé, dirty updates)  
6. **Phase 5** — Exports Python openpyxl (templates)  
7. **Phase 6 / 6b** — Graphiques (Taxe masse, Target-Forecast, Waterline)  
8. **Phase 7** — Écrans modules restants  
9. **Phase 8** — IA optimisation Python  
---
## Onglets du maître — visibilité app (synthèse)
| Visibles utilisateur | Masqués (intermédiaires / autre fichier) |
|---------------------|------------------------------------------|
| `BD`, `SYNTHESIS` | `OUTPUT FOR CDC`, `SYNTHESIS CDC V19`, `Bhatch optim`, `MNS`, `Options SP2`, `CPVpourRepartDec5`, `RepartitionDec5`, … |
| `Waterline` (+ graphe généré) | Contenu détaillé de `generation-graphes-waterline.xlsm` |
| `Weight Tax `, `Graph PTF limits`, `Target status forecast` | `Portfolio ` (→ `portfolio.xlsx`) |
| | `Z-CdG` (→ `z-cdg.xlsm`) |
---
## Prochaine étape côté dev
- **Phase 0** : analyser `base-de-donnees-complete-avec-liens.xlsm` (formules, liens BD ↔ Synthesis, volume, macros).
- Aligner exports JSON / manifest sur `base-de-donnees-complete-avec-liens.xlsm` (dossier `livrables/`).
- Spike Waterline : cartographier `Waterline tool` + « Adjust axis » pour reproduction app.
---
## Documents liés
| Fichier | Contenu |
|---------|---------|
| **`.cursorrules`** (racine dépôt) | Règles agents : stack, moteur, phases, checklist |
| **`workbooks/README.md`** | Scripts d’export JSON, emplacement classeurs |
*Dernière mise à jour : plan détaillé phases 0–8, stack, moteur de calcul, décisions produit (Split Cascading, Portfolio, Z-CdG, Waterline, intermédiaires masqués).*
