# Golden cells — bench recalcul BD → Synthesis

5 cas témoin à valider après chaque changement moteur / perf.

| # | Action | Cellule à vérifier | Attendu | Cible perf |
|---|--------|-------------------|---------|------------|
| 1 | Edit masse BD ligne 50 col masse | Syn ligne blue ~50, col H | Masse recalculée (SUMPRODUCT) | < 100 ms |
| 2 | Edit filtre Syn row 15 col H | Syn row 15 col H | Filtre appliqué, masses blue mises à jour | < 100 ms |
| 3 | Edit Syn adaptation row 30 col D | Syn row 25 col C (somme) | Total bande ADAPTATION | < 100 ms |
| 4 | F5 cold start | BD visible | Grille BD sans freeze | < 3 s |
| 5 | Menu BD → Synthesis (2e fois) | Syn scroll row 1 | Navigation instantanée | < 200 ms |

## BD → Synthesis live dependency (non négociable)

La grille Synthesis doit afficher des valeurs **live** calculées depuis l’index BD en mémoire (`workbookSession`) et non des valeurs copiées dans le JSON.

- Démarrage: l’index BD se charge dès que `bdSheet` est prêt (pas besoin de cliquer).
- Invalidation: un edit BD sur la masse (`V`), le L2 (`AU`) ou les filtres (`O`/`P`) doit provoquer un bump `synCalcTick` et un recalcul des cellules bleues (ex: **M30**).

## Console bench

Après chargement de l'app (Ctrl+F5) :

```js
await window.__runPerfBench?.()
```

## Trace perf par edit (optionnel)

Ouvrir l'app avec `?perf=1` (ex: `http://127.0.0.1:5173/?perf=1`) puis faire les edits des cas 1–3.
La console affichera des lignes du type:

- `[perf] edit BD ... total ...ms | invalidate ...ms | engine ...ms | post ...ms`
- `[perf] edit SYNTHESIS ...: ...ms`

## Notes

- Colonnes **K+** : valeurs preset / JSON — pas encore recalcul live moteur complet.
- Si un cas échoue, noter machine + navigateur + temps perçu.
