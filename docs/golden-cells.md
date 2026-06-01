# Golden cells — bench recalcul BD → Synthesis

5 cas témoin à valider après chaque changement moteur / perf.

| # | Action | Cellule à vérifier | Attendu | Cible perf |
|---|--------|-------------------|---------|------------|
| 1 | Edit masse BD ligne 50 col masse | Syn ligne blue ~50, col H | Masse recalculée (SUMPRODUCT) | < 100 ms |
| 2 | Edit filtre Syn row 15 col H | Syn row 15 col H | Filtre appliqué, masses blue mises à jour | < 100 ms |
| 3 | Edit Syn adaptation row 30 col D | Syn row 25 col C (somme) | Total bande ADAPTATION | < 100 ms |
| 4 | F5 cold start | BD visible | Grille BD sans freeze | < 3 s |
| 5 | Menu BD → Synthesis (2e fois) | Syn scroll row 1 | Navigation instantanée | < 200 ms |

## Console bench

Après chargement de l'app (Ctrl+F5) :

```js
await window.__runPerfBench?.()
```

## Notes

- Colonnes **K+** : valeurs preset / JSON — pas encore recalcul live moteur complet.
- Si un cas échoue, noter machine + navigateur + temps perçu.
