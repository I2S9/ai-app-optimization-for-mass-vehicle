# Base de donnees vs calculs temps reel

## Ce que fait la base (plus rapide au demarrage)

- Charge **BD / Synthesis par blocs** (pas 12 Mo JSON d'un coup)
- **Sauvegarde** matrix + edits (table `workbook_sessions`)
- Partage entre sessions (Postgres local ou Supabase cloud)

## Ce qui reste dans le navigateur (plus rapide a l'edition)

- **Synthesis SUMPRODUCT** apres edit Database
- **Bookmark Matrix** (reorganisation)
- Moteur `WorkbookSession` + cache

Mettre chaque calcul dans la base = 1 requete reseau par cellule = **plus lent**, pas plus rapide.

## Sur reseau Stellantis

| Option | Fichier |
|--------|---------|
| Supabase cloud | Bloque (port 6543) |
| **Postgres Docker local** | `DEMARRER-avec-base-locale.bat` |
| JSON seul | `run-bd-server.bat` + `DATA_BACKEND=default` |

## Demarrage avec base locale

1. Docker Desktop installe
2. Double-clic `DEMARRER-avec-base-locale.bat`
3. Application = `run-bd-server.bat` a la fin

Verif : http://127.0.0.1:8000/api/v1/config → `"mode": "postgres"`
