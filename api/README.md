# API (Laravel) — coming next
PHP/Laravel is not installed on the initial dev machine. The **Database (BD)** UI runs from `web/` with static JSON until the API is scaffolded.
When PHP 8.2+ and Composer are available:
```bash
composer create-project laravel/laravel api
```
Then: migrations for `mass_lines`, REST endpoints, and replace `web/public/data/bd-sheet.json` with API pagination.
