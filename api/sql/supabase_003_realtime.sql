-- Supabase SQL Editor -> New query -> coller -> Run
-- Active le Realtime (websocket) sur les cellules, pour que toute modification
-- de la feuille Synthesis (ligne 16 "Curb mass") soit poussee en direct vers
-- la page "Options SP2" (ligne 14) — sans rechargement, multi-onglets/utilisateurs.

-- 1) Le navigateur doit recevoir l'ancienne ET la nouvelle ligne (filtre sheet=SYNTHESIS,
--    suppressions, etc.) -> REPLICA IDENTITY FULL.
ALTER TABLE public.sheet_cells REPLICA IDENTITY FULL;

-- 2) Ajouter la table a la publication Realtime de Supabase.
--    (ignore l'erreur "already member of publication" si deja fait.)
ALTER PUBLICATION supabase_realtime ADD TABLE public.sheet_cells;

-- 3) Le client se connecte avec la cle "anon". RLS est desactive sur cette table
--    (voir supabase_001_schema.sql), donc l'anon peut lire les changements.
--    Si vous reactivez RLS un jour, il faudra une policy SELECT pour 'anon'.

-- Verification : la table doit apparaitre dans la publication.
-- SELECT * FROM pg_publication_tables WHERE pubname = 'supabase_realtime';
