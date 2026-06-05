/**
 * Supabase Realtime — pousse en direct les modifications de la feuille Synthesis
 * (table public.sheet_cells) vers le navigateur, sans rechargement.
 *
 * Utilise par la page "Options SP2" : la ligne 14 (Curb mass) est liee a la ligne
 * 16 de Synthesis. Quand une cellule R16..AF16 change cote Supabase (edit d'un autre
 * onglet / utilisateur), on recoit l'evenement et on met a jour synRaw -> la ligne 14
 * se rafraichit instantanement.
 *
 * Tout est best-effort : si la lib Supabase n'est pas joignable (pare-feu, CDN bloque)
 * ou si la config manque, on log un avertissement et on ne fait rien — l'app continue
 * de fonctionner avec la synchro intra-session deja en place.
 *
 * Prerequis Supabase (SQL Editor) : voir startSynthesisRealtime() plus bas.
 */

const SUPABASE_ESM = 'https://esm.sh/@supabase/supabase-js@2';
const SYN_SHEET = 'SYNTHESIS';

let clientPromise = null;

async function loadSupabaseClient(url, anonKey) {
  if (!clientPromise) {
    clientPromise = (async () => {
      const mod = await import(/* @vite-ignore */ SUPABASE_ESM);
      const createClient = mod.createClient || (mod.default && mod.default.createClient);
      if (typeof createClient !== 'function') {
        throw new Error('createClient introuvable dans @supabase/supabase-js');
      }
      return createClient(url, anonKey, {
        auth: { persistSession: false, autoRefreshToken: false },
        realtime: { params: { eventsPerSecond: 10 } },
      });
    })();
  }
  return clientPromise;
}

/**
 * Demarre l'abonnement Realtime sur public.sheet_cells (feuille SYNTHESIS).
 *
 * @param {object} opts
 * @param {string} opts.supabaseUrl    https://<ref>.supabase.co
 * @param {string} opts.supabaseAnonKey  cle publique "anon"
 * @param {string} [opts.projectId='default']
 * @param {(change: { row:number, col:string, v:(string|null) }) => void} opts.onSynCellChange
 *        appele a chaque insert/update/delete d'une cellule Synthesis.
 * @returns {Promise<{ stop: () => void } | null>} null si non demarre.
 */
export async function startSynthesisRealtime(opts) {
  const {
    supabaseUrl,
    supabaseAnonKey,
    projectId = 'default',
    onSynCellChange,
  } = opts || {};

  if (!supabaseUrl || !supabaseAnonKey) {
    // Mode statique / Supabase non configure : pas de Realtime, ce n'est pas une erreur.
    return null;
  }
  if (typeof onSynCellChange !== 'function') {
    console.warn('[syn-realtime] onSynCellChange manquant — abonnement annule.');
    return null;
  }

  let supabase;
  try {
    supabase = await loadSupabaseClient(supabaseUrl, supabaseAnonKey);
  } catch (e) {
    console.warn(
      '[syn-realtime] Lib @supabase/supabase-js injoignable (CDN/pare-feu ?) — ' +
        'la synchro temps reel Supabase est desactivee.',
      e
    );
    return null;
  }

  const handlePayload = (payload) => {
    try {
      const rec = payload.new || payload.old;
      if (!rec) return;
      if (String(rec.sheet) !== SYN_SHEET) return;
      const row = Number(rec.excel_row);
      if (!Number.isFinite(row)) return;
      onSynCellChange({
        row,
        col: String(rec.col),
        v: payload.eventType === 'DELETE' ? null : (rec.v ?? null),
      });
    } catch (err) {
      console.warn('[syn-realtime] payload ignore:', err);
    }
  };

  const channel = supabase
    .channel(`syn-cells-${projectId}`)
    .on(
      'postgres_changes',
      {
        event: '*',
        schema: 'public',
        table: 'sheet_cells',
        filter: `sheet=eq.${SYN_SHEET}`,
      },
      handlePayload
    )
    .subscribe((status) => {
      if (status === 'SUBSCRIBED') {
        console.info('[syn-realtime] Abonne aux modifs Synthesis (Supabase Realtime).');
      } else if (status === 'CHANNEL_ERROR' || status === 'TIMED_OUT') {
        console.warn(
          '[syn-realtime] Realtime indisponible (' +
            status +
            '). Verifiez que la table public.sheet_cells est ajoutee a la ' +
            'publication "supabase_realtime" (voir SQL).'
        );
      }
    });

  return {
    stop() {
      try {
        supabase.removeChannel(channel);
      } catch {
        /* noop */
      }
    },
  };
}
