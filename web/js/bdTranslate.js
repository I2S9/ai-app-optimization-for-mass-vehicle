/** French → English for BD sheet values (headers, sections, cell text). */
export const HEADER_FR_EN = {
  Date: 'Date',
  Projet: 'Project',
  Silhouette: 'Silhouette',
  'Plaque de conception': 'Design plate',
  Hybridation: 'Hybridization',
  Sièges: 'Seats',
  'Spécificité technique': 'Technical spec',
  'Charge utile': 'Payload',
  'Pack technique': 'Technical pack',
  Roues: 'Wheels',
  Pôle: 'Pole',
  Energie: 'Energy',
  Moteur: 'Engine',
  Boite: 'Gearbox',
  Finition: 'Trim',
  'Curb mass :': 'Curb mass:',
  'Curb mass : last update': 'Curb mass: last update',
  Roues: 'Wheels',
  Equipts: 'Equipment',
  Option: 'Option',
  Codification: 'Codification',
  Découpage: 'Breakdown',
  Intitulé: 'Title',
  BVH: 'BVH',
  Masse: 'Mass',
  AV: 'Front (AV)',
  AR: 'Rear (AR)',
  'Désignation technique': 'Technical designation',
  Source: 'Source',
  Pack: 'Pack',
  Reference: 'Reference',
  Metier: 'Trade',
  'Positionnement en X': 'X position',
  'Positionnement en Y': 'Y position',
  'Positionnement en Z': 'Z position',
  'Champ libre': 'Free field',
  'CODE MODULE ': 'Module code',
  'LOT DECPSA': 'Lot DECPSA',
  'Type modulaire': 'Modular type',
  'Attribut technique': 'Technical attribute',
  'Sub-System Level1': 'Sub-system L1',
  'Sub-System Level2': 'Sub-system L2',
  'Sub-System Design Dpt': 'Sub-System Design Dpt',
  'Ligne N°': 'Line #',
  'Ligne avec formules': 'Row with formulas',
  'A recopier': 'To copy',
  'Fin de Lot - Ne pas supprimer': 'End of lot - do not delete',
};
/** CA chapter bands only (yellow). */
export const CA_BAND_EN = {
  '-ADAPTATION': '-ADAPTATION',
  '-ADTH': '-CABIN CLIMATE TREATMENT SYSTEM',
};
/** Sub-system L1 labels (blue) — column AP / W / A. */
export const L1_SECTION_EN = {
  AILES: 'FENDERS',
  ALTERNATEUR: 'ALTERNATOR',
  ASSISES: 'SEATS',
  ATTELAGE: 'TOWING',
  BATTERIE: 'BATTERY',
  'BATTERIE TRACTION': 'HIGH-VOLTAGE BATTERY',
  'BOUCLIER AR': 'REAR BUMPER',
  'BOUCLIER AV': 'FRONT',
  BV: 'GEARBOX',
  'CAISSE EN BLANC': 'BODY IN WHITE',
  CAPOT: 'HOOD',
  'CIELO / TO': 'HEADLINER / SR',
  COLONNE: 'STEERING COLUMN',
  CONSOLE: 'CONSOLE',
  DAE: 'EPS = Electric Power Steering',
  EPF: 'EPS = Electric Power Steering',
  DEMARREUR: 'STARTER',
  'DIVERS CAISSE': 'BODY MISC',
  'DIVERS EE': 'EE MISC',
  'DIVERS INT': 'INTERIOR MISC',
  'DIVERS LAS': 'CHASSIS MISC',
  ECLAIRAGE: 'LIGHTING',
  'ECRAN AERO': 'AERO SCREENS',
  ESSUYAGE: 'WIPERS',
  'FACADE AV': 'FACADE',
  FAISCEAUX: 'HARNESS',
  FREINAGE: 'BRAKING',
  GARNISSAGE: 'TRIM',
  GMP: 'POWERTRAIN',
  'HAYON / COFFRE / PORTE COFFRE': 'TAILGATE / TRUNK / TRUNK LID',
  INSONORISATION: 'INSULATION',
  LAS: 'CHASSIS',
  'MOYENS DE RETENU': 'RESTRAINT SYSTEMS',
  'PARE BRISE': 'WINDSHIELD',
  'PLANCHE DE BORD': 'INSTRUMENT PANEL',
  'PORTES AR': 'REAR DOORS',
  'PORTES AV': 'FRONT DOORS',
  'SAC AR': 'REAR AIRBAG',
  'SAC AV': 'FRONT AIRBAG',
  TAPIS: 'CARPET',
  TRANS: 'TRANSMISSION',
  'VITRAGE CAISSE': 'BODY GLAZING',
  VOLANT: 'STEERING WHEEL',
  'SPECIFIQUE PICK-UP': 'PICK-UP SPECIFIC',
  '-Non affecté': '-Unassigned',
  // Legacy EN section titles present in the BD template — unify with Synthesis canonical names.
  'TRACTION BATTERY': 'HIGH-VOLTAGE BATTERY',
  // NB: do NOT map 'FRONT BUMPER' here — it is used as the L2 sub-section "_FRONT BUMPER"
  // under the FRONT section (the FRONT section itself comes from 'BOUCLIER AV': 'FRONT').
  'FRONT END': 'FACADE',
};
/** Yellow section rows — exact labels from the BD template (FR + EN). */
export const SECTION_ALLOWLIST = new Set([
  ...Object.keys(L1_SECTION_EN),
  ...Object.values(L1_SECTION_EN),
  // Legacy EN section titles still present in some templates / snapshots.
  // Keep them here so Database/Synthesis/Matrix detect the same bookmark rows.
  'WINGS',
  'FRONT END',
  'FRONT BUMPER',
  'TRACTION BATTERY',
  // BD template stores this L1 in upper-case; the L1_SECTION_EN value is title-case.
  'EPS = ELECTRIC POWER STEERING',
  '-Non affecté',
  '-Unassigned',
  '_Non affecté',
  '_Unassigned',
  // Upper-case variant stored in the BD template for the catch-all bucket.
  '-UNASSIGNED',
]);
/** Sub-system L2 (_prefix) and common AS labels. */
export const L2_SECTION_EN = {
  _ADDBLUE: '_ADDBLUE',
  _CARBURANT: '_FUEL',
  '_CIRCUIT AIR': '_AIR CIRCUIT',
  '_CIRCUIT EAU': '_WATER CIRCUIT',
  '_COMMANDE VITESSE': '_SPEED CONTROL',
  '_COMPRESSEUR CLIM': '_AIR CONDITIONING COMPRESSOR',
  '_DIVERS ADTH': '_MISCELLANEOUS CABIN CLIMATE TREATMENT SYSTEM',
  _ECHAPPEMENT: '_EXHAUST',
  '_ECRANS THERMIQUE': '_THERMAL SCREENS',
  _FACADE: '_FACADE',
  '_GROUP CLIM': '_AIR CONDITIONING GROUP',
  '_GROUPE ADDITIONNEL': '_ADDITIONAL GROUP',
  '_LIQUIDE ADDITIF': '_ADDITIVE LIQUID',
  '_LIQUIDE REFROIDISSEMENT': '_COOLING LIQUID',
  _PEDALIER: '_PEDAL',
  '_RESERVOIR ADDITIF': '_ADDITIVE TANK',
  '_RESERVOIR CARBURANT': '_FUEL TANK',
  '_Supports GMP': '_ENGINE MOUNTS',
  '_Système de refroidissement par huile': '_OIL COOLING SYSTEM',
  '_Non affecté': '_Unassigned',
};
/**
 * English L2 aliases in BD templates (already translated from FR but not canonical).
 * Maps to the same labels as Synthesis / Bookmark Matrix.
 */
export const L2_ENGLISH_ALIASES = {
  '_COOLANT CIRCUIT': '_WATER CIRCUIT',
  '_COOLANT': '_COOLING LIQUID',
  '_GEAR SHIFT CONTROL': '_SPEED CONTROL',
  '_THERMAL SHIELDS': '_THERMAL SCREENS',
  '_FRONT END': '_FACADE',
  '_ADDITIVE FLUID': '_ADDITIVE LIQUID',
  '_PEDAL ASSEMBLY': '_PEDAL',
  '_POWERTRAIN SUPPORTS': '_ENGINE MOUNTS',
  '_ADTH MISC': '_MISCELLANEOUS CABIN CLIMATE TREATMENT SYSTEM',
  '_HVAC GROUP': '_AIR CONDITIONING GROUP',
  '_AC COMPRESSOR': '_AIR CONDITIONING COMPRESSOR',
  '_ADD-ON GROUP': '_ADDITIONAL GROUP',
  '_ADD6ON GROUP': '_ADDITIONAL GROUP',
  '_ADD ON GROUP': '_ADDITIONAL GROUP',
  '_MISCELLANEOUS CABIN HVACATE TREATMENT SYSTEM':
    '_MISCELLANEOUS CABIN CLIMATE TREATMENT SYSTEM',
};

function titleCasePhrase(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/\b\w/g, (c) => c.toUpperCase());
}

/** Map stripped L2 text via ENGLISH_LABEL_OVERRIDES → canonical _UPPERCASE bookmark. */
function l2FromEnglishOverride(stripped) {
  const candidates = [stripped, titleCasePhrase(stripped), stripped.toLowerCase()];
  for (const c of candidates) {
    if (ENGLISH_LABEL_OVERRIDES[c]) {
      const en = String(ENGLISH_LABEL_OVERRIDES[c]).trim();
      if (!en) continue;
      if (en.startsWith('_')) return en.toUpperCase().replace(/\s+/g, ' ');
      return `_${en.toUpperCase().replace(/\s+/g, ' ')}`;
    }
  }
  return null;
}
/** Frequent cell values (exact match). */
export const CELL_VALUE_EN = {
  'Ligne avec formules': 'Row with formulas',
  'A recopier': 'To copy',
  'Fin de Lot - Ne pas supprimer': 'End of lot - do not delete',
  FIN: 'END',
  TT: 'TT',
  'STLA/S': 'STLA/S',
  'STLA-S': 'STLA-S',
  'Ailes avant': 'FRONT FENDERS',
  'Ailes arrière': 'REAR FENDERS',
  // FENDERS: last two sub-sections use "OLD PROCESS" wording (not "legacy").
  'Avant gauche (ancien process)': 'FRONT LEFT (OLD PROCESS)',
  'Avant droite (ancien process)': 'FRONT RIGHT (OLD PROCESS)',
  'Système batterie 12V': '12V battery system',
  'Système batterie traction': 'Battery system',
  'Caisse en blanc (avec peinture)': 'Body in white (with paint)',
  'Système barre de direction': 'Steering column system',
  'electric power steering': 'Electric Power Steering',
  'Electric power steering': 'Electric Power Steering',
  'Système téléphone': 'Telephone system',
  'Système volant': 'Steering wheel system',
  'Système rétroviseur': 'Mirror system',
  'Câblage de puissance': 'Power wiring',
  'Support non utilisé': 'Unused support',
  'Fixation non utilisée': 'Unused fixing',
  'Ecrans aérodynamiques': 'Aerodynamic screens',
  'SYSTÈME LAVAGE': 'WASH SYSTEM',
  'Fixations ISOFIX': 'ISOFIX mountings',
  'Système module de charge HT (IDCM)': 'High-voltage charge module system (IDCM)',
  'ASSISE AR RG2': 'BACK SEAT RG2',
  'ASSISE AR RG3': 'BACK SEAT RG3',
  'ASSISE AV': 'FRONT SEAT',
  'Système alternateur': 'Alternator system',
  'SYSTÈME ALTERNATEUR': 'ALTERNATOR SYSTEM',
  ENJO: 'HUB',
  'ENJO CAISSE': 'BODY HUB',
  TOLERIE: 'SHEET METAL',
  EQUIPEMENT: 'EQUIPMENT',
  PAVILLON: 'ROOF',
  HABITACLE: 'CABIN',
  COFFRE: 'TRUNK',
  CARROSSERIE: 'BODY',
  'Plaque de conception': 'Design plate',
  Projet: 'Project',
  'Spécificité technique': 'Technical spec',
  'Pack technique': 'Technical pack',
  'Butée amortissement': 'Bump stop',
  'Garniture arrière': 'Rear trim',
  'Garniture arrière 2': 'Rear trim 2',
  'Seconde garniture arrière': 'Second rear trim',
  'Garnitures latérales': 'Side trim',
  'Partie supérieure': 'Upper part',
  'Enjoliveur arrière': 'Rear molding',
  'Enjoliveur de fixation côté droit': 'Right side fixing molding',
  'Enjoliveur de fixation côté gauche': 'Left side fixing molding',
  'Commande de déverrouillage de console centrale': 'Center console unlock control',
  'Deuxiéme support calculateur gestion batterie HT':
    'Second high-voltage battery management computer support',
  'Système crochet d attelage / Porte vélo':
    'Trailer hitch system / bike carrier',
  'Système de contôle avec téléphone / NFC': 'Control system with phone / NFC',
  'Système de déverrouillage': 'Unlocking system',
  'Système frein a main': 'Handbrake system',
  'Système gestion fermeture centralisée': 'Central locking management system',
  'Système poignée': 'Handle system',
  'Système rails de guidage': 'Guide rail system',
  'Système support téléphone': 'Phone support system',
  'Système toit ouvrant': 'Sunroof system',
  'Système ventilation': 'Ventilation system',
  "Système éclairage d'ambiance": 'Ambient lighting system',
  'Tablier coté conducteur / unique': 'Driver side / single dashboard panel',
  'Tablier coté passager': 'Passenger side dashboard panel',
  'X repère à la roue (facultatif)': 'X wheel reference (optional)',
  'Receptacle clé': 'Key receptacle',
  'Rangement arrière (accoudoir)': 'Rear storage (armrest)',
  'Bac de rangement arrière (mass removed)': 'Rear storage tray (mass removed)',
  'Support porte pièces': 'Parts door support',
  Aérateur: 'Vent',
  Arrière: 'Rear',
  'Ailes arrière': 'REAR FENDERS',
};
const LABEL_COLS = new Set(['A', 'AP', 'AS', 'AR', 'AU', 'W']);
// Canonical overrides for already-English labels (used by bookmark/matrix text).
// Keep these aligned with the expected English terminology.
const ENGLISH_LABEL_OVERRIDES = {
  // Legacy EN labels → canonical EN labels (keeps BD/SYN/Matrix identical).
  WINGS: 'FENDERS',
  Wings: 'Fenders',
  'Front wings': 'Front fenders',
  'Rear wings': 'Rear fenders',
  // Some sources use uppercase section titles.
  'FRONT WINGS': 'FRONT FENDERS',
  'REAR WINGS': 'REAR FENDERS',
  'Front fenders': 'FRONT FENDERS',
  'Rear fenders': 'REAR FENDERS',
  'Front left (legacy process)': 'FRONT LEFT (OLD PROCESS)',
  'Front right (legacy process)': 'FRONT RIGHT (OLD PROCESS)',

  'Coolant circuit': 'Water circuit',
  'Thermal shields': 'Thermal screens',
  'Front end': 'Facade',
  'Adaptive fluid': 'Additive liquid',
  'Adaptive liquid': 'Additive liquid',
  Coolant: 'Cooling liquid',
  'Pedal assembly': 'Pedal',
  'Powertrain supports': 'Engine mounts',
  'Gear shift control': 'Speed control',
  'System alternator': 'Alternator system',
  'Rear seat RG2': 'Back seat RG2',
  'Rear seat RG3': 'Back seat RG3',
  'Towing hook / bicycle rack system': 'Trailer hitch system / bike carrier',
  'Traction battery': 'High-voltage battery',
  'Traction battery system': 'Battery system',
  'HV charging module system (IDCM)': 'High-voltage charge module system (IDCM)',
  'HV Charging module system (IDCM)': 'High-voltage charge module system (IDCM)',
  'HV Charging module system (idcm)': 'High-voltage charge module system (IDCM)',
  'High-voltage charge module system (idcm)':
    'High-voltage charge module system (IDCM)',

  ADTH: 'Cabin climate treatment system',
  'ADTH MISC': 'Miscellaneous cabin climate treatment system',
  'HVAC Group': 'Air conditioning group',
  'AC Compressor': 'Air conditioning compressor',
  'ADD6on group': 'Additional group',
  'ADD-on group': 'Additional group',
  'ADD-ON GROUP': 'ADDITIONAL GROUP',

  // Explicit underscore labels (often appear in bookmark/matrix structure rows).
  _DURITE_DE_VENTILATION_MODULE_CHARGE_HT:
    'High-voltage charge module ventilation hose',
  ['_DURITE_DE_VENTILATION_MODULE_CHARGE_HT'.replaceAll('_', ' ')]:
    'High-voltage charge module ventilation hose',
  '_DURITE DE VENTILATION MODULE CHARGE HT':
    'High-voltage charge module ventilation hose',
  _CONTROLLER_MANAGEMENT_BATTERY_HT:
    'High-voltage battery management computer',
  '_CONTROLLER MANAGEMENT BATTERY HT':
    'High-voltage battery management computer',
  _SECOND_SUPPORT_MODULE_CHARGE_HT:
    'Second high-voltage charge module support',
  '_SECOND SUPPORT MODULE CHARGE HT':
    'Second high-voltage charge module support',
  _SUPPORT_CONTROLLER_MANAGEMENT_BATTERY_HT:
    'High-voltage battery management computer support',
  '_SUPPORT CONTROLLER MANAGEMENT BATTERY HT':
    'High-voltage battery management computer support',
  '_SECOND HV BATTERY MANAGEMENT CONTROLLER SUPPORT':
    'Second high-voltage battery management computer support',
};
/** Nomination / label columns visible on structure rows (yellow & blue). */
export function isLabelColumn(col) {
  return LABEL_COLS.has(col);
}

/** Undo Clim→HVAC rule damage inside English "climate" (HVACATE → CLIMATE). */
export function repairClimateTypo(raw) {
  if (raw == null || raw === '') return raw;
  return String(raw).replace(/HVACATE/gi, 'CLIMATE');
}

/** Yellow L1 section title for grid / matrix (no "_" prefix, no leading "-"). */
export function formatSectionDisplayLabel(raw) {
  let t = repairClimateTypo(String(raw != null ? raw : '').trim());
  if (!t) return '';
  if (CA_BAND_EN[t]) t = CA_BAND_EN[t];
  else if (ENGLISH_LABEL_OVERRIDES[t]) t = ENGLISH_LABEL_OVERRIDES[t];
  t = t.replace(/^_+/, '').replace(/^-+/, '').trim();
  if (!t) return '';
  return translateSubsystemLabel(t);
}

/** CA chapter band value stored in BD column W (-ADAPTATION / -CABIN CLIMATE…). */
export function formatCaBandStoredLabel(raw) {
  const disp = formatSectionDisplayLabel(raw);
  if (!disp) return '';
  return `-${disp}`;
}

export function translateValue(raw) {
  if (raw == null || raw === '') return raw;
  const v = repairClimateTypo(String(raw).trim());
  if (!v) return v;
  if (ENGLISH_LABEL_OVERRIDES[v]) return ENGLISH_LABEL_OVERRIDES[v];
  if (HEADER_FR_EN[v]) return HEADER_FR_EN[v];
  if (CA_BAND_EN[v]) return CA_BAND_EN[v];
  if (L1_SECTION_EN[v]) return L1_SECTION_EN[v];
  if (L2_SECTION_EN[v]) return L2_SECTION_EN[v];
  if (CELL_VALUE_EN[v]) return CELL_VALUE_EN[v];
  if (v.startsWith('_')) {
    const key = v.toUpperCase().replace(/\s+/g, ' ');
    for (const [fr, en] of Object.entries(L2_SECTION_EN)) {
      if (fr.toUpperCase() === key || fr === v) return en;
    }
    if (L2_ENGLISH_ALIASES[key]) return L2_ENGLISH_ALIASES[key];
    const fromOverride = l2FromEnglishOverride(v.slice(1));
    if (fromOverride) return fromOverride;
    // Already a canonical uppercase L2 bookmark — skip French phrase rules (e.g. Clim→HVAC).
    if (/^_[A-Z0-9][A-Z0-9 /\-–]*$/.test(key)) return key;
    return '_' + translateFrenchPhrase(v.slice(1));
  }
  if (v === v.toUpperCase() && /^[A-Z0-9][A-Z0-9 /\-–]+$/.test(v) && v.length >= 3) {
    if (L1_SECTION_EN[v]) return L1_SECTION_EN[v];
    // Canonical English section / CA band title — never run French phrase rules.
    if (/CLIMATE|ADAPTATION|FENDERS|POWERTRAIN|ALTERNATOR/i.test(v)) return v;
    return translateFrenchPhrase(v);
  }
  return translateFrenchPhrase(v);
}
/** Memoize pure phrase translation — same label recurs across thousands of rows. */
const _frenchPhraseCache = new Map();
/** Rule-based translation for remaining French phrases. */
export function translateFrenchPhrase(text) {
  const cacheKey = repairClimateTypo(String(text));
  const cached = _frenchPhraseCache.get(cacheKey);
  if (cached !== undefined) return cached;
  let s = cacheKey;
  const rules = [
    [/Non affecté/gi, 'Unassigned'],
    [/ancien process/gi, 'old process'],
    [/Système/g, 'System'],
    [/système/g, 'system'],
    [/Arrière/gi, 'Rear'],
    [/Avant/gi, 'Front'],
    [/avant/gi, 'front'],
    [/arrière/gi, 'rear'],
    [/gauche/gi, 'left'],
    [/droite/gi, 'right'],
    [/Caisse en blanc/gi, 'Body in white'],
    [/caisse/gi, 'body'],
    [/Carrosserie/gi, 'body'],
    [/Porte(s)?/gi, (m) => (m[1] ? 'Doors' : 'Door')],
    [/Hayon/gi, 'Tailgate'],
    [/Coffre/gi, 'Trunk'],
    [/Capot/gi, 'Hood'],
    [/Bouclier/gi, 'Bumper'],
    [/Batterie/gi, 'Battery'],
    [/Volant/gi, 'Steering wheel'],
    [/Freinage/gi, 'Braking'],
    [/Garnissage/gi, 'Trim'],
    [/Console/gi, 'Console'],
    [/Essuyage/gi, 'Wipers'],
    [/Eclairage/gi, 'Lighting'],
    [/éclairage/gi, 'lighting'],
    [/Insonorisation/gi, 'Insulation'],
    [/Vitrage/gi, 'Glazing'],
    [/Tapis/gi, 'Carpet'],
    [/Planche de bord/gi, 'Instrument panel'],
    [/Pare brise/gi, 'Windshield'],
    [/Pare-brise/gi, 'Windshield'],
    [/Attelage/gi, 'Towing'],
    [/Alternateur/gi, 'Alternator'],
    [/Démarreur/gi, 'Starter'],
    [/Colonne/gi, 'Column'],
    [/Ailes/gi, 'Fenders'],
    [/Assise/gi, 'Seat'],
    [/Assises/gi, 'Seats'],
    [/Fixation/gi, 'Mounting'],
    [/Fixations/gi, 'Mountings'],
    [/Support/gi, 'Support'],
    [/Réfroidissement/gi, 'Cooling'],
    [/refroidissement/gi, 'cooling'],
    [/Carburant/gi, 'Fuel'],
    [/Additif/gi, 'Additive'],
    [/additif/gi, 'additive'],
    [/Huile/gi, 'Oil'],
    [/\bClim\b/gi, 'HVAC'],
    [/Groupe/gi, 'Group'],
    [/Circuit/gi, 'Circuit'],
    [/Réservoir/gi, 'Tank'],
    [/Commande/gi, 'Control'],
    [/Calandre/gi, 'GRILLE'],
    [/calandre/gi, 'grille'],
    [/Vitesse/gi, 'Speed'],
    [/Échappement/gi, 'Exhaust'],
    [/Echappement/gi, 'Exhaust'],
    [/Pédalier/gi, 'Pedal'],
    [/Façade/gi, 'Facade'],
    [/Facade/gi, 'Facade'],
    [/Circuit eau/gi, 'Water circuit'],
    [/circuit eau/gi, 'water circuit'],
    [/Ecrans thermique/gi, 'Thermal screens'],
    [/écrans thermique/gi, 'thermal screens'],
    [/Liquide additif/gi, 'Additive liquid'],
    [/liquide additif/gi, 'additive liquid'],
    [/Liquide refroidissement/gi, 'Cooling liquid'],
    [/liquide refroidissement/gi, 'cooling liquid'],
    [/Supports GMP/gi, 'Engine mounts'],
    [/supports GMP/gi, 'engine mounts'],
    [/Écran/gi, 'Screen'],
    [/Ecran/gi, 'Screen'],
    [/aérodynamique/gi, 'aerodynamic'],
    [/Aéro/gi, 'Aero'],
    [/Divers/gi, 'Misc'],
    [/Equipement/gi, 'Equipment'],
    [/Équipement/gi, 'Equipment'],
    [/Tolerie/gi, 'Sheet metal'],
    [/Habitacle/gi, 'Cabin'],
    [/Pavillon/gi, 'Roof'],
    [/Câblage/gi, 'Wiring'],
    [/puissance/gi, 'power'],
    [/utilisé(e)?/gi, 'used'],
    [/non utilisé/gi, 'unused'],
    [/avec peinture/gi, 'with paint'],
    [/barre de direction/gi, 'steering column'],
    [/téléphone/gi, 'telephone'],
    [/rétroviseur/gi, 'mirror'],
    [/lavage/gi, 'wash'],
    [/modulaire/gi, 'modular'],
    [/technique/gi, 'technical'],
    [/thermique/gi, 'thermal'],
    [/additionnel/gi, 'additional'],
    [/Supports/gi, 'Supports'],
    [/par /gi, ' / '],
    [/Module de réparation/gi, 'Repair module'],
    [/latéraux/gi, 'lateral'],
    [/latérales/gi, 'lateral'],
    [/latéral/gi, 'lateral'],
    [/coté/gi, 'side'],
    [/côté/gi, 'side'],
    [/passager/gi, 'passenger'],
    [/conducteur/gi, 'driver'],
    [/piéton/gi, 'pedestrian'],
    [/nocturne/gi, 'night'],
    [/sonore/gi, 'sound'],
    [/caméra/gi, 'camera'],
    [/réparation/gi, 'repair'],
    [/électrique/gi, 'electric'],
    [/centrale/gi, 'central'],
    [/unité/gi, 'unit'],
    [/Enjoliveur/gi, 'Molding'],
    [/enjoliveur/gi, 'molding'],
    [/vélo/gi, 'bicycle'],
    [/crochet/gi, 'hook'],
    [/Garnitures/gi, 'Trim'],
    [/garnitures/gi, 'trim'],
    [/supérieure/gi, 'upper'],
    [/intermédiaire/gi, 'intermediate'],
    [/Déflecteurs/gi, 'Deflectors'],
    [/déflecteurs/gi, 'deflectors'],
    [/assistée/gi, 'assisted'],
    [/Liquide de/gi, 'Fluid'],
    [/direction/gi, 'steering'],
    [/Deuxième/gi, 'Second'],
    [/Deuxiéme/gi, 'Second'],
    [/calculateur/gi, 'controller'],
    [/gestion/gi, 'management'],
    [/alerte/gi, 'alert'],
    [/Custodes/gi, 'custodes'],
    [/Rangement/gi, 'Storage'],
    [/accoudoir/gi, 'armrest'],
    [/Aérateur/gi, 'Vent'],
    [/aérateur/gi, 'vent'],
    [/Bac de/gi, 'Tray'],
    [/rangemen/gi, 'storage'],
    [/toit ouvrant/gi, 'sunroof'],
    [/Tablier/gi, 'Dashboard panel'],
    [/unique/gi, 'single'],
    [/Partie/gi, 'Part'],
    [/d'unité/gi, 'unit'],
    [/d'/gi, ' '],
    [/é/g, 'e'],
    [/è/g, 'e'],
    [/ê/g, 'e'],
    [/ë/g, 'e'],
    [/à/g, 'a'],
    [/â/g, 'a'],
    [/ù/g, 'u'],
    [/û/g, 'u'],
    [/ô/g, 'o'],
    [/î/g, 'i'],
    [/ï/g, 'i'],
    [/ç/g, 'c'],
  ];
  for (const [re, rep] of rules) {
    s = s.replace(re, rep);
  }
  s = s.replace(/\s{2,}/g, ' ').trim();
  _frenchPhraseCache.set(cacheKey, s);
  return s;
}
/** Sub-system L1 / L2 / Design Dpt — always surface English in the grid. */
export function translateSubsystemLabel(raw) {
  if (raw == null || raw === '') return raw;
  // IMPORTANT: Bookmark breakdown (sections + sub-sections) must be identical
  // across BD, Synthesis and Bookmark Matrix. Business convention: ALL CAPS.
  return translateValue(String(raw).trim()).toUpperCase();
}

/**
 * Normalized key for SYN F ↔ BD L2 matching (Excel AS = AU on Database page).
 * Applies translation twice (Syn French → EN, BD title case) then uppercases.
 */
export function canonicalL2MatchKey(raw) {
  let s = translateSubsystemLabel(String(raw != null ? raw : '').trim());
  if (!s) return '';
  s = translateSubsystemLabel(s);
  return s.toUpperCase();
}
