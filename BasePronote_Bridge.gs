/**
 * ===================================================================
 * 🌉 BASE PRONOTE — PONT IMPORT → INTERFACE V2 (VERSION SIMPLIFIÉE)
 * ===================================================================
 *
 * Ce fichier relie l'import Pronote (onglets sources "5°1", "5°2"…) à
 * InterfaceV2, sans le moteur d'optimisation automatique.
 *
 * Rôle :
 *  - "Base Pronote"  : prépare des onglets de travail "<classe> TEST"
 *                      à partir des onglets importés (l'import reste INTACT).
 *  - _STRUCTURE auto : génère les règles (capacité / options) si absentes,
 *                      pour que les "déplacements sous conditions" fonctionnent.
 *  - Sauvegardes      : écriture COHÉRENTE et SÛRE des dispositions vers des
 *                      onglets de travail (jamais sur l'onglet source importé).
 *  - Matrice modèle   : un onglet type qui simule le fichier CSV Pronote attendu.
 *
 * Dépendances (toutes dans Code.gs, conservé) :
 *  getActiveSpreadsheetCached, collectClassesDataByMode, SHEET_PATTERNS, DEFAULTS.
 */

// ==================== CONSTANTES ====================

/** Nom de l'onglet matrice / modèle CSV Pronote */
const BP_MATRICE_SHEET = 'MODELE_PRONOTE';

/** Capacité par défaut d'une classe destination (modifiable dans _STRUCTURE) */
const BP_DEFAULT_CAPACITY = 28;

/** En-têtes attendus côté import (1ʳᵉ colonne = identifiant élève) */
const BP_ID_HEADER = 'ID_ELEVE';

// ==================== POINT D'ENTRÉE : BASE PRONOTE ====================

/**
 * Prépare la base de travail depuis l'import Pronote.
 * Pour chaque onglet source ("5°1"…), crée un onglet de travail "5°1 TEST"
 * (uniquement s'il n'existe pas déjà → ne détruit jamais un travail en cours).
 * Garantit aussi la présence de _STRUCTURE (règles).
 *
 * Appelé depuis la carte "Base Pronote" de la modale de démarrage d'InterfaceV2.
 * @returns {{success:boolean, created?:string[], existing?:string[], count?:number, error?:string}}
 */
function preparerBasePronote() {
  try {
    const ss = getActiveSpreadsheetCached();
    const sources = bp_findSourceClassSheets_();

    if (!sources.length) {
      return {
        success: false,
        error: 'Aucun onglet de classe importé (ex : "5°1"). Importez d\'abord vos données depuis Pronote.'
      };
    }

    const created = [];
    const existing = [];
    const sizeByClasse = {};

    sources.forEach(sheet => {
      const name = sheet.getName();
      sizeByClasse[name] = Math.max(0, sheet.getLastRow() - 1);

      const testName = name + ' TEST';
      if (ss.getSheetByName(testName)) {
        existing.push(testName); // travail déjà en cours → on préserve
        return;
      }
      const copy = sheet.copyTo(ss);
      copy.setName(testName);
      copy.showSheet();
      created.push(testName);
    });

    // Règles de répartition (capacité / options) pour les "déplacements sous conditions"
    ensureStructureFromClasses_(sources.map(s => s.getName()), sizeByClasse);

    SpreadsheetApp.flush();

    return {
      success: true,
      created,
      existing,
      count: created.length + existing.length
    };
  } catch (e) {
    console.log('❌ preparerBasePronote: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Réinitialise la base de travail depuis l'import (force la recréation des
 * onglets "<classe> TEST"). ⚠️ Écrase les modifications manuelles en cours.
 * @returns {Object} Résultat
 */
function reinitialiserBasePronote() {
  try {
    const ss = getActiveSpreadsheetCached();
    const sources = bp_findSourceClassSheets_();
    if (!sources.length) {
      return { success: false, error: 'Aucun onglet de classe importé trouvé.' };
    }
    sources.forEach(sheet => {
      const testName = sheet.getName() + ' TEST';
      const old = ss.getSheetByName(testName);
      if (old) ss.deleteSheet(old);
      const copy = sheet.copyTo(ss);
      copy.setName(testName);
      copy.showSheet();
    });
    SpreadsheetApp.flush();
    return { success: true, count: sources.length };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==================== _STRUCTURE AUTOMATIQUE ====================

/**
 * Crée l'onglet _STRUCTURE (règles) à partir des classes détectées, s'il
 * n'existe pas déjà. Schéma compatible avec loadStructureRules() :
 *   CLASSE_DEST | EFFECTIF | OPTIONS
 * @param {string[]} classeNames - Noms des classes (ex : ["5°1","5°2"])
 * @param {Object} sizeByClasse - Effectif courant par classe (optionnel)
 * @returns {boolean} true si créé, false si déjà présent
 */
function ensureStructureFromClasses_(classeNames, sizeByClasse) {
  const ss = getActiveSpreadsheetCached();
  if (ss.getSheetByName('_STRUCTURE')) return false; // ne jamais écraser une config existante

  const sheet = ss.insertSheet('_STRUCTURE');
  const headers = ['CLASSE_DEST', 'EFFECTIF', 'OPTIONS'];
  const rows = [headers];

  (classeNames || []).forEach(name => {
    const current = (sizeByClasse && sizeByClasse[name]) || 0;
    const capacity = Math.max(BP_DEFAULT_CAPACITY, current);
    // OPTIONS laissé vide par défaut : la capacité suffit aux "conditions".
    // L'utilisateur peut ajouter des quotas (ex : "LATIN=10,CHANT=8").
    rows.push([name, capacity, '']);
  });

  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#d9e1f2');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  return true;
}

// ==================== SAUVEGARDE COHÉRENTE & SÛRE ====================

/**
 * Construit une carte { ID_ELEVE -> ligne complète } à partir des onglets de
 * travail existants (TEST, CACHE, FIN puis SOURCE en repli). Les attributs
 * d'un élève ne changent pas lors d'un déplacement : seule sa classe (= onglet)
 * change. On peut donc reconstituer chaque ligne où qu'elle se trouve.
 * @returns {{headers:Array, map:Object}}
 * @private
 */
function bp_buildIdRowMap_() {
  const modes = ['TEST', 'CACHE', 'FIN', 'SOURCE'];
  let headers = null;
  const map = {};

  modes.forEach(mode => {
    const classes = collectClassesDataByMode(mode); // {sheetName:{headers,students}}
    Object.keys(classes).forEach(sn => {
      const entry = classes[sn];
      if (!entry || !entry.headers) return;
      if (!headers) headers = entry.headers;

      let idIdx = entry.headers.findIndex(h => String(h).trim().toUpperCase() === BP_ID_HEADER);
      if (idIdx < 0) idIdx = 0;

      (entry.students || []).forEach(row => {
        const id = String(row[idIdx] || '').trim();
        if (id && !map[id]) map[id] = row; // 1ʳᵉ occurrence (priorité TEST) gagne
      });
    });
  });

  return { headers, map };
}

/** Suffixe d'onglet de travail selon le mode (jamais l'onglet source brut). */
function bp_suffixForMode_(mode) {
  const m = String(mode || '').trim().toUpperCase();
  if (m === 'FIN') return 'FIN';
  if (m === 'TEST') return 'TEST';
  return 'CACHE'; // SOURCE / SOURCES / CACHE / inconnu → CACHE (protège l'import)
}

/** Nom d'onglet de travail cible pour une classe normalisée + un suffixe. */
function bp_resolveWriteName_(classe, suffix) {
  const base = String(classe || '').replace(/\s*(TEST|FIN|CACHE|PREVIOUS)\s*$/i, '').trim();
  return base + ' ' + suffix;
}

/**
 * Écrit une disposition vers des onglets de travail, de façon cohérente et sûre.
 * Accepte les deux formes de disposition :
 *   { "5°1": ["id1","id2"] }                  (export léger d'InterfaceV2)
 *   { "5°1": { headers:[…], students:[[…]] } } (export complet)
 * Ne touche JAMAIS l'onglet source importé ("5°1") : écrit "5°1 TEST"/"5°1 CACHE"…
 *
 * @param {Object} disposition - Disposition par classe
 * @param {string} suffix - Suffixe cible ("TEST" | "CACHE" | "FIN")
 * @returns {Object} {success, saved, failed, errors, timestamp}
 * @private
 */
function bp_writeDisposition_(disposition, suffix) {
  const ss = getActiveSpreadsheetCached();
  const built = bp_buildIdRowMap_();
  const headers = built.headers;

  if (!disposition || typeof disposition !== 'object' || !Object.keys(disposition).length) {
    return { success: false, error: 'Disposition vide' };
  }
  if (!headers) {
    return { success: false, error: 'Aucune donnée de travail trouvée (préparez d\'abord la Base Pronote).' };
  }

  const width = headers.length;
  let saved = 0, failed = 0;
  const errors = [];

  Object.keys(disposition).forEach(classe => {
    try {
      let value = disposition[classe];
      let rows;

      if (Array.isArray(value)) {
        // Tableau d'IDs → reconstituer les lignes depuis la carte
        rows = value.map(id => built.map[String(id).trim()]).filter(Boolean);
      } else if (value && Array.isArray(value.students)) {
        // Forme complète : lignes déjà fournies
        rows = value.students.filter(Boolean);
      } else {
        rows = [];
      }

      const sheetName = bp_resolveWriteName_(classe, suffix);
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) sheet = ss.insertSheet(sheetName);
      sheet.clearContents();

      const normalized = [headers].concat(rows).map(r => {
        const rr = (Array.isArray(r) ? r.slice(0, width) : []);
        while (rr.length < width) rr.push('');
        return rr;
      });

      sheet.getRange(1, 1, normalized.length, width).setValues(normalized);
      saved++;
    } catch (e) {
      failed++;
      errors.push(classe + ': ' + e.message);
    }
  });

  SpreadsheetApp.flush();
  return {
    success: failed === 0,
    saved,
    failed,
    errors: errors.length ? errors : undefined,
    timestamp: new Date().toISOString()
  };
}

// ==================== MATRICE / MODÈLE CSV PRONOTE ====================

/**
 * Crée (ou recrée) l'onglet matrice qui simule le fichier CSV attendu de Pronote.
 * Sert de gabarit visuel : l'utilisateur voit les colonnes et un exemple de
 * remplissage avant d'importer.
 * @returns {Object} {success, created}
 */
function creerMatriceModele() {
  try {
    const ss = getActiveSpreadsheetCached();
    let sheet = ss.getSheetByName(BP_MATRICE_SHEET);
    if (sheet) ss.deleteSheet(sheet);
    sheet = ss.insertSheet(BP_MATRICE_SHEET, 0); // en première position

    const titre = '📋 MODÈLE — Export "Liste des élèves" de Pronote (à coller dans l\'Assistant Import)';
    const headers = ['NOM', 'PRENOM', 'CLASSE', 'SEXE', 'LV2', 'TOUTES_LES_OPTIONS', 'DISPOSITIF'];
    const exemples = [
      ['DURAND', 'Léa', '5°1', 'F', 'ESPAGNOL', 'LATIN', ''],
      ['MARTIN', 'Lucas', '5°1', 'M', 'ALLEMAND', '', ''],
      ['BERNARD', 'Emma', '5°2', 'F', 'ESPAGNOL', 'CHANT', 'PAP'],
      ['PETIT', 'Noah', '5°2', 'M', 'ITALIEN', 'LATIN', ''],
      ['ROBERT', 'Jade', '5°3', 'F', 'ALLEMAND', 'CHANT, LATIN', 'PPRE']
    ];

    // Ligne 1 : titre fusionné
    sheet.getRange(1, 1, 1, headers.length).merge()
      .setValue(titre)
      .setFontWeight('bold').setFontSize(11)
      .setBackground('#1f6feb').setFontColor('#ffffff')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    sheet.setRowHeight(1, 30);

    // Ligne 2 : en-têtes
    sheet.getRange(2, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#d9e1f2');

    // Lignes d'exemple
    sheet.getRange(3, 1, exemples.length, headers.length).setValues(exemples);

    // Mise en forme
    sheet.setFrozenRows(2);
    sheet.autoResizeColumns(1, headers.length);
    sheet.getRange(2, 1, 1, headers.length).getCell(1, 3) // note sur CLASSE
      .setNote('Format attendu : NIVEAU°NUMÉRO (ex : 5°1). "4E 1" est aussi accepté et converti automatiquement.');
    sheet.getRange(2, 4).setNote('F ou M.');
    sheet.getRange(2, 6).setNote('Plusieurs options séparées par une virgule (ex : "CHANT, LATIN").');

    SpreadsheetApp.flush();
    return { success: true, created: BP_MATRICE_SHEET };
  } catch (e) {
    console.log('❌ creerMatriceModele: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Crée la matrice modèle au démarrage UNIQUEMENT si le classeur est "vierge"
 * (aucun onglet de classe importé et pas déjà de matrice). Appelé depuis onOpen
 * en best-effort (échoue silencieusement si non autorisé).
 * @private
 */
function ensureMatriceModele_() {
  try {
    const ss = getActiveSpreadsheetCached();
    if (ss.getSheetByName(BP_MATRICE_SHEET)) return;        // déjà présente
    if (bp_findSourceClassSheets_().length > 0) return;     // classeur déjà utilisé
    creerMatriceModele();
  } catch (e) {
    // best-effort : on ignore (droits onOpen limités selon le contexte)
  }
}

// ==================== UTILITAIRES ====================

/**
 * Retourne les onglets "classe source" issus de l'import (ex : "5°1"),
 * en excluant les onglets de travail (TEST/CACHE/FIN/PREVIOUS) et système (_…).
 * @returns {Sheet[]}
 * @private
 */
function bp_findSourceClassSheets_() {
  const ss = getActiveSpreadsheetCached();
  return ss.getSheets().filter(sheet => {
    const name = sheet.getName();
    if (name.indexOf('_') === 0) return false;                       // système
    if (/(TEST|CACHE|FIN|PREVIOUS)$/i.test(name)) return false;      // onglet de travail
    return SHEET_PATTERNS.SOURCE.test(name);                         // termine par °chiffre
  });
}

// ==================== COMPAT : HELPERS PRÉSERVÉS ====================
// Ces fonctions provenaient de fichiers retirés (moteur d'optimisation) mais
// restent utilisées par le chemin d'import. Rapatriées ici pour le conserver.

/**
 * Logger uniforme (anciennement dans App.Core.gs). Utilisé par RunAudit_log().
 * @param {string} level - INFO / WARN / ERROR
 * @param {string} msg - message
 */
function logLine(level, msg) {
  try {
    const stamp = new Date().toLocaleString('fr-FR');
    Logger.log(stamp + ' ' + String(level || 'INFO').padEnd(7) + ' ' + msg);
  } catch (e) {
    Logger.log(level + ' ' + msg);
  }
}

// ==================== STUBS — FONCTIONNALITÉS RETIRÉES ====================
// Les modules Optimisation automatique / Analytics / Contraintes ont été retirés
// (version simplifiée = déplacements MANUELS sous conditions). Leurs boutons
// déclencheurs ont été supprimés, donc ces fonctions sont normalement
// inatteignables. Ces stubs garantissent une dégradation propre si un appel
// résiduel survient (au lieu d'une erreur "fonction introuvable").

const BP_FEATURE_REMOVED_MSG =
  'Fonctionnalité retirée dans la version simplifiée. Utilisez les déplacements manuels.';

/** @returns {Object} snapshot analytique vide */
function getAnalyticsSnapshotForUI() {
  return { timestamp: null, disabled: true, message: BP_FEATURE_REMOVED_MSG };
}

/** @returns {Array} historique analytique vide */
function getAnalyticsHistoryForUI() {
  return [];
}

/** Optimisation automatique retirée. */
function lancerOptimisationClaudeMotor(options) {
  return { success: false, error: BP_FEATURE_REMOVED_MSG };
}

/** Panneau de contraintes : retourne une structure vide (ouverture propre). */
function chargerContraintes() {
  return { success: true, constraints: {}, message: BP_FEATURE_REMOVED_MSG };
}
