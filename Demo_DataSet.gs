/**
 * ===================================================================
 * 🎲 JEU DE DONNÉES DE DÉMONSTRATION (150 élèves fictifs)
 * ===================================================================
 *
 * Génère des classes fictives entièrement remplies pour tester / démontrer
 * l'interface sans passer par un import Pronote réel.
 *
 * Produit des onglets "5°1".."5°5" (~30 élèves chacun) au schéma exact attendu
 * par l'interface, avec de la variété pour exercer TOUTES les fonctionnalités :
 *  - scores COM/TRA/PART/ABS (graphiques) ;
 *  - SEXE équilibré (parité) ;
 *  - LV2 / OPTIONS (répartition et quotas) ;
 *  - MOBILITE (FIXE = bloqué, LIBRE, PERMUT) → déplacements sous conditions ;
 *  - quelques codes ASSO (à garder ensemble) et DISSO (à séparer).
 *
 * Dépendances : getActiveSpreadsheetCached (Code.gs), ensureStructureFromClasses_
 * non utilisé ici (on écrit _STRUCTURE avec une marge de capacité pour permettre
 * les déplacements de démonstration).
 */

const DEMO_CLASSES = ['5°1', '5°2', '5°3', '5°4', '5°5'];
const DEMO_HEADERS = ['ID_ELEVE', 'NOM', 'PRENOM', 'NOM_PRENOM', 'SEXE', 'LV2', 'OPT',
  'COM', 'TRA', 'PART', 'ABS', 'DISPO', 'ASSO', 'DISSO', 'MOBILITE', 'SOURCE'];

const DEMO_NOMS = ['MARTIN', 'BERNARD', 'DUBOIS', 'THOMAS', 'ROBERT', 'PETIT', 'DURAND', 'LEROY',
  'MOREAU', 'SIMON', 'LAURENT', 'LEFEBVRE', 'MICHEL', 'GARCIA', 'DAVID', 'BERTRAND', 'ROUX',
  'VINCENT', 'FOURNIER', 'MOREL', 'GIRARD', 'ANDRE', 'MERCIER', 'DUPONT', 'LAMBERT', 'BONNET',
  'FRANCOIS', 'LEGRAND', 'GARNIER', 'FAURE', 'ROUSSEAU', 'BLANC', 'GUERIN', 'HENRY', 'ROUSSEL',
  'NICOLAS', 'PERRIN', 'MORIN', 'MATHIEU', 'CLEMENT', 'GAUTHIER', 'DUMONT', 'FONTAINE',
  'CHEVALIER', 'ROBIN', 'MASSON', 'MARCHAND', 'DUVAL', 'DENIS', 'LEMAIRE'];

const DEMO_PRENOMS_F = ['Emma', 'Jade', 'Louise', 'Alice', 'Chloé', 'Lina', 'Léa', 'Rose', 'Anna',
  'Mila', 'Inès', 'Ambre', 'Julia', 'Manon', 'Zoé', 'Eva', 'Léna', 'Camille', 'Sarah', 'Nina',
  'Romane', 'Juliette', 'Clara', 'Agathe', 'Lou'];

const DEMO_PRENOMS_M = ['Gabriel', 'Léo', 'Raphaël', 'Arthur', 'Louis', 'Lucas', 'Adam', 'Jules',
  'Hugo', 'Maël', 'Noah', 'Liam', 'Ethan', 'Nathan', 'Tom', 'Sacha', 'Aaron', 'Naël', 'Gabin',
  'Timéo', 'Mathis', 'Eden', 'Paul', 'Marius', 'Victor'];

/** Pioche aléatoire dans un tableau. */
function demo_pick_(arr) {
  return arr[Math.floor(Math.random() * arr.length)];
}

/** Score 1..5 avec une distribution centrée (un peu plus de 3/4). */
function demo_score_() {
  const r = Math.random();
  if (r < 0.10) return 1;
  if (r < 0.30) return 2;
  if (r < 0.60) return 3;
  if (r < 0.85) return 4;
  return 5;
}

/**
 * Génère le jeu de démonstration et écrit les onglets de classe.
 * @param {number} [nbEleves=150] - Nombre total d'élèves
 * @returns {Object} {success, total, classes}
 */
function genererJeuDeDonneesDemo(nbEleves) {
  try {
    const ss = getActiveSpreadsheetCached();
    const total = nbEleves || 150;
    const perClass = Math.ceil(total / DEMO_CLASSES.length);

    const lv2Pool = ['ESP', 'ESP', 'ESP', 'ALL', 'ALL', 'ITA'];          // pondéré
    const optPool = ['', '', '', '', '', 'LATIN', 'LATIN', 'CHANT', 'GREC'];
    const dispoPool = ['', '', '', '', '', '', '', '', 'PAP', 'PPRE'];
    const mobPool = ['LIBRE', 'LIBRE', 'LIBRE', 'LIBRE', 'LIBRE', 'LIBRE', 'LIBRE', 'FIXE', 'FIXE', 'PERMUT'];

    const byClass = {};
    DEMO_CLASSES.forEach(c => { byClass[c] = []; });

    for (let i = 0; i < total; i++) {
      const classe = DEMO_CLASSES[Math.min(Math.floor(i / perClass), DEMO_CLASSES.length - 1)];
      const sexe = Math.random() < 0.5 ? 'F' : 'M';
      const prenom = sexe === 'F' ? demo_pick_(DEMO_PRENOMS_F) : demo_pick_(DEMO_PRENOMS_M);
      const nom = demo_pick_(DEMO_NOMS);
      const id = 'E' + String(i + 1).padStart(3, '0');

      byClass[classe].push([
        id, nom, prenom, nom + ' ' + prenom, sexe,
        demo_pick_(lv2Pool), demo_pick_(optPool),
        demo_score_(), demo_score_(), demo_score_(), demo_score_(),
        demo_pick_(dispoPool), '', '', demo_pick_(mobPool), classe
      ]);
    }

    // Démonstration des contraintes ASSO / DISSO (colonnes 13 et 14, index 12/13)
    if (byClass['5°1'].length >= 2) {
      byClass['5°1'][0][12] = 'AMIS1';   // ASSO : ces deux élèves doivent rester ensemble
      byClass['5°1'][1][12] = 'AMIS1';
    }
    if (byClass['5°2'].length >= 2) {
      byClass['5°2'][0][13] = 'SEP1';     // DISSO : ces deux élèves ne doivent pas être ensemble
      byClass['5°2'][1][13] = 'SEP1';
    }

    // Écriture des onglets de classe (remplace s'ils existent)
    DEMO_CLASSES.forEach(classe => {
      const old = ss.getSheetByName(classe);
      if (old) ss.deleteSheet(old);
      const sheet = ss.insertSheet(classe);
      const rows = byClass[classe];
      const data = [DEMO_HEADERS].concat(rows);
      sheet.getRange(1, 1, data.length, DEMO_HEADERS.length).setValues(data);
      sheet.getRange(1, 1, 1, DEMO_HEADERS.length).setFontWeight('bold').setBackground('#d9e1f2');
      sheet.setFrozenRows(1);
      sheet.autoResizeColumns(1, DEMO_HEADERS.length);
    });

    // _STRUCTURE avec une marge de capacité (+2) pour autoriser des déplacements
    demo_ecrireStructure_(byClass);

    SpreadsheetApp.flush();
    return { success: true, total, classes: DEMO_CLASSES.length };
  } catch (e) {
    console.log('❌ genererJeuDeDonneesDemo: ' + e.message);
    return { success: false, error: e.message };
  }
}

/** Écrit un _STRUCTURE de démo (capacité = effectif + 2, pour laisser de la place). */
function demo_ecrireStructure_(byClass) {
  const ss = getActiveSpreadsheetCached();
  const old = ss.getSheetByName('_STRUCTURE');
  if (old) ss.deleteSheet(old);
  const sheet = ss.insertSheet('_STRUCTURE');
  const headers = ['CLASSE_DEST', 'EFFECTIF', 'OPTIONS'];
  const rows = [headers];
  DEMO_CLASSES.forEach(classe => {
    const capacity = (byClass[classe] ? byClass[classe].length : 28) + 2;
    rows.push([classe, capacity, '']); // OPTIONS vide = capacité seule (modifiable)
  });
  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#d9e1f2');
  sheet.setFrozenRows(1);
}

/**
 * Lanceur de menu (avec confirmation) : génère le jeu de démo puis prépare
 * directement la base de travail pour ouvrir l'interface aussitôt.
 */
function genererJeuDeDonneesDemoUI() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    '🎲 Générer un jeu de démonstration',
    'Crée 5 classes fictives (5°1 à 5°5, 150 élèves) entièrement remplies pour tester l\'interface.\n\n' +
    '⚠️ Les onglets 5°1 à 5°5 et _STRUCTURE seront REMPLACÉS s\'ils existent.\n\n' +
    'Continuer ?',
    ui.ButtonSet.OK_CANCEL);
  if (resp !== ui.Button.OK) return;

  const result = genererJeuDeDonneesDemo(150);
  if (!result.success) {
    ui.alert('❌ ' + (result.error || 'Échec de la génération.'));
    return;
  }

  // Prépare aussitôt les onglets de travail "… TEST"
  if (typeof preparerBasePronote === 'function') preparerBasePronote();

  ui.alert(
    '✅ Jeu de démo créé',
    result.total + ' élèves répartis sur ' + result.classes + ' classes.\n\n' +
    'Ouvrez maintenant « 🎯 Ouvrir l\'interface de répartition » → carte « Base Pronote » (ou « Base préparée »).',
    ui.ButtonSet.OK);
}
