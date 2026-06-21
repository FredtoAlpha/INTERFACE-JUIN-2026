# Outil de répartition — version simplifiée

Version allégée centrée sur l'essentiel : **importer un CSV Pronote → mettre en forme → répartir les élèves à la main dans InterfaceV2** (déplacements sous conditions, graphiques, sauvegardes automatiques, création d'onglets).

Le moteur d'optimisation automatique, la console d'administration et le module de groupes ont été retirés. La répartition se fait **manuellement**, avec garde-fous (capacités, options, ASSO/DISSO).

## Le menu « 🏫 Répartition »

À l'ouverture du classeur, un menu apparaît :

| Entrée | Rôle |
|---|---|
| 📋 Créer l'onglet modèle (CSV Pronote) | Génère l'onglet `MODELE_PRONOTE` : un exemple de la liste d'élèves attendue (colonnes + données fictives). |
| 📥 Importer depuis Pronote | Ouvre l'assistant d'import (élèves, notes, absences, vie scolaire) → crée les onglets de classe et calcule les scores COM/TRA/PART/ABS. |
| ⚙️ Configurer les classes | Ouvre `_STRUCTURE` (capacité max + quotas d'options par classe). Ces règles pilotent les déplacements sous conditions. |
| 🎯 Ouvrir l'interface de répartition | Lance InterfaceV2. |
| ➕ Intégrer un nouvel élève | Ajoute un élève à une classe existante. |
| 🔁 Réinitialiser la base depuis l'import | Recrée les onglets de travail `… TEST` depuis l'import (écrase les modifications en cours). |

## Déroulé type

1. **Créer le modèle** (facultatif) pour voir le format attendu.
2. **Importer depuis Pronote** : colle les exports Pronote dans l'assistant. Cela crée les onglets de classe (`5°1`, `5°2`, …) mis en forme.
3. **Configurer les classes** (facultatif) : ajuste capacités/quotas dans `_STRUCTURE` (généré automatiquement sinon).
4. **Ouvrir l'interface de répartition**, puis cliquer sur la carte **« Base Pronote »**.

## Modes de l'interface (carte de démarrage)

- **Base Pronote** 🔵 — point de départ : prépare des onglets de travail `<classe> TEST` à partir de l'import (l'import d'origine reste **intact**), puis ouvre la répartition.
- **Base préparée (TEST)** 🟢 — reprend les onglets de travail déjà préparés.
- **Brouillon (CACHE)** 🟡 — reprend la dernière auto-sauvegarde.

## Comment les données sont protégées

- Les onglets importés (`5°1`, …) ne sont **jamais** modifiés par l'interface.
- L'édition se fait sur les onglets `… TEST`.
- L'auto-sauvegarde écrit dans le navigateur **et** dans des onglets `… CACHE`.
- « Réinitialiser » permet de repartir de l'import à tout moment.

## Architecture (fichiers clés)

- `Code.gs` — menu, point d'entrée web (`doGet`), backend d'InterfaceV2.
- `BasePronote_Bridge.gs` — pont import → interface : `preparerBasePronote()`, `_STRUCTURE` auto, sauvegardes cohérentes, matrice modèle.
- `ImportAssistant_Server.gs` + `Backend_ImportDB.gs` — import Pronote et calcul des scores.
- `Consolidation.gs`, `ListesDeroulantes.gs`, `GenereNOMprenomID.gs` — mise en forme.
- `Structure.gs`, `Config.gs` — règles et listes déroulantes.
- `InterfaceV2*.html` — interface de répartition (déplacements, graphiques, sauvegardes, onglets).

> Sauvegarde de l'ancienne version complète : tag git `sauvegarde-avant-simplification`.
