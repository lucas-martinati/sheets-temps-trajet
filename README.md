# 🚗 Sheets Temps de Trajet

Extension **Google Apps Script** pour Google Sheets qui calcule automatiquement les
**temps de trajet** entre une adresse de départ et une liste d'adresses, directement
dans une feuille de calcul. Pratique pour comparer des logements, des offres d'emploi,
des lieux de rendez-vous, etc.

Le calcul tourne **en arrière-plan** : on peut fermer la fenêtre, il continue tout seul.

## ✨ Fonctionnalités

- Menu personnalisé dans Google Sheets (**🚗 Outils Candidatures**).
- Choix du **point de départ** parmi une liste d'adresses enregistrées.
- Choix du **mode de transport** : 🚗 voiture · 🚶 à pied · 🚆 train · 🚌 bus.
- Choix libre des **colonnes** utilisées (colonne de l'adresse, colonne du résultat).
- **Traitement en arrière-plan** via un déclencheur temporel : la fermeture de la
  fenêtre n'interrompt pas le calcul ; on peut rouvrir le menu pour revoir la progression.
- Barre de progression, journal des dernières lignes traitées, et bouton **Arrêter**.
- Gestion propre des erreurs Google Maps (quota journalier, limite momentanée,
  clé invalide, adresse introuvable…).

## 🔧 Installation

1. Ouvrir le Google Sheet cible → menu **Extensions → Apps Script**.
2. Coller le contenu de [`Code.gs`](./Code.gs) dans l'éditeur, puis **Enregistrer**.
3. Recharger le Google Sheet : un menu **🚗 Outils Candidatures** apparaît.

## 🔑 Configuration de la clé API Google Maps (obligatoire)

Le script utilise la **Directions API** avec **ta propre clé** (quota personnel, et non
le quota partagé quasi nul du service Maps intégré à Apps Script).

1. Aller sur [console.cloud.google.com](https://console.cloud.google.com) → créer/choisir un projet.
2. **APIs & Services → Bibliothèque** → activer **« Directions API »**.
3. **Facturation** : associer une carte.
   ⚠️ Obligatoire, mais Google offre **200 $ de crédit gratuit/mois** (~40 000 trajets) —
   aucun débit en usage normal.
4. **APIs & Services → Identifiants → Créer des identifiants → Clé API**.
5. (Recommandé) Restreindre la clé à la **Directions API**.
6. Dans le Sheet : menu **🚗 Outils Candidatures → 🔑 Configurer la clé API**, coller la clé.

La clé est stockée dans les *Script Properties* du projet — **jamais** en dur dans le code
ni visible dans la feuille.

## 📋 Utilisation

1. Préparer l'onglet cible avec au moins une colonne d'**adresses de destination**
   et une colonne (vide) qui recevra le **résultat**.
2. Menu **🚗 Outils Candidatures → Calculer les temps de trajet**.
3. Dans la fenêtre : choisir le point de départ, le mode de transport, la colonne
   de l'adresse et la colonne du résultat, puis **Lancer le calcul**.
4. La progression s'affiche. On peut fermer la fenêtre : le calcul continue.

Le premier onglet **« Adresses de départ »** est créé automatiquement au premier
lancement : y renseigner les points de départ possibles (nom + adresse exacte).

### Notes

- Seules les lignes dont la colonne résultat est **vide** ou contient
  « Erreur » / « Introuvable » sont (re)calculées. Pour recalculer, vider la cellule
  ou utiliser une autre colonne de résultat (ex. une pour la voiture, une pour le train).
- Les modes **train / bus** dépendent des horaires réels : une adresse sans desserte
  peut renvoyer « Introuvable » même si l'adresse est correcte.

## 🗂️ Structure

| Fichier | Rôle |
|---|---|
| `Code.gs` | Tout le script (serveur Apps Script + interface HTML de la fenêtre). |

## ⚙️ Détails techniques

- **Arrière-plan** : `demarrerCalculFond` prépare un « job » (stocké dans les
  *Document Properties*) ; `executerLotFond`, lancé par déclencheur temporel et/ou
  depuis la fenêtre, traite les lignes par tranches de ~5 min et se re-planifie tant
  qu'il reste du travail. Un `LockService` garantit qu'un seul lot tourne à la fois.
- **Progression** : la fenêtre interroge `getEtatProgression` toutes les 1,5 s.
- **Directions API** : appelée via `UrlFetchApp` ; le mode (`driving` / `walking` /
  `transit` + `transit_mode=train|bus`) est passé en paramètre.
