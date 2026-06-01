# Recherche d'occurrences dans Excel

Ce programme permet de rechercher des occurrences dans un fichier Excel et de créer une liste de liens cliquables dans un nouvel onglet.

## Prérequis

- Python 3.6+
- Bibliothèque `openpyxl`

## Installation

```bash
pip install -r requirements.txt
```

## Utilisation

### Recherche dans le premier onglet
```bash
python recherche_occurrences_excel.py EdT_2026-27_FISE_2A_V2.xlsx "NomDuCours"
```

### Recherche dans tous les onglets
```bash
python recherche_occurrences_excel.py EdT_2026-27_FISE_2A_V2.xlsx "NomDuCours" --tous_onglets
```

### Spécifier un nom d'onglet de sortie personnalisé
```bash
python recherche_occurrences_excel.py EdT_2026-27_FISE_2A_V2.xlsx "NomDuCours" --onglet_sortie "MesRésultats"
```

## Fonctionnalités

- Recherche insensible à la casse
- Création d'un onglet "Résultats" (ou nom personnalisé) avec :
  - Nom de l'onglet source
  - Coordonnées de la cellule
  - Contenu de la cellule (tronqué à 100 caractères si trop long)
  - Lien cliquable vers la cellule
- Style appliqué aux en-têtes
- Les liens sont en bleu et soulignés

## Exemple de sortie

Le programme crée un onglet avec le format suivant :

| Onglet | Cellule | Contenu | Lien |
|--------|---------|---------|------|
| Feuille1 | A1 | Cours de Math | [Lien cliquable] |
| Feuille1 | B5 | TP de Math | [Lien cliquable] |
| Feuille2 | C3 | Examen Math | [Lien cliquable] |

## Notes

- Le fichier Excel est modifié directement (sauvegarde automatique)
- Si un onglet avec le nom de sortie existe déjà, il sera remplacé
- Les liens fonctionnent dans Excel et permettent de naviguer directement vers les cellules
