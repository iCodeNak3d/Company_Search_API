# Script de Recherche d'Entreprises

Ce script permet d'enrichir un fichier Excel contenant des noms d'entreprises avec des informations obtenues via l'API Recherche d'Entreprises.

## Prérequis

- Python 3.6 ou supérieur
- Packages Python listés dans `requirements.txt`

## Installation

```bash
pip install -r requirements.txt
```

## Utilisation

Le fichier d'entrée doit être un fichier Excel (.xlsx) avec au moins deux colonnes :
- `Company` : Le nom de l'entreprise à rechercher
- `Address` : L'adresse de l'entreprise (optionnelle mais améliore la précision)

Pour exécuter le script :

```bash
python3 enrich_simple.py --input votre_fichier.xlsx --output resultat.xlsx
```

Options disponibles :
- `--input`, `-i` : Chemin vers le fichier Excel d'entrée (par défaut : "data.xlsx")
- `--output`, `-o` : Chemin vers le fichier Excel de sortie (par défaut : "data_enrichi.xlsx")
- `--token`, `-t` : Token d'API Recherche d'Entreprises (un token par défaut est inclus dans le script)
- `--debug`, `-d` : Active le mode debug pour plus d'informations

## Résultats

Le script génère :
- Un fichier Excel enrichi (.xlsx) avec toutes les informations trouvées
- Un fichier CSV (.csv) avec les mêmes données

La mise en forme conditionnelle est appliquée au fichier Excel pour surligner en rouge les lignes où l'adresse ne correspond pas.

## Particularités du script

- Extraction des premiers prénoms des dirigeants pour éviter les doublons
- Supression de la colonne Code_Effectif qui n'est pas nécessaire
- Mise en forme conditionnelle des lignes où l'adresse ne correspond pas (uniquement dans le fichier Excel)