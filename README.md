# Convertisseur Cartographie → CSV

Transforme la feuille **« Cartographie »** d'un classeur Excel (`.xlsm` ou `.xlsx`)
en fichier CSV prêt pour le **chargement en masse** de l'outil d'analyse
réglementaire.

## Installation

Il faut Python 3.8 ou plus récent et la bibliothèque `openpyxl` :

```
pip install openpyxl
```

Sous Windows, si `python3` n'est pas reconnu, utiliser `python` (ou `py`) dans
les commandes ci-dessous.

## Utilisation

Un CSV par classeur, écrit à côté de chaque classeur (`base1.csv`, `base2.csv`…) :

```
python3 convertir_cartographie.py base1.xlsm base2.xlsm
```

Plusieurs classeurs réunis dans un seul CSV :

```
python3 convertir_cartographie.py base1.xlsm base2.xlsm -o tout.csv
```

Réunir les bases est conseillé quand un même article peut figurer dans plusieurs
d'entre elles : chargé une seconde fois depuis un autre fichier, il serait
rejeté par l'outil (« déjà dans la base »).

Le programme affiche le nombre d'articles et de substances produits, ainsi que
les éventuels avertissements.

## Fichier source attendu

- Une feuille nommée **Cartographie**.
- Une ligne d'en-têtes dans les 10 premières lignes. Les colonnes sont repérées
  par leur titre (majuscules et espaces indifférents), quelle que soit leur
  position :

  | En-tête dans le classeur | Obligatoire |
  |---|:---:|
  | Référence Fabricant | ✅ |
  | Fabricant/Fournisseur | ✅ |
  | Nom substance (s) ou Composé(s) | ✅ |
  | N°CAS | |
  | N°CE | |
  | Présence ou concentration sur article (%) | |

- **Un article s'étend sur plusieurs lignes** : la première porte la référence et
  le fabricant ; les lignes suivantes, où ces deux colonnes sont vides, ajoutent
  chacune une substance à cet article.

Les autres colonnes (Référence ID, désignation, pays, analyses calculées,
justificatifs, code déchet, niveaux d'ensemble…) ne sont pas reprises.

## CSV produit

- Mêmes 15 colonnes que l'export de la liste des articles de l'outil, de « Nom
  fabricant » à « Statut actuel ».
- UTF-8 avec marque d'ordre (lisible directement par Excel), séparateur `;`.
- Colonnes remplies :
  - **Nom fabricant** ← Fabricant/Fournisseur ;
  - **Réf fabricant** ← Référence Fabricant ;
  - **Substances** : une substance par ligne dans la cellule,
    `nom | n° CAS | n° CE | %` (champs vides de fin omis).
- Les autres colonnes (conformités, certificats, directives…) sont présentes
  mais vides.

Règles de conversion :

- Les noms de substances sont recopiés tels quels (espaces doublés réduits).
- **Aucune valeur n'est corrigée** : un n° CAS invalide (par exemple un code
  d'alliage comme `AISI_303`) est recopié ; l'analyse du fichier dans l'outil
  le signalera.
- Un même article (fabricant + réf, majuscules et espaces indifférents) présent
  plusieurs fois n'apparaît qu'une fois, avec ses substances réunies sans
  doublon.
- Une référence saisie comme nombre sort sans décimale (`2170144`). Si la
  cellule a un format à zéros (`000000`), les zéros de tête sont conservés.
- Une cellule de pourcentage au format % d'Excel (valeur 0,003) donne `0,3`.

## Avertissements

Le fichier `<sortie>_avertissements.txt`, écrit à côté du CSV, liste :

- les substances placées avant le premier article (ignorées) ;
- les articles sans référence ou sans fabricant (gardés, l'outil les rejettera) ;
- les articles fusionnés, avec la ligne où ils ont été vus en premier ;
- un dépassement de 20 000 articles, le maximum accepté par l'outil en un
  chargement (le fichier est alors à découper).

Ce fichier n'est pas créé s'il n'y a rien à signaler.

## Ensuite, dans l'outil

1. Page « Chargement en masse » : **Analyser un fichier** avec le CSV produit.
2. Corriger dans le CSV ce que le rapport d'analyse signale (n° CAS invalides,
   noms de fabricant proches, etc.), puis relancer l'analyse.
3. **Simuler le chargement**, puis charger.
