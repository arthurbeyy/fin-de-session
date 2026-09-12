# Prototype Fin de session

Ce prototype transforme une liste Excel de participants en certificats de réalisation, au format Word et au format PDF.

## Ce qu'il faut installer une seule fois

1. Python.
2. Les deux briques qui lisent le tableur et écrivent le Word : `pip install openpyxl python-docx`.
3. LibreOffice, gratuit, qui fabrique les PDF à partir des Word. Sans lui le programme produit quand même les Word et le dit clairement. Si LibreOffice est absent mais que Microsoft Word est installé, le programme sait aussi passer par Word à condition d'avoir `pip install docx2pdf`.

## Utilisation

1. Ouvrir `donnees-participants.xlsx`.
2. Modifier les lignes de l'onglet `Participants` sans changer les en-têtes.
3. Depuis ce dossier, lancer `python generer_certificats.py donnees-participants.xlsx certificats`.

Le programme écrit dans le dossier `certificats` un Word et un PDF par participant, plus un fichier `journal.txt`.

## Ce que fait le programme sur les lignes imparfaites

Il ne s'arrête jamais en cours de route. Il traite ce qu'il peut et note le reste dans `journal.txt`, qui sépare deux choses :

- les participants au statut autre que `Terminé`, ignorés volontairement, ce qui est normal ;
- les lignes à corriger : durée nulle ou négative, durée réalisée supérieure à la durée prévue, durée qui n'est pas un nombre, colonne obligatoire vide.

Il suffit de corriger le tableur et de relancer. Les certificats déjà produits sont réécrits.

## Adapter au client

Les cinq premières lignes du programme, sous le commentaire prévu à cet effet, portent le nom de l'organisme, le signataire, sa fonction, la ville de signature et le pied de page. Ce sont les seules à changer pour un nouveau client. La mise en page du certificat, elle, se règle au cas par cas à partir du modèle existant du client.

La date de signature est celle du jour où les certificats sont générés.

## Données

Toutes les identités, entreprises et références légales de ce dépôt sont fictives. Le prototype doit être adapté au modèle et aux règles du client avant un usage professionnel.
