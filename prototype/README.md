# Prototype Fin de session

Ce prototype transforme une liste Excel de participants en certificats de réalisation Word. Des PDF de démonstration sont inclus dans `certificats/`.

## Utilisation

1. Ouvrir `donnees-participants.xlsx`.
2. Modifier les lignes de l'onglet `Participants` sans changer les en-têtes.
3. Installer Python, `openpyxl` et `python-docx`.
4. Depuis ce dossier, lancer `python generer_certificats.py donnees-participants.xlsx certificats`.

Seules les lignes au statut `Terminé` sont générées. Le générateur refuse une durée nulle ou supérieure à la durée prévue.

Toutes les identités, entreprises et références légales de ce dépôt sont fictives. Le prototype doit être adapté au modèle et aux règles du client avant un usage professionnel.
