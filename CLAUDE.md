# Contexte de reprise — Fin de session

## Produit

Fin de session automatise la création en série de certificats de réalisation à partir d'un tableau Excel. Chaque ligne éligible produit un document Word modifiable et un PDF final.

Ce projet est distinct de CADENCE. Ne pas réintroduire la marque CADENCE, son identité visuelle ou son positionnement reporting.

## Décisions validées

- Nom : Fin de session
- Cible : petits organismes de formation français
- Offre d'entrée : 490 € HT en paiement unique
- Cas d'usage volontairement étroit : certificats Word et PDF depuis Excel
- Promesse : conserver les outils et modèles existants du client
- Positionnement : prestation sur mesure, pas SaaS générique

## Direction visuelle

Le site doit rester éditorial, franc et utile : grille visible, typographie serif expressive, aplats orange/bleu/vert et presque aucun arrondi. Éviter les gradients, les cartes interchangeables, les halos et le vocabulaire creux typiques des landing pages générées par IA.

La scène 3D du hero est fonctionnelle : elle montre une feuille Excel qui devient une pile de certificats. Elle est faite en CSS, sans bibliothèque ni dépendance externe.

## Contraintes

- Site statique sans build dans `dist/`
- Le domaine principal est publié par GitHub Pages à chaque mise à jour de `main`. L'ancienne adresse ChatGPT Sites possède une source séparée qui ne doit rester qu'une redirection.
- HTML, CSS et JavaScript natifs
- Accessible au clavier et lisible sur mobile
- Ne jamais publier de vraies données de prospect
- Ne pas affirmer une conformité Qualiopi garantie
- Conserver l'exemple PDF et le tableur téléchargeables

## Prochaine amélioration possible

Brancher un vrai formulaire de qualification, compléter les mentions légales déjà publiées avec le SIREN dès réception, puis instrumenter les conversions avec un outil respectueux de la vie privée. La page actuellement publiée indique « immatriculation en cours » : ne pas inventer de numéro.
