# Coordination Codex vers Claude

Le dépôt GitHub est le relais commun. Aucun fait indispensable ne doit rester dans un chat.

## Début de relève

1. Mettre à jour la copie locale depuis main.
2. Lire PROJECT.md, STATE.md, DECISIONS.md, BACKLOG.md et ce fichier.
3. Choisir la première action P0 non bloquée.

## Fin de relève

1. Mettre à jour STATE.md avec ce qui a réellement été fait, les preuves, les blocages et l'action suivante.
2. Ajouter une décision dans DECISIONS.md seulement si elle change durablement le projet.
3. Mettre à jour BACKLOG.md.
4. Créer un seul commit descriptif et le pousser sur main, sauf si la publication ou l'écriture GitHub n'est pas autorisée.

## Passage à Claude

Quand Codex approche de sa limite de consommation, il ne commence pas de chantier complexe. Il termine l'étape atomique en cours, met à jour les trois fichiers ci-dessus, pousse le commit puis laisse cette instruction :

> Claude, reprends depuis main. Lis les fichiers de reprise, puis exécute la première action P0 non bloquée. Mets à jour l'état et pousse un unique commit à la fin.

Claude n'a pas besoin du chat précédent pour reprendre.
