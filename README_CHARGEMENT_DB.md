# Continuer la saisie sur Render gratuit

Cette version repart de l'application originale et ajoute une seule fonction : `user1` et `admin` peuvent charger une base SQLite sauvegardée depuis la page Saisie.

## Cycle obligatoire

1. Au début : charger la dernière copie de `sotreg_saisie.db`.
2. Effectuer les saisies et cliquer sur Sauvegarder.
3. Avant de quitter : télécharger la nouvelle copie de `sotreg_saisie.db`.
4. Conserver cette copie en lieu sûr et la recharger à la prochaine session.

La base chargée est contrôlée avec `PRAGMA integrity_check` et doit contenir les tables principales. Un fichier invalide est refusé sans écraser la base active.

## GitHub

Pour cette modification, remplacez uniquement `sotreg_web.py`. Les autres fichiers sont fournis afin de disposer d'un paquet complet cohérent.

Commande Render : `gunicorn sotreg_web:app`
