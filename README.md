Alertes leboncoin v2
====================

Script d'alertes email leboncoin.fr via Google Docs / Drive

**Prérequis :** *vous devez avoir un compte Google et y être connecté.*

### Installation en 4 étapes
1. Créer une copie de cette feuille de calcul : https://docs.google.com/spreadsheet/ccc?key=0Atof5tNmg-CYdC1hVTkybGxOYkFhM0Qxd0tIYldneVE&newcopy  

2. Renseignez votre adresse email dans la cellule concernée  

3. Pour chaque requête que vous souhaitez effectuer sur leboncoin.fr, copiez simplement son url dans la colonne "url" (une url par ligne).  

4. Pour que le script soit executé de manière automatique, vous devez programmer un trigger sur la fonction alerteLeBonCoin().  
(Dans outils > editeur de scripts, puis Ressources > déclencheur du script actuel).  
Il est conseillé de régler le trigger sur "toutes les 2 heures".

5. Vous pouvez faire un test en cliquant sur LBC Alertes > Lancer manuellement. (à côté du menu outils)

Créé par Just docs it : http://justdocsit.blogspot.fr  
Code optimisé par Tom  
CSS par [mlb](http://www.maximelebreton.com)  

Modifications par rapport à la v1
* code optimisé
* email avec les styles css du site original
* envoi par email séparé

