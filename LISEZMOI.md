# Convertisseur OST vers PST (gratuit)

Convertit un fichier OST « orphelin » (Outlook 2013, 2016, 2019 ou 365) en vrai fichier PST, ouvert directement dans Outlook.

## Ce qu'il faut sur le PC

- **Windows 10 ou 11** avec **Outlook** installé. Outlook doit avoir été ouvert au moins une fois.
- **Une connexion Internet au premier lancement.** Le convertisseur télécharge alors un Python portable (python.org) et la bibliothèque open source libpff (pypi.org) dans `%LOCALAPPDATA%\ost2pst`.

Aucun droit administrateur n'est nécessaire, et WSL n'est pas nécessaire non plus.

## Utilisation

1. Mettez les deux fichiers `Convertir-OST-en-PST.ps1` et `Lancer-le-convertisseur.bat` dans le même dossier.
2. Double-cliquez sur **Lancer-le-convertisseur.bat**. Ne le lancez pas en administrateur, sinon Outlook ne pourra pas communiquer avec le script.
3. Choisissez le fichier OST, puis l'emplacement du PST, et cliquez sur **Convertir**.
4. À la fin, le PST apparaît dans Outlook sous le nom « Converti - … ». Un journal `.log.txt` est créé à côté du PST.

Pendant la conversion, une copie temporaire du contenu est créée dans `C:\ost2pst_…`. Elle est supprimée automatiquement à la fin.

## Ce qui est repris

- L'arborescence des dossiers.
- Les e-mails : expéditeur, destinataires, dates d'envoi et de réception, statut lu ou non lu, corps HTML, pièces jointes et en-têtes Internet d'origine.
- Les contacts, rendez-vous, tâches et notes, avec leurs champs principaux.
- En option, les éléments supprimés que l'outil arrive à récupérer.

## Limites connues

- Les rendez-vous récurrents sont importés comme une seule occurrence. Le texte de la récurrence d'origine est ajouté dans leur description.
- Les listes de distribution ne sont pas reprises.
- Les mails joints à un autre mail (transfert « en pièce jointe ») ne sont pas repris. Le mail qui les contenait est bien importé.
- Selon la version d'Outlook, la colonne « À » peut rester vide dans la liste des messages. Les destinataires restent visibles dans les en-têtes d'origine du message (Fichier > Propriétés).
- Les fichiers OST chiffrés ou très endommagés peuvent n'être récupérés qu'en partie.

## En cas de problème

- **« Téléchargement impossible »** : le réseau (proxy ou pare-feu d'entreprise) bloque peut-être python.org ou pypi.org. Essayez depuis une autre connexion.
- **Utilisation en ligne de commande** :
  `powershell -ExecutionPolicy Bypass -File .\Convertir-OST-en-PST.ps1 -OstPath C:\ost\boite.ost -PstPath C:\ost\boite.pst -IncludeRecovered`
