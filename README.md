# ConvertDirectoryToPDF

ConvertDirectoryToPDF est un outil qui permet de convertir des fichiers "office", ou le contenu de répertoire, en pdf. Il se base au choix sur le moteur MS-Office ou Libre Office en tâche de fond.

# Configuration

1. Ouvrez le fichier settings.json à la racine du répertoire.
1. Modifier les clés suivantes:
1. `officeBGExe` : définit le moteur. Les valeurs attendues sont soit `LIBRE_OFFICE` soit `MS_OFFICE`
1. `libreOfficePath` : est l'emplacement de l'exécutable de libre office. S'il est installé sur votre machine, il est probable que celui que j'ai pré-défini soit le bon.
1. `inputPath` : est le chemin du répertoire contenant les fichiers que vous voulez convertir
1. `outputPath` : est le chemin du répertoire où vous voulez que les fichiers convertis soient créés
