# ConvertDirectoryToPDF

ConvertDirectoryToPDF est un outil qui permet de convertir des fichiers "office", ou le contenu de répertoire, en pdf. Il se base au choix sur le moteur MS-Office ou Libre Office en tâche de fond.

# Configuration

1. Ouvrez le fichier settings.json à la racine du répertoire.
1. Modifier les clés suivantes:
1. `officeBGExe` : définit le moteur. Les valeurs attendues sont soit `LIBRE_OFFICE` soit `MS_OFFICE`
1. `libreOfficePath` : est l'emplacement de l'exécutable de libre office.
1. `inputPath` : est le chemin du répertoire contenant les fichiers que vous voulez convertir
1. `outputPath` : est le chemin du répertoire où vous voulez que les fichiers convertis soient créés

### exemple

```JSON
{
"officeBGExe":"LIBRE_OFFICE",
"libreOfficePath":"Include\\LibreOffice\\program\\soffice.exe",
"inputPath":"C:\\Workspace\\Python\\ConvertDirectoryToPDF\\tests\\sources_folder",
"outputPath":"C:\\temp"
}
```

# Evolutions à venir

1. Gérer les accents dans le nom des paths
1. Pouvoir inclure/exclure : des fichiers, des répertoires et des extensions

# Pense bête

## Env

Activer/Désactiver l'environement virtuel :
"env/Scripts/activate.bat"
"deactivate.bat"

## Run Tests

python -m unittest /tests

## Dependance

pip install nose
pip install docx2pdf
