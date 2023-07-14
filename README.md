# ConvertDirectoryToPDF

ConvertDirectoryToPDF est un outil qui permet de convertir des fichiers "office", ou le contenu de répertoire, en pdf. Il se base au choix sur le moteur MS-Office ou Libre Office en tâche de fond.

# Configuration

1. Ouvrez le fichier settings.json à la racine du répertoire.
1. Modifier les clés suivantes:

- `officeBGExe` : définit le moteur. Les valeurs attendues sont soit `LIBRE_OFFICE` soit `MS_OFFICE`
- `libreOfficePath` : est l'emplacement de l'exécutable de libre office.
- `inputPath` : est le chemin du répertoire contenant les fichiers que vous voulez convertir
- `outputPath` : est le chemin du répertoire où vous voulez que les fichiers convertis soient créés

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
1. Passer par une API pour ce passer de la dépendance avec les exe

# Pense bête

## Env

Activer/Désactiver l'environement virtuel :
```PowerShell
"./env/Scripts/activate"
deactivate
```

## Executer les Tests

```Powershell
python -m unittest /tests
```

## Dépendances

```Powershell
pip install nose
pip install docx2pdf
pip install comtypes
```
