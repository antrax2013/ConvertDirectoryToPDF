Remove-Item -Path "dist" -Force;
pyinstaller --clean -c -n ConvertDirectoryToPDF -w -F 'src/main.py'; 
Copy-Item "src/settings.json",  -Destination "dist"; 
