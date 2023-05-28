from LibreOfficeConverter import *

converter = LibreOfficeConverter(r"C:\Program Files\LibreOffice\program\soffice.exe")

print(f"main:{converter.exe_path}")
converter.convert_files_to_pdf("tests/sources_folder","tests/results_folder")