#pip install nose
import unittest
import sys
sys.path.append('../')
from LibreOfficeConverter import *
import shutil


class TestLibreOfficeConverter(unittest.TestCase):
    
    def tearDown(self) -> None:
        try :
            path ="results_folder"
            if os.path.exists(path):
                shutil.rmtree("results_folder")
        except OSError as e:
            print("Error: %s - %s." % (e.filename, e.strerror))
        return super().tearDown()

    def test_constructor_libreOfficeExe_FileNotFoundError(self):
        with self.assertRaises(FileNotFoundError): LibreOfficeConverter("")

    #Region test_convert_to_pdf
    def test_convert_to_pdf_returns_correct_result(self):
        converter = LibreOfficeConverter()
        #print(f"test => {converter.libre_office_path}")
        converter.convert_to_pdf("sources_folder/file_1.docx", "results_folder")

        generatedFile = "results_folder/file_1.pdf"
        self.assertGreater(os.path.getsize(generatedFile),0)


    def test_convert_to_pdf_returns_raise_FileNotFoundError(self):
        converter = LibreOfficeConverter()
        with self.assertRaises(FileNotFoundError):
            converter.convert_to_pdf("sources_folder/unknown.file", "results_folder")
    #EndRegion test_convert_to_pdf

    #Region convert_files_to_pdf
    def test_convert_files_to_pdf_returns_raise_FileNotFoundError(self):
        converter = LibreOfficeConverter()
        with self.assertRaises(FileNotFoundError):
            converter.convert_files_to_pdf("unknownFolder","unknownFolder")

    def test_convert_files_to_pdf_returns_1_file(self):
        converter = LibreOfficeConverter()
        converter.convert_files_to_pdf("sources_folder","results_folder", False)        
        files = os.listdir("results_folder")
        self.assertEqual(6, len(files))
       
    #EndRegion convert_files_to_pdf

if __name__ == '__main__':
    unittest.main()
