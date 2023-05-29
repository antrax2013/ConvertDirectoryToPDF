#pip install nose
import unittest
import sys
sys.path.append('../')
from MsOfficeConverter import *
import shutil


class TestMsOfficeConverter(unittest.TestCase):
    
    def tearDown(self) -> None:
        try :
            path ="results_folder"
            if os.path.exists(path):
                shutil.rmtree("results_folder")
        except OSError as e:
            print("Error: %s - %s." % (e.filename, e.strerror))
        return super().tearDown()

    def test_constructor_MsOfficeExe_FileNotFoundError(self):
        with self.assertRaises(ModuleNotFoundError): MsOfficeConverter()

    def test_constructor_MsOfficeExe_skipeCheck(self):
        converter = MsOfficeConverter(True)
        self.assertEqual(0, converter.versionNum)

    #Region test_convert_to_pdf
    def test_convert_word_to_pdf_returns_correct_result(self):
        converter = MsOfficeConverter()
        converter.convert_to_pdf("sources_folder/file_1.docx", "results_folder")

        generatedFile = "results_folder/file_1.pdf"
        self.assertGreater(os.path.getsize(generatedFile),0)

    def test_convert_excel_to_pdf_returns_correct_result(self):
        converter = MsOfficeConverter()
        converter.convert_to_pdf("sources_folder/tableur_2.xlsx", "results_folder")

        generatedFile = "results_folder/tableur_2.pdf"
        self.assertGreater(os.path.getsize(generatedFile),0)

    def test_convert_to_pdf_returns_raise_FileNotFoundError(self):
        converter = MsOfficeConverter()
        with self.assertRaises(FileNotFoundError):
            converter.convert_to_pdf("sources_folder/unknown.file", "results_folder")

    def test_convert_to_pdf_returns_raise_UknowExtensionError(self):
        converter = MsOfficeConverter()
        with self.assertRaises(UknowExtensionError):
            converter.convert_to_pdf("sources_folder/tableur_1.ods", "results_folder")
 
    #EndRegion test_convert_to_pdf   

if __name__ == '__main__':
    unittest.main()
