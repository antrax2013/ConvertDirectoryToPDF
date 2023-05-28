from subprocess import  Popen
import os
from os.path import exists
from AbstractPdfConverter import *


##
## Class to convert file to pdf based on Libre office
class LibreOfficeConverter(AbstractPdfConverter)  :   

  def __init__(self, libre_office_path = r"C:\Program Files\LibreOffice\program\soffice.exe"):
    self.libre_office_path = libre_office_path

    if(not exists(self.libre_office_path)) :
        raise FileNotFoundError(f"LibreOffice not found : {self.libre_office_path}")

  ##
  ## Function to convert file to pdf
  ## input_file : the path of document
  ## output_folder : the output path folder
  def convert_to_pdf(self, input_file, output_folder):
    ## Checks
    if(self.libre_office_path=="") :
        raise ValueError(f"Path to LibreOffice.exe not defined")
    
    if(not exists(input_file)) :
        raise FileNotFoundError(f"File {input_file} not found")
    
    p = Popen([self.libre_office_path, '--headless', '--convert-to', 'pdf', '--outdir',
               output_folder, input_file])
    p.communicate()

