from os.path import exists
from docx2pdf import convert #pip install docx2pdf
from comtypes import client #pip install comtypes
import winreg

#https://github.com/robsonlimadeveloper/msoffice2pdf/blob/develop/src/msoffice2pdf/__init__.py

from AbstractPdfConverter import *

class UknowExtensionError(Exception): ...

##
## Class to convert file to pdf based on MS-Office
class MsOfficeConverter(AbstractPdfConverter)  :   

  def __init__(self, skipeCheck=False):
    AbstractPdfConverter.__init__(self)
    self.versionNum = self.__getMicrosoftWordVersion()

    if(not skipeCheck) :
      if(self.versionNum == 0) :
        raise ModuleNotFoundError(f"MS-Office not found")

  def __getMicrosoftWordVersion(self):
    key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, "SOFTWARE\\Microsoft\\Office", 0, winreg.KEY_READ)
    versionNum = 0
    i = 0
    while True:
        try:
            subkey = winreg.EnumKey(key, i)
            i+=1
            if versionNum < float(subkey):
                versionNum = float(subkey)
        except: #relies on error handling WindowsError as e as well as type conversion when we run out of numbers
            break
    return versionNum

  ##
  ## Function to convert file to pdf
  ## input_file : the path of document
  ## output_folder : the output path folder
  def convert_to_pdf(self, input_file, output_folder):

    ## Checks
    if(self.versionNum == 0) :
        raise ValueError(f"MS-Office not found")
    
    if(not exists(input_file)) :
        raise FileNotFoundError(f"File {input_file} not found")
 
    fileName, ext = os.path.splitext(input_file)
    outputFile = output_folder+"/"+fileName+".pdf"
    
    if ext in [".doc", ".docx", ".txt", ".xml"]:
      self.__convert_Word_to_pdf(input_file, outputFile)
    elif ext in [".xls", ".xlsx"]:
      self.__convert_Excel_to_pdf(input_file, outputFile)
    elif ext in [".ppt", ".pptx"]:
      self.__convert_PowerPoint_to_pdf(input_file, outputFile)
    else :
      raise UknowExtensionError(f"File {input_file} not found")

  ##
  ## Function to convert Word document to pdf
  ## input_file : the path of document
  ## output_folder : the output path folder
  def __convert_Word_to_pdf(self, input_file, outputFile):
    # convert(input_file, outputFile)
    ws_pdf_format: int = 17
    app = client.CreateObject("Word.Application")
    try:
      doc = app.Documents.Open(input_file)
      doc.ExportAsFixedFormat(outputFile, ws_pdf_format, Item=7, CreateBookmarks=0)
    finally:
      app.Quit()
     
  ##
  ## Function to convert Excel document to pdf
  ## input_file : the path of document
  ## output_folder : the output path folder
  def __convert_Excel_to_pdf(self, input_file, outputFile):
    app = client.CreateObject("Excel.Application")
    try:
      sheets = app.Workbooks.Open(input_file)
      sheets.ExportAsFixedFormat(0, outputFile)
    finally:
      app.Quit()
     
  ##
  ## Function to convert powerPoint document to pdf
  ## input_file : the path of document
  ## output_folder : the output path folder
  def __convert_PowerPoint_to_pdf(self, input_file, outputFile):
    app = client.CreateObject("PowerPoint.Application")
    try:
      obj = app.Presentations.Open(input_file, False, False, False)
      obj.ExportAsFixedFormat(outputFile, 2, PrintRange=None)
    finally:
      app.Quit()
