from LibreOfficeConverter import *
from MsOfficeConverter import *
import json
from enum import Enum

class OfficeBGExe(Enum):
    LIBRE_OFFICE = 'LIBRE_OFFICE'
    MS_OFFICE = 'MS_OFFICE'

settings = None
with open("settings.json", 'r') as f:
  settings = json.load(f)


if(settings != None) :
  if(settings["officeBGExe"] == OfficeBGExe.LIBRE_OFFICE.value ) :
    converter = LibreOfficeConverter(settings["libreOfficePath"])
  else :
    converter = MsOfficeConverter()

  converter.convert_files_to_pdf(settings["inputPath"],settings["outputPath"])
