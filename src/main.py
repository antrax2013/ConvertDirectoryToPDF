from LibreOfficeConverter import *
from MsOfficeConverter import *
import logging

import json
from enum import Enum

class OfficeBGExe(Enum):
    LIBRE_OFFICE = 'LIBRE_OFFICE'
    MS_OFFICE = 'MS_OFFICE'

logging.basicConfig(filename='ConvertDirectoryToPDF.log', encoding='utf-8', level=logging.INFO)

logging.info('Process started...')

try:
  settings = None
  with open("settings.json", 'r') as f:
    settings = json.load(f)


  if(settings != None) :
    if(settings["officeBGExe"] == OfficeBGExe.LIBRE_OFFICE.value ) :
      converter = LibreOfficeConverter(settings["libreOfficePath"])
    else :
      converter = MsOfficeConverter()
except Exception as e: # work on python 3.x
    logging.error('Failed to settings: '+ str(e))

try:
  converter.convert_files_to_pdf(settings["inputPath"],settings["outputPath"])
except Exception as e: # work on python 3.x
    logging.error('On error occured: '+ str(e))
finally:
   logging.info('Process finished')
