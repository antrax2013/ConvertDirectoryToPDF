from abc import ABC, abstractmethod
import os
from os.path import exists

##
## Class to convert file to pdf based on Libre office
class AbstractPdfConverter(ABC):
  
  ##
  ## Function to convert file to pdf
  ## input_docx : the path of document
  ## output_folder : the output path folder
  @abstractmethod
  def convert_to_pdf(self, input_file, output_folder):
      pass
  #endOf
  
  #
  # Function to convert all files in the folder to pdf
  # input_folder : the path folder
  # output_folder : the output path folder
  def convert_files_to_pdf(self,input_folder, output_folder, recursively = True):
      ## Checks
      if not exists(input_folder):
        raise FileNotFoundError("Could not find path: %s"%(input_folder))
      
      for element in os.listdir(input_folder) :
        path = input_folder+"/"+element

        if os.path.isdir(path) and recursively :
          sub_output_folder = output_folder+"/"+element
          self.convert_files_to_pdf(path, sub_output_folder, recursively)
        
        elif os.path.isfile(path):
          self.convert_to_pdf(path, output_folder)
        
        else : 
           print(f"{path} is unknown")
    #endOf convert_files_to_pdf