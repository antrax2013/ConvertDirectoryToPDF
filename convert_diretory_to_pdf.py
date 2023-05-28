# from docx2pdf import convert

# inputFile = "Les_themes.docx"
# outputFile = "Les_themes.pdf"
# file = open(outputFile, "w")
# file.close()

# convert(inputFile)



from subprocess import  Popen
LIBRE_OFFICE = r"C:\Program Files\LibreOffice\program\soffice.exe"

def convert_to_pdf(input_docx, out_folder):
    p = Popen([LIBRE_OFFICE, '--headless', '--convert-to', 'pdf', '--outdir',
               out_folder, input_docx])
    print([LIBRE_OFFICE, '--convert-to', 'pdf', input_docx])
    p.communicate()


sample_doc = 'Les_themes.docx'
out_folder = 'some_folder'
convert_to_pdf(sample_doc, out_folder)