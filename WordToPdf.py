import os
import comtypes.client
import docx
import argparse

def main():
    parser = argparse.ArgumentParser(description='')
    parser.add_argument('Archivo_Doc',type=str)
    parser.add_argument('Archivo_Pdf',type=str)
    
    args = parser.parse_args()

    # Set the paths for the Word and PDF files
    word_path = args.Archivo_Doc
    pdf_path = args.Archivo_Pdf

    #wpath = os.path.dirname(word_path) #Ruta del archivo
    #wname = os.path.splitext(os.path.basename(word_path))[0] #Nombre del archivo

    #pdf_path = wpath + "\\" + wname + ".pdf" #Misma ruta que el docx

    try:
        # Load the Word document using the docx library
        doc = docx.Document(word_path)
        
        # Save the Word document as a PDF using Microsoft Word
        word = comtypes.client.CreateObject("Word.Application")
        docx_path = os.path.abspath(word_path)
        pdf_path = os.path.abspath(pdf_path)
        
        pdf_format = 17  # PDF file format code
        word.Visible = False
        in_file = word.Documents.Open(docx_path)
        in_file.SaveAs(pdf_path, FileFormat=pdf_format)
        in_file.Close()
        
        # Quit Microsoft Word
        word.Quit()

    except Exception as e:
        # Manejar cualquier otra excepción no especificada anteriormente
        print("No fue posible convertir a PDF.", e)

if __name__ == "__main__":
    main()