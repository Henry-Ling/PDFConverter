##############################################
## Script to convert .doc or .docx to .pdf  ##
##         Written by Henry Ling            ##
##############################################

import os
from time import strftime
from sys import platform, exc_info
from win32com import client

# Counts the number of files in the directory that can be converted
def n_files(directory):
    total = 0
    for file in os.listdir(directory):
        if (file.endswith('.doc') or file.endswith('.docx') or file.endswith('.tmd')):
            total += 1
    return total

# Creates a new directory within current directory called PDFs
def createFolder(directory):
    if not os.path.exists(directory + '/PDFs'):
        os.makedirs(directory + '/PDFs')

def doc2pdf_libreoffice(doc, ending, newdic):
    """
    convert a doc/docx document to pdf format (linux only, requires libreoffice)
    :param doc: path to document
    """
    cmd = f"lowriter --convert-to pdf:writer_pdf_Export '{doc}'"
    os.system(cmd)
    new_file = doc.replace(ending, r".pdf")
    if platform =='linux':
        cmdmove = f"mv '{new_file}' '{newdic}'"

    elif platform == 'win32':
        new_file = new_file.replace("/", "\\")
        cmdmove = f"move '{new_file}' '{newdic}'"
    
    os.system(cmdmove)
    print(new_file)

def is_tool(name):
    """Check whether `name` is on PATH and marked as executable."""

    # from whichcraft import which
    from shutil import which

    return which(name) is not None


if __name__ == "__main__":
    print('\nPlease note that this will overwrite any existing PDF files')
    print('For best results, close Microsoft Word before proceeding')
    input('Press enter to continue.')
	
    directory = os.getcwd()
	
    if n_files(directory) == 0:
        print('There are no files to convert')
        exit()
		
    createFolder(directory)
	
    print('Starting conversion... \n')

    # Opens each file with Microsoft Word and saves as a PDF
    try:
        if(is_tool('libreoffice') == False):
            
            word = client.DispatchEx('Word.Application')
        for file in os.listdir(directory):
            if (file.endswith('.doc') or file.endswith('.docx') or file.endswith('.tmd')):
                ending = ""
                if file.endswith('.doc'):
                    ending = '.doc'
                if file.endswith('.docx'):
                    ending = '.docx'
                if file.endswith('.tmd'):
                    ending = '.tmd'
                if is_tool('libreoffice'):
                    in_file = os.path.abspath(directory + '/' + file)
                    new_file = os.path.abspath(directory + '/PDFs')
                    doc2pdf_libreoffice(in_file, ending, new_file)

                if(is_tool('libreoffice') == False):
                    new_name = file.replace(ending,r".pdf")
                    in_file = os.path.abspath(directory + '\\' + file)
                    new_file = os.path.abspath(directory + '\\PDFs' + '\\' + new_name)
                    doc = word.Documents.Open(in_file)
                    print(new_name)
                    doc.SaveAs(new_file,FileFormat = 17)
                    doc.Close()
                
    except Exception as e:
        print(e)

    print('\nConversion finished at ' + strftime("%H:%M:%S"))
