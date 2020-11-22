##############################################
## Script to convert .doc or .docx to .pdf  ##
##         Written by Henry Ling            ##
##############################################

import os
import argparse
from time import strftime
from win32com import client

parser = argparse.ArgumentParser(
	description = "Convert doc/docx to pdf"
)

# Argument --path for cli use
parser.add_argument(
	"--path",
	type = str,
	default = "."
	help = "Directory path were document files are"
)

args = parser.parse_args()
path = args.path

dir_content = os.listdir(path)
path_dir_content = [os.path.join(path, doc) for doc in dir_content]
docs = [doc for doc in path_dir_content if os.path.isfile(doc)]

# Counts the number of files in the directory that can be converted
def n_files(path):
    return len(docs)

# Creates a new directory within current directory called PDFs
def createFolder(path):
	if not os.path.exists(path + '\\PDFs'):
		try:
			os.mkdir(path + '\\PDFs')
			print("\nfolder created.")
		except FileExistsError as er:
			print(f"\nFolder already exists at {path}.. {er}")
    
		
if __name__ == "__main__":
    print('\nPlease note that this will overwrite any existing PDF files')
    print('For best results, close Microsoft Word before proceeding')
    input('Press enter to continue.')
	
    if n_files(path) == 0:
        print('\nThere are no files to convert')
        exit()
		
    createFolder(path)
	
    print(f"\nStarting conversion of {len(docs)} docs... \n")
	
    # Opens each file with Microsoft Word and saves as a PDF
    try:
        word = client.DispatchEx('Word.Application')
        for file in dir_content:
            if (file.endswith('.doc') or file.endswith('.docx') or file.endswith('.tmd')):
                ending = ""
                if file.endswith('.doc'):
                    ending = '.doc'
                if file.endswith('.docx'):
                    ending = '.docx'
                if file.endswith('.tmd'):
                    ending = '.tmd'
                new_name = file.replace(ending,r".pdf")
                in_file = os.path.abspath(path + '\\' + file)
                new_file = os.path.abspath(path + '\\PDFs' + '\\' + new_name)
                doc = word.Documents.Open(in_file)
                print(new_name)
                doc.SaveAs(new_file,FileFormat = 17)
                doc.Close()
    except :
        print("Error: Aborting")
    finally:
        word.Quit()

    print('\nConversion finished at ' + strftime("%H:%M:%S"))
