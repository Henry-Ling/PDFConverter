# PDFConverter

### Introduction

Python script to convert a large number of .doc, .docx, and .tmd (Textmaker doc) to .pdf. I made this while having to convert a number of Word documents to PDF format to send in for job applications.

### Required Modules

The only module required that is not included in the default Python 3.5.1 installation is [win32com](http://starship.python.net/~skippy/win32/Downloads.html)

### Usage

To use this script, put it in a directory of files you want to convert, and then run it. It will create a new folder within the given directory called 'PDFs', and save all converted files there. This only works on Windows, due to the usage of win32com.

### Future Changes & Notes

Could (and should) be written to include a case for UNIX environments

### Resources

[win32com](http://starship.python.net/~skippy/win32/Downloads.html)
