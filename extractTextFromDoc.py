#!/usr/local/bin/python3.7
import sys
import libs.OldDoc as od
def main(argv):
    inputFile=argv[1]
    extension = inputFile.split(".")[-1]
    
    myod = od.OldDoc()
    if (extension == 'doc'):
        print(myod.extractText(inputFile))
    elif (extension == 'docx'):
        from docx import Document
        document = Document(inputFile)
        text = ''
        for para in document.paragraphs:
            text = text + para.text
        print(text)
    else:
        print("Not a Word Document!!!")
        
   
if __name__ == "__main__":
    main(sys.argv)