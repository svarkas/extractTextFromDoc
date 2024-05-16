#!/usr/local/bin/python3.7
import sys
import libs.OldDoc as OD
import libs.DocXtoTXT as DXT


def main(argv):
    inputFile=argv[1]
    extension = inputFile.split(".")[-1]

    if extension == 'doc':
        od = OD.OldDoc()
        print(od.extractText(inputFile))
    elif extension == 'docx':
        dxt = DXT.DocXtoTXT()
        print(dxt.get_docx_txt(inputFile))
    else:
        print("Not a Word Document!!!")
        
   
if __name__ == "__main__":
    main(sys.argv)
