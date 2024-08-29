#!/usr/local/bin/python3.7
import sys
import libs.OldDoc as OD
import libs.DocXtoTXT as DXT
import magic

def main(argv):
    inputFile=argv[1]
    extension = inputFile.split(".")[-1]
    file_type_info = magic.from_file(inputFile)
    if "Composite Document File V2 Document" in file_type_info: 
        od = OD.OldDoc()
        print(od.extractText(inputFile))
    elif file_type_info == "Microsoft Word 2007+":
        dxt = DXT.DocXtoTXT()
        print(dxt.get_docx_txt(inputFile))
    else:
        print("Not a Word Document!!!")
        
   
if __name__ == "__main__":
    main(sys.argv)
