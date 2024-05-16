import xml.etree.ElementTree as ET
from xml.etree.ElementTree import XML
import zipfile


class DocXtoTXT:

    def get_docx_txt(self, inputFile):
        NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        PAR_TAG = NAMESPACE + 'p'
        NODE_IN_PAR_TAG = NAMESPACE + 't'

        document = zipfile.ZipFile(inputFile)
        xml_content = document.read('word/document.xml')
        document.close()
        '''
        ET.parse('word/document.xml') # reads the xml tree from file and requires to xml_tree.getroot
        ET.fromstring(xml_content)  , reads from string and fetches only the root
        '''
        root = ET.fromstring(xml_content)

        paragraphs = []
        for element in root.iter():
            if element.tag == PAR_TAG:
                text = [node.text
                        for node in element.iter(NODE_IN_PAR_TAG)
                        if node.text]
                if text:
                    paragraphs.append((''.join(text)))
        result = '\n\n'.join(paragraphs)
        return result


