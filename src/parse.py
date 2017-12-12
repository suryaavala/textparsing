import os
import sys

from xml.etree.ElementTree import XML
import zipfile

def read_file():
    '''
    Read the doc file
    Return xml tree
    '''
    if len(sys.argv) > 1:
        doc_name = sys.argv[1]
    else:
        doc_name = 'cv.docx'

    # print(doc_name)
    doc = zipfile.ZipFile(doc_name)
    xml_content = doc.read('word/document.xml')
    doc.close()

    #xml string to tree
    tree = XML(xml_content)

    return tree

if __name__ == '__main__':
    tree = read_file()
    print(type(tree), tree)
