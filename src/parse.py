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

def parse_xml(tree):
    '''
    Parse xml document and
    Output list of paragraphs
    '''

    WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    PARA = WORD_NAMESPACE + 'p'
    TEXT = WORD_NAMESPACE + 't'

    paragraphs= []
    for paragraph in tree.getiterator(PARA):
        texts = [node.text
                for node in paragraph.getiterator(TEXT)
                if node.text]
        joined = ''.join(texts)

        if texts:
            paragraphs[-1].append(''.join(texts))
        else:
            paragraphs.append([])


    return paragraphs

if __name__ == '__main__':
    tree = read_file()
    paragraphs = parse_xml(tree)

    for i in range(0,len(paragraphs)):
        print('{}.  {}'.format(i, paragraphs[i]))
