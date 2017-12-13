import os
import sys

from xml.etree.ElementTree import XML
import zipfile

import json

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
    Output list of sections
    '''

    WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    PARA = WORD_NAMESPACE + 'p'
    TEXT = WORD_NAMESPACE + 't'

    paragraphs= []
    previous = False
    for paragraph in tree.getiterator(PARA):
        texts = [node.text
                for node in paragraph.getiterator(TEXT)
                if node.text]
        joined = ''.join(texts)

        # if joined and previous:
        #     paragraphs[-1][-1].append(joined[:-1])
        #     previous = False
        # elif joined and joined[-1]==';':
        #     paragraphs[-1].append([joined[:-1]])
        #     previous = True
        # elif joined and not previous:
        #     paragraphs[-1].append(joined[:-1])
        # else:
        #     paragraphs.append([])
        if joined and previous:
            paragraphs[-1][-1].append(joined[:-1])
            if joined[-1] == '.':
                previous = False
        elif joined and joined[-1] == ';':
                paragraphs[-1].append([joined[:-1]])
                previous = True
        elif joined:
            if joined[-1] == 'A':
                paragraphs[-1].append(joined)
            else:
                paragraphs[-1].append(joined[:-1])
        elif not joined:
            paragraphs.append([])

    return paragraphs

def to_json(paragraphs):
    '''

    '''
    title = paragraphs[0][0]

    contents = []
    doc = {}
    for p in paragraphs[1:]:
        sub = {}
        sub[p[0]] = p[1:]
        contents.append(sub)

    doc[title] = contents
    #
    j_str = json.dumps(doc, sort_keys=True, indent=4)

    return j_str

def json_html(json_data):
    doc = json.loads(json_data)
    html = json2html.convert(json = doc)
    file_name = '/'.join(sys.path[1].split('/')[:-1]) + '/index.html'
    with open(file_name, 'w') as f:
        f.write(html)
    f.close()
    print('index.html updated')

if __name__ == '__main__':
    from json2html import *
    tree = read_file()
    paragraphs = parse_xml(tree)

    # for i in range(0,len(paragraphs)):
    #     print('{}.  {}'.format(i, paragraphs[i]))
    # import numpy as np
    # print(np.matrix(paragraphs))
    json_html(to_json(paragraphs))
