import logging
import ooxml
from configobj import ConfigObj
from docx import Document
from bs4 import BeautifulSoup
from dictdiffer import diff
from ooxml import serialize

logging.basicConfig(filename='ooxml.log', level=logging.INFO)

docx_path = 'test.docx'
config_path = 'config.ini'
html_path = 'sample.html'

PARAGRAPH = 'p'
SPAN = 'span'
BOLD = 'b'


def read_config():
    config = ConfigObj('config.ini')
    properties = dict()
    for section in config.sections:
        properties[section] = dict(config[section])

    return properties


def convert_docx_to_html():
    file_name = docx_path

    dfile = ooxml.read_from_file(file_name)
    output = serialize.serialize(dfile.document).decode("utf-8")
    print(output)

    file = open('sample.html', 'w')
    file.write(output)


def get_properties_from_html():
    file = open(html_path)
    soup = BeautifulSoup(file, 'html.parser')
    paragraph_properties = [[dict([i.split(': ') for i in r['style'][:-1].split('; ')])]
                            for r in soup.find_all(PARAGRAPH)]
    run_properties = [[dict([k.split(': ') for k in j['style'][:-1].split('; ')]) for j in r.find_all_next(SPAN)]
                      for r in soup.find_all(PARAGRAPH)]
    properties = list(zip(paragraph_properties, run_properties))

    return properties


# only for debugging
def write_xml():
    document = Document(docx_path)
    document_xml = document.element.xml
    file = open('sample.xml', 'w')
    file.write(document_xml)


def check_differences(config, properties):
    difference = diff(config, properties)
    return list(difference)


def print_difference(difference, section_name):
    print(f'{section_name} difference:')
    if not any(difference):
        print('ok')
    else:
        for i in difference:
            if len(i) > 0:
                print(i)


def run():
    write_xml()
    convert_docx_to_html()
    config = read_config()['HEADER']
    print('config', config)
    properties = get_properties_from_html()
    print('properties', properties)
    # paragraph_difference = [check_differences(config['PARAGRAPH'], i) for i in properties['PARAGRAPH']]
    # run_difference = [check_differences(config['FONT'], i) for i in properties['FONT']]
    # print_difference(paragraph_difference, 'paragraph')
    # print_difference(run_difference, 'run')
