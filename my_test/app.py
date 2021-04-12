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
    PARAGRAPH = 'p'
    SPAN = 'span'

    file = open(html_path)
    soup = BeautifulSoup(file, 'html.parser')
    paragraph_properties = [dict([i.split(': ') for i in r['style'][:-1].split('; ')])
                            for r in soup.find_all(PARAGRAPH)]
    run_properties = [[dict([k.split(': ') for k in j['style'][:-1].split('; ')]) for j in r.find_all(SPAN)]
                      for r in soup.find_all(PARAGRAPH)]
    properties = list(zip(paragraph_properties, run_properties))

    return properties


# only for debugging
def write_xml():
    document = Document(docx_path)
    document_xml = document.element.xml
    file = open('sample.xml', 'w')
    file.write(document_xml)


def get_paragraph_count(config):
    return [int(i) for i in config['PARAGRAPH_COUNT']['paragraphs']]


def get_difference(config, properties):
    difference = diff(properties, config)
    return list(difference)


def compare_paragraphs(config, properties):
    paragraph_difference = [get_difference(config['PARAGRAPH'], i[0]) for i in properties]
    return paragraph_difference


def compare_runs(config, properties):
    run_differences = [[get_difference(config['FONT'], j) for j in i[1]] for i in properties]
    return run_differences


def print_difference(difference, section_name):
    print(f'{section_name} DIFFERENCE:')
    if not any(difference):
        print('ok')
    else:
        for i in range(len(difference)):
            if len(difference[i]) > 0:
                print(f'ERROR IN PARAGRAPH {i + 1}\n', difference[i])


def run():
    write_xml()
    convert_docx_to_html()
    config = read_config()['HEADER']
    count = get_paragraph_count(config)
    print('config', config)
    properties = get_properties_from_html()
    print('properties', properties)
    print('===========================================')
    print('CHECKING...')
    paragraph_difference = compare_paragraphs(config, properties)
    run_difference = compare_runs(config, properties)
    print_difference(paragraph_difference, 'PARAGRAPH')
    print('===========================================')
    print_difference(run_difference, 'RUN')
