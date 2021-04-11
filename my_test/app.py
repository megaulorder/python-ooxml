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
TEXT = 'span'
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
    properties = {'PARAGRAPH': {}, 'FONT': {}}
    paragraph_properties = dict(item.split(": ") for item in soup.__getattr__(PARAGRAPH).attrs['style'][:-1].split(";"))
    text_properties = dict(item.split(": ") for item in soup.__getattr__('b').attrs['style'][:-1].split(";"))
    properties['PARAGRAPH'] = paragraph_properties
    properties['FONT'] = text_properties

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


def run():
    write_xml()
    convert_docx_to_html()
    # config = read_config()
    # print('config', config['HEADER'])
    # properties = get_properties_from_html()
    # print('properties', properties)
    # difference = check_differences(config['HEADER'], properties)
    #
    # if len(difference) == 0:
    #     print('ok')
    # else:
    #     print(difference)
