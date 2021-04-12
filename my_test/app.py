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


def get_paragraphs_count(style):
    return [int(i) - 1 for i in style['PARAGRAPH_COUNT']['paragraphs']]


def get_paragraphs_for_styles(config):
    return dict(zip(config.keys(), [get_paragraphs_count(style) for style in config.values()]))


def get_difference(style, properties):
    return list(diff(properties, style))


def compare_paragraphs(style, properties, paragraph_count):
    selected_paragraphs = [properties[i] for i in paragraph_count]
    paragraph_difference = [get_difference(style['PARAGRAPH'], i[0]) for i in selected_paragraphs]

    return paragraph_difference


def compare_runs(style, properties, paragraph_count):
    selected_paragraphs = [properties[i] for i in paragraph_count]
    run_differences = [[get_difference(style['FONT'], j) for j in i[1]] for i in selected_paragraphs]

    return run_differences


def compare_styles(config, paragraphs_for_styles):
    properties = get_properties_from_html()
    paragraph_difference = [compare_paragraphs(config[style], properties, paragraphs_for_styles[style])
                            for style in config.keys()]
    run_difference = [compare_runs(config[style], properties, paragraphs_for_styles[style])
                      for style in config.keys()]

    return dict(zip(config.keys(), [paragraph_difference, run_difference]))


def print_difference(difference, paragraphs_for_styles):
    for style in paragraphs_for_styles.keys():
        paragraph = dict(zip(paragraphs_for_styles[style], difference[style][0]))
        font = dict(zip(paragraphs_for_styles[style], difference[style][1]))

        for key, value in paragraph.items():
            print(key + 1, ' : ', paragraph[key]) if paragraph[key] != [[]] else print(key + 1, ' : ', 'ok')

        for key, value in font.items():
            print(key + 1, ' : ', font[key]) if font[key] != [[]] else print(key + 1, ' : ', 'ok')


def run():
    write_xml()
    convert_docx_to_html()
    config = read_config()
    paragraphs_for_styles = get_paragraphs_for_styles(config)
    difference = compare_styles(config, paragraphs_for_styles)
    print_difference(difference, paragraphs_for_styles)
