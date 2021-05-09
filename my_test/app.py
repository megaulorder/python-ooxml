import logging

from bs4 import BeautifulSoup
from configobj import ConfigObj
from dictdiffer import diff
from docx import Document

import ooxml
from ooxml import serialize

logging.basicConfig(filename='ooxml.log', level=logging.INFO)

docx_path = 'test.docx'
config_path = 'config.ini'
html_path = 'sample.html'


def flatten(l, ltypes=(list, tuple)):
    ltype = type(l)
    l = list(l)
    i = 0
    while i < len(l):
        while isinstance(l[i], ltypes):
            if not l[i]:
                l.pop(i)
                i -= 1
                break
            else:
                l[i:i + 1] = l[i]
        i += 1
    return ltype(l)


# only for debugging
def write_xml():
    document = Document(docx_path)
    document_xml = document.element.xml
    file = open('sample.xml', 'w')
    file.write(document_xml)


def convert_docx_to_html():
    file_name = docx_path

    dfile = ooxml.read_from_file(file_name)
    output = serialize.serialize(dfile.document).decode("utf-8")

    file = open('sample.html', 'w')
    file.write(output)


def get_data_from_html():
    PARAGRAPH = 'p'
    SPAN = 'span'
    LIST_ELEMENT = 'li'
    ORDERED_LIST = 'ol'
    UNORDERED_LIST = 'ul'

    file = open(html_path)
    soup = BeautifulSoup(file, 'html.parser')

    paragraph_properties = [dict([i.split(': ') for i in r['style'][:-1].split('; ')])
                            for r in soup.find_all(PARAGRAPH)]
    run_properties = [[dict([k.split(': ') for k in j['style'][:-1].split('; ')]) for j in r.find_all(SPAN)]
                      for r in soup.find_all(PARAGRAPH)]

    properties = list(zip(paragraph_properties, run_properties))

    paragraphs_text = [text.replace('\n', '')[:30] for text in soup.get_text().split('\n\n\n')[:-1] if text != '']
    paragraphs_parents = ['ol' if par.find_parent(ORDERED_LIST)
                          else 'ul' if par.find_parent(UNORDERED_LIST)
                          else False for par in soup.find_all(PARAGRAPH)]

    return properties, paragraphs_text, paragraphs_parents


def read_config():
    config = ConfigObj('config.ini')
    properties = dict()
    for section in config.sections:
        properties[section] = dict(config[section])

    return properties


def parse_config(config):
    style_names = [style for style in config.keys() if 'sub-' not in style.lower()]
    substyle_names = [style for style in config.keys() if 'sub-' in style.lower()]

    styles = {style: config[style] for style in style_names}
    substyles = {substyle.split('SUB-', 1)[1]: config[substyle] for substyle in substyle_names}

    lists = {style: styles[style]['LISTS'] for style in styles.keys()}

    return styles, substyles, lists


def get_paragraphs_count(style):
    return [int(i) - 1 for i in style['PARAGRAPH_COUNT']['paragraphs']]


def get_paragraphs_for_styles(config):
    return dict(zip(config.keys(), [get_paragraphs_count(style) for style in config.values()]))


def compare(style, properties):
    return list(diff(properties, style))


def compare_paragraphs(style, properties, paragraph_count):
    selected_paragraphs = [properties[i] for i in paragraph_count]
    paragraph_difference = [compare(style['PARAGRAPH'], i[0]) for i in selected_paragraphs]

    return paragraph_difference


def compare_runs(style, properties, paragraph_count):
    selected_paragraphs = [properties[i] for i in paragraph_count]
    run_differences = [[compare(style['FONT'], j) for j in i[1]] for i in selected_paragraphs]

    return run_differences


def compare_styles(styles, paragraphs_for_styles, properties):
    paragraph_difference = [compare_paragraphs(styles[style], properties, paragraphs_for_styles[style])
                            for style in styles.keys()]
    run_difference = [compare_runs(styles[style], properties, paragraphs_for_styles[style])
                      for style in styles.keys()]

    paragraph_difference_by_style = dict(zip(styles.keys(), paragraph_difference))
    run_difference_by_style = dict(zip(styles.keys(), run_difference))

    return paragraph_difference_by_style, run_difference_by_style


def compare_substyles(substyles, paragraphs_for_substyles, properties):
    run_difference = [compare_runs(substyles[substyle], properties, paragraphs_for_substyles[substyle])
                      for substyle in substyles.keys()]

    run_difference_by_substyle = dict(zip(substyles.keys(), run_difference))

    for substyle in run_difference_by_substyle.keys():
        run_difference_by_substyle[substyle] = dict(zip(paragraphs_for_substyles[substyle],
                                                        [any(
                                                            [not any(a in (['add', 'change']) for a in flatten(run)) for
                                                             run in paragraph])
                                                            for paragraph in run_difference_by_substyle[substyle]]))

    substyles_by_paragraph = dict.fromkeys(set(flatten([paragraphs_for_substyles[substyle]
                                                        for substyle in run_difference_by_substyle.keys()])), [])
    for i in substyles_by_paragraph.keys():
        substyles_by_paragraph[i] = [style for style in run_difference_by_substyle.keys()
                                     if i in list(run_difference_by_substyle[style].keys())
                                     and run_difference_by_substyle[style][i] is True]

    return run_difference_by_substyle, substyles_by_paragraph


def check_lists(lists, paragraphs_for_styles, paragraphs_parents):
    config_lists_in_paragraphs = [False for _ in range(len(paragraphs_parents))]
    for i in range(len(config_lists_in_paragraphs)):
        for style in lists.keys():
            if lists[style]:
                for par in paragraphs_for_styles[style]:
                    if lists[style].keys()[0] == 'ordered-list':
                        config_lists_in_paragraphs[par] = 'ol'
                    elif lists[style].keys()[0] == 'unordered-list':
                        config_lists_in_paragraphs[par] = 'ul'

    out = []
    for par in range(len(config_lists_in_paragraphs)):
        if config_lists_in_paragraphs[par] and paragraphs_parents[par]:
            if config_lists_in_paragraphs[par] == paragraphs_parents[par]:
                out.append('paragraph is a list element (ok)')
            else:
                out.append(f'change list type from {paragraphs_parents[par]} to {config_lists_in_paragraphs[par]}')
        elif config_lists_in_paragraphs[par] and not paragraphs_parents[par]:
            out.append('paragraph is not a list element (ok)')
        elif not config_lists_in_paragraphs[par] and paragraphs_parents[par]:
            out.append('remove list from paragraph')
        else:
            out.append(None)

    return out


def paragraph_diff_to_string(difference):
    out_string = ''
    if difference != 'paragraph properties ok':
        for d in difference:
            if d[0] == 'change':
                out_string += f'{d[0]} {d[1]} from {d[2][0]} to {d[2][1]}; '
            elif d[0] in (['add', 'remove']):
                out_string += ''.join([f'{d[0]} property {i[0]} with value {i[1]}; ' for i in d[2]])

    return out_string if out_string != '' else difference


def font_diff_to_string(difference, substyles=None, substyles_by_paragraph=None):
    out_string = ''

    if substyles:
        ignored_values = dict(zip(
            flatten([substyles[style]['FONT'].keys() for style in substyles.keys() if style in substyles_by_paragraph]),
            flatten(
                [substyles[style]['FONT'].values() for style in substyles.keys() if style in substyles_by_paragraph])))
        if difference != 'font properties ok':
            for run in difference:
                for d in run:
                    if d[0] == 'change':
                        if d[1] in ignored_values.keys() and d[2][0] == ignored_values[d[1]]:
                            continue
                        else:
                            out_string += f'{d[0]} {d[1]} from {d[2][0]} to {d[2][1]};\n '
                    elif d[0] in (['add', 'remove']):
                        out_string += ''.join([f'{d[0]} property {i[0]} with value {i[1]};\n'
                                               for i in d[2] if i[0] not in ignored_values.keys()])
        out_string = ' '.join(filter(None, list(set([s[1:] if s[0] == ' ' else s for s in out_string.splitlines()]))))

    else:
        if difference != 'font properties ok':
            for run in difference:
                for d in run:
                    if d[0] == 'change':
                        out_string += f'{d[0]} {d[1]} from {d[2][0]} to {d[2][1]};\n '
                    elif d[0] in (['add', 'remove']):
                        out_string += ''.join([f'{d[0]} property {i[0]} with value {i[1]};\n'
                                               for i in d[2]])
        out_string = ' '.join(
            filter(None, list(set([s[1:] if s[0] == ' ' else s for s in out_string.splitlines()]))))

    return out_string if out_string != '' else difference


def get_paragraph_difference_for_styles(difference_by_styles, paragraphs_for_styles):
    output = {}
    for style in paragraphs_for_styles.keys():
        paragraph = dict(zip(paragraphs_for_styles[style], difference_by_styles[0][style]))
        font = dict(zip(paragraphs_for_styles[style], difference_by_styles[1][style]))

        empty_values = [[], [[]], [[], [], []]]

        for key, value in paragraph.items():
            output[key + 1] = [paragraph[key]] if paragraph[key] not in empty_values else ['paragraph properties ok']

        for key, value in font.items():
            output[key + 1].append(
                list(filter(None, font[key])) if font[key] not in empty_values else 'font properties ok')

    return output


def print_result(paragraph_difference, text, lists, substyles=None, difference_by_substyles=None,
                 paragraphs_for_substyles=None, substyles_by_paragraphs=None):
    print('Checking paragraphs...')

    if substyles:
        for key in sorted(paragraph_difference.keys()):
            print(f'\nParagraph #{key} {text[key - 1]}...')
            print(f'\t{paragraph_diff_to_string(paragraph_difference[key][0])}'
                  f'\n\t{font_diff_to_string(paragraph_difference[key][1], substyles, substyles_by_paragraphs[key - 1])}'
                  f'\n\tSubstyles {", ".join(["SUB-" + i for i in substyles_by_paragraphs[key - 1]])} are used')  # debug
            if lists[key - 1]:
                print(f'\t{lists[key - 1]}')

        for substyle in paragraphs_for_substyles.keys():
            font = dict(zip(paragraphs_for_substyles[substyle], difference_by_substyles[substyle].values()))
            for key, value in font.items():
                print(f'\nSubstyle SUB-{substyle} is not in use in paragraph #{key + 1}') if value is False else None

    else:
        for key in sorted(paragraph_difference.keys()):
            print(f'\nParagraph #{key} {text[key - 1]}...')
            print(f'\t{paragraph_diff_to_string(paragraph_difference[key][0])}'
                  f'\n\t{font_diff_to_string(paragraph_difference[key][1])}')
            if lists[key-1]:
                print(f'\t{lists[key - 1]}')


def run():
    write_xml()

    convert_docx_to_html()
    properties, text, paragraphs_parents = get_data_from_html()

    config = read_config()
    styles, substyles, lists = parse_config(config)

    paragraphs_for_styles = get_paragraphs_for_styles(styles)
    difference_by_styles = compare_styles(styles, paragraphs_for_styles, properties)

    lists_difference = check_lists(lists, paragraphs_for_styles, paragraphs_parents)

    if substyles:
        paragraphs_for_substyles = get_paragraphs_for_styles(substyles)
        difference_by_substyles, substyles_by_paragraphs = compare_substyles(substyles, paragraphs_for_substyles,
                                                                             properties)
        print_result(get_paragraph_difference_for_styles(difference_by_styles, paragraphs_for_styles),
                     text, lists_difference, substyles, difference_by_substyles, paragraphs_for_substyles,
                     substyles_by_paragraphs)
    else:
        print_result(get_paragraph_difference_for_styles(difference_by_styles, paragraphs_for_styles),
                     text, lists_difference)
