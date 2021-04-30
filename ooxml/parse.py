# -*- coding: utf-8 -*-

"""Parse OOXML structure.

.. moduleauthor:: Aleksandar Erkalovic <aerkalov@gmail.com>

"""

import logging
from lxml import etree
from . import doc, NAMESPACES

logger = logging.getLogger('ooxml')


def _name(name):
    """Returns full name for the attribute.

    It checks predefined namespaces used in OOXML documents.

    > _name('{{{w}}}rStyle')
    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rStyle'
    """
    return name.format(**NAMESPACES)


def is_on(value):
    return value in ['true', 'on', '1']


def return_underline_value(value):
    return value in ['single', 'words', 'double', 'thick', 'dotted',
                    'dottedHeavy', 'dash', 'dashedHeavy', 'dashLong',
                    'dashLongHeavy', 'dotDash', 'dashDotHeavy', 'dotDotDash',
                    'dashDotDotHeavy', 'wave', 'wavyHeavy', 'wavyDouble']


def parse_run_properties(document, paragraph, properties):
    if not paragraph:
        return

    run_style = properties.find(_name('{{{w}}}rStyle'))

    if run_style is not None:
        paragraph.run_properties['style'] = run_style.attrib[_name('{{{w}}}val')]
        document.add_style_as_used(paragraph.run_properties['style'])

    text_color = properties.find(_name('{{{w}}}color'))

    if text_color is not None:
        paragraph.run_properties['color'] = text_color.attrib[_name('{{{w}}}val')]

    right_to_left = properties.find(_name('{{{w}}}rtl'))

    if right_to_left is not None:
        if is_on(right_to_left.attrib.get(_name('{{{w}}}val'), 'on')):
            paragraph.run_properties['rtl'] = True

    font_family = properties.find(_name('{{{w}}}rFonts'))

    if font_family is not None:
        paragraph.run_properties['font_family'] = font_family.attrib[_name('{{{w}}}ascii')]

    text_size = properties.find(_name('{{{w}}}sz'))

    if text_size is not None:
        paragraph.run_properties['sz'] = text_size.attrib[_name('{{{w}}}val')]

        if isinstance(paragraph, doc.Text):
            if not ('dropcap' in paragraph.paragraph_properties and paragraph.paragraph_properties['dropcap']):
                if paragraph.parent and hasattr(paragraph.parent, 'ppr') and (
                        not ('dropcap' in paragraph.parent.paragraph_properties and
                             paragraph.parent.paragraph_properties['dropcap'])):
                    document.add_font_as_used(paragraph.run_properties['sz'])
        elif isinstance(paragraph, doc.Paragraph):
            if not ('dropcap' in paragraph.paragraph_properties and paragraph.paragraph_properties['dropcap']):
                document.add_font_as_used(paragraph.run_properties['sz'])
        else:
            document.add_font_as_used(paragraph.run_properties['sz'])

    bold = properties.find(_name('{{{w}}}b'))

    if bold is not None:
        if is_on(bold.attrib.get(_name('{{{w}}}val'), 'on')):
            paragraph.run_properties['b'] = True

    italic = properties.find(_name('{{{w}}}i'))

    if italic is not None:
        if is_on(italic.attrib.get(_name('{{{w}}}val'), 'on')):
            paragraph.run_properties['i'] = True

    underline = properties.find(_name('{{{w}}}u'))

    if underline is not None:
        value = underline.attrib[_name('{{{w}}}val')]
        if return_underline_value(value):
            paragraph.run_properties['u'] = value

    strikethrough = properties.find(_name('{{{w}}}strike'))

    if strikethrough is not None:
        # todo
        # check b = on and not off
        if is_on(strikethrough.attrib.get(_name('{{{w}}}val'), 'on')):
            paragraph.run_properties['strike'] = True

    vertical_align = properties.find(_name('{{{w}}}vertAlign'))

    if vertical_align is not None:
        paragraph.run_properties['vertical_align'] = vertical_align.attrib[_name('{{{w}}}val')]
        # value = vertical_align.attrib[_name('{{{w}}}val')]
        #
        # if return_underline_value(value):
        #     paragraph.run_properties['u'] = value
        #
        # if value == 'superscript':
        #     paragraph.run_properties['superscript'] = True
        #
        # if value == 'subscript':
        #     paragraph.run_properties['subscript'] = True

    small_caps = properties.find(_name('{{{w}}}smallCaps'))

    if small_caps is not None:
        if is_on(small_caps.attrib.get(_name('{{{w}}}val'), 'on')):
            paragraph.run_properties['small_caps'] = True

    highlight = properties.find(_name('{{{w}}}highlight'))

    if highlight is not None:
        paragraph.run_properties['highlight'] = highlight.attrib[_name('{{{w}}}val')]


def parse_paragraph_properties(document, paragraph, properties):
    if not paragraph:
        return

    paragraph_style = properties.find(_name('{{{w}}}pStyle'))

    if paragraph_style is not None:
        paragraph.style_id = paragraph_style.attrib[_name('{{{w}}}val')]
        document.add_style_as_used(paragraph.style_id)

    numpr = properties.find(_name('{{{w}}}numPr'))

    if numpr is not None:
        ilvl = numpr.find(_name('{{{w}}}ilvl'))

        if ilvl is not None:
            paragraph.ilvl = int(ilvl.attrib[_name('{{{w}}}val')])

        numid = numpr.find(_name('{{{w}}}numId'))

        if numid is not None:
            paragraph.numid = int(numid.attrib[_name('{{{w}}}val')])

    alignment = properties.find(_name('{{{w}}}jc'))

    if alignment is not None:
        paragraph.paragraph_properties['jc'] = alignment.attrib[_name('{{{w}}}val')]

    spacing = properties.find(_name('{{{w}}}spacing'))

    if spacing is not None:
        paragraph.paragraph_properties['spacing'] = {}

        if _name('{{{w}}}line') in spacing.attrib:
            paragraph.paragraph_properties['spacing']['line'] = spacing.attrib[_name('{{{w}}}line')]

        if _name('{{{w}}}before') in spacing.attrib:
            paragraph.paragraph_properties['spacing']['before'] = spacing.attrib[_name('{{{w}}}before')]

        if _name('{{{w}}}after') in spacing.attrib:
            paragraph.paragraph_properties['spacing']['after'] = spacing.attrib[_name('{{{w}}}after')]

    indentation = properties.find(_name('{{{w}}}ind'))

    if indentation is not None:
        paragraph.paragraph_properties['ind'] = {}

        if _name('{{{w}}}left') in indentation.attrib:
            paragraph.paragraph_properties['ind']['left'] = indentation.attrib[_name('{{{w}}}left')]

        if _name('{{{w}}}right') in indentation.attrib:
            paragraph.paragraph_properties['ind']['right'] = indentation.attrib[_name('{{{w}}}right')]

        if _name('{{{w}}}firstLine') in indentation.attrib:
            paragraph.paragraph_properties['ind']['first_line'] = indentation.attrib[_name('{{{w}}}firstLine')]

    background_color = properties.find(_name('{{{w}}}shd'))

    if background_color is not None:
        paragraph.paragraph_properties['shd'] = background_color.attrib[_name('{{{w}}}fill')]

    frame_pr = properties.find(_name('{{{w}}}framePr'))

    if frame_pr is not None:
        if _name('{{{w}}}dropCap') in frame_pr.attrib:
            drop_cap = frame_pr.attrib[_name('{{{w}}}dropCap')]

            if drop_cap.lower() in ['drop', 'margin']:
                paragraph.paragraph_properties['dropcap'] = True

    # Do not show run properties in paragraph
    # run_properties = properties.find(_name('{{{w}}}rPr'))
    #
    # if run_properties is not None:
    #     parse_run_properties(document, paragraph, run_properties)


def parse_drawing(document, container, elem):
    """Parse drawing element.

    We don't do much with drawing element. We can find embeded image but we don't do more than that.
    """

    _blip = elem.xpath('.//a:blip', namespaces=NAMESPACES)

    if len(_blip) > 0:
        blip = _blip[0]
        _rid = blip.attrib[_name('{{{r}}}embed')]

        img = doc.Image(_rid)
        container.elements.append(img)


def parse_footnote(document, container, elem):
    """Parse the footnote element."""

    _rid = elem.attrib[_name('{{{w}}}id')]
    foot = doc.Footnote(_rid)
    container.elements.append(foot)


def parse_endnote(document, container, elem):
    """Parse the endnote element."""

    _rid = elem.attrib[_name('{{{w}}}id')]
    note = doc.Endnote(_rid)
    container.elements.append(note)


def parse_alternate(document, container, elem):
    txtbx = elem.find('.//' + _name('{{{w}}}txbxContent'))
    paragraphs = []

    if txtbx is None:
        return

    for el in txtbx:
        if el.tag == _name('{{{w}}}p'):
            paragraphs.append(parse_paragraph(document, el))

    textbox = doc.TextBox(paragraphs)
    container.elements.append(textbox)


def parse_text(document, container, element):
    """Parse text element."""

    txt = None

    alternate = element.find(_name('{{{mc}}}AlternateContent'))

    if alternate is not None:
        parse_alternate(document, container, alternate)

    br = element.find(_name('{{{w}}}br'))

    if br is not None:
        if _name('{{{w}}}type') in br.attrib:
            _type = br.attrib[_name('{{{w}}}type')]
            brk = doc.Break(_type)
        else:
            brk = doc.Break()

        container.elements.append(brk)

    t = element.find(_name('{{{w}}}t'))

    if t is not None:
        txt = doc.Text(t.text)
        txt.parent = container

        container.elements.append(txt)

    rpr = element.find(_name('{{{w}}}rPr'))

    if rpr is not None:
        # Notice it is using txt as container
        parse_run_properties(document, txt, rpr)

    for r in element.findall(_name('{{{w}}}r')):
        parse_text(document, container, r)

    foot = element.find(_name('{{{w}}}footnoteReference'))

    if foot is not None:
        parse_footnote(document, container, foot)

    end = element.find(_name('{{{w}}}endnoteReference'))

    if end is not None:
        parse_endnote(document, container, end)

    sym = element.find(_name('{{{w}}}sym'))

    if sym is not None:
        _font = sym.attrib[_name('{{{w}}}font')]
        _char = sym.attrib[_name('{{{w}}}char')]

        container.elements.append(doc.Symbol(font=_font, character=_char))

    image = element.find(_name('{{{w}}}drawing'))

    if image is not None:
        parse_drawing(document, container, image)

    refe = element.find(_name('{{{w}}}commentReference'))

    if refe is not None:
        _m = doc.Comment(refe.attrib[_name('{{{w}}}id')], 'reference')
        container.elements.append(_m)

    return


def parse_smarttag(document, container, tag_elem):
    """Parse the endnote element."""

    tag = doc.SmartTag()

    tag.element = tag_elem.attrib[_name('{{{w}}}element')]

    for elem in tag_elem:
        if elem.tag == _name('{{{w}}}r'):
            parse_text(document, tag, elem)

        if elem.tag == _name('{{{w}}}smartTag'):
            parse_smarttag(document, tag, elem)

    container.elements.append(tag)

    return


def parse_paragraph(document, par):
    """Parse paragraph element.

    Some other elements could be found inside of paragraph element (math, links).
    """

    paragraph = doc.Paragraph()
    paragraph.document = document

    for elem in par:
        if elem.tag == _name('{{{w}}}pPr'):
            parse_paragraph_properties(document, paragraph, elem)

        if elem.tag == _name('{{{w}}}r'):
            parse_text(document, paragraph, elem)

        if elem.tag == _name('{{{m}}}oMath'):
            _m = doc.Math()
            paragraph.elements.append(_m)

        if elem.tag == _name('{{{m}}}oMathPara'):
            _m = doc.Math()
            paragraph.elements.append(_m)

        if elem.tag == _name('{{{w}}}commentRangeStart'):
            _m = doc.Comment(elem.attrib[_name('{{{w}}}id')], 'start')
            paragraph.elements.append(_m)

        if elem.tag == _name('{{{w}}}commentRangeEnd'):
            _m = doc.Comment(elem.attrib[_name('{{{w}}}id')], 'end')
            paragraph.elements.append(_m)

        if elem.tag == _name('{{{w}}}hyperlink'):
            try:
                t = doc.Link(elem.attrib[_name('{{{r}}}id')])

                parse_text(document, t, elem)

                paragraph.elements.append(t)
            except:
                logger.error('Error with with hyperlink [%s].', str(elem.attrib.items()))

        if elem.tag == _name('{{{w}}}smartTag'):
            parse_smarttag(document, paragraph, elem)

    return paragraph


def parse_table_properties(doc, table, prop):
    """Parse table properties."""

    if not table:
        return

    style = prop.find(_name('{{{w}}}tblStyle'))

    if style is not None:
        table.style_id = style.attrib[_name('{{{w}}}val')]
        doc.add_style_as_used(table.style_id)


def parse_table_column_properties(doc, cell, prop):
    """Parse table column properties."""

    if not cell:
        return

    grid = prop.find(_name('{{{w}}}gridSpan'))

    if grid is not None:
        cell.grid_span = int(grid.attrib[_name('{{{w}}}val')])

    vmerge = prop.find(_name('{{{w}}}vMerge'))

    if vmerge is not None:
        if _name('{{{w}}}val') in vmerge.attrib:
            cell.vmerge = vmerge.attrib[_name('{{{w}}}val')]
        else:
            cell.vmerge = ""


def parse_table(document, tbl):
    """Parse table element."""

    def _change(rows, pos_x):
        if len(rows) == 1:
            return rows

        count_x = 1

        for x in rows[-1]:
            if count_x == pos_x:
                x.row_span += 1

            count_x += x.grid_span

        return rows

    table = doc.Table()

    tbl_pr = tbl.find(_name('{{{w}}}tblPr'))

    if tbl_pr is not None:
        parse_table_properties(document, table, tbl_pr)

    for tr in tbl.xpath('./w:tr', namespaces=NAMESPACES):
        columns = []
        pos_x = 0

        for tc in tr.xpath('./w:tc', namespaces=NAMESPACES):
            cell = doc.TableCell()

            tc_pr = tc.find(_name('{{{w}}}tcPr'))

            if tc_pr is not None:
                parse_table_column_properties(doc, cell, tc_pr)

            # maybe after
            pos_x += cell.grid_span

            if cell.vmerge is not None and cell.vmerge == "":
                table.rows = _change(table.rows, pos_x)
            else:
                for p in tc.xpath('./w:p', namespaces=NAMESPACES):
                    cell.elements.append(parse_paragraph(document, p))

                columns.append(cell)

        table.rows.append(columns)

    return table


def parse_document(xmlcontent):
    """Parse document with content.

    Content is placed in file 'document.xml'.
    """

    document = etree.fromstring(xmlcontent)

    body = document.xpath('.//w:body', namespaces=NAMESPACES)[0]

    document = doc.Document()

    for elem in body:
        if elem.tag == _name('{{{w}}}p'):
            document.elements.append(parse_paragraph(document, elem))

        if elem.tag == _name('{{{w}}}tbl'):
            document.elements.append(parse_table(document, elem))

        if elem.tag == _name('{{{w}}}sdt'):
            document.elements.append(doc.TOC())

    return document


def parse_relationship(document, xmlcontent, rel_type):
    """Parse relationship document.

    Relationships hold information like external or internal references for links.

    Relationships are placed in file '_rels/document.xml.rels'.
    """

    doc = etree.fromstring(xmlcontent)

    for elem in doc:
        if elem.tag == _name('{{{pr}}}Relationship'):
            rel = {'target': elem.attrib['Target'],
                   'type': elem.attrib['Type'],
                   'target_mode': elem.attrib.get('TargetMode', 'Internal')}

            document.relationships[rel_type][elem.attrib['Id']] = rel


def parse_style(document, xmlcontent):
    """Parse styles document.

    Styles are defined in file 'styles.xml'.
    """

    styles = etree.fromstring(xmlcontent)

    _r = styles.xpath('.//w:rPrDefault', namespaces=NAMESPACES)

    if len(_r) > 0:
        rpr = _r[0].find(_name('{{{w}}}rPr'))

        if rpr is not None:
            st = doc.Style()
            parse_run_properties(document, st, rpr)
            document.default_style = st

    # rest of the styles
    for style in styles.xpath('.//w:style', namespaces=NAMESPACES):
        st = doc.Style()

        st.style_id = style.attrib[_name('{{{w}}}styleId')]

        style_type = style.attrib[_name('{{{w}}}type')]
        if style_type is not None:
            st.style_type = style_type

        if _name('{{{w}}}default') in style.attrib:
            is_default = style.attrib[_name('{{{w}}}default')]
            if is_default is not None:
                st.is_default = is_default == '1'

        name = style.find(_name('{{{w}}}name'))
        if name is not None:
            st.name = name.attrib[_name('{{{w}}}val')]

        based_on = style.find(_name('{{{w}}}basedOn'))

        if based_on is not None:
            st.based_on = based_on.attrib[_name('{{{w}}}val')]

        document.styles.styles[st.style_id] = st

        if st.is_default:
            document.styles.default_styles[st.style_type] = st.style_id

        rpr = style.find(_name('{{{w}}}rPr'))

        if rpr is not None:
            parse_run_properties(document, st, rpr)

        ppr = style.find(_name('{{{w}}}pPr'))

        if ppr is not None:
            parse_paragraph_properties(document, st, ppr)


def parse_comments(document, xmlcontent):
    """Parse comments document.

    Comments are defined in file 'comments.xml'
    """

    comments = etree.fromstring(xmlcontent)
    document.comments = {}

    for comment in comments.xpath('.//w:comment', namespaces=NAMESPACES):
        # w:author
        # w:id
        # w: date
        comment_id = comment.attrib[_name('{{{w}}}id')]

        comm = doc.CommentContent(comment_id)
        comm.author = comment.attrib.get(_name('{{{w}}}author'), None)
        comm.date = comment.attrib.get(_name('{{{w}}}date'), None)

        comm.elements = [parse_paragraph(document, para) for para in comment.xpath('.//w:p', namespaces=NAMESPACES)]

        document.comments[comment_id] = comm


def parse_footnotes(document, xmlcontent):
    """Parse footnotes document.

    Footnotes are defined in file 'footnotes.xml'
    """

    footnotes = etree.fromstring(xmlcontent)
    document.footnotes = {}

    for footnote in footnotes.xpath('.//w:footnote', namespaces=NAMESPACES):
        _type = footnote.attrib.get(_name('{{{w}}}type'), None)

        # don't know what to do with these now
        if _type in ['separator', 'continuationSeparator', 'continuationNotice']:
            continue

        paragraphs = [parse_paragraph(document, para) for para in footnote.xpath('.//w:p', namespaces=NAMESPACES)]

        document.footnotes[footnote.attrib[_name('{{{w}}}id')]] = paragraphs


def parse_endnotes(document, xmlcontent):
    """Parse endnotes document.

    Endnotes are defined in file 'endnotes.xml'
    """

    endnotes = etree.fromstring(xmlcontent)
    document.endnotes = {}

    for note in endnotes.xpath('.//w:endnote', namespaces=NAMESPACES):
        paragraphs = [parse_paragraph(document, para) for para in note.xpath('.//w:p', namespaces=NAMESPACES)]

        document.endnotes[note.attrib[_name('{{{w}}}id')]] = paragraphs


def parse_numbering(document, xmlcontent):
    """Parse numbering document.

    Numbering is defined in file 'numbering.xml'.
    """

    numbering = etree.fromstring(xmlcontent)

    document.abstruct_numbering = {}
    document.numbering = {}

    for abstruct_num in numbering.xpath('.//w:abstractNum', namespaces=NAMESPACES):
        numb = {}
        for lvl in abstruct_num.xpath('./w:lvl', namespaces=NAMESPACES):
            ilvl = int(lvl.attrib[_name('{{{w}}}ilvl')])

            fmt = lvl.find(_name('{{{w}}}numFmt'))
            numb[ilvl] = {'numFmt': fmt.attrib[_name('{{{w}}}val')]}

        document.abstruct_numbering[abstruct_num.attrib[_name('{{{w}}}abstractNumId')]] = numb

    for num in numbering.xpath('.//w:num', namespaces=NAMESPACES):
        num_id = num.attrib[_name('{{{w}}}numId')]

        abs_num = num.find(_name('{{{w}}}abstractNumId'))

        if abs_num is not None:
            number_id = abs_num.attrib[_name('{{{w}}}val')]
            document.numbering[int(num_id)] = number_id


def parse_from_file(file_object):
    """Parses existing OOXML file.

    :Args:
      - file_object (:class:`ooxml.docx.DOCXFile`): OOXML file object

    :Returns:
      Returns parsed document of type :class:`ooxml.doc.Document`
    """

    logger.info('Parsing %s file.', file_object.file_name)

    # Read the files
    doc_content = file_object.read_file('document.xml')

    # Parse the document
    document = parse_document(doc_content)

    try:
        style_content = file_object.read_file('styles.xml')
        parse_style(document, style_content)
    except KeyError:
        logger.warning('Could not read styles.')

    try:
        doc_rel_content = file_object.read_file('_rels/document.xml.rels')
        parse_relationship(document, doc_rel_content, 'document')
    except KeyError:
        logger.warning('Could not read document relationships.')

    try:
        doc_rel_content = file_object.read_file('_rels/endnotes.xml.rels')
        parse_relationship(document, doc_rel_content, 'endnotes')
    except KeyError:
        logger.warning('Could not read endnotes relationships.')

    try:
        doc_rel_content = file_object.read_file('_rels/footnotes.xml.rels')
        parse_relationship(document, doc_rel_content, 'footnotes')
    except KeyError:
        logger.warning('Could not read footnotes relationships.')

    try:
        comments_content = file_object.read_file('comments.xml')
        parse_comments(document, comments_content)
    except KeyError:
        logger.warning('Could not read comments.')

    try:
        footnotes_content = file_object.read_file('footnotes.xml')
        parse_footnotes(document, footnotes_content)
    except KeyError:
        logger.warning('Could not read footnotes.')

    try:
        endnotes_content = file_object.read_file('endnotes.xml')
        parse_endnotes(document, endnotes_content)
    except KeyError:
        logger.warning('Could not read endnotes.')

    try:
        numbering_content = file_object.read_file('numbering.xml')
        parse_numbering(document, numbering_content)
    except KeyError:
        logger.warning('Could not read numbering.')

    return document
