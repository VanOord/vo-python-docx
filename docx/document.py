# encoding: utf-8

"""|Document| and closely related objects"""

from __future__ import absolute_import, division, print_function, unicode_literals

from docx.blkcntnr import BlockItemContainer
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_BREAK
from docx.section import Section, Sections
from docx.shared import ElementProxy, Emu
from .oxml.shared import qn
from .oxml import CT_P
from docx.oxml import OxmlElement

class Document(ElementProxy):
    """WordprocessingML (WML) document.

    Not intended to be constructed directly. Use :func:`docx.Document` to open or create
    a document.
    """

    __slots__ = ('_part', '__body')

    def __init__(self, element, part):
        super(Document, self).__init__(element)
        self._part = part
        self.__body = None

    def add_heading(self, text="", level=1):
        """Return a heading paragraph newly added to the end of the document.

        The heading paragraph will contain *text* and have its paragraph style
        determined by *level*. If *level* is 0, the style is set to `Title`. If *level*
        is 1 (or omitted), `Heading 1` is used. Otherwise the style is set to `Heading
        {level}`. Raises |ValueError| if *level* is outside the range 0-9.
        """
        if not 0 <= level <= 9:
            raise ValueError("level must be in range 0-9, got %d" % level)
        style = "Title" if level == 0 else "Heading %d" % level
        return self.add_paragraph(text, style)

    def add_page_break(self):
        """Return newly |Paragraph| object containing only a page break."""
        paragraph = self.add_paragraph()
        paragraph.add_run().add_break(WD_BREAK.PAGE)
        return paragraph

    def add_paragraph(self, text='', style=None):
        """
        Return a paragraph newly added to the end of the document, populated
        with *text* and having paragraph style *style*. *text* can contain
        tab (``\\t``) characters, which are converted to the appropriate XML
        form for a tab. *text* can also include newline (``\\n``) or carriage
        return (``\\r``) characters, each of which is converted to a line
        break.
        """
        return self._body.add_paragraph(text, style)

    def add_citation(self, bookmark, par=None):
        if par is not None:
            paragraph = par
        else:
            paragraph = self.add_paragraph()
        run = paragraph.add_run()
        fldChar = OxmlElement('w:fldChar')  # creates a new element
        fldChar.set(qn('w:fldCharType'), 'begin')  # sets attribute on element

        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')  # sets attribute on element
        instrText.text = ' CITATION {:s} \l 1033'.format(bookmark)

        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'separate')
        fldChar3 = OxmlElement('w:t')
        fldChar3.text = "Right-click to update field."
        fldChar2.append(fldChar3)

        fldChar4 = OxmlElement('w:fldChar')
        fldChar4.set(qn('w:fldCharType'), 'end')

        r_element = run._r
        r_element.append(fldChar)
        r_element.append(instrText)
        r_element.append(fldChar2)
        r_element.append(fldChar4)
        p_element = paragraph._p


    def add_crossreference(self, bookmark, par=None, obj_type='figure'):
#        print('input', bookmark)
#        print('identifier', bookmark[:3])
#
        if bookmark[:3] not in ['fig', 'tab', 'equ']:
            str_ = '{:s} is an incorrect bookmark, bookmark should start with fig, tab or equ'.format(bookmark)
            raise ValueError(str_)
        if par is not None:
            paragraph = par
        else:
            paragraph = self.add_paragraph()

        fldChar = OxmlElement('w:fldChar')  # creates a new element
        fldChar.set(qn('w:fldCharType'), 'begin')  # sets attribute on element

        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')
        # sets attribute on element
        instrText.text = ' REF _Ref{:s} \# 0 \h'.format(bookmark)
#        instrText.text = ' REF {:s} \h'.format(bookmark)
        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'separate')
        fldChar3 = OxmlElement('w:t')
        fldChar3.text = "1"
#        instrTextFig.append(fldChar3)

        fldChar4 = OxmlElement('w:fldChar')
        fldChar4.set(qn('w:fldCharType'), 'end')


        run = paragraph.add_run()
        r_element = run._r
        fig = OxmlElement('w:t')
        fig.set(qn('xml:space'), 'preserve')

        if bookmark[:3] == 'fig':
            fig.text = 'figure '
        if bookmark[:3] == 'tab':
            fig.text = 'table '
        if bookmark[:3] == 'equ':
            fig.text = 'equation '
        r_element.append(fig)

        run = paragraph.add_run()
        r_element = run._r
        r_element.append(fldChar)

        run = paragraph.add_run()
        r_element = run._r
        r_element.append(instrText)

        run = paragraph.add_run()
        r_element = run._r
        r_element.append(fldChar2)

        run = paragraph.add_run()
        r_element = run._r
        r_element.append(fldChar3)

        run = paragraph.add_run()
        r_element = run._r
        r_element.append(fldChar4)
        p_element = paragraph._p

    def create_caption(self, entry, paragraph, obj_type, bmark='1'):
        obj_type = obj_type.lower()
        run = paragraph.add_run()
        r = run._r

        par = paragraph._p
        run = paragraph.add_run()
        r = run._r

        bmrk = OxmlElement('w:bookmarkStart')
        bmrk.set(qn('w:id'), '1')
        bmrk.set(qn('w:name'), '_Ref{:s}'.format(bmark))
        par.append(bmrk)

        fig = OxmlElement('w:t')
        fig.set(qn('xml:space'), 'preserve')
        if obj_type == 'figure':
            fig.text = 'Figure '
        if obj_type == 'table':
            fig.text = 'Table '
        r.append(fig)

        fldChar_1 = OxmlElement('w:fldSimple')
        fldChar_1.set(qn('xml:space'), 'preserve')
        if obj_type == 'figure':
            fldChar_1.set(qn('w:instr'), ' SEQ Figure \* ARABIC ')
        if obj_type == 'table':
            fldChar_1.set(qn('w:instr'), ' SEQ Table \* ARABIC ')
        par.append(fldChar_1)

        run = paragraph.add_run()
        r = run._r
        fldChar_2 = OxmlElement('w:fldChar')
        fldChar_2.set(qn('w:fldCharType'), "end")
        r.append(fldChar_2)

        bmrk_1 = OxmlElement('w:bookmarkEnd')
        bmrk_1.set(qn('w:id'), '1')
        par.append(bmrk_1)

        run = paragraph.add_run()
        caption = OxmlElement('w:t')
        caption.set(qn('xml:space'), 'preserve')
        caption.text = ' {:s}'.format(entry)
        run._r.append(caption)

    def add_caption(self, caption, bmark, obj_type='figure'):
        """

        """
        par = self.add_paragraph(style='caption')
        self.create_caption(caption, par,  obj_type, bmark)
        return par

    def add_picture(self, image_path_or_stream, width=None, height=None, style=None):
        """
        Return a new picture shape added in its own paragraph at the end of
        the document. The picture contains the image at
        *image_path_or_stream*, scaled based on *width* and *height*. If
        neither width nor height is specified, the picture appears at its
        native size. If only one is specified, it is used to compute
        a scaling factor that is then applied to the unspecified dimension,
        preserving the aspect ratio of the image. The native size of the
        picture is calculated using the dots-per-inch (dpi) value specified
        in the image file, defaulting to 72 dpi if no value is specified, as
        is often the case.
        """
        if style is not None:
            run = self.add_paragraph(style=style).add_run()
        else:
            run = self.add_paragraph().add_run()
        return run.add_picture(image_path_or_stream, width, height)

    def add_section(self, start_type=WD_SECTION.NEW_PAGE):
        """
        Return a |Section| object representing a new section added at the end
        of the document. The optional *start_type* argument must be a member
        of the :ref:`WdSectionStart` enumeration, and defaults to
        ``WD_SECTION.NEW_PAGE`` if not provided.
        """
        new_sectPr = self._element.body.add_section_break()
        new_sectPr.start_type = start_type
        return Section(new_sectPr, self._part)

    def add_table(self, rows, cols, style=None):
        """
        Add a table having row and column counts of *rows* and *cols*
        respectively and table style of *style*. *style* may be a paragraph
        style object or a paragraph style name. If *style* is |None|, the
        table inherits the default table style of the document.
        """
        table = self._body.add_table(rows, cols, self._block_width)
        table.style = style
        return table

    def bookmark_text(self, bookmark_name, text, underline = False, italic = False, bold = False, style = None):
        doc_element = self._part._element
        bookmarks_list = doc_element.findall('.//' + qn('w:bookmarkStart'))
        bookmarks_list = doc_element.findall('.//' + qn('wp:docPr'))
        for bookmark in bookmarks_list:
            name = bookmark.get(qn('wp:name'))
            print('asdf?', bookmark.name)
            if  bookmark.name == bookmark_name:
                par = bookmark.getparent()

#                if not isinstance(par, CT_P):
#                    return False
#                else:
                print('test')
                i = par.index(bookmark) + 1
                print(i)
                p = self.add_paragraph()
                run = p.add_run(text, style)
                run.underline = underline
                run.italic = italic
                run.bold = bold
                par.insert(i, run._element)
                p = p._element
                p.getparent().remove(p)
                p._p = p._element = None
                return True
        return False

    def bookmark_table(self, bookmark_name, rows, cols, style=None):
        tb = self.add_table(rows=rows, cols=cols, style=style)
        doc_element = self._part._element
        bookmarks_list = doc_element.findall('.//' + qn('w:bookmarkStart'))
        for bookmark in bookmarks_list:
            name = bookmark.get(qn('w:name'))
            if name == bookmark_name:
                par = bookmark.getparent()
                if not isinstance(par, CT_P):
                    return False
                else:
                    i = par.index(bookmark) + 1
                    par.addnext(tb._element)
                    return tb
        return tb

    def bookmark_picture(self, bookmark_name, picture, width=None, height=None, caption='', bmark=''):
        doc_element = self._part._element
        bookmarks_list = doc_element.findall('.//' + qn('w:bookmarkStart'))
        for bookmark in bookmarks_list:
            name = bookmark.get(qn('w:name'))
            if name == bookmark_name:
                par = bookmark.getparent()
                if not isinstance(par, CT_P):
                    return False
                else:
                    i = par.index(bookmark) + 1
                    p = self.add_paragraph()
                    run = p.add_run()
                    run.add_picture(picture, width, height)

#                    run.add_caption(caption, bmark)
                    par.insert(i, run._element)
                    p = p._element
                    p.getparent().remove(p)
                    p._p = None
                    p._element = None
                    return True

    def delete_paragraph(self, paragraph):
        p = paragraph._element
        p.getparent().remove(p)
        p._p = p._element = None


    def add_table_of_contents(self):
        paragraph = self.add_paragraph()
        run = paragraph.add_run()
        fldChar = OxmlElement('w:fldChar')  # creates a new element
        fldChar.set(qn('w:fldCharType'), 'begin')  # sets attribute on element
        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')  # sets attribute on element
        instrText.text = r'TOC \o "1-3" \h \z \u'   # change 1-3 depending on heading levels you need

        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'separate')
        fldChar3 = OxmlElement('w:t')
        fldChar3.text = "Right-click to update field."
        fldChar2.append(fldChar3)

        fldChar4 = OxmlElement('w:fldChar')
        fldChar4.set(qn('w:fldCharType'), 'end')

        r_element = run._r
        r_element.append(fldChar)
        r_element.append(instrText)
        r_element.append(fldChar2)
        r_element.append(fldChar4)
        p_element = paragraph._p

    def add_bibliography(self):
        paragraph = self.add_paragraph()
        run = paragraph.add_run()
        fldChar = OxmlElement('w:fldChar')  # creates a new element
        fldChar.set(qn('w:fldCharType'), 'begin')  # sets attribute on element
        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')  # sets attribute on element
        instrText.text = r'BIBLIOGRAPHY'   # change 1-3 depending on heading levels you need

        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'separate')
        fldChar3 = OxmlElement('w:t')
        fldChar3.text = "Right-click to update field."
        fldChar2.append(fldChar3)

        fldChar4 = OxmlElement('w:fldChar')
        fldChar4.set(qn('w:fldCharType'), 'end')

        r_element = run._r
        r_element.append(fldChar)
        r_element.append(instrText)
        r_element.append(fldChar2)
        r_element.append(fldChar4)
        p_element = paragraph._p

    def remove_picture(self, bookmark_name):
        doc_element = self._part._element
        bookmarks_list = doc_element.findall('.//' + qn('w:bookmarkStart'))
        for bookmark in bookmarks_list:
            name = bookmark.get(qn('w:name'))
            if name == bookmark_name:
                par = bookmark.getparent()


#                print(doc_element)
                if not isinstance(par, CT_P):
                    return False
                else:
                    i = par.index(bookmark) + 1
                    p = self.add_paragraph()
                    run = p.add_run()
#                    run.add_picture(picture, width, height)

#                    run.add_caption(caption, bmark)
                    par.insert(i, run._element)
                    p = p._element
                    p.getparent().remove(p)
                    p._p = None
                    p._element = None
                    return True
#    def bookmark_picture(self, bookmark_name, picture):
#        doc_element = self._part._element
#        bookmarks_list = doc_element.findall('.//' + qn('wp:docPr'))
#        caption_list = doc_element.findall('.//' + qn('w:fldSimple'))
##        print(caption_list)
#        for bookmark in bookmarks_list:
#            name = bookmark.get(qn('wp:name'))
##            print('bmark_id', bookmark.id)
##            print('bmark_name', bookmark.name)
#            print(bookmark.name, bookmark_name)
#            if bookmark.name == bookmark_name:
#                print(bookmark.name)
#                par = bookmark.getparent()
#                print(par)
#
#                test = _Body(par, CT_P())
#                #test.clear_content()
#
#                run = test.add_paragraph().add_run()
#
#
#                run.add_picture(picture)
#
##                if not isinstance(par, CT_P):
##                    return False
##                else:
#                print('Doe iets')
#
#                return True
#

    @property
    def core_properties(self):
        """
        A |CoreProperties| object providing read/write access to the core
        properties of this document.
        """
        return self._part.core_properties

    @property
    def inline_shapes(self):
        """
        An |InlineShapes| object providing access to the inline shapes in
        this document. An inline shape is a graphical object, such as
        a picture, contained in a run of text and behaving like a character
        glyph, being flowed like other text in a paragraph.
        """
        return self._part.inline_shapes

    @property
    def paragraphs(self):
        """
        A list of |Paragraph| instances corresponding to the paragraphs in
        the document, in document order. Note that paragraphs within revision
        marks such as ``<w:ins>`` or ``<w:del>`` do not appear in this list.
        """
        return self._body.paragraphs

    @property
    def part(self):
        """
        The |DocumentPart| object of this document.
        """
        return self._part

    def save(self, path_or_stream):
        """
        Save this document to *path_or_stream*, which can be either a path to
        a filesystem location (a string) or a file-like object.
        """
        self._part.save(path_or_stream)

    @property
    def sections(self):
        """|Sections| object providing access to each section in this document."""
        return Sections(self._element, self._part)

    @property
    def settings(self):
        """
        A |Settings| object providing access to the document-level settings
        for this document.
        """
        return self._part.settings

    @property
    def styles(self):
        """
        A |Styles| object providing access to the styles in this document.
        """
        return self._part.styles

    @property
    def tables(self):
        """
        A list of |Table| instances corresponding to the tables in the
        document, in document order. Note that only tables appearing at the
        top level of the document appear in this list; a table nested inside
        a table cell does not appear. A table within revision marks such as
        ``<w:ins>`` or ``<w:del>`` will also not appear in the list.
        """
        return self._body.tables

    @property
    def _block_width(self):
        """
        Return a |Length| object specifying the width of available "writing"
        space between the margins of the last section of this document.
        """
        section = self.sections[-1]
        return Emu(
            section.page_width - section.left_margin - section.right_margin
        )

    @property
    def _body(self):
        """
        The |_Body| instance containing the content for this document.
        """
        if self.__body is None:
            self.__body = _Body(self._element.body, self)
        return self.__body


class _Body(BlockItemContainer):
    """
    Proxy for ``<w:body>`` element in this document, having primarily a
    container role.
    """
    def __init__(self, body_elm, parent):
        super(_Body, self).__init__(body_elm, parent)
        self._body = body_elm

    def clear_content(self):
        """
        Return this |_Body| instance after clearing it of all content.
        Section properties for the main document story, if present, are
        preserved.
        """
        self._body.clear_content()
        return self
