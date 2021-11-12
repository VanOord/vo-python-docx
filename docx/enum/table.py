# encoding: utf-8

"""
Enumerations related to tables in WordprocessingML files
"""

from __future__ import absolute_import, division, print_function, unicode_literals

from .base import Enumeration, EnumMember, XmlEnumeration, XmlMappedEnumMember, alias


@alias("WD_ALIGN_VERTICAL")
class WD_CELL_VERTICAL_ALIGNMENT(XmlEnumeration):
    """
    alias: **WD_ALIGN_VERTICAL**

    Specifies the vertical alignment of text in one or more cells of a table.

    Example::

        from docx.enum.table import WD_ALIGN_VERTICAL

        table = document.add_table(3, 3)
        table.cell(0, 0).vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM
    """

    __ms_name__ = "WdCellVerticalAlignment"

    __url__ = "https://msdn.microsoft.com/en-us/library/office/ff193345.aspx"

    __members__ = (
        XmlMappedEnumMember(
            "TOP", 0, "top", "Text is aligned to the top border of the cell."
        ),
        XmlMappedEnumMember(
            "CENTER", 1, "center", "Text is aligned to the center of the cel" "l."
        ),
        XmlMappedEnumMember(
            "BOTTOM",
            3,
            "bottom",
            "Text is aligned to the bottom border of " "the cell.",
        ),
        XmlMappedEnumMember(
            "BOTH",
            101,
            "both",
            "This is an option in the OpenXml spec, but"
            " not in Word itself. It's not clear what Word behavior this se"
            "tting produces. If you find out please let us know and we'll u"
            "pdate this documentation. Otherwise, probably best to avoid thi"
            "s option.",
        ),
    )


@alias("WD_ROW_HEIGHT")
class WD_ROW_HEIGHT_RULE(XmlEnumeration):
    """
    alias: **WD_ROW_HEIGHT**

    Specifies the rule for determining the height of a table row

    Example::

        from docx.enum.table import WD_ROW_HEIGHT_RULE

        table = document.add_table(3, 3)
        table.rows[0].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
    """

    __ms_name__ = "WdRowHeightRule"

    __url__ = "https://msdn.microsoft.com/en-us/library/office/ff193620.aspx"

    __members__ = (
        XmlMappedEnumMember(
            "AUTO",
            0,
            "auto",
            "The row height is adjusted to accommodate th"
            "e tallest value in the row.",
        ),
        XmlMappedEnumMember(
            "AT_LEAST",
            1,
            "atLeast",
            "The row height is at least a minimum " "specified value.",
        ),
        XmlMappedEnumMember("EXACTLY", 2, "exact", "The row height is an exact value."),
    )


class WD_TABLE_ALIGNMENT(XmlEnumeration):
    """
    Specifies table justification type.

    Example::

        from docx.enum.table import WD_TABLE_ALIGNMENT

        table = document.add_table(3, 3)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
    """

    __ms_name__ = "WdRowAlignment"

    __url__ = " http://office.microsoft.com/en-us/word-help/HV080607259.aspx"

    __members__ = (
        XmlMappedEnumMember("LEFT", 0, "left", "Left-aligned"),
        XmlMappedEnumMember("CENTER", 1, "center", "Center-aligned."),
        XmlMappedEnumMember("RIGHT", 2, "right", "Right-aligned."),
    )


class WD_TABLE_BORDER(XmlEnumeration):
    """
    Specifies table border type.

    Example::

        from docx.enum.table import WD_TABLE_BORDER

        table = document.add_table(3, 3)

    """

    __ms_name__ = "WdLineStyle"

    __url__ = " http://office.microsoft.com/en-us/word-help/HV080607259.aspx"

    __members__ = (
        XmlMappedEnumMember("NONE", 0, "none", "No border."),
        XmlMappedEnumMember("SINGLE", 1, "single", "A single solid line."),
        XmlMappedEnumMember("DOT", 2, "dot", "Dots."),
        XmlMappedEnumMember(
            "DASH_SMALL_GAP", 3, "dashSmallGap", "A dash followed by a small gap."
        ),
        XmlMappedEnumMember(
            "DASH_LARGE_GAP", 4, "dashLargeGap", "A dash followed by a large gap."
        ),
        XmlMappedEnumMember("DASH_DOT", 5, "dashDot", "A dash followed by a dot."),
        XmlMappedEnumMember(
            "DASH_DOT_DOT", 6, "dashDotDot", "A dash followed by two dots."
        ),
        XmlMappedEnumMember("DOUBLE", 7, "double", "Double solid lines."),
        XmlMappedEnumMember("TRIPLE", 8, "triple", "Three solid thin lines."),
        XmlMappedEnumMember(
            "THIN_THICK_SMALL_GAP",
            9,
            "thinThickSmallGap",
            "An internal single thin solid line surrounded by a single thick solid line with a small gap between them.",
        ),
        XmlMappedEnumMember(
            "THICK_THIN_SMALL_GAP",
            10,
            "thickThinSmallGap",
            "An internal single thick solid line surrounded by a single thin solid line with a small gap between them.",
        ),
        XmlMappedEnumMember(
            "THIN_THICK_THIN_SMALLGAP",
            11,
            "thinThickThinSmallGap",
            "An internal single thin solid line surrounded by a single thick solid line surrounded by a single thin solid line with a small gap between all lines.",
        ),
        XmlMappedEnumMember(
            "THIN_THICK_MED_GAP",
            12,
            "thinThickMedGap",
            "An internal single thin solid line surrounded by a single thick solid line with a medium gap between them.",
        ),
        XmlMappedEnumMember(
            "THICK_THIN_MED_GAP",
            13,
            "thickThinMedGap",
            "An internal single thick solid line surrounded by a single thin solid line with a medium gap between them.",
        ),
        XmlMappedEnumMember(
            "THIN_THICK_THIN_MED_GAP",
            14,
            "thinThickThinMedGap",
            "An internal single thin solid line surrounded by a single thick solid line surrounded by a single thin solid line with a medium gap between all lines.",
        ),
        XmlMappedEnumMember(
            "THIN_THICK_LARGE_GAP",
            15,
            "thinThickLargeGap",
            "An internal single thin solid line surrounded by a single thick solid line with a large gap between them.",
        ),
        XmlMappedEnumMember(
            "THICK_THIN_LARGE_GAP",
            16,
            "thickThinLargeGap",
            "An internal single thick solid line surrounded by a single thin solid line with a large gap between them.",
        ),
        XmlMappedEnumMember(
            "THIN_THICK_THIN_LARGE_GAP",
            17,
            "thinThickThinLargeGap",
            "An internal single thin solid line surrounded by a single thick solid line surrounded by a single thin solid line with a large gap between all lines.",
        ),
        XmlMappedEnumMember(
            "SINGLE_WAVY", 18, "singleWavy", "A single wavy solid line."
        ),
        XmlMappedEnumMember(
            "DOUBLE_WAVY", 19, "doubleWavy", "Double wavy solid lines."
        ),
        XmlMappedEnumMember(
            "DASH_DOT_STROKED",
            20,
            "dashDotStroked",
            "A dash followed by a dot stroke, thus rendering a border similar to a barber pole.",
        ),
        XmlMappedEnumMember(
            "EMBOSS_3D",
            21,
            "emboss3D",
            "The border appears to have a 3D embossed look.",
        ),
        XmlMappedEnumMember(
            "ENGRAVE_3D",
            22,
            "engrave3D",
            "The border appears to have a 3D engraved look.",
        ),
        XmlMappedEnumMember("OUTSET", 23, "outset", "The border appears to be outset."),
        XmlMappedEnumMember("INSET", 24, "inset", "The border appears to be inset."),
    )


class WD_TABLE_DIRECTION(Enumeration):
    """
    Specifies the direction in which an application orders cells in the
    specified table or row.

    Example::

        from docx.enum.table import WD_TABLE_DIRECTION

        table = document.add_table(3, 3)
        table.direction = WD_TABLE_DIRECTION.RTL
    """

    __ms_name__ = "WdTableDirection"

    __url__ = "https://docs.microsoft.com/en-us/office/vba/api/word.wdlinestyle"

    __members__ = (
        EnumMember(
            "LTR",
            0,
            "The table or row is arranged with the first column "
            "in the leftmost position.",
        ),
        EnumMember(
            "RTL",
            1,
            "The table or row is arranged with the first column "
            "in the rightmost position.",
        ),
    )
