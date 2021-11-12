"""Objects related to borders.

There can be multiple objects with borders, they appear in tables 
and table cells but can also be found on sections.
"""
from __future__ import absolute_import, division, print_function, unicode_literals

from docx.shared import ElementProxy, lazyproperty


class Border(ElementProxy):
    """
    Proxy wrapping the different border elements.
    """

    __slots__ = "_border"

    def __init__(self, element):
        super(Border, self).__init__(element, None)
        self._border = element

    @property
    def color(self):
        return self._border.color

    @color.setter
    def color(self, value):
        self._border.color = value

    @property
    def frame(self):
        return self._border.frame

    @frame.setter
    def frame(self, value):
        self._border.frame = value

    @property
    def theme_color(self):
        return self._border.themeColor

    @theme_color.setter
    def theme_color(self, value):
        self._border.themeColor = value

    @property
    def line_style(self):
        return self._border.val

    @line_style.setter
    def line_style(self, value):
        self._border.val = value

    @property
    def line_width(self):
        return self._border.sz

    @line_width.setter
    def line_width(self, value):
        self._border.sz = value

    @property
    def space(self):
        return self._border.space

    @space.setter
    def space(self, value):
        self._border.space = value

    @property
    def shadow(self):
        return self._border.shadow

    @shadow.setter
    def shadow(self, value):
        self._border.shadow = value


class Borders(ElementProxy):
    def __init__(self, element):
        super(Borders, self).__init__(element, None)
        self._tblBorders = self._element = element

    def __iter__(self):
        for border in self._element.getchildren():
            yield Border(border)

    @property
    def bottom(self):
        bottom = self._element.bottom_border
        if bottom is None:
            return None
        return Border(bottom)

    @bottom.setter
    def bottom(self, value):
        self._element.bottom_border = value

    @lazyproperty
    def borders(self):
        return [border for border in self]

    @property
    def inside_horizontal(self):
        insideH = self._element.insideH_border
        if insideH is None:
            return None
        return Border(insideH)

    @inside_horizontal.setter
    def inside_horizontal(self, value):
        self._element.insideH_border = value

    @property
    def inside_vertical(self):
        insideV = self._element.insideV_border
        if insideV is None:
            return None
        return Border(insideV)

    @inside_vertical.setter
    def inside_vertical(self, value):
        self._element.insideV_border = value

    @property
    def left(self):
        left = self._element.left_border
        if left is None:
            return None
        return Border(left)

    @left.setter
    def left(self, value):
        self._element.left_border = value

    @property
    def right(self):
        right = self._element.right_border
        if right is None:
            return None
        return Border(right)

    @right.setter
    def right(self, value):
        self._element.right_border = value

    @property
    def top(self):
        top = self._element.top_border
        if top is None:
            return None
        return Border(top)

    @top.setter
    def top(self, value):
        self._element.top_border = value


class CellBorders(Borders):
    @property
    def top_left_bottom_right(self):
        tl2br = self._element.top_left_bottom_right
        if tl2br is None:
            return None
        return Border(tl2br)

    @top_left_bottom_right.setter
    def top_left_bottom_right(self, value):
        self._element.top_left_bottom_right = value

    @property
    def top_right_bottom_left(self):
        tr2bl = self._element.top_right_bottom_left
        if tr2bl is None:
            return None
        return Border(tr2bl)

    @top_right_bottom_left.setter
    def top_right_bottom_left(self, value):
        self._element.top_right_bottom_left = value
