# encoding: utf-8

"""
Run-related proxy objects for python-docx, Run in particular.
"""

from __future__ import absolute_import, print_function, unicode_literals

from ..shared import Parented


class Equation(Parented):
    """ """

    def __init__(self, eq, parent):
        super(Equation, self).__init__(parent)
        self._eq = self._element = self.element = eq

    @property
    def raw_xml(self):
        return self._eq.xml
