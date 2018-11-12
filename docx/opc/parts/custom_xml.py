# encoding: utf-8

"""
Core properties part, corresponds to ``/docProps/core.xml`` part in package.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from datetime import datetime

from ..constants import CONTENT_TYPE as CT
from ..customxml import CustomXML, CustomXmlBase
from ...oxml.customxml import CT_CustomXML
from ..packuri import PackURI
from ..part import XmlPart


class CustomXmlPart(XmlPart):
    """

    """
    @property
    def custom_xml(self):
        """
        A |CoreProperties| object providing read/write access to the core
        properties contained in this core properties part.
        """
        if type(self._element) == CT_CustomXML:
            return CustomXML(self.element)
        else:
            return CustomXmlBase(self.element)     

    @classmethod
    def _new(cls, package):
        partname = PackURI('/customXml/item2.xml')
        content_type = CT.XML
        custom_Xml = CT_CustomXML.new()
        return CustomXmlPart(
            partname, content_type, custom_Xml, package
        )
