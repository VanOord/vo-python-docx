# encoding: utf-8
"""
Custom Xml part, corresponds to ``/docProps/core.xml`` part in package.
"""

from __future__ import absolute_import, division, print_function, unicode_literals

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.customxml import CustomXML, CustomXmlBase
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.oxml.customxml import CT_CustomXML


class CustomXmlPart(XmlPart):
    """ """

    @classmethod
    def default(cls, package):
        """
        Return a new |CustomXML| object initialized with default
        values for its base properties.
        """
        custom_xml_part = cls._new(package)
        content_control = custom_xml_part.custom_xml
        content_control.revision = "A"
        return custom_xml_part

    @property
    def custom_xml(self):
        """
        A |CustomXmlPart| object providing read/write access to the custom
        document properties contained in this customxml part.
        """
        print(type(self._element))
        if type(self._element) == CT_CustomXML:
            return CustomXML(self.element)
        else:
            return CustomXmlBase(self.element)

    @classmethod
    def _new(cls, package):
        partname = PackURI("/customXml/item2.xml")
        content_type = CT.XML
        custom_xml = CT_CustomXML.new()
        return CustomXmlPart(partname, content_type, custom_xml, package)
