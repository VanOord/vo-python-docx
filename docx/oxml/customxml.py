# encoding: utf-8

"""
lxml custom element classes for core properties-related XML elements.
"""

from __future__ import absolute_import, division, print_function, unicode_literals

import re

from datetime import datetime, timedelta

from . import parse_xml
from .ns import nsdecls, qn
from .xmlchemy import BaseOxmlElement, ZeroOrOne


class CT_CustomXML(BaseOxmlElement):
    """ """

    address = ZeroOrOne("vo:Address", successors=())
    apiversion = ZeroOrOne("vo:ApiVersion", successors=())
    approvedby = ZeroOrOne("vo:ApprovedBy", successors=())
    authorization = ZeroOrOne("vo:Authorization", successors=())
    checkedby = ZeroOrOne("vo:CheckedBy", successors=())
    clientname = ZeroOrOne("vo:ClientName", successors=())
    company = ZeroOrOne("vo:Company", successors=())
    date = ZeroOrOne("vo:Date", successors=())
    department = ZeroOrOne("vo:Department", successors=())
    documentnumber = ZeroOrOne("vo:DocumentNumber", successors=())
    documenttype = ZeroOrOne("vo:DocumentType", successors=())
    email = ZeroOrOne("vo:Email", successors=())
    function = ZeroOrOne("vo:Function", successors=())
    functionexcerpt = ZeroOrOne("vo:FunctionExcerpt", successors=())
    location = ZeroOrOne("vo:Location", successors=())
    name = ZeroOrOne("vo:Name", successors=())
    preparedby = ZeroOrOne("vo:PreparedBy", successors=())
    projectdirector = ZeroOrOne("vo:ProjectDirector", successors=())
    projectname = ZeroOrOne("vo:ProjectName", successors=())
    projecttitle = ZeroOrOne("vo:ProjectTitle", successors=())
    projectcode = ZeroOrOne("vo:Projectcode", successors=())
    recipient = ZeroOrOne("vo:Recipient", successors=())
    reference = ZeroOrOne("vo:Reference", successors=())
    reportdate = ZeroOrOne("vo:ReportDate", successors=())
    reportnumber = ZeroOrOne("vo:ReportNumber", successors=())
    revision = ZeroOrOne("vo:Revision", successors=())
    surname = ZeroOrOne("vo:Surname", successors=())
    telephone = ZeroOrOne("vo:Telephone", successors=())
    yourreference = ZeroOrOne("vo:YourReference", successors=())

    _coreProperties_tmpl = "<vo:customContent %s/>\n" % nsdecls("vo")

    @classmethod
    def new(cls):
        """
        Return a new ``<cp:coreProperties>`` element
        """
        xml = cls._coreProperties_tmpl
        coreProperties = parse_xml(xml)
        return coreProperties

    @property
    def address_text(self):
        return self._text_of_element("address")

    @address_text.setter
    def address_text(self, value):
        self._set_element_text("address", value)

    @property
    def apiversion_text(self):
        return self._text_of_element("apiversion")

    @apiversion_text.setter
    def apiversion_text(self, value):
        self._set_element_text("apiversion", value)

    @property
    def approvedby_text(self):
        return self._text_of_element("approvedby")

    @approvedby_text.setter
    def approvedby_text(self, value):
        self._set_element_text("approvedby", value)

    @property
    def authorization_text(self):
        return self._text_of_element("authorization")

    @authorization_text.setter
    def authorization_text(self, value):
        self._set_element_text("authorization", value)

    @property
    def checkedby_text(self):
        return self._text_of_element("checkedby")

    @checkedby_text.setter
    def checkedby_text(self, value):
        self._set_element_text("checkedby", value)

    @property
    def clientname_text(self):
        return self._text_of_element("clientname")

    @clientname_text.setter
    def clientname_text(self, value):
        self._set_element_text("clientname", value)

    @property
    def company_text(self):
        return self._text_of_element("company")

    @company_text.setter
    def company_text(self, value):
        self._set_element_text("company", value)

    @property
    def date_text(self):
        return self._text_of_element("date")

    @date_text.setter
    def date_text(self, value):
        self._set_element_text("date", value)

    @property
    def department_text(self):
        return self._text_of_element("department")

    @department_text.setter
    def department_text(self, value):
        self._set_element_text("department", value)

    @property
    def documentnumber_text(self):
        return self._text_of_element("documentnumber")

    @documentnumber_text.setter
    def documentnumber_text(self, value):
        self._set_element_text("documentnumber", value)

    @property
    def documenttype_text(self):
        return self._text_of_element("documenttype")

    @documenttype_text.setter
    def documenttype_text(self, value):
        self._set_element_text("documenttype", value)

    @property
    def email_text(self):
        return self._text_of_element("email")

    @email_text.setter
    def email_text(self, value):
        self._set_element_text("email", value)

    @property
    def function_text(self):
        return self._text_of_element("function")

    @function_text.setter
    def function_text(self, value):
        self._set_element_text("function", value)

    @property
    def functionexcerpt_text(self):
        return self._text_of_element("functionexcerpt")

    @functionexcerpt_text.setter
    def functionexcerpt_text(self, value):
        self._set_element_text("functionexcerpt", value)

    @property
    def location_text(self):
        return self._text_of_element("location")

    @location_text.setter
    def location_text(self, value):
        self._set_element_text("location", value)

    @property
    def name_text(self):
        return self._text_of_element("name")

    @name_text.setter
    def name_text(self, value):
        self._set_element_text("name", value)

    @property
    def preparedby_text(self):
        return self._text_of_element("preparedby")

    @preparedby_text.setter
    def preparedby_text(self, value):
        self._set_element_text("preparedby", value)

    @property
    def projectdirector_text(self):
        return self._text_of_element("projectdirector")

    @projectdirector_text.setter
    def projectdirector_text(self, value):
        self._set_element_text("projectdirector", value)

    @property
    def projectname_text(self):
        return self._text_of_element("projectname")

    @projectname_text.setter
    def projectname_text(self, value):
        self._set_element_text("projectname", value)

    @property
    def projecttitle_text(self):
        return self._text_of_element("projecttitle")

    @projecttitle_text.setter
    def projecttitle_text(self, value):
        self._set_element_text("projecttitle", value)

    @property
    def projectcode_text(self):
        return self._text_of_element("projectcode")

    @projectcode_text.setter
    def projectcode_text(self, value):
        self._set_element_text("projectcode", value)

    @property
    def recipient_text(self):
        return self._text_of_element("recipient")

    @recipient_text.setter
    def recipient_text(self, value):
        self._set_element_text("recipient", value)

    @property
    def reference_text(self):
        return self._text_of_element("reference")

    @reference_text.setter
    def reference_text(self, value):
        self._set_element_text("reference", value)

    @property
    def reportdate_text(self):
        return self._text_of_element("reportdate")

    @reportdate_text.setter
    def reportdate_text(self, value):
        self._set_element_text("reportdate", value)

    @property
    def reportnumber_text(self):
        return self._text_of_element("reportnumber")

    @reportnumber_text.setter
    def reportnumber_text(self, value):
        self._set_element_text("reportnumber", value)

    @property
    def revision_text(self):
        return self._text_of_element("revision")

    @revision_text.setter
    def revision_text(self, value):
        self._set_element_text("revision", value)

    @property
    def surname_text(self):
        return self._text_of_element("surname")

    @surname_text.setter
    def surname_text(self, value):
        self._set_element_text("surname", value)

    @property
    def telephone_text(self):
        return self._text_of_element("telephone")

    @telephone_text.setter
    def telephone_text(self, value):
        self._set_element_text("telephone", value)

    @property
    def yourreference_text(self):
        return self._text_of_element("yourreference")

    @yourreference_text.setter
    def yourreference_text(self, value):
        self._set_element_text("yourreference", value)

    def _get_or_add(self, prop_name):
        """
        Return element returned by 'get_or_add_' method for *prop_name*.
        """
        get_or_add_method_name = "get_or_add_%s" % prop_name
        get_or_add_method = getattr(self, get_or_add_method_name)
        element = get_or_add_method()
        return element

    def _set_element_text(self, prop_name, value):
        """
        Set string value of *name* property to *value*.
        """
        value = str(value)
        if len(value) > 255:
            tmpl = "exceeded 255 char limit for property, got:\n\n'%s'"
            raise ValueError(tmpl % value)
        element = self._get_or_add(prop_name)
        element.text = value

    def _text_of_element(self, property_name):
        """
        Return the text in the element matching *property_name*, or an empty
        string if the element is not present or contains no text.
        """
        element = getattr(self, property_name)
        if element is None:
            return ""
        if element.text is None:
            return ""
        return element.text
