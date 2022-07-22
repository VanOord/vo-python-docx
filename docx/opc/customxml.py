# encoding: utf-8

"""
The :mod:`pptx.packaging` module coheres around the concerns of reading and
writing presentations to and from a .pptx file.
"""

from __future__ import absolute_import, division, print_function, unicode_literals


class CustomXmlBase(object):
    """
    Corresponds to part named ``/CustomXml/item2.xml``, containing the core
    document properties for this document package.
    """

    def __init__(self, element):
        self._element = element


class CustomXML(object):
    """
    Corresponds to part named ``/CustomXml/item2.xml``, containing the
    custom content control elements for this document package.
    """

    def __init__(self, element):
        self._element = element

    @property
    def address(self):
        return self._element.address_text

    @address.setter
    def address(self, value):
        self._element.address_text = value

    @property
    def api_version(self):
        return self._element.apiversion_text

    @api_version.setter
    def api_version(self, value):
        self._element.apiversion_text = value

    @property
    def approved_by(self):
        return self._element.approvedby_text

    @approved_by.setter
    def approved_by(self, value):
        self._element.approvedby_text = value

    @property
    def authorization(self):
        return self._element.authorization_text

    @authorization.setter
    def authorization(self, value):
        self._element.authorization_text = value

    @property
    def checked_by(self):
        return self._element.checkedby_text

    @checked_by.setter
    def checked_by(self, value):
        self._element.checkedby_text = value

    @property
    def client_name(self):
        return self._element.clientname_text

    @client_name.setter
    def client_name(self, value):
        self._element.clientname_text = value

    @property
    def company(self):
        return self._element.company_text

    @company.setter
    def company(self, value):
        self._element.company_text = value

    @property
    def date(self):
        return self._element.date_text

    @date.setter
    def date(self, value):
        self._element.date_text = value

    @property
    def department(self):
        return self._element.department_text

    @department.setter
    def department(self, value):
        self._element.department_text = value

    @property
    def document_number(self):
        return self._element.documentnumber_text

    @document_number.setter
    def document_number(self, value):
        self._element.documentnumber_text = value

    @property
    def document_type(self):
        return self._element.documenttype_text

    @document_type.setter
    def document_type(self, value):
        self._element.documenttype_text = value

    @property
    def email(self):
        return self._element.email_text

    @email.setter
    def email(self, value):
        self._element.email_text = value

    @property
    def function(self):
        return self._element.function_text

    @function.setter
    def function(self, value):
        self._element.function_text = value

    @property
    def function_excerpt(self):
        return self._element.functionexcerpt_text

    @function_excerpt.setter
    def function_excerpt(self, value):
        self._element.functionexcerpt_text = value

    @property
    def location(self):
        return self._element.location_text

    @location.setter
    def location(self, value):
        self._element.location_text = value

    @property
    def name(self):
        return self._element.name_text

    @name.setter
    def name(self, value):
        self._element.name_text = value

    @property
    def prepared_by(self):
        return self._element.preparedby_text

    @prepared_by.setter
    def prepared_by(self, value):
        self._element.preparedby_text = value

    @property
    def project_director(self):
        return self._element.projectdirector_text

    @project_director.setter
    def project_director(self, value):
        self._element.projectdirector_text = value

    @property
    def project_name(self):
        return self._element.projectname_text

    @project_name.setter
    def project_name(self, value):
        self._element.projectname_text = value

    @property
    def project_title(self):
        return self._element.projecttitle_text

    @project_title.setter
    def project_title(self, value):
        self._element.projecttitle_text = value

    @property
    def project_code(self):
        return self._element.projectcode_text

    @project_code.setter
    def project_code(self, value):
        self._element.projectcode_text = value

    @property
    def recipient(self):
        return self._element.recipient_text

    @recipient.setter
    def recipient(self, value):
        self._element.recipient_text = value

    @property
    def reference(self):
        return self._element.reference_text

    @reference.setter
    def reference(self, value):
        self._element.reference_text = value

    @property
    def report_date(self):
        return self._element.reportdate_text

    @report_date.setter
    def report_date(self, value):
        self._element.reportdate_text = value

    @property
    def report_number(self):
        return self._element.reportnumber_text

    @report_number.setter
    def report_number(self, value):
        self._element.reportnumber_text = value

    @property
    def revision(self):
        return self._element.revision_text

    @revision.setter
    def revision(self, value):
        self._element.revision_text = value

    @property
    def surname(self):
        return self._element.surname_text

    @surname.setter
    def surname(self, value):
        self._element.surname_text = value

    @property
    def telephone(self):
        return self._element.telephone_text

    @telephone.setter
    def telephone(self, value):
        self._element.telephone_text = value

    @property
    def your_reference(self):
        return self._element.yourreference_text

    @your_reference.setter
    def your_reference(self, value):
        self._element.yourreference_text = value
