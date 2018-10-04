# encoding: utf-8

"""
Custom element classes related to paragraphs (CT_P).
"""


from docx.oxml.xmlchemy import BaseOxmlElement


class CT_OMath(BaseOxmlElement):
    """
    ``<m:oMath>`` element, containing math contents.
    """
    
class CT_OMathPara(BaseOxmlElement):
    """
    ``<m:oMathPara>`` element, containing math contents.
    """
    