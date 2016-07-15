# -*- coding: utf-8 -*-
#pylint: disable=I0011,W0703,R0903,W0110,W0141,R0904
""" DocDriver class
"""

import os

class Driver(object):
    """ Driver abstract class
    """

    def __init__(self, path):
        self._path = os.path.abspath(path)
        self._basename = os.path.basename(path)
        self._name = os.path.splitext(self._basename)[0]

        extension = os.path.splitext(self._basename)[1]
        self._extension = None if len(extension) <= 1 else extension[1:]

    # -------------------------- PATH INFORMATION -----------------------------

    def get_name(self):
        """ Gets the document name
        """
        return self._name

    def get_path(self):
        """ Gets the document path
        """
        return self._path

    def get_extension(self):
        """ Gets the document filename extension
        """
        return self._extension

    def get_type(self):
        """ Gets the document type
        """
        raise NotImplementedError("get_title has not been implemented yet")

    # ------------------------- DOCUMENT PROPERTIES ---------------------------

    def get_title(self):
        """ Gets the document title propperty
        """
        raise NotImplementedError("get_title has not been implemented yet")

    def get_author(self):
        """ Gets the document author propperty
        """
        raise NotImplementedError("get_author has not been implemented yet")

    def get_last_author(self):
        """ Gets the document last_author propperty
        """
        raise NotImplementedError("get_last_author has not been implemented yet")

    def get_manager(self):
        """ Gets the document manager propperty
        """
        raise NotImplementedError("get_manager has not been implemented yet")

    def get_company(self):
        """ Gets the document company propperty
        """
        raise NotImplementedError("get_company has not been implemented yet")

    def get_category(self):
        """ Gets the document category propperty
        """
        raise NotImplementedError("get_category has not been implemented yet")

    def get_keywords(self):
        """ Gets the document keywords propperty
        """
        raise NotImplementedError("get_keywords has not been implemented yet")

    def get_comments(self):
        """ Gets the document comments propperty
        """
        raise NotImplementedError("get_comments has not been implemented yet")

    def get_template(self):
        """ Gets the document template propperty
        """
        raise NotImplementedError("get_template has not been implemented yet")

    def get_revision_number(self):
        """ Gets the document revision_number propperty
        """
        raise NotImplementedError("get_revision_number has not been implemented yet")

    def get_last_print_date(self):
        """ Gets the document last_print_date propperty
        """
        raise NotImplementedError("get_last_print_date has not been implemented yet")

    def get_creation_date(self):
        """ Gets the document creation_date propperty
        """
        raise NotImplementedError("get_creation_date has not been implemented yet")

    def get_last_save_time(self):
        """ Gets the document last_save_time propperty
        """
        raise NotImplementedError("get_last_save_time has not been implemented yet")

    def get_total_editing_time(self):
        """ Gets the document total_editing_time propperty
        """
        raise NotImplementedError("get_total_editing_time has not been implemented yet")

    def get_number_of_pages(self):
        """ Gets the document number_of_pages propperty
        """
        raise NotImplementedError("get_number_of_pages has not been implemented yet")

    def get_number_of_words(self):
        """ Gets the document number_of_words propperty
        """
        raise NotImplementedError("get_number_of_words has not been implemented yet")

    def get_number_of_characters(self):
        """ Gets the document number_of_characters propperty
        """
        raise NotImplementedError("get_number_of_characters has not been implemented yet")

    def get_number_of_lines(self):
        """ Gets the document number_of_lines propperty
        """
        raise NotImplementedError("get_number_of_lines has not been implemented yet")

    def get_number_of_paragraphs(self):
        """ Gets the document number_of_paragraphs propperty
        """
        raise NotImplementedError("get_number_of_paragraphs has not been implemented yet")

    def get_number_of_chars_with_spaces(self):
        """ Gets the document umber_of_characters_with_spaces propperty
        """
        raise NotImplementedError("get_number_of_characters has not been implemented yet")

    def get_number_of_bytes(self):
        """ Gets the document number_of_bytes propperty
        """
        raise NotImplementedError("get_number_of_bytes has not been implemented yet")

    def get_hyperlink_base(self):
        """ Gets the document hyperlink_base propperty
        """
        raise NotImplementedError("get_hyperlink_base has not been implemented yet")

    # --------------------------- STATIC METHODS ------------------------------

    @staticmethod
    def get_allowed_types():
        """ Returns a dictionary with allowed application types
        """
        raise NotImplementedError("get_title has not been implemented yet")
