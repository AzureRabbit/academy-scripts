# -*- coding: utf-8 -*-
#pylint: disable=I0011,W0703,R0903,W0110,W0141,R0904,W0403
""" DocDriver class
"""

# --------------------------- REQUIRED LIBRARIES ------------------------------

from driver_class import Driver
import win32com.client

# ---------------------------- CLASS DEFINITION -------------------------------

class DocDriver(Driver):
    """ DocDriver object
    """

    def __init__(self, path):
        super(DocDriver, self).__init__(path)
        self._word = win32com.client.DispatchEx("Word.Application")
        self._doc = self._word.Documents.Open(path)

    def __del__(self):
        self._doc.Close()
        self._word.Quit()

    def _get_builtindocumentproperty(self, name, default=None):
        """ Try to get a BuiltInDocumentProperty from document and convert its
        value to an string.
        """

        value = None

        try:
            prop = self._doc.BuiltInDocumentProperties[name]
            value = str(prop)
            value = value if value and value != 'None' else default
        except Exception as ex:
            print name, ex

        return value


    def get_type(self):
        """ Gets the document type
        """
        allowed_types = self.get_allowed_types()
        return allowed_types[self._extension]

    # ------------------------- DOCUMENT PROPERTIES ---------------------------

    def get_title(self):
        """ Gets the document title propperty
        """
        return self._get_builtindocumentproperty('Title')

    def get_author(self):
        """ Gets the document author propperty
        """
        return self._get_builtindocumentproperty('Author')

    def get_last_author(self):
        """ Gets the document last_author propperty
        """
        return self._get_builtindocumentproperty('Manager')

    def get_manager(self):
        """ Gets the document manager propperty
        """
        return self._get_builtindocumentproperty('Manager')

    def get_company(self):
        """ Gets the document company propperty
        """
        return self._get_builtindocumentproperty('Company')

    def get_category(self):
        """ Gets the document category propperty
        """
        return self._get_builtindocumentproperty('Category')

    def get_keywords(self):
        """ Gets the document keywords propperty
        """
        return self._get_builtindocumentproperty('Keywords')

    def get_comments(self):
        """ Gets the document comments propperty
        """
        return self._get_builtindocumentproperty('Comments')

    def get_template(self):
        """ Gets the document template propperty
        """
        return self._get_builtindocumentproperty('Template')

    def get_revision_number(self):
        """ Gets the document revision_number propperty
        """
        return self._get_builtindocumentproperty('Revision number')

    def get_last_print_date(self):
        """ Gets the document last_print_date propperty
        """
        return self._get_builtindocumentproperty('Last print date')

    def get_creation_date(self):
        """ Gets the document creation_date propperty
        """
        return self._get_builtindocumentproperty('Creation date')

    def get_last_save_time(self):
        """ Gets the document last_save_time propperty
        """
        return self._get_builtindocumentproperty('Last save time')

    def get_total_editing_time(self):
        """ Gets the document total_editing_time propperty
        """
        return self._get_builtindocumentproperty('Total editing time')

    def get_number_of_pages(self):
        """ Gets the document number_of_pages propperty
        """
        return self._get_builtindocumentproperty('Number of pages')

    def get_number_of_words(self):
        """ Gets the document number_of_words propperty
        """
        return self._get_builtindocumentproperty('Number of words')

    def get_number_of_characters(self):
        """ Gets the document number_of_characters propperty
        """
        return self._get_builtindocumentproperty('Number of characters')

    def get_number_of_lines(self):
        """ Gets the document number_of_lines propperty
        """
        return self._get_builtindocumentproperty('Number of lines')

    def get_number_of_paragraphs(self):
        """ Gets the document number_of_paragraphs propperty
        """
        return self._get_builtindocumentproperty('Number of paragraphs')

    def get_number_of_chars_with_spaces(self):
        """ Gets the document umber_of_characters_with_spaces propperty
        """
        return self._get_builtindocumentproperty('Number of characters (with spaces)')

    def get_number_of_bytes(self):
        """ Gets the document number_of_bytes propperty
        """
        return self._get_builtindocumentproperty('Number of bytes')


    def get_hyperlink_base(self):
        """ Gets the document hyperlink_base propperty
        """
        return self._get_builtindocumentproperty('Hyperlink base')

    # --------------------------- STATIC METHODS ------------------------------

    @staticmethod
    def get_allowed_types():
        return {
            u'doc': u'Documento de Word 97-2003',
            u'docm': u'Documento habilitado con macros de Word',
            u'docx': u'Documento de Word',
            u'dot': u'Plantilla de Word 97-2003',
            u'dotm': u'Plantilla habilitada con macros de Word',
            u'dotx': u'Plantilla de Word',
            u'mht': u'Página web de un solo archivo',
            u'html': u'Página web',
            u'mhtml': u'Página web de un solo archivo',
            u'odt': u'Texto de OpenDocument',
            u'pdf': u'PDF',
            u'rtf': u'Formato de texto enriquecido',
            u'txt': u'Texto sin formato',
            u'wps': u'Documento de Works 6-9',
            u'xml': u'Documento XML de Word',
            u'xps': u'Documento XPS',
            u'dic': u'Corrector ortográfico personalizado',
            u'thmx': u'Tema de Office'
        }
