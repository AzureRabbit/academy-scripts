# -*- coding: utf-8 -*-
#pylint: disable=I0011,W0703,R0903,W0110,W0141,C0301
""" Builds tests about documents
"""

# --------------------------- REQUIRED LIBRARIES ------------------------------

from lib.answer_class import Answer
from lib.question_class import Question
from lib.docdriver_class import DocDriver
from lib.random_values import RandomValues
from lib.words import WORDS
from lib.idioms import IDIOMS

import argparse
import os
import win32com.client
import random
import locale

from datetime import datetime, timedelta, date

from pprint import pprint
import locale
import re

# -------------------------------- CONSTANTS ----------------------------------

XMLOUTER = """<?xml version="1.0" encoding="UTF-8"?>
<?xml-stylesheet type="text/css" href="test.css"?>

<test>
{}
</test>

"""

# -------------------------- MAIN SCRIPT BEHAVIOR -----------------------------


class App(object):
    """ Application main controller, this class has been defined following the
    singleton pattern to ensures only one object can be instantiated.
    """

    __instance = None


    def __new__(cls):
        """ Prevent multiple instances from self (Singleton Pattern)
        """

        if cls.__instance == None:
            cls.__instance = object.__new__(cls)
            cls.__instance.name = "The one"
        return cls.__instance


    def __init__(self):
        self.abspath = None
        self.basename = None
        self.dirname = None
        self.filename = None

        self._os_encoding = locale.getpreferredencoding()
        locale.setlocale(locale.LC_ALL, '')

        self._random = RandomValues()
        self._questions = []
        self._driver = None


    def _argparse(self):
        """ Detines an user-friendly command-line interface and proccess its
        arguments.
        """

        description = u'Convert a Microsoft Word DOCX format to PDF document.'

        parser = argparse.ArgumentParser(description)
        parser.add_argument('file', metavar='file', type=str,
                            help='path of the .docx file will be converted')

        args = parser.parse_args()

        self.abspath = os.path.abspath(args.file)
        self.basename = os.path.basename(self.abspath)
        self.dirname = os.path.dirname(self.abspath)
        self.filename = self.basename and os.path.splitext(self.basename)[0]

# -----------------------------------------------------------------------------


    def _q_path_name(self, docname):
        """ Build question about filename with answers:
            - True name without extension
            - Extension only
            - Literal "Fichero"
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_name()

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Cuál es la cadena de texto que se utiliza para '
                    u'designar el archivo')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            extension = self._driver.get_extension() or u'.rtf'
            question.add_answer(extension, False)

            question.add_answer(u'Fichero', False)


    def _q_path_extension(self, docname):
        """ Build question about extension with answers:
            - True file extension
            - Size in bytes
            - Random file extension
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_extension()

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Cuál es la extension del archivo')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            size = locale.format(
                "%d", os.path.getsize(self.abspath), grouping=True)
            question.add_answer(size + ' Bytes', False)

            extensions = self._driver.get_allowed_types().keys()
            random_extension = self._random.values(extensions, right_value)
            question.add_answer(random_extension[0], False)


    def _q_path_type(self, docname):
        """ Build question about type with answers:
            - Rigth type for file
            - Random application file type
            - Random application file type
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_type()

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Qué tipo de archivo corresponde al documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            types = self._driver.get_allowed_types().values()
            random_types = self._random.values(types, right_value, 2)

            question.add_answer(random_types[0], False)

            question.add_answer(random_types[1], False)

    # ------------------------- DOCUMENT PROPERTIES ---------------------------


    def _q_prop_title(self, docname):
        """ Build question about title with answers:
            - True document title or empty sentence
            - The document name
            - Literal "El texto del primer párrafo del documento"
        """

        # STEP 1: Get value for rigth answer
        itdoesnothave = u'El documento carece de título'
        right_value = self._driver.get_title(itdoesnothave)

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Cuál es el título del documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            name = self._driver.get_name()
            question.add_answer(name, False)

            wrong2 = u'El texto del primer párrafo del documento'
            question.add_answer(wrong2, False)


    def _q_prop_author(self, docname):
        """ Build question about author with answers:
            - True document author or empty sentence
            - Random name from list
            - Random name from list
        """

        # STEP 1: Get value for rigth answer
        itdoesnothave = u'No se ha indicado autor para el documento'
        right_value = self._driver.get_author(itdoesnothave)

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Quién figura como autor del documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            authors = self._random.person(right_value, 2)

            question.add_answer(authors[0], False)
            question.add_answer(authors[1], False)


    def _q_prop_last_author(self, docname):
        """ Build question about last_author with answers:
            - True document author (mandatory)
            - Random name from list
            - Random name from list
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_last_author(None)

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Quién figura como último autor del documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            authors = self._random.person(right_value, 2)

            question.add_answer(authors[0], False)
            question.add_answer(authors[1], False)


    def _q_prop_manager(self, docname):
        """ Build question about manager with answers:
            - True value for administrador or empty sentence
            - Random name from list
            - Random name from list
        """

        # STEP 1: Get value for rigth answer
        itdoesnothave = u'No se ha indicado administrador para el documento'
        right_value = self._driver.get_manager(itdoesnothave)

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Quién figura como administrador del documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            authors = self._random.person(right_value, 2)

            question.add_answer(authors[0], False)
            question.add_answer(authors[1], False)


    def _q_prop_company(self, docname):
        """ Build question about company with answers:
            - True value for company or empty sentence
            - Random value from list
            - Random value form list
        """

        # STEP 1: Get value for rigth answer
        itdoesnothave = u'El valor de la propiedad aparece vacío'
        right_value = self._driver.get_company(itdoesnothave)

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Qué compañía figura en las propiedades del documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            companies = self._random.company(right_value, 2)

            question.add_answer(companies[0], False)
            question.add_answer(companies[1], False)


    def _q_prop_subject(self, docname):
        """ Build question about subject with answers
            - True value for subject or empty sentence
            - Random value from list
            - Random value form list
        """

        # STEP 1: Get value for rigth answer
        itdoesnothave = u'No se ha asignado asunto al documento'
        right_value = self._driver.get_subject(itdoesnothave)

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Cuál es el asunto del documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            subjects = self._random.subject(right_value, 2)

            question.add_answer(subjects[0], False)
            question.add_answer(subjects[1], False)

    def _q_prop_category(self, docname):
        """ Build question about category with answers:
            - True value for category or empty sentence
            - Random value from list
            - Random value form list
        """

        # STEP 1: Get value for rigth answer
        itdoesnothave = u'No se ha asignado categoría al documento'
        right_value = self._driver.get_category(itdoesnothave)

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Qué categoría ha sido indicada para el documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            categories = self._random.subject(right_value, 2)

            question.add_answer(categories[0], False)
            question.add_answer(categories[1], False)


    def _q_prop_keywords(self, docname):
        """ Build question about keywords with answers:
            - The real keywords from document or empty sentence
            - Random keywords
            - Random keywords
        """

        # STEP 1: Get value for rigth answer
        itdoesnothave = u'El valor de la propiedad aparece vacío'
        right_value = self._driver.get_keywords(itdoesnothave)

        if not right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Cuáles son las palabras clave empleadas en el '
                    u'documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            exclude_words = re.findall(r"[\w']+", right_value)

            keywords1 = self._random.topic(exclude_words, 2)
            question.add_answer(u' '.join(keywords1), False)

            keywords2 = self._random.topic(exclude_words, 3)
            question.add_answer(u' '.join(keywords2), False)


    def _q_prop_comments(self, docname):
        """ Build question about comments with answers:
            - Real comment text string or empty sentence
            - Random topic
            - Random subject
        """

        # STEP 1: Get value for rigth answer
        itdoesnothave = u'El valor de la propiedad aparece vacío'
        right_value = self._driver.get_comments(itdoesnothave)

        if not right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Qué comentarios figuran en las propiedades del '
                    u'documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            subject = self._random.subject(right_value, 1)
            question.add_answer(subject[0], False)

            topic = self._random.topic(right_value, 1)
            question.add_answer(topic[0], False)


    def _q_prop_template(self, docname):
        """ Build question about template with answers:
            - Real template name or empty sentence
            - Random template name
            - Random template name
        """

        # STEP 1: Get value for rigth answer
        itdoesnothave = u'El nombre de la plantilla no aparece reflejado'
        right_value = self._driver.get_template(itdoesnothave)

        if not right_value:

            # STEP 3: Create question and add register it in test
            text = (u'A partir de qué plantilla ha sido creado el documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)
            tplname = os.path.splitext(right_value)[0]

            # STEP 5: Build and add wrong answers
            exception = os.path.splitext(right_value)
            extensions = self._random.values(
                ['dot', 'dotm', 'dotx', 'docx', 'doc'], exception, 2)

            question.add_answer(u'{}.{}'.format(tplname, extensions[0]), False)
            question.add_answer(u'{}.{}'.format(tplname, extensions[1]), False)


    def _q_prop_revision_number(self, docname):
        """ Build question about revision_number with answers:
            - Real revision number or empty sentence
            - Random revision number
            - Random revision number
        """

        # STEP 1: Get value for rigth answer
        itdoesnothave = u'El valor de la propiedad no aparece indicado'
        right_value = self._driver.get_revision_number(itdoesnothave)

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Cuál es el número de revisión que figura en el '
                    u'documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(1, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    def _q_prop_last_print_date(self, docname):
        """ Build question about last_print_date with answers:
            - Real last print date (mandatory)
            - Random last print date
            - Random last print date
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_last_print_date(None)

        if right_value:
            date_sn = datetime.strptime(right_value, '%x %X').date()
            right_value = date_sn.strftime('%x')

            # STEP 3: Create question and add register it in test
            text = (u'Cuál de las siguientes figura como la fecha de última '
                    u'impresión del documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            dates = self._random.date(date_sn, 10, date_sn, 2)

            date_str = dates[0].strftime('%x')
            question.add_answer(date_str, False)

            date_str = dates[1].strftime('%x')
            question.add_answer(date_str, False)


    def _q_prop_creation_date(self, docname):
        """ Build question about creation_date with answers:
            - Real creation date (mandatory)
            - Random creation date
            - Random creation date
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_creation_date(None)

        if right_value:
            date_sn = datetime.strptime(right_value, '%x %X').date()
            right_value = date_sn.strftime('%x')

            # STEP 3: Create question and add register it in test
            text = (u'En qué fecha fue creado el documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            dates = self._random.date(date_sn, 10, date_sn, 2)

            date_str = dates[0].strftime('%x')
            question.add_answer(date_str, False)

            date_str = dates[1].strftime('%x')
            question.add_answer(date_str, False)


    def _q_prop_last_save_time(self, docname):
        """ Build question about last save time with answers:
            - Real last save date (mandatory)
            - Random date
            - Random date
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_last_save_time(None)

        if right_value:

            date_sn = datetime.strptime(right_value, '%x %X').date()
            right_value = date_sn.strftime('%x')

            # STEP 3: Create question and add register it in test
            text = (u'En qué fecha fue guardado por última vez el documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            dates = self._random.date(date_sn, 10, date_sn, 2)

            date_str = dates[0].strftime('%x')
            question.add_answer(date_str, False)

            date_str = dates[1].strftime('%x')
            question.add_answer(date_str, False)


    # def _q_prop_total_editing_time(self, docname):
    #     """ Build question about total editing time with answers:
    #         - Real editing time (mandatory)
    #         - Random number
    #         - Random number
    #     """

    #     # STEP 1: Get value for rigth answer
    #     right_value = self._driver.get_total_editing_time(None)

    #     if right_value:

    #         # STEP 3: Create question and add register it in test
    #         text = (u'Cuál ha sido el tiempo total de edición del documento')
    #         question = Question(u'¿%s %s?' % (text, docname))
    #         self._questions.append(question)

    #         # STEP 4: Add the right answer
    #         rightans = u'{} minutos'.format(right_value)
    #         question.add_answer(rightans, True)

    #         # STEP 5: Build and add wrong answers
    #         number = int(right_value)
    #         wrongs = self._random.integer(1, number+10, number, 2)

    #         wrong1 = u'{} minutos'.format(wrongs[0])
    #         question.add_answer(wrong1, False)

    #         wrong2 = u'{} minutos'.format(wrongs[1])
    #         question.add_answer(wrong2, False)


    def _q_prop_number_of_pages(self, docname):
        """ Build question about number_of_pages with answers:
            - Real number of pages (mandatory)
            - Random number
            - Random number
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_number_of_pages(None)

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Cuál es el número total de páginas del documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(1, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    def _q_prop_number_of_words(self, docname):
        """ Build question about number_of_words with answers:
            - Real number of words (mandatory)
            - Random number of words
            - Random number of words
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_number_of_words(None)

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Cuál es el número total de palabras del documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(1, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    def _q_prop_number_of_characters(self, docname):
        """ Build question about number of characters with answers:
            - Real number of characters (mandatory)
            - Random number of characters
            - Random number of characters
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_number_of_characters(None)

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Cuál es el número total de caracteres, sin contar '
                    u'espacios, del documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(1, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    def _q_prop_number_of_lines(self, docname):
        """ Build question about number_of_lines with answers:
            - Real number of lines (mandatory)
            - Random number
            - Random number
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_number_of_lines(None)

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Cuál es el número total de líneas existentes en el '
                    u'documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(1, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    def _q_prop_number_of_paragraphs(self, docname):
        """ Build question about number_of_paragraphs with answers:
            - Real number of paragraphs (mandatory)
            - Random number
            - Random number
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_number_of_paragraphs(None)

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Cuál es el número total de párrafos existentes en '
                    u'el documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(1, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    def _q_prop_numchars_with_spaces(self, docname):
        """ Build question about number_of_characters with answers:
            - Real number of characters (mandatory)
            - Random number
            - Random number
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_number_of_chars_with_spaces(None)

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Cuál es el número total de caracteres, contando '
                    u'espacios, del documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(1, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    def _q_prop_number_of_bytes(self, docname):
        """ Build question about number_of_bytes with answers:
            - Real size in bytes (mandatory)
            - Random number
            - Random number
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_number_of_bytes(None)

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Cuál es, en bytes, el tamaño del documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(1, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    def _q_prop_hyperlink_base(self, docname):
        """ Build question about hyperlink_base with answers:
            - Real hyperlink base text string
            - Random hyperlink url
            - Random hyperlink url
        """

        # STEP 1: Get value for rigth answer
        itdoesnothave = u'No se ha indicado base para el hipervínculo'
        right_value = self._driver.get_hyperlink_base(itdoesnothave)

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Qué dirección figura como base del hipervínculo '
                    u'para el documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value.lower(), True)

            # STEP 5: Build and add wrong answers
            subjects = self._random.subject(right_value, 2)

            hbase1 = u'www.%s.com' % re.sub(r'[^a-zA-Z0-9]+', '_', subjects[0])
            question.add_answer(hbase1.lower(), False)

            hbase2 = u'www.%s.es' % re.sub(r'[^a-zA-Z0-9]+', '_', subjects[0])
            question.add_answer(hbase2.lower(), False)


    def _q_prop_number_of_sections(self, docname):
        """ Build question about number_of_sections with answers:
            - Real number of sections (mandatory)
            - Random number
            - Random number
        """

        # STEP 1: Get value for rigth answer
        itdoesnothave = u'No hay secciones en el documento'
        right_value = self._driver.get_number_of_sections(itdoesnothave)

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Cuál es el número total de secciones del documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(1, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    def _q_prop_number_of_bookmarks(self, docname):
        """ Build question about number_of_bookmarks with answers:
            - Real number of bookmarks or empty sentence
            - Random number
            - Random number
        """

        # STEP 1: Get value for rigth answer
        itdoesnothave = u'No hay marcadores en el documento'
        right_value = self._driver.get_number_of_bookmarks(itdoesnothave)

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Cuál es el número total de marcadores del documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(1, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    def _q_prop_number_of_comments(self, docname):
        """ Build question about number_of_comments with answers:
            - Real number of comments or empty sentence
            - Random number
            - Random number
        """

        # STEP 1: Get value for rigth answer
        itdoesnothave = u'No hay comentarios en el documento'
        right_value = self._driver.get_number_of_comments(itdoesnothave)

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Cuál es el número total de comentarios del documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(1, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    def _q_prop_default_tabstop(self, docname):
        """ Build question about default_tabstop with answers:
            - Real distance for tabstop (mandatory)
            - Random number
            - Random number
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_default_tabstop(None)

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Cuál es la distancia de tabulación predeterminada para el documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            rigthstr = u'{:.2f} cm'.format(round(right_value, 2))
            question.add_answer(rigthstr, True)

            # STEP 5: Build and add wrong answers
            number = right_value
            wrongans = self._random.float(0.25, number+1.0, number, 2, 0.01)

            wrongstr = u'{:.2f} cm'.format(round(wrongans[0], 2))
            question.add_answer(wrongstr, False)

            wrongstr = u'{:.2f} cm'.format(round(wrongans[1], 2))
            question.add_answer(wrongstr, False)


    def _q_prop_number_of_endnotes(self, docname):
        """ Build question about number_of_endnotes with answers:
            - Real number of end notes or empty sentence
            - Random number
            - Random number
        """

        # STEP 1: Get value for rigth answer
        itdoesnothave = u'No hay notas al final en el documento'
        right_value = self._driver.get_number_of_endnotes(itdoesnothave)

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Cuál es el número total de notas al final en el documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(1, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    def _q_prop_number_of_footnotes(self, docname):
        """ Build question about number_of_footnotes with answers:
            - Real number of notas al pie or empty sentence
            - Random number
            - Random number
        """

        # STEP 1: Get value for rigth answer
        itdoesnothave = u'No hay notas al pie en el documento'
        right_value = self._driver.get_number_of_footnotes(itdoesnothave)

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Cuál es el número total de notas al pie en el documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(1, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    def _q_prop_number_of_revisions(self, docname):
        """ Build question about number_of_revisions with answers:
            - Real number of revisions or empty sentence
            - Random number
            - Random number
        """

        # STEP 1: Get value for rigth answer
        itdoesnothave = u'No hay revisiones en el documento'
        right_value = self._driver.get_number_of_revisions(itdoesnothave)

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Cuál es el número total de revisiones en el documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(1, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    def _q_prop_number_of_shapes(self, docname):
        """ Build question about number_of_shapes with answers:
            - Real number of shapes or empty sentence
            - Random number
            - Random number
        """

        # STEP 1: Get value for rigth answer
        itdoesnothave = u'No hay formas en el documento'
        right_value = self._driver.get_number_of_formas(itdoesnothave)

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Cuál es el número total de formas en el documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(1, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    def _q_prop_number_of_spellingerrors(self, docname):
        """ Build question about number_of_spellingerrors with answers:
            - Real number of spellingerrors or empty sentence
            - Random number
            - Random number
        """

        # STEP 1: Get value for rigth answer
        itdoesnothave = u'No hay errores de ortografía en el documento'
        right_value = self._driver.get_number_of_spellingerrors(itdoesnothave)

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Cuál es el número total de errores ortográficos en el documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(1, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    # def _q_prop_number_of_styles(self, docname):
    #     """ Build question about number_of_styles with answers:
    #         - Real number of styles or empty sentence
    #         - Random number
    #         - Random number
    #     """

    #     # STEP 1: Get value for rigth answer
    #     itdoesnothave = u'No hay estilos definidos en el documento'
    #     right_value = self._driver.get_number_of_styles(itdoesnothave)

    #     if right_value:

    #         # STEP 3: Create question and add register it in test
    #         text = (u'Cuál es el número total de estilos disponibles en el documento')
    #         question = Question(u'¿%s %s?' % (text, docname))
    #         self._questions.append(question)

    #         # STEP 4: Add the right answer
    #         question.add_answer(right_value, True)

    #         # STEP 5: Build and add wrong answers
    #         number = int(right_value)
    #         wrongans = self._random.integer(1, number+10, number, 2)

    #         question.add_answer(unicode(wrongans[0]), False)
    #         question.add_answer(unicode(wrongans[1]), False)


    def _q_prop_number_of_tables(self, docname):
        """ Build question about number_of_tables with answers:
            - Real number of tables or empty sentence
            - Random number
            - Random number
        """

        # STEP 1: Get value for rigth answer
        itdoesnothave = u'No hay tablas en el documento'
        right_value = self._driver.get_number_of_tables(itdoesnothave)

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Cuál es el número total de tablas en el documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(1, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)

# ---------------------------- FIND AND REPLACE -------------------------------


    def _q_find_text_count_matches(self, docname):
        """ Find text and count matches:
            - True number of matches
            - Random number of matches
            - Random number of matches
        """


        # STEP 1: Get value for rigth answer
        for attempt in range(1, 25):
            searchstr = self._random.values(IDIOMS, name="words")[0]
            right_value = self._driver.find_text(
                text=searchstr, sensitive=False, wholeword=False, likeness=False)
            print u'Search attempt: %5d => %s' % (attempt, searchstr)
            if right_value and right_value > 0:
                break

        if right_value:

            # STEP 3: Create question and add register it in test
            pattern = (u'¿Cuántas veces aparece el texto «%s» en el cuerpo '
                       u'del documento %s?')
            question = Question(pattern % (searchstr, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            rightstr = right_value or u'No aparece en el documento'
            question.add_answer(rightstr, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(0, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    def _q_find_text_count_chars(self, docname):
        """ Find text and count characters (with spaces)
            - True number of characters (with spaces) in all matches
            - Random number of matches
            - Random number of matches
        """


        # STEP 1: Get value for rigth answer
        for attempt in range(1, 25):
            searchstr = self._random.values(IDIOMS, name="words")[0]
            right_value = self._driver.find_text(
                text=searchstr, sensitive=False, wholeword=False, likeness=False)
            print u'Search attempt: %5d => %s' % (attempt, searchstr)
            if right_value and right_value > 0:
                break

        if right_value:
            right_value = right_value * len(searchstr)

            # STEP 3: Create question and add register it in test
            pattern = (u'¿Cuántas caracteres, con espacios, suman todas las '
                       u'apariciones del texto «%s» en el cuerpo del documento %s?')
            question = Question(pattern % (searchstr, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            rightstr = right_value or u'No aparece en el documento'
            question.add_answer(rightstr, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(0, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    def _q_find_text_count_notspaces(self, docname):
        """ Find text and count characters (without spaces)
            - True number of characters (without spaces) in all matches
            - Random number of matches
            - Random number of matches
        """


        # STEP 1: Get value for rigth answer
        for attempt in range(1, 25):
            searchstr = self._random.values(IDIOMS, name="words")[0]
            right_value = self._driver.find_text(
                text=searchstr, sensitive=False, wholeword=False, likeness=False)
            print u'Search attempt: %5d => %s' % (attempt, searchstr)
            if right_value and right_value > 0:
                break

        if right_value:
            right_value = right_value * len(searchstr.replace(' ', ''))

            # STEP 3: Create question and add register it in test
            pattern = (u'¿Cuántas caracteres, sin espacios, suman todas las '
                       u'apariciones del texto «%s» en el cuerpo del documento %s?')
            question = Question(pattern % (searchstr, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            rightstr = right_value or u'No aparece en el documento'
            question.add_answer(rightstr, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(0, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    def _q_find_text_count_words(self, docname):
        """ Find text and count words:
            - True number of words in all matches
            - Random number of matches
            - Random number of matches
        """


        # STEP 1: Get value for rigth answer
        for attempt in range(1, 25):
            searchstr = self._random.values(IDIOMS, name="words")[0]
            numwords = len(searchstr.split(' '))
            right_value = self._driver.find_text(
                text=searchstr, sensitive=False, wholeword=False, likeness=False)
            print u'Search attempt: %5d => %s' % (attempt, searchstr)
            if right_value and right_value > 0:
                break

        if right_value:
            right_value = right_value * numwords

            # STEP 3: Create question and add register it in test
            pattern = (u'¿Cuántas palabras suman todas las apariciones del '
                       u'texto «%s» en el cuerpo del documento %s?')
            question = Question(pattern % (searchstr, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            rightstr = right_value or u'No aparece en el documento'
            question.add_answer(rightstr, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(0, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    def _q_find_word_count_words(self, docname):
        """ Find text and count words:
            - True number of words in all matches
            - Random number of matches
            - Random number of matches
        """

        # STEP 1: Get value for rigth answer
        for attempt in range(1, 25):
            searchstr = self._random.values(WORDS, name="words")[0]
            right_value = self._driver.find_text(
                text=searchstr, sensitive=False, wholeword=True, likeness=False)
            print u'Search attempt: %5d => %s' % (attempt, searchstr)
            if right_value and right_value > 0:
                break

        if right_value:
            right_value = right_value

            # STEP 3: Create question and add register it in test
            pattern = (u'¿Cuántas veces aparece la palabra «%s» en el cuerpo '
                       u'del documento %s?')
            question = Question(pattern % (searchstr, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            rightstr = right_value or u'No aparece en el documento'
            question.add_answer(rightstr, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(0, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    def _q_find_word_count_notspaces(self, docname):
        """ Find text and count words:
            - True number of words in all matches
            - Random number of matches
            - Random number of matches
        """

        # STEP 1: Get value for rigth answer
        for attempt in range(1, 25):
            searchstr = self._random.values(WORDS, name="words")[0]
            right_value = self._driver.find_text(
                text=searchstr, sensitive=False, wholeword=True, likeness=False)
            print u'Search attempt: %5d => %s' % (attempt, searchstr)
            if right_value and right_value > 0:
                break

        if right_value:
            right_value = right_value * len(searchstr.replace(' ', ''))

            # STEP 3: Create question and add register it in test
            pattern = (u'¿Cuántos caracteres suman todas las apariciones '
                       u'de la palabra «%s» en el cuerpo del documento %s?')
            question = Question(pattern % (searchstr, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            rightstr = right_value or u'No aparece en el documento'
            question.add_answer(rightstr, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(0, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    def _q_find_case_text_count_matches(self, docname):
        """ Find text and count matches (case sensitive)
            - True number of matches
            - Random number of matches
            - Random number of matches
        """

        # STEP 1: Get value for rigth answer
        for attempt in range(1, 25):
            searchstr = self._random.values(IDIOMS, name="words")[0]
            right_value = self._driver.find_text(
                text=searchstr, sensitive=True, wholeword=False, likeness=False)
            print u'Search attempt: %5d => %s' % (attempt, searchstr)
            if right_value and right_value > 0:
                break

        if right_value:

            # STEP 3: Create question and add register it in test
            pattern = (u'¿Cuántas veces aparece, haciendo distinción entre '
                       u'mayúsculas de minúsculas, el texto «%s» en el cuerpo '
                       u'del documento %s?')
            question = Question(pattern % (searchstr, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            rightstr = right_value or u'No aparece en el documento'
            question.add_answer(rightstr, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(0, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    def _q_find_case_text_count_chars(self, docname):
        """ Find text and count characters (with spaces) (case sensitive)
            - True number of characters (with spaces) in all matches
            - Random number of matches
            - Random number of matches
        """


        # STEP 1: Get value for rigth answer
        for attempt in range(1, 25):
            searchstr = self._random.values(IDIOMS, name="words")[0]
            right_value = self._driver.find_text(
                text=searchstr, sensitive=True, wholeword=False, likeness=False)
            print u'Search attempt: %5d => %s' % (attempt, searchstr)
            if right_value and right_value > 0:
                break

        if right_value:
            right_value = right_value * len(searchstr)

            # STEP 3: Create question and add register it in test
            pattern = (u'¿Cuántas caracteres, con espacios, suman todas las '
                       u'apariciones del texto «%s» en el cuerpo del documento '
                       u'%s haciendo distinción entre mayúsculas y minúsculas?')
            question = Question(pattern % (searchstr, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            rightstr = right_value or u'No aparece en el documento'
            question.add_answer(rightstr, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(0, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    def _q_find_case_text_count_notspaces(self, docname):
        """ Find text and count characters (without spaces) (case sensitive)
            - True number of characters (without spaces) in all matches
            - Random number of matches
            - Random number of matches
        """


        # STEP 1: Get value for rigth answer
        for attempt in range(1, 25):
            searchstr = self._random.values(IDIOMS, name="words")[0]
            right_value = self._driver.find_text(
                text=searchstr, sensitive=True, wholeword=False, likeness=False)
            print u'Search attempt: %5d => %s' % (attempt, searchstr)
            if right_value and right_value > 0:
                break

        if right_value:
            right_value = right_value * len(searchstr.replace(' ', ''))

            # STEP 3: Create question and add register it in test
            pattern = (u'¿Cuántas caracteres, sin espacios, suman todas las '
                       u'apariciones del texto «%s» en el cuerpo del documento '
                       u'%s haciendo distinción entre mayúsculas y minúsculas?')
            question = Question(pattern % (searchstr, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            rightstr = right_value or u'No aparece en el documento'
            question.add_answer(rightstr, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(0, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    def _q_find_case_text_count_words(self, docname):
        """ Find text and count words (case sensitive)
            - True number of words in all matches
            - Random number of matches
            - Random number of matches
        """


        # STEP 1: Get value for rigth answer
        for attempt in range(1, 25):
            searchstr = self._random.values(IDIOMS, name="words")[0]
            numwords = len(searchstr.split(' '))
            right_value = self._driver.find_text(
                text=searchstr, sensitive=True, wholeword=False, likeness=False)
            print u'Search attempt: %5d => %s' % (attempt, searchstr)
            if right_value and right_value > 0:
                break

        if right_value:
            right_value = right_value * numwords

            # STEP 3: Create question and add register it in test
            pattern = (u'¿Cuántas palabras suman todas las apariciones del '
                       u'texto «%s» en el cuerpo del documento %s '
                       u'haciendo distinción entre mayúsculas y minúsculas?')
            question = Question(pattern % (searchstr, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            rightstr = right_value or u'No aparece en el documento'
            question.add_answer(rightstr, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(0, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    def _q_find_case_word_count_words(self, docname):
        """ Find text and count words (case sensitive)
            - True number of words in all matches
            - Random number of matches
            - Random number of matches
        """

        # STEP 1: Get value for rigth answer
        for attempt in range(1, 25):
            searchstr = self._random.values(WORDS, name="words")[0]
            right_value = self._driver.find_text(
                text=searchstr, sensitive=True, wholeword=True, likeness=False)
            print u'Search attempt: %5d => %s' % (attempt, searchstr)
            if right_value and right_value > 0:
                break

        if right_value:
            right_value = right_value

            # STEP 3: Create question and add register it in test
            pattern = (u'¿Cuántas veces aparece, haciendo distinción entre '
                       u'mayúsculas y minúsculas, la palabra «%s» en el cuerpo '
                       u'del documento %s?')
            question = Question(pattern % (searchstr, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            rightstr = right_value or u'No aparece en el documento'
            question.add_answer(rightstr, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(0, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    def _q_find_case_word_count_notspaces(self, docname):
        """ Find text and count words (case sensitive)
            - True number of words in all matches
            - Random number of matches
            - Random number of matches
        """

        # STEP 1: Get value for rigth answer
        for attempt in range(1, 25):
            searchstr = self._random.values(WORDS, name="words")[0]
            right_value = self._driver.find_text(
                text=searchstr, sensitive=True, wholeword=True, likeness=False)
            print u'Search attempt: %5d => %s' % (attempt, searchstr)
            if right_value and right_value > 0:
                break

        if right_value:
            right_value = right_value * len(searchstr.replace(' ', ''))

            # STEP 3: Create question and add register it in test
            pattern = (u'¿Cuántos caracteres suman todas las apariciones '
                       u'de la palabra «%s» en el cuerpo del documento %s '
                       u'haciendo distinción entre mayúsculas y minúsculas?')
            question = Question(pattern % (searchstr, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            rightstr = right_value or u'No aparece en el documento'
            question.add_answer(rightstr, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(0, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    # def _q_find_likeness_count_word(self, docname):
    #     """ Find similar words and count words
    #         - True number of words in all matches
    #         - Random number of matches
    #         - Random number of matches
    #     """

    #     # STEP 1: Get value for rigth answer
    #     for attempt in range(1, 25):
    #         searchstr = self._random.values(WORDS, name="words")[0]
    #         right_value = self._driver.find_text(
    #             text=searchstr, sensitive=False, wholeword=False, likeness=True)
    #         print u'Search attempt: %5d => %s' % (attempt, searchstr)
    #         if right_value and right_value > 0:
    #             break

    #     if right_value:
    #         right_value = right_value * len(searchstr.replace(' ', ''))

    #         # STEP 3: Create question and add register it in test
    #         pattern = (u'¿Cuántas palabras similares a «%s» aparecen en el '
    #                    u'cuerpo del documento %s?')
    #         question = Question(pattern % (searchstr, docname))
    #         self._questions.append(question)

    #         # STEP 4: Add the right answer
    #         rightstr = right_value or u'No aparece en el documento'
    #         question.add_answer(rightstr, True)

    #         # STEP 5: Build and add wrong answers
    #         number = int(right_value)
    #         wrongans = self._random.integer(0, number+10, number, 2)

    #         question.add_answer(unicode(wrongans[0]), False)
    #         question.add_answer(unicode(wrongans[1]), False)


    def _q_find_likeness_count_chars(self, docname):
        """ Search for text and count chars with spaces
            - True number of words in all matches
            - Random number of matches
            - Random number of matches
        """

        # STEP 1: Get value for rigth answer
        for attempt in range(1, 25):
            searchstr = self._random.values(IDIOMS, name="words")[0]
            right_value = self._driver.find_text(
                text=searchstr, sensitive=False, wholeword=False, likeness=True)
            print u'Search attempt: %5d => %s' % (attempt, searchstr)
            if right_value and right_value > 0:
                break

        if right_value:
            right_value = right_value * len(searchstr.replace(' ', ''))

            # STEP 3: Create question and add register it in test
            pattern = (u'¿Cuántos caracteres suman todas las apariciones de '
                       u'textos similares a «%s» en el cuerpo del documento %s?')
            question = Question(pattern % (searchstr, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            rightstr = right_value or u'No aparece en el documento'
            question.add_answer(rightstr, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(0, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    def _q_find_likeness_count_nospaces(self, docname):
        """ Search for text and count chars without spaces
            - True number of words in all matches
            - Random number of matches
            - Random number of matches
        """

        # STEP 1: Get value for rigth answer
        for attempt in range(1, 25):
            searchstr = self._random.values(IDIOMS, name="words")[0]
            right_value = self._driver.find_text(
                text=searchstr, sensitive=False, wholeword=False, likeness=True)
            print u'Search attempt: %5d => %s' % (attempt, searchstr)
            if right_value and right_value > 0:
                break

        if right_value:
            right_value = right_value * len(searchstr.replace(' ', ''))

            # STEP 3: Create question and add register it in test
            pattern = (u'¿Cuántos caracteres, sin espacios, suman todas las '
                       u'apariciones de las palabras similares a «%s» '
                       u'existentesen el cuerpo del documento %s?')
            question = Question(pattern % (searchstr, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            rightstr = right_value or u'No aparece en el documento'
            question.add_answer(rightstr, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            wrongans = self._random.integer(0, number+10, number, 2)

            question.add_answer(unicode(wrongans[0]), False)
            question.add_answer(unicode(wrongans[1]), False)


    def _q_find_special(self, docname):
        """ Search for special chars in the document
        """
        specials = self._driver.get_search_specials()
        for special, description in specials.iteritems():

            print u'=>', special, description
            right_value = self._driver.find_special(special)

            if right_value and (right_value > 0 or self._random.bool(0.25)):
                pattern = (u'¿Cuántas apariciones totales de «%s» podemos '
                           u'hallar en el documento %s?')
                question = Question(pattern % (description, docname))
                self._questions.append(question)

                # STEP 4: Add the right answer
                rightstr = right_value or u'No aparece en el documento'
                question.add_answer(rightstr, True)

                # STEP 5: Build and add wrong answers
                number = int(right_value)
                wrongans = self._random.integer(0, number+10, number, 2)

                question.add_answer(unicode(wrongans[0]), False)
                question.add_answer(unicode(wrongans[1]), False)

# ------------------------------- PAGE DESIGN ---------------------------------

    def _q_page_margins(self, docname):
        """ The four questions for margins """

        margins = self._driver.get_page_margins()

        translations = {
            u'top' : u'superior',
            u'right' : u'derecho',
            u'bottom' : u'inferior',
            u'left' : u'izquierdo'
        }

        for name, value in margins.iteritems():
            print 'margin ', name, ':', value
            translation = translations[name]
            pattern = (u'¿Cuál es el margen %s del documento %s?')
            question = Question(pattern % (translation, docname))
            self._questions.append(question)

            valuestr = u'{} cm'.format(value)
            question.add_answer(valuestr, True)

            number = value
            wrongans = self._random.float(0.25, number+1.0, number, 2, 0.01)

            valuestr = u'{} cm'.format(wrongans[0])
            question.add_answer(valuestr, False)

            valuestr = u'{} cm'.format(wrongans[1])
            question.add_answer(valuestr, False)

    # def _q_page_gutter_margin(self, docname):
    #     """ Gutter margin distance and position """

    #     gutters = self._driver.get_gutter_margin()

    #     # STEP 1: Distance question
    #     pattern = u'¿Cuál es el margen de encuadernación del documento %s?'
    #     question = Question(pattern % docname)
    #     self._questions.append(question)

    #     valuestr = u'{} cm'.format(gutters['distance'])
    #     question.add_answer(valuestr, True)

    #     number = float(gutters['distance'])
    #     wrongans = self._random.float(0.25, number+1.0, number, 2, 0.01)

    #     valuestr = u'{} cm'.format(wrongans[0])
    #     question.add_answer(valuestr, False)

    #     valuestr = u'{} cm'.format(wrongans[1])
    #     question.add_answer(valuestr, False)

    #     # # STEP 2: Position question
    #     # pattern = u'¿Cuál es la posición del margen interno en el documento %s?'
    #     # question = Question(pattern % docname)
    #     # self._questions.append(question)

    #     # valuestr = u'{} cm'.format(gutters['position'])
    #     # question.add_answer(gutters['position'], True)

    #     # valuestr = u'{} cm'.format(wrongans[0])
    #     # question.add_answer(valuestr, False)

    #     # valuestr = u'{} cm'.format(wrongans[1])
    #     # question.add_answer(valuestr, False)



# -----------------------------------------------------------------------------


    def _docx2test(self):
        """ Performs the conversion from docx to pdf
        """

        new_path = os.path.join(self.dirname, self.filename+'.pdf')
        self._driver = DocDriver(self.abspath)
        docname = os.path.basename(new_path)

        methods = [k for k in dir(self) if k[:3] == '_q_']

        for name in methods:
            try:
                getattr(self, name)(self.basename)
            except Exception as ex:
                pass

        print 'finish'

        for question in self._questions:
            question.shuffle()
            question.add_answer(u'Ninguna de las anteriores es correcta', False)

        random.shuffle(self._questions)

        xml = '\n'.join([k.to_xml() for k in self._questions]).encode(u'utf8')

        textname = os.path.splitext(self.basename)[0]
        text_file = open(textname + u'.xml', "w")
        text_file.write(XMLOUTER.format(xml))
        text_file.close()


    def _is_file(self):
        """ Check if given path is a valid file path
        """
        return os.path.isfile(self.abspath)


    def main(self):
        """ The main application behavior, this method should be used to
        start the application.
        """

        self._argparse()
        if self._is_file():
            self._docx2test()




# --------------------------- SCRIPT ENTRY POINT ------------------------------

App().main()
