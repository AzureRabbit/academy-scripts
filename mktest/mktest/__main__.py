# -*- coding: utf-8 -*-
#pylint: disable=I0011,W0703,R0903,W0110,W0141,C0301
""" Builds tests about documents
"""

# --------------------------- REQUIRED LIBRARIES ------------------------------

from answer_class import Answer
from question_class import Question
from docdriver_class import DocDriver
from random_values import RandomValues

import argparse
import os
import win32com.client
import random
import locale

from datetime import datetime, timedelta, date

from pprint import pprint
import locale
import re

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

    def _q_prop_title(self, docname):
        """ Build question about title with answers:
            - True document title
            - The document name
            - Literal "El texto del primer párrafo del documento"
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_title()
        if not right_value:
            right_value = self._random.unspecified(female=False)

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
            - True document author
            - Random name from list
            - Random name from list
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_author()

        if not right_value:
            right_value = self._random.unspecified(female=False)

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
            - True document author
            - Random name from list
            - Random name from list
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_last_author()
        if not right_value:
            right_value = self._random.unspecified(female=False)

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
            - True value for administrador
            - Random name from list
            - Random name from list
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_manager()

        if not right_value:
            right_value = self._random.unspecified(female=False)

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
            - True value for company
            - Random value from list
            - Random value form list
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_company()
        if not right_value:
            right_value = self._random.unspecified(female=True)

        # STEP 3: Create question and add register it in test
        text = (u'Qué compañía fugura en las propiedades del documento')
        question = Question(u'¿%s %s?' % (text, docname))
        self._questions.append(question)

        # STEP 4: Add the right answer
        question.add_answer(right_value, True)

        # STEP 5: Build and add wrong answers
        companies = self._random.company(right_value, 2)

        question.add_answer(companies[0], False)
        question.add_answer(companies[1], False)

    def _q_prop_category(self, docname):
        """ Build question about category with answers:
            - True value for category
            - Random value from list
            - Random value form list
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_category()

        if not right_value:
            right_value = self._random.unspecified(female=True)

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
            - The real keywords from document
            - Random keywords
            - Random keywords
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_keywords()
        if not right_value:
            right_value = self._random.unspecified(female=True, plural=True)

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
            - Real comment text string
            - Random topic
            - Random subject
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_comments()
        if not right_value:
            right_value = self._random.unspecified(female=False, plural=True)

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
            - Real template name
            - Random template name
            - Random template name
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_template()
        if not right_value:
            right_value = self._random.unspecified(female=True, plural=False)

        # STEP 3: Create question and add register it in test
        text = (u'A partir de qué plantilla ha sido creado el documento')
        question = Question(u'¿%s %s?' % (text, docname))
        self._questions.append(question)

        # STEP 4: Add the right answer
        question.add_answer(right_value, True)

        # STEP 5: Build and add wrong answers
        exception = os.path.splitext(right_value)
        extensions = self._random.values(
            ['dot', 'dotm', 'dotx', 'docx', 'doc'], exception, 2)

        question.add_answer(u'Normal' + extensions[0], True)
        question.add_answer(u'Normal' + extensions[1], True)

    def _q_prop_revision_number(self, docname):
        """ Build question about revision_number with answers:
            - Real revision number
            - Random revision number
            - Random revision number
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_revision_number()

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
            self._random.integer(1, number+10, number, 2)

            question.add_answer(unicode(right_value), True)
            question.add_answer(unicode(right_value), True)

    def _q_prop_last_print_date(self, docname):
        """ Build question about last_print_date with answers:
            - Real last print date
            - Random last print date
            - Random last print date
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_last_print_date()
        if not right_value:
            date_sn = date.today()
            right_value = self._random.unspecified(female=True, plural=False)
        else:
            date_sn = datetime.strptime(right_value, '%x %X').date()
            right_value = date_sn.strftime('%x')


        # STEP 3: Create question and add register it in test
        text = (u'Qué fecha figura como la última en la que se imprimió el documento')
        question = Question(u'¿%s %s?' % (text, docname))
        self._questions.append(question)

        # STEP 4: Add the right answer
        question.add_answer(right_value, True)

        # STEP 5: Build and add wrong answers
        dates = self._random.date(date_sn, 10, date_sn, 2)

        date_str = dates[0].strftime('%x')
        question.add_answer(date_str, True)

        date_str = dates[1].strftime('%x')
        question.add_answer(date_str, True)

    def _q_prop_creation_date(self, docname):
        """ Build question about creation_date with answers:
            - Real creation date
            - Random creation date
            - Random creation date
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_creation_date()
        if not right_value:
            date_sn = date.today()
            right_value = self._random.unspecified(female=True, plural=False)
        else:
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
        question.add_answer(date_str, True)

        date_str = dates[1].strftime('%x')
        question.add_answer(date_str, True)

    def _q_prop_last_save_time(self, docname):
        """ Build question about last save time with answers:
            - Real last save date
            - Random date
            - Random date
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_last_save_time()
        if not right_value:
            date_sn = date.today()
            right_value = self._random.unspecified(female=True, plural=False)
        else:
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
        question.add_answer(date_str, True)

        date_str = dates[1].strftime('%x')
        question.add_answer(date_str, True)


    def _q_prop_total_editing_time(self, docname):
        """ Build question about total editing time with answers:
            - Real editing time
            - Random number
            - Random number
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_total_editing_time()

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Cuál ha sido el tiempo total de edición del documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            self._random.integer(1, number+10, number, 2)

            question.add_answer(unicode(right_value), True)
            question.add_answer(unicode(right_value), True)

    def _q_prop_number_of_pages(self, docname):
        """ Build question about number_of_pages with answers:
            - Real number of pages
            - Random number
            - Random number
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_number_of_pages()

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Cuál es el número total de páginas del documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            self._random.integer(1, number+10, number, 2)

            question.add_answer(unicode(right_value), True)
            question.add_answer(unicode(right_value), True)


    def _q_prop_number_of_words(self, docname):
        """ Build question about number_of_words with answers:
            - Real number of words
            - Random number of words
            - Random number of words
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_number_of_words()

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Cuál es el número total de palabras del documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            self._random.integer(1, number+10, number, 2)

            question.add_answer(unicode(right_value), True)
            question.add_answer(unicode(right_value), True)

    def _q_prop_number_of_characters(self, docname):
        """ Build question about number of characters with answers:
            - Real number of characters
            - Random number of characters
            - Random number of characters
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_number_of_characters()

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
            self._random.integer(1, number+10, number, 2)

            question.add_answer(unicode(right_value), True)
            question.add_answer(unicode(right_value), True)

    def _q_prop_number_of_lines(self, docname):
        """ Build question about number_of_lines with answers:
            - Real number of lines
            - Random number
            - Random number
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_number_of_lines()

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
            self._random.integer(1, number+10, number, 2)

            question.add_answer(unicode(right_value), True)
            question.add_answer(unicode(right_value), True)


    def _q_prop_number_of_paragraphs(self, docname):
        """ Build question about number_of_paragraphs with answers:
            - Real number of paragraphs
            - Random number
            - Random number
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_number_of_paragraphs()

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
            self._random.integer(1, number+10, number, 2)

            question.add_answer(unicode(right_value), True)
            question.add_answer(unicode(right_value), True)


    def _q_prop_numchars_with_spaces(self, docname):
        """ Build question about number_of_characters with answers:
            - Real number of characters
            - Random number
            - Random number
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_number_of_chars_with_spaces()

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
            self._random.integer(1, number+10, number, 2)

            question.add_answer(unicode(right_value), True)
            question.add_answer(unicode(right_value), True)

    def _q_prop_number_of_bytes(self, docname):
        """ Build question about number_of_bytes with answers:
            - Real size in bytes
            - Random number
            - Random number
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_number_of_bytes()

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Cuál es, en bytes, el tamaño del documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            number = int(right_value)
            self._random.integer(1, number+10, number, 2)

            question.add_answer(unicode(right_value), True)
            question.add_answer(unicode(right_value), True)


    def _q_prop_hyperlink_base(self, docname):
        """ Build question about hyperlink_base with answers:
            - Real hyperlink base text string
            - Random hyperlink url
            - Random hyperlink url
        """

        # STEP 1: Get value for rigth answer
        right_value = self._driver.get_hyperlink_base()

        if right_value:

            # STEP 3: Create question and add register it in test
            text = (u'Qué dirección figura como base del hipervínculo '
                    u'para el documento')
            question = Question(u'¿%s %s?' % (text, docname))
            self._questions.append(question)

            # STEP 4: Add the right answer
            question.add_answer(right_value, True)

            # STEP 5: Build and add wrong answers
            subjects = self._random.subject(right_value, 2)

            hbase1 = u'www.%s.com' % re.sub(r'[^a-zA-Z0-9]+', '_', subjects[0])
            question.add_answer(hbase1, True)

            hbase2 = u'www.%s.es' % re.sub(r'[^a-zA-Z0-9]+', '_', subjects[0])
            question.add_answer(hbase2, True)

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

        xml = '\n'.join([k.to_xml() for k in self._questions]).encode(u'utf8')

        text_file = open("prueba.xml", "w")
        text_file.write(xml)
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
