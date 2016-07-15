# -*- coding: utf-8 -*-
#pylint: disable=I0011,W0703,R0903,W0110,W0141,C0301
""" Builds tests about documents
"""

# --------------------------- REQUIRED LIBRARIES ------------------------------

from answer_class import Answer
from question_class import Question
from docdriver_class import DocDriver

import argparse
import os
import win32com.client
import random
import locale

from datetime import datetime, timedelta

from pprint import pprint


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


    def _docx2test(self):
        """ Performs the conversion from docx to pdf
        """

        new_path = os.path.join(self.dirname, self.filename+'.pdf')
        questions = []
        driver = DocDriver(self.abspath)

        try:

            authors = (
                u'Pilar Lopez Perez',
                u'Jose Luis Mendez Diaz',
                u'Jaime Abilleira Moldes',
                u'Jorge Soto Garcia'
            )

            companies = (
                u'Academia Postal',
                u'Academia Postal 3',
                u'Postal',
                u'Postal 3'
            )

            categories = {
                u'Oposiciones',
                u'Ofimatica',
                u'Informatica'
            }

            name = driver.get_name()
            extension = driver.get_extension()
            docname = u'{}.{}'.format(name, extension)

            if name:
                question = Question(u'Cuál es la cadena de texto que se utiliza para designar el archivo %s?' % docname)
                question.add_answer(name, True)
                question.add_answer(extension, True)
                question.add_answer('Fichero', True)
                questions.append(question)

            import locale
            locale.setlocale(locale.LC_ALL, '')

            if extension:
                question = Question(u'¿Cuál es la extension del archivo %s?' % docname)
                question.add_answer(extension, True)
                question.add_answer(u'No tiene extensión', False)
                size = locale.format("%d", os.path.getsize(self.abspath), grouping=True)
                question.add_answer(size + ' Bytes', False)
                questions.append(question)

            doc_type = driver.get_type()
            if doc_type:
                question = Question(u'¿Qué tipo de archivo corresponde al documento %s?' % docname)
                question.add_answer(doc_type, True)
                allowed = driver.get_allowed_types()
                allowed = {k:v for k, v in allowed.iteritems() if k != extension}
                for item in list(allowed.values())[:2]:
                    question.add_answer(item, False)
                questions.append(question)

            title = driver.get_title()
            if title:
                question = Question(u'¿Cuál es el título del documento %s?' % docname)
                question.add_answer(title, True)
                question.add_answer(name, True)
                question.add_answer(u'El texto del primer párrafo del documento', True)
                questions.append(question)

            author = driver.get_author()
            if author:
                question = Question(u'¿Quién figura como autor del documento %s?' % docname)
                question.add_answer(author, True)
                for auth in filter(lambda x: x != author, authors)[:2]:
                    question.add_answer(auth, True)
                questions.append(question)

            last_author = driver.get_last_author()
            if last_author and last_author != 'None':
                question = Question(u'¿Quién figura como último autor del documento %s?' % docname)
                question.add_answer(last_author, True)
                for auth in filter(lambda x: x != last_author, authors)[:2]:
                    question.add_answer(auth, True)
                questions.append(question)

            manager = driver.get_manager()
            if manager and manager != 'None':
                question = Question(u'¿Quién figura como administrador del documento %s?' % docname)
                question.add_answer(manager, True)
                for auth in filter(lambda x: x != manager, authors)[:2]:
                    question.add_answer(auth, True)
                questions.append(question)

            company = driver.get_company()
            if company and company != 'None':
                question = Question(u'¿Qué compañía aparece asociada en las propiedades del documento %s?' % docname)
                question.add_answer(company, True)
                for comp in filter(lambda x: x != company, companies)[:2]:
                    question.add_answer(comp, True)
                questions.append(question)

            category = driver.get_category()
            if category and category != 'None':
                question = Question(u'¿Qué categoría ha sido asignada al documento %s?' % docname)
                question.add_answer(category, True)
                for cat in filter(lambda x: x != category, categories)[:2]:
                    question.add_answer(cat, True)
                questions.append(question)

            keywords = driver.get_keywords()
            if keywords:
                question = Question(u'¿Cuáles son las palabras clave empleadas en el documento %s?' % docname)
                question.add_answer(keywords, True)
                question.add_answer(u'No figuran palabras clave asociadas al documento', True)
                questions.append(question)

            comments = driver.get_comments()
            if comments and comments != 'None':
                question = Question(u'¿Qué comentarios figuran en las propiedades del documento %s?' % docname)
                question.add_answer(comments, True)
                question.add_answer(u'No figuran comentarios en las propiedades del documento', True)
                questions.append(question)

            template = driver.get_template()
            if template and template != 'None':
                question = Question(u'¿A partir de qué plantilla ha sido creado el documento %s?' % docname)
                tname = os.path.splitext(template)[0]
                tnames = [tname + '.dotx', tname + '.dotm', tname + '.dot']
                for tpl in tnames:
                    question.add_answer(tpl, True)
                questions.append(question)

            revision_number = driver.get_revision_number()
            if revision_number and revision_number != 'None':
                question = Question(u'¿Cuál es el número de revisión del documento %s?' % docname)
                question.add_answer(revision_number, True)
                available = random.sample(xrange(int(revision_number)* 2 + 5), 3)
                for num in available[:2]:
                        question.add_answer(unicode(str(num)), False)
                questions.append(question)

            last_print_date = driver.get_last_print_date()
            if last_print_date and last_print_date != 'None':
                question = Question(u'¿Cuándo se imprimió por última vez el documento %s?' % docname)
                question.add_answer(last_print_date, True)
                questions.append(question)

            creation_date = driver.get_creation_date()
            if creation_date and creation_date != 'None':
                question = Question(u'¿En qué fecha fue creado el documento %s?' % docname)
                # 15/07/2016 0:26:00
                date_sn = datetime.strptime(creation_date, '%x %X').date()
                date_str = date_sn.strftime('%x')
                question.add_answer(date_str, True)
                date_str = (date_sn + timedelta(days=random.randint(-15, 1))).strftime('%x')
                question.add_answer(date_str, True)
                date_str = (date_sn + timedelta(days=random.randint(1, 15))).strftime('%x')
                question.add_answer(date_str, True)
                questions.append(question)

            last_save_time = driver.get_last_save_time()
            if last_save_time and last_save_time != 'None':
                question = Question(u'¿Cuándo se guardó por última vez el documento %s?' % docname)
                date_sn = datetime.strptime(last_save_time, '%x %X').date()
                date_str = date_sn.strftime('%x')
                question.add_answer(date_str, True)
                date_str = (date_sn + timedelta(days=random.randint(-15, 0))).strftime('%x')
                question.add_answer(date_str, True)
                date_str = (date_sn + timedelta(days=random.randint(0, 15))).strftime('%x')
                question.add_answer(date_str, True)
                questions.append(question)

            total_editing_time = driver.get_total_editing_time()
            if total_editing_time and total_editing_time != 'None':
                question = Question(u'¿Cuál ha sido el tiempo total de edición del documento %s?' % docname)
                total_editing_time = int(total_editing_time)
                question.add_answer(total_editing_time, True)
                total_editing_time = total_editing_time + random.randint(1, 15)
                question.add_answer(total_editing_time, True)
                total_editing_time = total_editing_time + random.randint(1, 15)
                question.add_answer(total_editing_time, True)
                questions.append(question)

            number_of_pages = driver.get_number_of_pages()
            if number_of_pages and number_of_pages != 'None':
                question = Question(u'¿Cuál es el número total de páginas del documento %s?' % docname)
                nval = int(number_of_pages)
                question.add_answer(nval, True)
                question.add_answer(nval + (- 1 if  nval > 1 else 1), True)
                question.add_answer(nval + (+ 1 if  nval > 1 else 2), True)
                questions.append(question)

            number_of_words = driver.get_number_of_words()
            if number_of_words and number_of_words != 'None':
                question = Question(u'¿Cuál es el número total de palabras del documento %s?' % docname)
                number_of_words = int(number_of_words)
                question.add_answer(number_of_words, True)
                nval = 999 if not number_of_words < 1 else random.randint(0, number_of_words - 1)
                question.add_answer(nval, True)
                nval = 9999 if not number_of_words else number_of_words + random.randint(10, 100)
                question.add_answer(nval, True)
                questions.append(question)

            number_of_characters = driver.get_number_of_characters()
            if number_of_characters and number_of_characters != 'None':
                question = Question(u'¿Cuál es el número total de caracteres, sin contar espacios, del documento %s?' % docname)
                number_of_characters = int(number_of_characters)
                question.add_answer(number_of_characters, True)
                nval = 999 if not number_of_characters < 1 else random.randint(0, number_of_characters - 1)
                question.add_answer(nval, True)
                nval = 9999 if not number_of_characters else number_of_characters + random.randint(10, 100)
                question.add_answer(nval, True)
                questions.append(question)

            number_of_lines = driver.get_number_of_lines()
            if number_of_lines and number_of_lines != 'None':
                question = Question(u'¿Cuál es el número total de líneas existentes en el documento %s?' % docname)
                number_of_lines = int(number_of_lines)
                question.add_answer(number_of_lines, True)
                nval = 999 if not number_of_lines < 1 else random.randint(0, number_of_lines - 1)
                question.add_answer(nval, True)
                nval = 9999 if not number_of_lines else number_of_lines + random.randint(10, 100)
                question.add_answer(nval, True)
                questions.append(question)

            number_of_paragraphs = driver.get_number_of_paragraphs()
            if number_of_paragraphs and number_of_paragraphs != 'None':
                question = Question(u'¿Cuál es el número total de párrafos existentes en el documento %s?' % docname)
                number_of_paragraphs = int(number_of_paragraphs)
                question.add_answer(number_of_paragraphs, True)
                nval = 999 if not number_of_paragraphs < 1 else random.randint(0, number_of_paragraphs - 1)
                question.add_answer(nval, True)
                nval = 9999 if not number_of_paragraphs else number_of_paragraphs + random.randint(10, 100)
                question.add_answer(nval, True)
                questions.append(question)

            number_of_characters = driver.get_number_of_characters()
            if number_of_characters and number_of_characters != 'None':
                question = Question(u'¿Cuál es el número total de caracteres, contando espacios, del documento %s?' % docname)
                number_of_characters = int(number_of_characters)
                question.add_answer(number_of_characters, True)
                nval = 999 if not number_of_characters < 1 else random.randint(0, number_of_characters - 1)
                question.add_answer(nval, True)
                nval = 9999 if not number_of_characters else number_of_characters + random.randint(10, 100)
                question.add_answer(nval, True)

            number_of_bytes = driver.get_number_of_bytes()
            if number_of_bytes and number_of_bytes != 'None':
                question = Question(u'¿Cuál es, en bytes, el tamaño del documento %s?' % docname)
                number_of_bytes = int(number_of_bytes)
                question.add_answer(number_of_paragraphs, True)
                nval = 999 if not number_of_bytes < 1 else random.randint(0, number_of_bytes - 1)
                question.add_answer(nval, True)
                nval = 9999 if not number_of_bytes else number_of_bytes + random.randint(10, 100)
                question.add_answer(nval, True)
                questions.append(question)

            hyperlink_base = driver.get_hyperlink_base()
            if hyperlink_base and hyperlink_base != 'None':
                question = Question(u'¿Qué dirección figura como base del hipervínculo para el documento %s?' % docname)
                question.add_answer(hyperlink_base, True)
                question.add_answer(u'http://www.google.com', True)
                question.add_answer(u'http://www.postal3.es', True)
                questions.append(question)

            index = 0
            for question in questions:
                question.add_answer(u'Ninguna de las anteriores es correcta.', False)
                print question.to_xml(index).encode(self._os_encoding, errors='replace')
                index = index + 1

        except Exception as ex:
            print ex
        else:
            str_new_path = new_path.decode('utf-8', 'ignore')
            print u'New file %s has been written.' % str_new_path

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
