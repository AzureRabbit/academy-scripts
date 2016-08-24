# -*- coding: utf-8 -*-
#pylint: disable=I0011,W0703
""" Set random document properties
"""

# --------------------------- REQUIRED LIBRARIES ------------------------------

import argparse
import lib.random_values
import os
import win32com.client

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
        self._file = None
        self._template = None

    def _argparse(self):
        """ Detines an user-friendly command-line interface and proccess its
        arguments.
        """

        # STEP 1: Define the arbument parser
        description = u'Set random properties for document'
        parser = argparse.ArgumentParser(description)

        # STEP 2: Determine positional arguments
        parser.add_argument('file', metavar='file', type=str,
                            help='The Microsoft Word document file')

        # STEP 3: Determine non positional arguments
        parser.add_argument('-t', '--template', type=str,
                            dest='template', default=None,
                            help='Template file to attach')

        args = parser.parse_args()

        self._file = os.path.abspath(args.file)
        print args.file
        self._template = os.path.abspath(args.template)

    @staticmethod
    def _has_prop(doc, propname):
        """ Try to get a BuiltInDocumentProperty from document
        @return: 1 => has property, 0 => has not property, -1 => error
        """
        result = 1

        try:
            prop = doc.BuiltInDocumentProperties[propname]
            result = 0 if prop == None or str(prop) == 'None' else 1
        except Exception as ex:
            result = -1
            print propname, ex

        return result

    @staticmethod
    def _set_prop(doc, propname, value):
        """ Try to get a BuiltInDocumentProperty from document
        @return: 1 => has property, 0 => has not property, -1 => error
        """
        try:
            doc.BuiltInDocumentProperties[propname] = value
        except Exception as ex:
            print u'{} could not be set to {} because {}'.format(
                propname, value, ex)

    def _set_random_props(self):
        """ Set random document properties
        """
        random = lib.random_values.RandomValues()

        try:

            word = win32com.client.DispatchEx("Word.Application")
            doc = word.Documents.Open(self._file)

            if self._has_prop(doc, u'Title') == 0:
                value = random.person()[0]
                self._set_prop(doc, u'Title', value)

            value = random.subject()[0]
            print value
            self._set_prop(doc, u'Subject', value)

            value = random.person()[0]
            self._set_prop(doc, u'Author', value)

            if self._has_prop(doc, u'Manager') == 0:
                value = random.person()[0]
                self._set_prop(doc, u'Manager', value)

            value = random.company()[0]
            self._set_prop(doc, u'Company', value)

            if self._has_prop(doc, u'Category') == 0:
                value = random.topic()[0]
                self._set_prop(doc, u'Category', value)

            keywords = random.topic(number=3)
            keywordsstr = u' '.join(keywords)
            self._set_prop(doc, u'Keywords', keywordsstr)

            if self._has_prop(doc, u'Comments') == 0:
                value = random.topic()[0]
                self._set_prop(doc, u'Comments', value)

            if self._has_prop(doc, u'Hyperlink base') == 0:
                urls = [
                    u'http://www.academiapostal.es',
                    u'http://vigo.academiapostal3.es',
                    u'http://www.postal3.es',
                    u'http://postal3.es',
                    u'http://sotogarcia.es',
                    u'http://sotogarcia.es'
                ]
                value = random.values(urls)[0]
                self._set_prop(doc, u'Hyperlink base', value)

            if not self._template == None:
                doc.AttachedTemplate = self._template

    #def get_default_tabstop(self, default=u''):

            doc.Save()

            doc.Close()
            word.Quit()

        except Exception as ex:
            print ex
        else:
            print u'File %s has been modified.' % os.path.basename(self._file)

    def main(self):
        """ The main application behavior, this method should be used to
        start the application.
        """

        self._argparse()

        self._set_random_props()


# --------------------------- SCRIPT ENTRY POINT ------------------------------

App().main()
