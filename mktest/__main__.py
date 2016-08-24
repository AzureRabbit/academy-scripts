# -*- coding: utf-8 -*-
###############################################################################
#    License, author and contributors information in:                         #
#    __openerp__.py file at the root folder of this module.                   #
###############################################################################
#pylint: disable=I0011,W0703,R0903


""" Build test from Word documents
"""

# --------------------------- REQUIRED LIBRARIES ------------------------------

import argparse
import os
import win32com.client

from lib.doc_search import DocSearch
from lib.rnd_values import RandomValues

from jinja2 import Template
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
        self._path = None
        self._word = None
        self._doc = None
        self._template = None
        self._out_file_path = None

    def _argparse(self):
        """ Detines an user-friendly command-line interface and proccess its
        arguments.
        """

        # STEP 1: Define the arbument parser
        description = (u'Build a test with questions about an existing '
                       u'word document')
        parser = argparse.ArgumentParser(description)

        # STEP 2: Determine positional arguments
        parser.add_argument('file', metavar='file', type=str,
                            help='path of the document will be used')

        # STEP 3: Determine non positional arguments
        parser.add_argument('-t', '--template', type=str, dest='template',
                            default=None, help='jinja2 template file')

        parser.add_argument('-o', '--out', type=str, dest='out',
                            default=None, help='output file path')

        args = parser.parse_args()

        self._path = os.path.abspath(args.file)

        if args.template:
            abspath = os.path.abspath(args.template)
            with open(abspath) as tfile:
                buff = tfile.read()
                self._template = Template(buff.decode('utf-8', errors='replace'))

        if args.out:
            self._out_file_path = os.path.abspath(args.out)


    def _main_method_name(self):
        """ Main method docstring
        """

        try:

            self._word = win32com.client.DispatchEx("Word.Application")
            self._doc = self._word.Documents.Open(self._path)

            docsearch = DocSearch(self._word, self._doc)

            # for prop in dir(docprops):
            #     if prop[:1] != '_':
            #         value = getattr(docprops, prop, None)
            #         print prop, value

            if self._template:
                out = self._template.render(
                    docname=self._doc.Name,
                    random=RandomValues(),
                    search=docsearch)

                if out and self._out_file_path:
                    with open(self._out_file_path, 'w') as ofile:
                        string_for_output = out.encode('utf8', 'replace')
                        ofile.write(string_for_output)


        except Exception as ex:
            print ex
        else:
            print u'New file %s has been readed.' % self._path
        finally:
            if self._doc:
                self._doc.Close()

            if self._word:
                self._word.Quit()

            self._doc, self._word = None, None


    def main(self):
        """ The main application behavior, this method should be used to
        start the application.
        """

        self._argparse()

        self._main_method_name()


# --------------------------- SCRIPT ENTRY POINT ------------------------------

App().main()
