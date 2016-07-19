# -*- coding: utf-8 -*-
#pylint: disable=I0011,W0703,R0903
""" Merge serveral PDF files into one file
"""

# --------------------------- REQUIRED LIBRARIES ------------------------------

import argparse
import os

from PyPDF2 import PdfFileMerger, PdfFileReader


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
        self.files = None
        self.title = None
        self.author = None
        self.subject = None
        self.keywords = None

    def _argparse(self):
        """ Detines an user-friendly command-line interface and proccess its
        arguments.
        """

        description = u'Merge serveral PDF files into one file'

        parser = argparse.ArgumentParser(description)
        parser.add_argument('file', metavar='file', type=str,
                            help='Output file path')


        parser.add_argument('-f', '--files', nargs='+', type=str, metavar='files',
                            help='PDF files will be merged')

        parser.add_argument('-t', '--title', type=str, metavar='title', default=u'',
                            help='Title will be used')

        parser.add_argument('-a', '--author', type=str, metavar='author', default=u'',
                            help='Author will be used')

        parser.add_argument('-s', '--subject', type=str, metavar='subject', default=u'',
                            help='Subject will be used')

        parser.add_argument('-k', '--keywords', type=str, metavar='keywords', default=u'',
                            help='Keywords will be used')

        args = parser.parse_args()

        self.abspath = os.path.abspath(args.file)
        self.basename = os.path.basename(self.abspath)
        self.dirname = os.path.dirname(self.abspath)
        self.filename = self.basename and os.path.splitext(self.basename)[0]

        self.files = args.files

        self.title = args.title
        self.author = args.author
        self.subject = args.subject
        self.keywords = args.keywords

    def _mergepdf(self):
        """ Performs the conversion from docx to pdf
        """

        new_path = os.path.join(self.dirname, self.filename+'.pdf')

        try:

            merger = PdfFileMerger()

            for fname in self.files:
                print fname
                fpath = os.path.abspath(fname)
                ffile = PdfFileReader(open(fpath, 'rb'))
                merger.append(ffile)

            metadata = {
                u'/Title': self.title,
                u'/Author': self.author,
                u'/Subject': self.subject,
                u'/Keywords': self.keywords
            }

            merger.addMetadata(metadata)

            merger.write(self.abspath)

        except Exception as ex:
            print ex
        else:
            str_new_path = new_path.decode('utf-8', 'ignore')
            print u'New file %s has been written.' % str_new_path

    def main(self):
        """ The main application behavior, this method should be used to
        start the application.
        """

        self._argparse()
        self._mergepdf()


# --------------------------- SCRIPT ENTRY POINT ------------------------------

App().main()
