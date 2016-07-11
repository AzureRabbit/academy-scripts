# -*- coding: utf-8 -*-
#pylint: disable=I0011,W0703,R0903
""" Merge serveral PDF files into one file
"""

# --------------------------- REQUIRED LIBRARIES ------------------------------

import argparse
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

        description = u'Converts files to PDF and merges them into one'

        parser = argparse.ArgumentParser(description)
        parser.add_argument('file', metavar='file', type=str,
                            help='Output file path')

        parser.add_argument('-f', '--files', nargs='+', type=str, metavar='files',
                            help='Files will be converted and merged')

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

            acrobat = win32com.client.Dispatch('Acroexch.app')
            outdoc = None
            pdout = None

            for pdf_file in self.files:

                abspath = os.path.abspath(pdf_file)
                if os.path.isfile(abspath):

                    if outdoc == None:
                        outdoc = win32com.client.Dispatch('AcroExch.AVDoc')
                        outdoc.open(abspath, '')
                        pdout = outdoc.GetPDDoc()
                    else:

                        avdoc = win32com.client.Dispatch('AcroExch.AVDoc')
                        avdoc.open(abspath, '')
                        pddoc = avdoc.GetPDDoc()

                        out_pages = pdout.GetNumPages()
                        this_pages = pddoc.GetNumPages()

                        pdout.InsertPages(out_pages - 1, pddoc, 0, this_pages, True)

                        avdoc.Close(True)

            if outdoc != None:
                pdout.SetInfo('Title', self.title)
                pdout.SetInfo('Author', self.author)
                pdout.SetInfo('Subject', self.subject)
                pdout.SetInfo('Keywords', self.keywords)

                pdout.SetPageMode(1) # One page

                pdout.Save(1, self.abspath)

                outdoc.SetTitle(self.title)
                outdoc.Close(True)

            acrobat.Exit()

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
