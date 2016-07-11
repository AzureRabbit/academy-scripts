# -*- coding: utf-8 -*-
#pylint: disable=I0011,W0703,R0903
""" Builds an exercise from serveral files
"""

# --------------------------- REQUIRED LIBRARIES ------------------------------

import argparse
import os
import win32com.client
import json
import  StringIO
import copy

from PyPDF2 import PdfFileReader, PdfFileWriter

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

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

        self.json_file = None

        self.data = None

    def _argparse(self):
        """ Detines an user-friendly command-line interface and proccess its
        arguments.
        """

        description = u'Builds an exercise from serveral files.'

        parser = argparse.ArgumentParser(description)
        parser.add_argument('file', metavar='file', type=str,
                            help='path for resulting .PDF file')

        parser.add_argument('-j', '--json', type=str, metavar='json',
                            default='exercise.json',
                            help='JSON file which contains exercise information')

        args = parser.parse_args()

        self.abspath = os.path.abspath(args.file)
        self.basename = os.path.basename(self.abspath)
        self.dirname = os.path.dirname(self.abspath)
        self.filename = self.basename and os.path.splitext(self.basename)[0]

        self.json_file = args.json

    @staticmethod
    def _create_pdf(relative_path):
        """ Performs a document conversion using Adobe Acrobat OLE Automation

        @param relative_path (string): relative path of the input source file
        @return (string): returns an absolute path of the output PDF file
        """

        # STEP 1: Get full path and filename (wihout extension) of source file
        abspath = os.path.abspath(relative_path)
        basename = os.path.basename(abspath)
        filename = basename and os.path.splitext(basename)[0]

        # STEP 2: Gets full path for output PDF filename
        outpath = os.path.abspath(filename+'.pdf')

        # STEP 3: Use Adobe Acrobat OLE Automation to convert the file
        acrobat = win32com.client.Dispatch('Acroexch.app')
        outdoc = win32com.client.Dispatch('AcroExch.AVDoc')

        outdoc.open(abspath, '')    # Opening file in Acrobat
        pdout = outdoc.GetPDDoc()   # Get document object reference
        pdout.Save(1, outpath)      # Save document

        outdoc.Close(True)
        acrobat.Exit()

        return outpath

    @staticmethod
    def _loadwatermark(wfilename):
        """ Loads the watermark PDF file
        """
        watermark = None

        if wfilename:
            wpath = os.path.abspath(wfilename)
            wpdf = PdfFileReader(open(wpath, 'rb'))
            watermark = wpdf.getPage(0)

        return watermark

    @staticmethod
    def _addwatermark(in_page, watermark):
        """ Adds a watermark
        """

        out_page = copy.copy(watermark)
        out_page.mergePage(in_page)

        return out_page

    @staticmethod
    def _addheader(page, item):
        """ Adds a watermark
        """

        return page

    @staticmethod
    def _addfooter(page, item):
        """ Adds a watermark
        """
        posx = item['margin']['left']
        posy = item['margin']['bottom']

        packet = StringIO.StringIO()

        can = canvas.Canvas(packet, pagesize=A4)
        can.drawString(posx, posy, item['header']['left'])
        can.save()

        packet.seek(0)
        new_pdf = PdfFileReader(packet)
        new_pdf.getPage(0)

        return page

    def _build(self):
        """ Performs the conversion from Excel to OpenDocument
        """

        new_path = os.path.join(self.dirname, self.filename+'.pdf')

        try:
            with open(self.json_file) as data_file:
                self.data = json.load(data_file)

            output = PdfFileWriter()

            for item in self.data['files']:
                fpath = self._create_pdf(item['src'])
                ffile = PdfFileReader(open(fpath, 'rb'))
                wpdf = self._loadwatermark(item['watermark'])

                for npage in xrange(ffile.getNumPages()):
                    page = ffile.getPage(npage)

                    if wpdf is not None:
                        page = self._addwatermark(page, wpdf)

                    if 'header' in item:
                        page = self._addheader(page, item)

                    if 'footer' in item:
                        page = self._addfooter(page, item)

                    output.addPage(page)

            output.write(open(new_path, "wb"))

        except Exception as ex:
            print ex
        else:
            print u'New file %s has been written.' % new_path

    def _is_file(self):
        """ Check if given path is a valid file path
        """
        return os.path.isfile(self.abspath)

    def main(self):
        """ The main application behavior, this method should be used to
        start the application.
        """

        self._argparse()
        self._build()


# --------------------------- SCRIPT ENTRY POINT ------------------------------

App().main()
