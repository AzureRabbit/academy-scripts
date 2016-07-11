# -*- coding: utf-8 -*-
#pylint: disable=I0011,W0703,R0903
""" Merge serveral PDF files into one file
"""

# --------------------------- REQUIRED LIBRARIES ------------------------------

import argparse
import os
import copy

from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
import StringIO
from reportlab.lib.units import cm

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
        self.input_file = None
        self.watermark_file = None
        self.output_file = None

    def _argparse(self):
        """ Detines an user-friendly command-line interface and proccess its
        arguments.
        """

        description = u'Merge serveral PDF files into one file'

        parser = argparse.ArgumentParser(description)
        parser.add_argument('input', metavar='input', type=str,
                            help='Input PDF file path')

        parser.add_argument('watermark', metavar='watermark', type=str,
                            help='Watermark PDF file path')

        parser.add_argument('output', metavar='output', type=str,
                            help='Output PDF file path')


        args = parser.parse_args()

        self.input_file = args.input
        self.watermark_file = args.watermark
        self.output_file = args.output

    def _mergepdf(self):
        """ Performs the conversion from docx to pdf
        """

        try:

            wpath = os.path.abspath(self.watermark_file)
            wpdf = PdfFileReader(open(wpath, 'rb'))
            watermark = wpdf.getPage(0)

            opath = os.path.abspath(self.output_file)
            output = PdfFileWriter()

            ipath = os.path.abspath(self.input_file)
            ipdf = PdfFileReader(open(ipath, 'rb'))

            for i in xrange(ipdf.getNumPages()):
                print ipdf.getPage(i).getContents()
                page = copy.copy(watermark)
                page.mergePage(ipdf.getPage(i))

                packet = StringIO.StringIO()
                can = canvas.Canvas(packet)
                can.setFont("Helvetica", 30)
                can.drawString(2.2*cm, 1.2*cm-30, "Hello world")
                can.save()
                packet.seek(0)
                new_pdf = PdfFileReader(packet)
                page.mergePage(new_pdf.getPage(0))

                output.addPage(page)

            with open(opath, 'wb') as fout:
                output.write(fout)

        except Exception as ex:
            print ex
        else:
            print u'The file has been watermarked'

    def main(self):
        """ The main application behavior, this method should be used to
        start the application.
        """

        self._argparse()
        self._mergepdf()


# --------------------------- SCRIPT ENTRY POINT ------------------------------

App().main()
