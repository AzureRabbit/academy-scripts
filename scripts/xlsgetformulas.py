# -*- coding: utf-8 -*-
#pylint: disable=I0011,W0703,R0903
""" Allows you to convert a Microsoft Excel XLS format to XLSX document
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
            cls.__instance.name = u"The one"
        return cls.__instance

    def __init__(self):
        self.abspath = None
        self.basename = None
        self.dirname = None
        self.filename = None

    def _argparse(self):
        """ Detines an user-friendly command-line interface and proccess its
        arguments.
        """

        description = u'Convert a Microsoft Excel XLX format to XLSX document.'

        parser = argparse.ArgumentParser(description)
        parser.add_argument(u'file', metavar='file', type=str,
                            help=u'path of the .xls file will be converted')

        args = parser.parse_args()

        self.abspath = os.path.abspath(args.file)
        self.basename = os.path.basename(self.abspath)
        self.dirname = os.path.dirname(self.abspath)
        self.filename = self.basename and os.path.splitext(self.basename)[0]

    def _get_formulas(self):
        """ Performs the conversion from xls to xlsx
        """

        try:
            excel_app = win32com.client.gencache.EnsureDispatch(u'Excel.Application')
            excel_app.DisplayAlerts = False
            workbook = excel_app.Workbooks.Open(self.abspath, ReadOnly=1)

            print u'Libro: {}\tHojas: {}\tGráficos: {}'.format(
                workbook.Name,
                workbook.Sheets.Count,
                workbook.Charts.Count
            ).encode('ascii', 'replace')

            if workbook.Sheets and len(workbook.Sheets) > 0:
                for sheet in workbook.Sheets:

                    conditionals = 0

                    if sheet.Type == -4167:
                        sheet.Activate()

                        try:
                            sheet.Select()
                            excel_app.Selection.SpecialCells(-4123, 23).Select()
                        except Exception:
                            pass

                        try:
                            sheet.Select()
                            excel_app.Selection.SpecialCells(-4172).Select()
                            conditionals = excel_app.Selection.Cells.Count
                        except Exception:
                            pass

                        if excel_app.Selection.Cells and len(excel_app.Selection.Cells):
                            print u'\tHoja: {}\tCondicionales: {}'.format(
                                sheet.Name,
                                conditionals
                            ).encode('ascii', 'replace')
                            for cell in excel_app.Selection.Cells:
                                print u'\t\tCelda {}: {}'.format(
                                    cell.Address,
                                    cell.FormulaLocal
                                ).encode('ascii', 'replace')
                    elif sheet.Type == -4100:
                        print u'\tHoja: {}'.format(sheet.Name).encode('ascii', 'replace')
                        print u'\t\tGráfico en hoja completa'.encode('ascii', 'replace')
                    else:
                        print u'\tHoja: {}'.format(sheet.Name).encode('ascii', 'replace')
                        print u'\t\tTipo desconocido'.encode('ascii', 'replace')

            workbook.Close(SaveChanges=False)
            excel_app.DisplayAlerts = True
            excel_app.Application.Quit()

        except Exception as ex:
            print ex
        else:
            print u'That\'s all'

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
            self._get_formulas()

# --------------------------- SCRIPT ENTRY POINT ------------------------------

App().main()
