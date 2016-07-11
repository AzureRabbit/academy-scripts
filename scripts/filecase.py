# -*- coding: utf-8 -*-
#pylint: disable=I0011,W0703,R0903
""" Allows you to convert a Microsoft Excel XLS format to XLSX document
"""

# --------------------------- REQUIRED LIBRARIES ------------------------------

import argparse
import os

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
            cls.__instance.name = "File Case"
        return cls.__instance

    def __init__(self):
        self.abspath = None
        self.basename = None
        self.dirname = None
        self.filename = None
        self.case = None

    def _argparse(self):
        """ Detines an user-friendly command-line interface and proccess its
        arguments.
        """

        description = u'Change case of the given filename.'

        parser = argparse.ArgumentParser(description)
        parser.add_argument('file', metavar='file', type=str,
                            help='path of the file which name will be changed')

        parser.add_argument('-c', '--case', dest='case', metavar='case',
                            choices=['upper', 'lower', 'capitalize'], default=u'capitalize',
                            help='type of case: upper, lower or capitalize')

        args = parser.parse_args()

        self.abspath = os.path.abspath(args.file)
        self.basename = os.path.basename(self.abspath)
        self.dirname = os.path.dirname(self.abspath)
        self.filename = self.basename and os.path.splitext(self.basename)[0]

        self.case = args.case

    def _is_file(self):
        """ Check if given path is a valid file path
        """
        return os.path.isfile(self.abspath) or os.path.isdir(self.abspath)

    def _change_case(self):
        """ Changes the case of the filename
        """

        if self.case == 'upper':
            new_name = self.abspath.replace(self.basename, self.basename.upper())
        elif self.case == 'lower':
            new_name = self.abspath.replace(self.basename, self.basename.lower())
        else:
            new_name = self.abspath.replace(self.basename, self.basename.capitalize())

        os.rename(self.abspath, new_name)

    def main(self):
        """ The main application behavior, this method should be used to
        start the application.
        """

        self._argparse()
        if self._is_file():
            self._change_case()

# --------------------------- SCRIPT ENTRY POINT ------------------------------

App().main()
