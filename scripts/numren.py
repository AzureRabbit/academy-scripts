# -*- coding: utf-8 -*-
#pylint: disable=I0011
""" Renames files with a counter
"""

# --------------------------- REQUIRED LIBRARIES ------------------------------

import argparse
import os
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
        self._path = None
        self._prefix = None
        self._digits = None
        self._regex = None
        self._start = None

    def _argparse(self):
        """ Detines an user-friendly command-line interface and proccess its
        arguments.
        """

        # STEP 1: Define the arbument parser
        description = u'Rename files adding counter'
        parser = argparse.ArgumentParser(description)

        # STEP 2: Determine positional arguments
        # parser.add_argument('command', metavar='command', type=str,
        #                     help='description for comamnd')

        # STEP 3: Determine non positional arguments
        parser.add_argument('-f', '--folder', type=str,
                            dest='folder', default='.',
                            help='Folder which contains the files')

        parser.add_argument('-p', '--prefix', type=str,
                            dest='prefix', default='',
                            help='Prefix for document name')

        parser.add_argument('-d', '--digits', type=int,
                            dest='digits', default=1,
                            help='Mininum number of digits for counter')

        parser.add_argument('-r', '--regex', type=str,
                            dest='regex', default='.*',
                            help='Folder to choose a random file')

        parser.add_argument('-s', '--start', type=int,
                            dest='start', default=1,
                            help='First value for counter')

        args = parser.parse_args()

        self._path = os.path.abspath(args.folder)
        self._prefix = args.prefix
        self._digits = args.digits
        self._regex = args.regex
        self._start = args.start


    def _rename(self):
        """ Renames the file
        """

        try:

            files = os.listdir(self._path)
            files = [f for f in files if bool(re.match(self._regex, f))]

            counter = self._start
            for filename in files:
                filepath = os.path.join(self._path, filename)
                if os.path.isfile(filepath):
                    ext = os.path.splitext(filename)[1]
                    strn = str(counter).zfill(self._digits)
                    new_name = u'{}{}{}'.format(self._prefix, strn, ext)
                    print u'{} ===> {}'.format(filename, new_name)
                    os.rename(filepath, new_name)
                    counter = counter + 1

        except Exception as ex:
            print ex

    def main(self):
        """ The main application behavior, this method should be used to
        start the application.
        """

        self._argparse()

        self._rename()


# --------------------------- SCRIPT ENTRY POINT ------------------------------

App().main()
