# -*- coding: utf-8 -*-
#pylint: disable=I0011
""" Choose random file
"""

# --------------------------- REQUIRED LIBRARIES ------------------------------

import argparse
import os
import re
import random

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
        self._regex = None

    def _argparse(self):
        """ Detines an user-friendly command-line interface and proccess its
        arguments.
        """

        # STEP 1: Define the arbument parser
        description = u'Something about this application'
        parser = argparse.ArgumentParser(description)

        # STEP 2: Determine positional arguments
        parser.add_argument('command', metavar='command', type=str,
                            choices=['file'],
                            help='Object type to perform a random choise')

        # STEP 3: Determine non positional arguments
        parser.add_argument('-p', '--path', type=str,
                            dest='path', default='.',
                            help='Folder to choose a random file')

        parser.add_argument('-r', '--regex', type=str,
                            dest='regex', default='.*',
                            help='Folder to choose a random file')

        args = parser.parse_args()

        self._path = os.path.abspath(args.path)
        self._regex = args.regex


    def _random_choice(self):
        """ Main method docstring
        """

        try:
            files = os.listdir(self._path)
            files = [f for f in files if bool(re.match(self._regex, f))]
        except Exception as ex:
            print ex
        else:
            result = random.sample(files, 1)
            if result and len(result) > 0:
                os.environ["RCHOICE"] = result[0]
                print result[0]

    def main(self):
        """ The main application behavior, this method should be used to
        start the application.
        """

        self._argparse()

        self._random_choice()


# --------------------------- SCRIPT ENTRY POINT ------------------------------

App().main()
