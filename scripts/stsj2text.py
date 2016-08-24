# -*- coding: utf-8 -*-
#pylint: disable=I0011
""" Remove cover and instructions from Justice exercises
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
        self._file = None

    def _argparse(self):
        """ Detines an user-friendly command-line interface and proccess its
        arguments.
        """

        # STEP 1: Define the arbument parser
        description = u'Remove cover and instructions from Justice exercises'
        parser = argparse.ArgumentParser(description)

        # STEP 2: Determine positional arguments
        parser.add_argument('file', metavar='file', type=str,
                            help='The Microsoft Word document file')

        # STEP 3: Determine non positional arguments
        # parser.add_argument('-m', '--modifier', type=str, dest='modifier',
        #             choices=['one', 'two', 'tree'], default='day',
        #             help='description for modifier')

        args = parser.parse_args()

        self._file = os.path.abspath(args.file)

    def _stsj2text(self):
        """ Main method docstring
        """

        try:

            word = win32com.client.DispatchEx("Word.Application")
            doc = word.Documents.Open(self._file)

            # STEP 1: Search for last section break
            doc.Content.Select()
            word.Selection.EndKey(Unit=6) # wdStory

            find = word.Selection.Find
            find.ClearFormatting()
            find.Text = "^b"
            find.Replacement.Text = ""
            find.Forward = False
            find.Wrap = 0 # wdFindStop
            find.Format = False
            find.MatchCase = False
            find.MatchWholeWord = False
            find.MatchWildcards = False
            find.MatchSoundsLike = False
            find.MatchAllWordForms = False

            word.Selection.Find.Execute()

            # STEP 2: Go back 2 chars, select and remove
            word.Selection.MoveLeft(Unit=1, Count=2) # wdCharacter
            word.Selection.EndKey(Unit=6, Extend=1) # wdStory, wdExtend
            word.Selection.Delete(Unit=1, Count=1) # wdCharacter
            word.Selection.TypeBackspace()


            # STEP 2: Remove first seccion
            doc.Sections(1).Range.Select()
            word.selection.Delete()

            # STEP 3: Save document
            doc.Save() # wdOriginalDocumentFormat

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

        self._stsj2text()


# --------------------------- SCRIPT ENTRY POINT ------------------------------

App().main()
