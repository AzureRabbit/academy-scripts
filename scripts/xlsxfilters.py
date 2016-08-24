# -*- coding: utf-8 -*-
#pylint: disable=I0011,W0703,R0903,W0702
""" Create questions with filters in an Excel File
"""

# --------------------------- REQUIRED LIBRARIES ------------------------------

import argparse
import os
import win32com.client
import datetime
import locale
import sys
import json

from jinja2 import Template
from lib.random_values import RandomValues

OLE_TIME_ZERO = datetime.datetime(1899, 12, 30, 0, 0, 0)

# ------------------------------ HEADER CLASS ---------------------------------

class Header(object):
    """ Stores a header name
        This class has been defined to store duplicate header names as dict keys
    """

    def __init__(self, name):
        self._name = name

    @property
    def name(self):
        """ Return name
        """
        return self._name

    def __str__(self):
        return self._name

    def __unicode__(self):
        return self._name

# ------------------------------- DATA CLASS ----------------------------------

class Data(object):
    """ Class to manage data range
    """
    #pylint: disable=I0011,R0902

    def __init__(self, workbook, name):
        #pylint: disable=I0011,R0914,W0612

        # STEP 1: Save: workbook, sheet, range name and range address
        self._workbook = workbook
        self._name = name
        sname, daddr = self._search_for_range()
        self._worksheet = workbook.Sheets(sname)

        # STEP 3: Split the range in headers and values
        drange, hrange, vrange = self._split_range(daddr)
        self._drange = drange
        self._hrange = hrange
        self._vrange = vrange

        # STEP 4: Save the current iterator position
        self._current = 1 if self._hrange.Cells.Count > 0 else 0

        self._values = {}
        self._headers = []
        self._iteritems = None


    # ----------------------------- PROPERTIES --------------------------------


    @property
    def headers(self):
        """ Return headers in Data range """

        if not self._headers:
            headers = self._hrange.Columns
            ncols = headers.Count

            self._headers = \
                [headers.Cells(1, ncol).Value for ncol in range(1, ncols)]

        return self._headers


    @property
    def values(self):
        """ Return values in Data range """

        if not self._values:
            colidxs = range(0, self.__len__() -1)
            rowidxs = range(1, self._vrange.Rows.Count)
            headers = self.headers

            for colidx in colidxs:
                top = self._vrange.Cells(1, colidx+1)
                bottom = self._vrange.Cells(self._vrange.Rows.Count, colidx+1)
                column = self._worksheet.Range(top, bottom)

                values = [column.Cells(ridx, 1).Value for ridx in rowidxs]
                if unicode(type(values[0])) == u'<type \'time\'>':
                    values = [self._valid_date(value) for value in values]

                self._values[headers[colidx]] = values

        return self._values


    @property
    def address(self):
        """ Return the global address of data the range
        """
        return self.global_address(self.addresslocal)


    @property
    def addresslocal(self):
        """ Return the local address of data the range
        """
        return self._drange.Address


    @property
    def worksheet(self):
        """ Return the Worksheet which contains the range
        """
        return self._worksheet


    @property
    def workbook(self):
        """ Return the Worksheet which contains the range
        """
        return self._workbook


    @property
    def vrange(self):
        """ Return the range with the values
        """
        return self._vrange


    @property
    def hrange(self):
        """ Return headers range
        """
        return self._hrange


    @property
    def drange(self):
        """ Return full range
        """
        return self._drange


    @property
    def name(self):
        """ Return given range name passed to the constructor
        """
        return self._name


    # --------------------------- USEFULL METHODS -----------------------------


    def global_address(self, localaddress):
        """ Builds a global address using sheet name and localaddress
        """
        return u'={}!{}'.format(self._worksheet.Name, localaddress)

    @staticmethod
    def _valid_date(value):
        """ Returns a valid python datetime
        """
        return OLE_TIME_ZERO + datetime.timedelta(max(float(value), 60))

    # ---------------------------- MAGIC METHODS ------------------------------


    def __getitem__(self, key):
        index = -1

        key = unicode(key, errors=u'replace') if not isinstance(key, unicode) else key

        for header in self.headers:
            index += 1
            if header == key:
                break

        top = self._vrange.Cells(1, index)
        bottom = self._vrange.Cells(self._vrange.Rows.Count, index)

        return self._worksheet.Range(top, bottom)


    def __len__(self):
        return self._drange.Columns.Count


    def __iter__(self):
        return self


    def next(self):
        """ Itterator
        """
        if self._current > self._vrange.Columns.Count:
            raise StopIteration
        else:
            top = self._vrange.Cells(1, self._current)
            bottom = self._vrange.Cells(self._vrange.Rows.Count, self._current)
            self._current += 1

            return self._worksheet.Range(top, bottom)


    def iteritems(self):
        """ Return iterator like dictionary
        """

        if not self._iteritems:
            idxs = range(0, self.__len__() -1)
            headers = self.headers

            tupleslist = []

            for index in idxs:
                top = self._vrange.Cells(1, index+1)
                bottom = self._vrange.Cells(self._vrange.Rows.Count, index+1)
                column = self._worksheet.Range(top, bottom)
                tupleslist.append((headers[index], column))

            self._iteritems = iter(tupleslist)

        return self._iteritems


    # -------------------------- AUXILIAR METHODS -----------------------------


    def _search_for_range(self):
        """ Search for range in Workbook
        """
        address = None
        sname = None

        for nom in self._workbook.Names:
            if nom.Name == self._name:
                address = nom.RefersTo
                sname = address.split(u'!')[0][1:]
                break

        try:
            if not address:
                sname = self._workbook.ActiveSheet.Name
                laddress = self._workbook.ActiveSheet.Range(self._name).Address
                address = self.global_address(laddress)
        except:
            pass

        return sname, address


    def _split_range(self, daddr):
        """ Split range in two parts returning range, headers and value
        """

        # STEP 1: Save the current active sheet
        current_sheet = self._workbook.ActiveSheet
        # STEP 2: Get the data range
        self._worksheet.Activate()
        drange = self._worksheet.Range(daddr)

        # STEP 3: Calculate the with and height
        dwidth = drange.Columns.Count
        dheight = drange.Rows.Count

        # STEP 4: Local address for headers and values
        haddr = u'%s:%s' % (
            drange.Cells(1, 1).Address,
            drange.Cells(1, dwidth).Address
        )

        vaddr = u'%s:%s' % (
            drange.Cells(2, 1).Address,
            drange.Cells(dheight, dwidth).Address
        )

        # STEP 5: Global address for headers and values
        haddr = self.global_address(haddr)
        vaddr = self.global_address(vaddr)

        # STEP 6: Get ranges for headers and values
        hrange = self._worksheet.Range(haddr)
        vrange = self._worksheet.Range(vaddr)

        # STEP 7: Restore the previous active sheet
        current_sheet.Activate()

        # STEP 8: Return ranges for: all, headers and values
        return drange, hrange, vrange


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
        self._abspath = None
        self._dname = None
        self._data = None

        self._name = None
        self._varname = None
        self._minimun = None
        self._maximun = None
        self._number = None
        self._attempts = None

        locale.setlocale(locale.LC_ALL, '')

        # STEP 5: Define text for signs
        self._ostrs = {
            u'>': u'mayor que',
            u'<': u'menor que',
            u'>=': u'superior a',
            u'<=': u'inferior a',
            u'<>': u'distinto de',
            u'=': u'igual a'
        }

        # STEP 6: Question text template
        self._template = u'Si consideramos únicamente aquellos {} donde '


    def _argparse(self):
        """ Detines an user-friendly command-line interface and proccess its
        arguments.
        """

        description = u'Build quiestions with filters in an Excel File.'

        parser = argparse.ArgumentParser(description)
        parser.add_argument(u'file', metavar=u'file', type=str,
                            help=u'path of the Excel file will be used')

        # STEP 3: Determine non positional arguments
        parser.add_argument(u'-n', u'--name', type=str, dest=u'name',
                            default=u'DATOS', help=u'Data range name or address')

        parser.add_argument(u'-v', u'--var-name', type=str, dest=u'varname',
                            default=u'registros',
                            help=u'Name of the variable represented by data')

        parser.add_argument(u'-m', u'--min', type=float, dest=u'minimun',
                            default=0.1,
                            help=u'Minimun number of records after filtering (so much per one)')

        parser.add_argument(u'-x', u'--max', type=float, dest=u'maximun',
                            default=0.75,
                            help=u'Maximun number of records after filtering (so much per one)')

        parser.add_argument(u'-b', u'--number', type=int, dest=u'number',
                            default=10,
                            help=u'Number of questions to perform')

        parser.add_argument(u'-a', u'--attempts', type=int, dest=u'attempts',
                            default=100,
                            help=u'Maximun number of attempts')

        args = parser.parse_args()

        self._abspath = os.path.abspath(args.file)
        self._dname = args.name

        self._name = args.name
        self._varname = args.varname
        self._minimun = args.minimun
        self._maximun = args.maximun
        self._number = args.number
        self._attempts = args.attempts


    def _bounds(self, values):
        timet, datet = datetime.time, datetime.date
        random = RandomValues()
        lower, higher = None, None

        greater = random.values((u'<', u'<='))[0]
        less = random.values((u'>', u'>='))[0]

        if isinstance(values[0], unicode) or isinstance(values[0], unicode):
            sortedlist = sorted(values)
            length = len(sortedlist)

            tmp = random.integer(int(length * 0.1), int(length * 0.4))[0]
            lower = (greater, sortedlist[tmp])
            tmp = random.integer(int(tmp + length * 0.1), int(length - length * 0.1))[0]
            higher = (less, sortedlist[tmp])
            equal = (random.values([u'=', u'<>'])[0], random.values(values)[0])
        elif isinstance(values[0], timet) or isinstance(values[0], datet):
            lower = min(values)
            higher = max(values)
            rank = (higher - lower).days

            lower = lower + datetime.timedelta(days=(rank * 0.1))
            higher = higher - datetime.timedelta(days=(rank * 0.1))

            tmp = datetime.timedelta(days=(rank * 0.4)).days
            lower = min(random.date(lower, tmp, number=3, bidirectional=False))
            tmp = lower + datetime.timedelta(days=(rank * 0.1))
            higher = random.date(tmp, (higher-tmp).days, bidirectional=False)[0]

            lower = (greater, lower.strftime(u'%d/%m/%Y'))
            higher = (less, higher.strftime(u'%d/%m/%Y'))

            equal = (u'<>', random.values(values)[0].strftime(u'%d/%m/%Y'))

        elif isinstance(values[0], int):
            lower = min(values)
            higher = max(values)
            rank = higher - lower

            lower += int(rank * 0.1)
            higher -= int(rank * 0.1)

            lower = min(random.float(lower, int(lower + rank * 0.4), number=3))
            higher = random.float(lower + int(rank * 0.1), higher)[0]

            lower = (greater, lower)
            higher = (less, higher)

            equal = (u'<>', random.values(values)[0])

        elif isinstance(values[0], float):
            lower = min(values)
            higher = max(values)
            rank = higher - lower

            lower += rank * 0.1
            higher -= rank * 0.1

            lower = min(random.float(lower, lower + rank * 0.4, number=3))
            higher = random.float(lower + (rank * 0.1), higher)[0]

            lower = (greater, unicode(round(lower, 2)).replace('.', ','))
            higher = (less, unicode(round(higher, 2)).replace('.', ','))

            equal = (u'<>', unicode(random.values(values)[0]).replace('.', ','))

        else:
            print type(values[0])
            raise TypeError(u'Unknown type for colum')


        return [lower, higher, equal]

    def _criteria(self):
        """ Returns criteria table
        """

        random = RandomValues()

        itemnumber = random.integer(1, 5)[0]
        headers = random.values(self._data.headers, number=itemnumber)
        values = {k: self._data.values[k] for k in self._data.values.keys() if k in headers}

        criteria = {}
        toand = {}
        for header in random.shuffle(headers):
            bounds = self._bounds(values[header])
            bound1 = random.values(bounds)[0]
            if random.float(0, 1)[0] > 0.60 or len(criteria) == 0:
                bound2 = random.values(bounds, exclude=[bound1])[0]
            else:
                bound2 = bound1
                toand[header] = bound2


            criteria[Header(header)] = [u'\'%s%s' % bound1, u'\'%s%s' % bound2]

        # print u'\tY: ',
        # for key, value in toand.iteritems():
        #     if len(criteria) >= 5:
        #         break

        #     bounds = self._bounds(values[key])
        #     if value[0] == u'<>':
        #         bounds = [b for b in bounds if b[0] != u'<>']
        #     else:
        #         bounds = [b for b in bounds if b[0] == u'<>']

        #     if bounds and bounds[0]:
        #         print key,
        #         bound = bounds[0]
        #         criteria[Header(key)] = [u'\'%s%s' % bound, u'\'%s%s' % bound]

        return criteria

    @staticmethod
    def _print_criteria(criteria):
        """ Print the criteria dictionary as table
        """
        #pylint: disable=I0011,W0141
        from tabulate import tabulate
        print tabulate(
            map(list, zip(*criteria.values())), headers=criteria.keys()), '\n'


    @staticmethod
    def _add_sheet(workbook):
        """ Add new Sheet into the book
        """
        numsheets = workbook.Sheets.Count
        lastsheet = workbook.Sheets(numsheets)
        workbook.Sheets.Add(After=lastsheet)
        return workbook.Sheets(numsheets + 1)

    @staticmethod
    def _write_criteria(worksheet, criteria):
        """ Write table of criteria
        """

        colums = len(criteria.keys())
        area = worksheet.Range(u'A1:Z100')

        for columnidx in range(1, colums+1):
            header = criteria.keys()[columnidx - 1]
            area.Cells(1, columnidx).Value = header.name
            area.Cells(2, columnidx).Value = criteria[header][0]
            area.Cells(3, columnidx).Value = criteria[header][1]

        return area.Range(area.Cells(1, 1), area.Cells(3, colums))

    @staticmethod
    def _get_result_range(worksheet):
        """ Return the A5 cell
        """
        area = worksheet.Range(u'A1:Z100')

        return area.Range(u'A5'), area.Range(u'A10')

    @staticmethod
    def _split_criteria(item):
        """ Split given criteria string returning sign and value
        """

        pointer = 3 if (item[2] == u'=' or item[2] == u'>') else 2
        return item[1:pointer], item[pointer:]

    def _build_question(self, criteria):
        """ Build a text for question
        """

        # STEP 1: First line for question, this have the name of the variable
        question = self._template
        question = question.format(self._varname)

        # STEP 3: Start loop over criterias
        index, last = 0, len(criteria) - 1
        keys = criteria.keys()

        for index in range(0, last + 1):

            key = keys[index]
            crit = criteria[key]

            # STEP 4: Text for the first criteria over the column
            sign, value = self._split_criteria(crit[0])
            fragment = u'«{}» es {} {}'.format(key.name, self._ostrs[sign], value)

            # STEP 5: Text for the second criteria over the colum if it is not
            # the same as the first criteria
            if crit[0] != crit[1]:
                sign, value = self._split_criteria(crit[1])
                fragment += u' o es {} {}'.format(self._ostrs[sign], value)

            # STEP 6: Separator: none, ',' or 'y'
            if index == last and index > 0:
                question += ' y '
            elif index > 0:
                question += ', '

            # STEP 7: Concatenate the fragment to the question and increase
            # the counter
            question += fragment
            index += 1

        return question + u'; ¿#####?'


    @staticmethod
    def _write_question(worksheet, question):
        """ Writes question text in sheet
        """
        worksheet.Range('A7').Value = question


    @staticmethod
    def _print_question(question):
        """ Prints given question
        """
        os_encoding = locale.getpreferredencoding()
        print question.encode(os_encoding, errors='replace'), u'\n'


    def _xlsxfilters(self):
        """ Performs the conversion from xls to xlsx
        """

        questions = []

        # try:

        excel_app = win32com.client.gencache.EnsureDispatch(u'Excel.Application')
        workbook = excel_app.Workbooks.Open(self._abspath)

        self._data = Data(workbook, self._dname)

        attempt = 0
        count = 0
        minimun = int(self._data.vrange.Rows.Count * self._minimun)
        maximun = int(self._data.vrange.Rows.Count * self._maximun)


        print 'Starting with \tmaximun:', maximun, '\tmaximun:', minimun, '\tnumber:', self._number

        while attempt < self._attempts and count < self._number:

            criteria = self._criteria()
            question = self._build_question(criteria)

            worksheet = self._add_sheet(workbook)
            self._write_question(worksheet, question)
            crange = self._write_criteria(worksheet, criteria)

            drange = self._data.drange
            orange, rrange = self._get_result_range(worksheet)

            result = excel_app.WorksheetFunction.DCountA(
                drange, self._data.headers[0], crange)

            print u'Attempt: ', attempt, u'\tResult: ', result, u'\tQuestions: ', len(questions)

            if result > minimun and result < maximun:
                orange.Value = result

                drange.AdvancedFilter(
                    Action=2, # xlFilterCopy,
                    CriteriaRange=crange,
                    CopyToRange=rrange,
                    Unique=False
                )

                questions.append(question)
                self._print_criteria(criteria)
                self._print_question(question)

                count += 1
            else:
                excel_app.DisplayAlerts = False
                worksheet.Select()
                excel_app.ActiveWindow.SelectedSheets.Delete()

            attempt += 1

        excel_app.DisplayAlerts = False
        workbook.SaveAs(
            self._abspath.replace('prueba', 'prueba2'),
            FileFormat=51,
            ConflictResolution=2
        )
        workbook.Close()
        excel_app.Application.Quit()

        # except Exception as ex:
        #     print ex
        # else:
        #     print u'File %s has been updated.' % self._abspath


    def _is_file(self):
        """ Check if given path is a valid file path
        """
        return os.path.isfile(self._abspath)


    def main(self):
        """ The main application behavior, this method should be used to
        start the application.
        """

        self._argparse()
        if self._is_file():
            self._xlsxfilters()


# --------------------------- SCRIPT ENTRY POINT ------------------------------

App().main()
