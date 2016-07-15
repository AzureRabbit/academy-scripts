# -*- coding: utf-8 -*-
#pylint: disable=I0011,W0703,R0903,W0110,W0141
""" Performs a Web Scraping to the SISGAP online web application
"""

# --------------------------- REQUIRED LIBRARIES ------------------------------

from collections import OrderedDict

import argparse
import datetime
import os

import sisgap_class
import google_class

import pprint
import sys

# -------------------------- MAIN SCRIPT BEHAVIOR -----------------------------

class SisgapApp(object):
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
        self._command = None
        self.date = None
        self.lapse = None
        self._groupid = None
        self._jsonfile = None
        self._gsync = None

        self._user = None
        self._password = None
        self._headquarter = None

        self._sisgap = sisgap_class.Sisgap()

    # --------------------------- ARGUMENT PARSE ------------------------------

    def _argparse(self):
        """ Detines an user-friendly command-line interface and proccess its
        arguments.
        """

        description = u'Performs a Web Scraping to the SISGAP online web application.'

        parser = argparse.ArgumentParser(description)
        parser.add_argument('command', metavar='command', type=str,
                            choices=['timetable', 'students'],
                            help='command will be executed')

        parser.add_argument('-u', '--user', type=str, dest='user',
                            metavar='username', default='wrong!',
                            help='username will be used to login on platform')

        parser.add_argument('-p', '--pass', type=str, dest='password',
                            metavar='password', default='wrong!',
                            help='password will be used to login on platform')

        parser.add_argument('-q', '--headq', type=str, dest='headquarter',
                            metavar='headquarter', default='VIGOZA',
                            help='headquarter to be accessed')

        parser.add_argument('-d', '--date', type=str, dest='datestr',
                            default=datetime.datetime.now().strftime(u'%d/%m/%Y'),
                            help='date will be used')

        parser.add_argument('-l', '--lapse', type=str,
                            dest='lapse',
                            choices=['day', 'week', 'month'], default='day',
                            help='interval type for timetable and gsync commands')

        parser.add_argument('-g', '--groupid', type=int,
                            dest='groupid', metavar='ID',
                            help='group identififer for students and fsync commands')

        parser.add_argument('-j', '--json', type=str, dest='jsonfile',
                            metavar='file', default='python-sisgap.json',
                            help='json file which contains needed script configuration')

        parser.add_argument('-s', '--sync', type=str, dest='gsync',
                            metavar='email', default=None,
                            help='synchronizes timetable (with Google) or folders')

        args = parser.parse_args()

        self._command = args.command
        self._argparse_date(args.datestr)
        self._argparse_lapse(args.lapse)
        self._groupid = args.groupid
        self._jsonfile = os.path.abspath(args.jsonfile)
        self._gsync = args.gsync

        self._user = args.user
        self._password = args.password
        self._headquarter = args.headquarter

    def _argparse_date(self, datestr):
        """ Parses date sting given as command line argument and stores
        a valid datetime in related class attribute

        @param datestr: date string will be parsed
        """

        try:
            date = datetime.datetime.strptime(datestr, u'%d/%m/%Y')
        except Exception:
            msg = u'Invalid date format (%d/%m/%Y), {} will be used instead'
            date = datetime.datetime.now()
            print  msg.format(date.strftime('%d/%m/%Y'))

        self.date = date

        return date

    def _argparse_lapse(self, lapse, date=None):
        """ Parses date sting given as command line argument and stores
        a valid datetime in related class attribute

        @param date: date between maximum and minimum time
        @param lapse: type of interval
        """

        date = date or self.date

        if lapse == 'week':
            first_day = date - datetime.timedelta(days=date.weekday())
            last_day = first_day + datetime.timedelta(days=6)
        elif lapse == 'month':
            first_day = datetime.datetime(date.year, date.month, 1)
            last_day = datetime.datetime(date.year, date.month + 1, 1) - datetime.timedelta(days=1)
        else:
            first_day = date
            last_day = date

        self.lapse = (first_day, last_day)

        return first_day, last_day

    def _print_values(self):
        """ Prints stored class attributes
        """

        print u'Date\t: %s' % self.date.strftime(u'%d/%m/%Y %H:%M:%S')
        print u'Interval: %s - %s' % (
            self.lapse[0].strftime(u'%d/%m/%Y %H:%M:%S'),
            self.lapse[1].strftime(u'%d/%m/%Y %H:%M:%S'),
        )
        print u'Group\t: %s' % self._groupid


    # ----------------------- SISGAP RELATED METHODS --------------------------

    def _search_group(self, start_date, end_date, group_id):
        """ Search for gropup using given identififer

        @param group_id (int): group identififer
        """
        # STEP 1: Set default value for group info, this will be returned by
        # this method
        groupinfo = None

        # STEP 2: Opens Sisgap session
        self._sisgap.open_session()

        # STEP 4: Retrieves timetable for each day between given dates
        for single_date in self._daterange(start_date, end_date):
            timetable = self._sisgap.solic_pasar_lista(single_date)
            groups = filter(lambda x: x['idGrupo'] == group_id, timetable)

            if groups:
                groupinfo = groups[0]
                break

        # STEP 5: Closes Sisgap session
        self._sisgap.close_session()

        # STEP 6: Returns the timetable
        return groupinfo

    def _get_time_table(self, start_date, end_date):
        """ Get the timetable for all dates between start_date and end_date
            [ {date1: [{...}, {...}, ...]}, {date2: [{...}, {...}, ...]}, ...]

        @param start_date (date): first date in range
        @param end_date (date): last date in range
        """

        # STEP 1: Set default timetable, this will be an empty list
        timetable = OrderedDict()

        # STEP 2: Opens Sisgap session
        self._sisgap.open_session()

        # STEP 4: Retrieves timetable for each day between given dates
        for single_date in self._daterange(start_date, end_date):
            timetable[single_date] = self._sisgap.solic_pasar_lista(single_date)

        # STEP 5: Closes Sisgap session
        self._sisgap.close_session()

        # STEP 6: Returns the timetable
        return timetable

    def _get_student_list(self, group_id):
        """ Returns the student list for a given group_id
            [{'firstname': u'', 'lastname': u'', 'name': u''}, ...]

        @param group_id (int): Group identifier
        """

        # STEP 1: Set default return value
        students = []

        # STEP 2: Search for group using identifier
        group = self._search_group(self.lapse[0], self.lapse[1], group_id)

        # STEP 3: If group has been found, student list will be retrieved
        if group:
            self._sisgap.open_session()
            students = self._sisgap.ed_asist_grupo(group)
            self._sisgap.close_session()

        # STEP 4: Return student list
        return students

    # -------------------------- AUXILIAR METHODS -----------------------------

    @staticmethod
    def _daterange(start_date, end_date):
        """ Builds a range of dates between start_date and end_date, this
        will include both given dates in range.

        @param start_date (date): first date in range
        @param end_date (date): last date in range
        """

        for day in range(int((end_date - start_date).days+1)):
            yield start_date + datetime.timedelta(day)

    # ---------------------------- VIEW METHODS -------------------------------

    @staticmethod
    def _draw_horizontal_line(width_list, mode=1):
        """ a
        """
        chs = [(u'┌', u'┬', u'┐'), (u'├', u'┼', u'┤'), (u'└', u'┴', u'┘')]

        line = u''

        cell = u'{{0:─<{0}}}'.format(width_list[0])
        line = line + cell.format(chs[mode][0])

        for width in width_list[1:]:
            cell = u'{{0:─<{0}}}'.format(width)
            line = line + cell.format(chs[mode][1])

        line = line + chs[mode][2]

        print line

    def _print_timetable_daily_header(self, _in_date):
        """ Prints header for timetable group
        """

        title = datetime.datetime.strftime(_in_date, u'%A').upper()
        title = title + datetime.datetime.strftime(_in_date, u' (%d-%m-%Y)')
        print '\n{0:^90s}'.format(title)

        self._draw_horizontal_line([6, 43, 8, 8, 23], 0)

        print( u'│ {:>3} │ {:<40} │ {:^5} │ {:^5} │ {:<20} │'.format(
            'ID', 'GROUP', 'START', 'END', 'SUBJECT'))

        self._draw_horizontal_line([6, 43, 8, 8, 23], 1)


    def _print_timetable(self, timetable):
        for datett, grouptt in timetable.iteritems():
            if grouptt and len(grouptt):
                self._print_timetable_daily_header(datett)
                for group in grouptt:
                    line = u'│ {:>3} │ {:<40} │ {:^5} │ {:^5} │ {:<20} │'.format(
                        group['idGrupo'],
                        group['grupo'],
                        group['hora_inicio'].strftime('%H:%M'),
                        group['hora_fin'].strftime('%H:%M'),
                        group['materia']
                    )
                    print line

                self._draw_horizontal_line([6, 43, 8, 8, 23], 2)

    @staticmethod
    def _print_students(students):

        sorted_students = sorted(students, key=lambda k: k['name'])

        index = 1
        for student in sorted_students:
            line = u'{:>2}.- {} {} {}'.format(
                index,
                student['name'],
                student['firstname'],
                student['lastname']
            )
            print line
            index = index + 1

    # ---------------------------- MAIN METHODS -------------------------------

    def _timetable_cmd(self):
        timetable = self._get_time_table(self.lapse[0], self.lapse[1])
        if self._gsync:
            gcalendar = google_class.GoogleCalendar(self._gsync)
            gcalendar.google_sync(timetable)

        self._print_timetable(timetable)

    def _students_cmd(self):
        students = self._get_student_list(self._groupid)
        self._print_students(students)

    def run(self):
        """ The main application behavior, this method should be used to
        start the application.
        """


        self._argparse()
        self._print_values()

        # start_date = datetime.date.today()
        # end_date = start_date + datetime.timedelta(days=1)
        # pprint.pprint(self._get_time_table(start_date, end_date))

        # pprint.pprint(self._get_student_list(907))

        self._sisgap.set_credentials(self._user, self._password, self._headquarter)

        if self._command == 'timetable':
            self._timetable_cmd()
        elif self._command == 'students':
            self._students_cmd()





# --------------------------- SCRIPT ENTRY POINT ------------------------------

if __name__ == "__main__":
    SisgapApp().run()

