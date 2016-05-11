# -*- coding: utf-8 -*-

#pylint: disable=W0110,W0141,W0702
from datetime import date, timedelta, datetime
from lxml import etree
from HTMLParser import HTMLParser
from ast import literal_eval

import urllib2
import urllib
import re

import httplib2
import os

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools

import uuid
import pytz

class Sisgap(object):
    """ Get resources from Sisgap
    """

    _header_accpet = (
        'text/html,'
        'application/xhtml+xml,'
        'application/xml;q=0.9,'
        'image/webp,'
        '*/*;q=0.8'
    )

    _header_accept_encoding = (
        'gzip, '
        'deflate, '
        'sdch'
    )

    _header_accept_language = (
        'es-ES,'
        'es;q=0.8,'
        'en;q=0.6,'
        'ru;q=0.4,'
        'de;q=0.2,'
        'pt;q=0.2,'
        'gl;q=0.2,'
        'zh-CN;q=0.2,'
        'zh;q=0.2,'
        'it;q=0.2,'
        'fr;q=0.2'
    )

    _header_connection = 'keep-alive'

    _header_host = 'vigo.academiapostal3.es:8095'

    _header_upg_insecure_req = '1'

    _header_user_agent = (
        'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML,'
        ' like Gecko) Chrome/50.0.2661.94 Safari/537.36'
    )

    # --------------------------- ESTATIC METHODS -----------------------------

    @staticmethod
    def _get_monday(in_date=None, week_type=0):
        """ Returns the first day from a week

        @note: first day of the week is needed to get a week timetable

        @param in_date (date): date for which the schedule will be consulted
        @param week_type: type of week: 0 for spanish or 1 for english
        """

        assert week_type in (0, 1), 'Invalid type of week'

        this_day = in_date or date.today()
        relative = this_day.weekday() + week_type

        return this_day - timedelta(days=relative)

    @staticmethod
    def _get_full_tag_text(element, unescape=True):
        """ Get the text from give tag and from the all of the child elements
        In addition, this method decode all the HTML entities.

        @note: some tags have other inline tags inside, this method removes
        these leaving the text clean.

        @param element: elementtree witch represents the HTML tag
        @param unescape: if true HTML entities will be decoded
        @return full tag text
        """

        text = etree.tostring(element)

        if unescape:
            parser = HTMLParser()
            text = parser.unescape(text)

        text = re.sub(r'(<[^>]*>|[\a\b\f\n\r\t\v]+)', '', text, re.UNICODE)

        text = u' '.join(text.split())
        text = text.strip()

        return text

    @staticmethod
    def _one_digit_date(_in_date):
        """ Returns an string with a formate date d/m/yyyy

        @note: the url "solicPasarLista.do" needs one argument, this must be
        a date formated as d/m/yyy

        @param _in_date (date): date will be formated
        """
        return u'{}/{}/{}'.format(_in_date.day, _in_date.month, _in_date.year)

    @staticmethod
    def _parse_action_html_cell(html_cell):
        """ Parses the cell which contains the javascript action getting its
        arguments.

        @note: this method has been detached from _parse_day_html_table becouse
        it exceeded the number of variables (pylint R0914).

        @param html_cell: etreeelement which represents the cell
        return (tuple): returns group_id and topic_id
        """

        act_link = html_cell.xpath('child::a')[0]
        act_attrs = re.search(r'\(([^)]+)\)', act_link.attrib['href'])
        act_values = act_attrs.groups(0)[0].split(',')

        group_id = literal_eval(act_values[0].strip())
        topic_id = literal_eval(act_values[1].strip())

        return (group_id, topic_id)

    # -------------------- CONSTRUCTORS AND DESTRUCTORS -----------------------

    def __init__(self):
        self._referer = None
        self._cookie = None
        self._user = 'user'
        self._password = 'password'
        self._headquarters = 'headquarters'
        self._htmlparser = etree.HTMLParser()

        # If modifying these scopes, delete your previously saved credentials
        # at ~/.credentials/calendar-python-quickstart.json
        _google_scopes = 'https://www.googleapis.com/auth/calendar'
        _google_client_secret_file = 'secret.json'
        _google_application_name = 'Sisgap'
        _google_mail_address = 'a@b.c'
        
    # --------------------------- PRIVATE METHODS -----------------------------

    def _build_request(self, url, values=None):
        """ Buils new request with needed headers and post arguments

        @param url (string): fully qualified url from resource
        @param values (dict): dictionary with request parameters
        """

        if values:
            data = urllib.urlencode(values)
            req = urllib2.Request(url, data)
        else:
            req = urllib2.Request(url)

        # Headers below have been retrieved from Google Chome Session
        req.add_header('Accept', self._header_accpet)
        req.add_header('Accept-Encoding', self._header_accept_encoding)
        req.add_header('Accept-Language', self._header_accept_language)
        req.add_header('Connection', self._header_connection)
        req.add_header('Host', self._header_host)
        req.add_header('Upgrade-Insecure-Requests', self._header_upg_insecure_req)
        req.add_header('User-Agent', self._header_user_agent)

        # First concection have not cookie or referer. Once the conection has
        # been stablished by object instance, _referer will have the las opened
        # url and the _cookie will have the session cookie.
        if self._referer:
            req.add_header('Referer', self._referer)
        if self._cookie:
            req.add_header('Cookie', self._cookie)

        return req

    def _build_url(self, resource_identifier):
        """ Builds a fully qualified url joining protocol, host and relative url
            @param resource_identifier: relative url
        """

        seq = ('http:/', self._header_host, resource_identifier or '/')

        return '/'.join(seq)


    def _open_url(self, resource_identifier=None, values=None):
        """ Opens a given relative resource from host

        @param resource_identifier (string): relative url
        @param values (dict): dictionary with request parameters
        @return (stream): response from Sisgap server
        """
        response = False

        url = self._build_url(resource_identifier)
        request = self._build_request(url, values)

        self._referer = url

        response = urllib2.urlopen(request)

        return response

    def _open_session(self):
        """ Navigates to the first page of Sisgap and performs a login in site
        Once it has loged in, stores the cookie and the last opened page to
        be used in the next requests

        @return (stream): response from Sisgap server
        """

        # STEP 1: Reset cookie and referer
        self._cookie = None
        self._referer = None

        # STEP 2: Perform a request to get new session cookie
        resource = 'sisgap/paginas/profesores/indice.jsp'
        response = self._open_url(resource, None)
        self._cookie = response.headers.getheader('set-cookie').split(';')[0]

        # STEP 3: Indentify with username and password
        values = {'usuario' : self._user,
                  'password' : self._password,
                  'centro' : self._headquarters}

        resource = 'sisgap/logon.do'
        response = self._open_url(resource, values)

        # STEP 4: Navigate to the next page after login
        resource = 'sisgap/iniciadaLogon.do'

        return self._open_url(resource, None)

    def _close_session(self):
        """ Navigates to the last page of Sisgap and performs a logout in site
        Once it has loged out, resets the cookie and the last opened page.

        @return (stream): response from Sisgap server
        """

        resource = 'sisgap/cerrarSesion.do'

        self._cookie = None
        self._referer = None

        return self._open_url(resource, None)

    def _get_day_html_table(self, in_date):
        """ Opens "solicPasarLista.do" HTML page and gets the table which
        contains the timetable

        @note: this method has been detached from _parse_day_html_table becouse
        it exceeded the number of variables (pylint R0914).

        @param in_date (date): date to get the schedule
        @return: elementtree which represents the table
        """

        # STEP 1: Something about
        str_date = '{dt.month}/{dt.day}/{dt.year}'.format(dt=in_date)

        # STEP 2: Gets the HTML page whith contains the timetible
        resource = 'sisgap/profesores/solicPasarLista.do?fecha=%s' % str_date
        response = self._open_url(resource, None)

        # STEP 3: Parse HTML and gets the table
        tree = etree.parse(response, self._htmlparser)
        xpathsel = ('//table[@class="tabla"][9]')

        return tree.xpath(xpathsel)[0]

    def _parse_time_html_cell(self, in_date, html_cell):
        """ Parses the HTML table cell which contains the javascript to navitate
        to the group assistance page.

        @note: this method has been detached from _parse_day_html_table becouse
        it exceeded the number of variables (pylint R0914).

        @param in_date (date): date to get the schedule
        @param html_cell: elementtree which represents the cell
        """

        timeslot = self._get_full_tag_text(html_cell)
        bounds = timeslot.split('-')
        time1 = datetime.strptime(bounds[0].strip(), '%H:%M')
        time2 = datetime.strptime(bounds[1].strip(), '%H:%M')
        time1 = datetime.combine(in_date, time1.time())
        time2 = datetime.combine(in_date, time2.time())

        return (time1, time2)

    def _parse_day_html_table(self, in_date):
        """ Parses the HTML table witch contains the list of groups in a day

        @note: this behavior has been detached from get_day_timetable because
        it must be used in several methods.

        @in_date (date): date to get the schedule

        return (list): returns a list of dictinaries with the following keys:
        (fecha, grupo, hora_fin, hora_inicio, idGrupo, idMateria, materia, tipo, tp)
        """

        table = self._get_day_html_table(in_date)

        # STEP 3: Parse all rows getting storing cell values in dictionaries
        index = 0
        items = []
        for row in table.xpath('child::tr'):
            cells = [cell for cell in row.xpath('child::td')]

            if index:
                group_name = self._get_full_tag_text(cells[0])
                topic_name = self._get_full_tag_text(cells[1])
                time_bounds = self._parse_time_html_cell(in_date, cells[2])
                group_id, topic_id = self._parse_action_html_cell(cells[3])

                values = {
                    u'grupo' : group_name,
                    u'materia' : topic_name,
                    u'tipo' : 'asistencias',
                    u'fecha' : in_date,
                    u'idMateria' : topic_id,
                    u'idGrupo' : group_id,
                    u'hora_inicio' : time_bounds[0],
                    u'hora_fin' : time_bounds[1],
                    u'tp' : ''
                }

                items.append(values)

            index = index + 1

        return items

    # --------------------------- PUBLIC METHODS ------------------------------

    def get_day_timetable(self, in_date=None):
        """ a
        """

        # STEP 1: Ensure a valid date and formated date string mm/dd/yyyy
        if not in_date:
            in_date = date.today()

        assert isinstance(in_date, date), 'Invalid date format: %s ' % in_date


        # STEP 2: Opens new session and performs the log in
        self._open_session()

        # STEP 3: Call private common _get_day_table method
        items = self._parse_day_html_table(in_date)

        # STEP 4: Close session
        self._close_session()

        return items

    def get_week_timetable(self, in_date=None):
        """ a
        """
        first_day = self._get_monday(in_date)

        # STEP 2: Opens new session and performs the log in
        self._open_session()

        # STEP 2: Loop over each one of the week days getting timetable
        items = []
        for index in range(0, 6):
            dayindex = first_day + timedelta(days=index)
            items = items + self._parse_day_html_table(dayindex)

        # STEP 1: Close session
        self._close_session()

        return items

    def get_next_group(self):

        items = self.get_day_table()
        items = sorted(items, key=lambda x: x['hora_inicio'], reverse=False)
        items = filter(lambda x: x['hora_inicio'] >= datetime.now(), items)

        return items[0] if items else None


    def get_student_list(self, in_date=None, group_id=None):
        """ a
        """

        in_date = in_date or date.today()

        items = self.get_day_table(in_date)
        items = sorted(items, key=lambda x: x['hora_inicio'], reverse=False)

        if group_id:
            items = filter(lambda x: x['idGrupo'] == group_id, items)
        else:
            items = filter(lambda x: x['hora_inicio'] <= datetime.now(), items)

        if items:
            values = {
                'tipo' : items[0]['tipo'],
                'fecha' : self._one_digit_date(items[0]['fecha']),
                'idMateria' : items[0]['idMateria'],
                'idGrupo' : items[0]['idGrupo'],
                'hora_inicio' : items[0]['hora_inicio'].strftime('%H:%M'),
                'hora_fin' : items[0]['hora_fin'].strftime('%H:%M'),
                'tp' : items[0]['tp'],
            }

            self._open_session()

            resource = 'sisgap/profesores/edAsistGrupo.do'
            response = self._open_url(resource, values)

            self._close_session()

            tree = etree.parse(response, self._htmlparser)
            xpathsel = ('//table[@class="tabla"][9]')

            table = tree.xpath(xpathsel)[0]

            # STEP 6: Parse all rows getting storing cell values in dictionaries
            index = 0
            items = []
            for row in table.xpath('child::tr'):
                cells = [cell for cell in row.xpath('child::td')]

                if index:
                    name = self._get_full_tag_text(cells[1])
                    surname = self._get_full_tag_text(cells[2])

                    values = {
                        'name' : name.strip().capitalize(),
                        'surname' : surname.strip().capitalize()
                    }

                    items.append(values)


                index = index + 1

            text_file = open("d:\\proyectos\\sisgap\\output.html", "w")
            text_file.write("Purchase Amount: %s" % etree.tostring(table))
            text_file.close()

            return items

    # ------------------------------- GOOGLE ----------------------------------

    @staticmethod
    def _generate_event_uuid(event):
        """ Generates an inmutable UUID (RFC 4122) using time values from event

        @note: this method is used to create a valid UUID will be used in Google
        Calendar events.
        @see: https://developers.google.com/google-apps/calendar/v3/reference/events
        @see: https://docs.python.org/2/library/uuid.html

        @param envent (dict): event retrieved from sisgap timetable
        @return (string): inmutable UUID
        """

        strfecha = event['fecha'].strftime('%m%d')
        strstart = event['hora_inicio'].strftime('%H%M')
        strend = event['hora_fin'].strftime('%H%M')
        strgroupid = unicode(event['idGrupo']).zfill(4)

        string = '{}{}{}{}'.format(strfecha, strstart, strend, strgroupid)

        uuid_obj = uuid.UUID(bytes=string)

        return unicode(uuid_obj).lower().replace('-', '')

    def _get_credentials(self):
        """Gets valid user credentials from storage.

        If nothing has been stored, or if the stored credentials are invalid,
        the OAuth2 flow is completed to obtain the new credentials.

        Returns:
            Credentials, the obtained credential.
        """
        home_dir = os.path.expanduser('~')
        credential_dir = os.path.join(home_dir, '.credentials')
        if not os.path.exists(credential_dir):
            os.makedirs(credential_dir)
        credential_path = os.path.join(credential_dir,
                                       'calendar-python-sisgap.json')

        store = oauth2client.file.Storage(credential_path)
        credentials = store.get()
        if not credentials or credentials.invalid:
            flow = client.flow_from_clientsecrets(
                self._google_client_secret_file, self._google_scopes)
            flow.user_agent = self._google_application_name
            credentials = tools.run_flow(flow, store, None)

        return credentials

    def _build_event(self, timetable_item):
        madrid = pytz.timezone('Europe/Madrid')
        hora_inicio = madrid.localize(timetable_item['hora_inicio']).isoformat()
        hora_fin = madrid.localize(timetable_item['hora_fin']).isoformat()
        uuid_str = self._generate_event_uuid(timetable_item)

        return {
            'id': uuid_str,
            'colorId': 1,
            'summary': timetable_item['grupo'],
            'location': 'Rúa Zamora, 106, 36211 Vigo, Ponteveda',
            'description': u'{} - {}'.format(timetable_item['grupo'], timetable_item['materia']),
            'start': {
                'dateTime': hora_inicio,
                'timeZone': 'Europe/Madrid',
            },
            'end': {
                'dateTime': hora_fin,
                'timeZone': 'Europe/Madrid',
            },
            'attendees': [
                {'email': self._google_mail_address},
            ],
            'reminders': {
                'useDefault': True,
            },
            'status': 'confirmed'
        }

    def _add_events(self, timetable):
        credentials = self._get_credentials()
        http = credentials.authorize(httplib2.Http())
        service = discovery.build('calendar', 'v3', http=http)

        for timetable_item in timetable:
            event_dict = self._build_event(timetable_item)

            try:
                event = service.events().get(
                    calendarId='primary',
                    eventId=event_dict['id']
                ).execute()
            except:
                event = False

            if event:
                service.events().update(
                    calendarId='primary',
                    eventId=event_dict['id'],
                    body=event_dict
                ).execute()
            else:
                service.events().insert(
                    calendarId='primary',
                    body=event_dict
                ).execute()

    def google_sync(self, in_date=None):
        timetable = self.get_week_timetable(in_date)

        self._add_events(timetable)

    def list_events(self):
        credentials = self._get_credentials()
        http = credentials.authorize(httplib2.Http())
        service = discovery.build('calendar', 'v3', http=http)

from pprint import pprint

sisgap = Sisgap()
sisgap.google_sync()
