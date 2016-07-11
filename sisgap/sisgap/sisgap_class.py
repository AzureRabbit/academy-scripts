# -*- coding: utf-8 -*-
#pylint: disable=I0011,W0703,R0903
""" Performs a Web Scraping to the SISGAP online web application
"""

# --------------------------- REQUIRED LIBRARIES ------------------------------

import urllib
import urllib2

import re

from datetime import datetime
from lxml import etree
from HTMLParser import HTMLParser
from ast import literal_eval

# ------------------------------ SISGAP CLASS ---------------------------------

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

    def __init__(self):
        self._referer = None
        self._cookie = None
        self._user = 'jsoto'
        self._password = '32693935X'
        self._headquarters = 'VIGOZA'
        self._htmlparser = etree.HTMLParser()

    # --------------------------- PRIVATE METHODS -----------------------------

    def _build_url(self, resource_identifier):
        """ Builds a fully qualified url joining protocol, host and relative url
            @param resource_identifier: relative url
        """

        seq = ('http:/', self._header_host, resource_identifier or '/')

        return '/'.join(seq)

    def _build_request(self, url, values=None):
        """ Buils new request with needed headers and post arguments

        @param url (string): fully qualified url from resource
        @param values (dict): dictionary with request parameters
        """

        # STEP 1: Creates an urllib2 request with or without values
        if values:
            data = urllib.urlencode(values)
            req = urllib2.Request(url, data)
        else:
            req = urllib2.Request(url)

        # STEP 2: Headers below have been retrieved from Google Chome Session
        req.add_header('Accept', self._header_accpet)
        req.add_header('Accept-Encoding', self._header_accept_encoding)
        req.add_header('Accept-Language', self._header_accept_language)
        req.add_header('Connection', self._header_connection)
        req.add_header('Host', self._header_host)
        req.add_header('Upgrade-Insecure-Requests', self._header_upg_insecure_req)
        req.add_header('User-Agent', self._header_user_agent)

        # STEP 3: First concection have not cookie or referer. Once the
        # conection has been stablished by object instance, _referer will have
        # the las opened url and the _cookie will have the session cookie.
        if self._referer:
            req.add_header('Referer', self._referer)
        if self._cookie:
            req.add_header('Cookie', self._cookie)

        return req

    def _open_url(self, resource_identifier=None, values=None):
        """ Opens a given relative resource from host

        @param resource_identifier (string): relative url
        @param values (dict): dictionary with request parameters
        @return (stream): response from Sisgap server
        """

        # STEP 1: Set default returned value
        response = False

        # STEP 2: build full URL and prepair the requests
        url = self._build_url(resource_identifier)
        request = self._build_request(url, values)

        # STEP 3: Store the URL as
        response = urllib2.urlopen(request)

        # STEP 4: If URL has been opened, the url will be stored to be used
        # as referer URL in future
        if response and response.getcode() == 200:
            self._referer = url

        return response

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

    def _parse_time_html_cell(self, in_date, html_cell):
        """ Parses the HTML table cell which contains the javascript to navitate
        to the group assistance page.

        @note: this method has been detached from _parse_day_html_table becouse
        it exceeded the number of variables (pylint R0914).

        @param in_date (date): date to get the schedule
        @param html_cell: elementtree which represents the cell
        """
        time1 = None
        time2 = None

        try:
            timeslot = self._get_full_tag_text(html_cell)

            bounds = timeslot.split('-')
            time1 = datetime.strptime(bounds[0].strip(), '%H:%M')
            time2 = datetime.strptime(bounds[1].strip(), '%H:%M')
            time1 = datetime.combine(in_date, time1.time())
            time2 = datetime.combine(in_date, time2.time())
        except Exception:
            pass

        return (time1, time2)

    @staticmethod
    def _one_digit_date(_in_date):
        """ Returns an string with a formate date d/m/yyyy

        @note: the url "solicPasarLista.do" needs one argument, this must be
        a date formated as d/m/yyy

        @param _in_date (date): date will be formated
        """
        return u'{}/{}/{}'.format(_in_date.day, _in_date.month, _in_date.year)

    # ------------------------------- SESSION ---------------------------------


    def open_session(self):
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

    def close_session(self):
        """ Navigates to the last page of Sisgap and performs a logout in site
        Once it has loged out, resets the cookie and the last opened page.

        @return (stream): response from Sisgap server
        """

        resource = 'sisgap/cerrarSesion.do'

        self._cookie = None
        self._referer = None

        return self._open_url(resource, None)

    # -------------------------- solicPasarLista.do ---------------------------

    def solic_pasar_lista(self, in_date):
        """ Parses solicPasarLista.do page. This page has the daily timetable

        @param _in_date: date will be passed in HTTP post argument
        @return: list of the dictinaries contain the values
        """

        items = []

        # STEP 1: Something about
        str_date = '{dt.month}/{dt.day}/{dt.year}'.format(dt=in_date)

        # STEP 2: Gets the HTML page whith contains the timetible
        resource = 'sisgap/profesores/solicPasarLista.do?fecha=%s' % str_date
        response = self._open_url(resource, None)

        # STEP 3: Gets the HTML table and all its rows except headers
        table = self._solic_pasar_lista_get_table(response)
        rows = self._solic_pasar_lista_get_rows(table)

        # STEP 4: Parses each one of data rows getting the cell values
        if rows:
            for row in rows:
                values = self._solic_pasar_lista_parse_row(in_date, row)
                if values:
                    items.append(values)

        return items

    def _solic_pasar_lista_get_table(self, response):
        """ Try to get the HTML table from solic_pasar_lista.do page, which
        contains the daily group list.

        @param response: response returned by urllib2.urlopen
        @return (etreeelement): return the HTML table
        """
        table = None

        try:
            xpathsel = ('//table[@class="tabla"][9]')
            tree = etree.parse(response, self._htmlparser)
            table = tree.xpath(xpathsel)[0]
        except Exception:
            pass

        return table


    @staticmethod
    def _solic_pasar_lista_get_rows(table):
        """ Get all data rows from table in solic_pasar_lista.do page except
        headers
        """
        data_rows = []

        if table is not None:
            rows = table.xpath('child::tr')

            if rows and len(rows) > 1:
                data_rows = rows[1:]

        return data_rows

    def _solic_pasar_lista_parse_row(self, in_date, row):
        """ Parses the given HTML row from  _solic_pasar_lista.do page getting
        contains from each one of the cells
        """
        values = None

        cells = [cell for cell in row.xpath('child::td')]

        if len(row.xpath('child::td')) >= 4:
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

        return values

    # --------------------------- edAsistGrupo.do -----------------------------

    def ed_asist_grupo(self, groupinfo):
        """ Parses edAsistGrupo.do page. This page has the group attendance on
        a particular day

        @param groupinfo: group information will be passed in HTTP post argument
        """

        # STEP 1: Set default student list, this will be returned by this method
        students = []

        # STEP 2: Build a dictionary to be used as response POST arguments
        values = {
            'tipo' : groupinfo['tipo'],
            'fecha' : self._one_digit_date(groupinfo['fecha']),
            'idMateria' : groupinfo['idMateria'],
            'idGrupo' : groupinfo['idGrupo'],
            'hora_inicio' : groupinfo['hora_inicio'].strftime('%H:%M'),
            'hora_fin' : groupinfo['hora_fin'].strftime('%H:%M'),
            'tp' : groupinfo['tp'],
        }

        # STEP 2: Open URL to get student list
        resource = 'sisgap/profesores/edAsistGrupo.do'
        response = self._open_url(resource, values)

        # STEP 3: Parse HTML to get the student list table
        tree = etree.parse(response, self._htmlparser)
        table = tree.xpath('//table[@class="tabla"][9]')[0]

        # STEP 4: Run over each one of the rows to get student info
        index = 0
        for row in table.xpath('child::tr'):
            cells = [cell for cell in row.xpath('child::td')]

            if index:
                name = self._get_full_tag_text(cells[1])
                firstname = self._get_full_tag_text(cells[2])
                lastname = self._get_full_tag_text(cells[3])

                values = {
                    'name' : name.strip().capitalize(),
                    'firstname' : firstname.strip().capitalize(),
                    'lastname' : lastname.strip().capitalize()
                }

                students.append(values)

            index = index + 1

        return students

    # -------------------------- modAsistencias.do ----------------------------

    def mod_asistencias(self, _in_date, group_id, presence_data):
        """ Parses modAsistencias.do page. This page fills the group attendance

        @param _in_date: date will be passed in HTTP post argument
        @param group_id: group identifier will be passed in HTTP post argument
        @presence_data: dictionary with the attendance data
        """
        pass

# id_grupo:912
# id_materia:5
# id_empate:0
# fecha:06/01/2016
# horaInicio:12:00:00
# horaFin:14:00:00
# obs_76702525D:
# hi_76702525D:12:00:00
# hf_76702525D:14:00:00
# desfase:0



