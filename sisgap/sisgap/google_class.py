# -*- coding: utf-8 -*-
#pylint: disable=I0011,W0703,R0903,W0702
""" Performs a Web Scraping to the SISGAP online web application
"""

# --------------------------- REQUIRED LIBRARIES ------------------------------


from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools

import os
import uuid
import pytz
import httplib2

# ------------------------------ SISGAP CLASS ---------------------------------

class GoogleCalendar(object):
    """ Updates google calendar
    """

    def __init__(self, email_address):

        self._google_scopes = 'https://www.googleapis.com/auth/calendar'
        self._google_client_secret_file = 'sisgap.google.json'
        self._google_application_name = 'Sisgap'
        self._google_mail_address = email_address


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
        """ Builds a new Google Calendar event
        """

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
        """ Register new calendar event in Google Calendar
        """

        credentials = self._get_credentials()
        http = credentials.authorize(httplib2.Http())
        service = discovery.build('calendar', 'v3', http=http)
        for timetable_day in timetable.values():
            for timetable_item in timetable_day:
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

    def google_sync(self, timetable):
        """ Syncronice Sisgap timetable with Google calendar
        """
        self._add_events(timetable)

    def list_events(self):
        """ List all registered events in Google Calendar
        """
        credentials = self._get_credentials()
        http = credentials.authorize(httplib2.Http())
        service = discovery.build('calendar', 'v3', http=http)
