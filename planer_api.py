from __future__ import print_function
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import calendar
from datetime import date, timedelta

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']

def main():
    method = Methods()
    # method.get_calendar_list()
    method.get_events_list("o685qbmudvk6su1j6di61utk8k@group.calendar.google.com", tmin='2020-03-03T07:33:24.149205Z', tmax='2020-05-03T07:33:24.149205Z')
    # method.create_event(
    #     summary="Внушительный Костя",
    #     location="РАБота",
    #     dateTime_time_start='19:00:00',
    #     dateTime_date_start='2020-03-05',
    #     dateTime_time_end='21:00:00',
    #     dateTime_date_end='2020-03-05',
    #     email='alexe0110@yandex.ru',
    #     visibility='public',
    #     description='Очень внушительынй',
    #     calendarId='o685qbmudvk6su1j6di61utk8k@group.calendar.google.com',
    # )


class Auth():
    """
    Атрибут service дальше нужен для авторизации, поэтому наследуюсь от этого класса
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)


class Methods(Auth):
    def get_events_list(self, calendarId='ru.russian#holiday@group.v.calendar.google.com', tmin='now', tmax='through_mounth', quantity=100):
        """
        :param calendarId: ID календаря
        :param tmin: Время начала в формате 2020-03-03T07:33:24.149205Z(необязательно) - по умолчанию сегодня
        :param tmax: Время окончанния в формате 2020-03-03T07:33:24.149205Z(необязательно) - по умолчанию через месяц
        :param quantity: Количество событий (необязательно) - по умолчанию 100
        :return: Предстоящие quantity событий в указанном календаре
        """

        # Сегодня
        now = datetime.datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time

        # Дата через месяц
        today = date.today()
        days = calendar.monthrange(today.year, today.month)[1]
        through_mounth = (today + timedelta(days=days)).isoformat() + 'T23:59:59.000000Z'

        if tmin == 'now': tmin = now
        if tmax == 'through_mounth': tmax = through_mounth

        events_result = self.service.events().list(calendarId=calendarId,
                                                   timeMin=tmin,
                                                   timeMax=tmax,
                                                   singleEvents=True,
                                                   maxResults=quantity,
                                                   orderBy='startTime').execute()
        events = events_result.get('items', [])

        if not events:
            print('No upcoming events found.')
        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            print(start, event['summary'])

    def get_calendar_list(self):
        """
        :return: Список календарей
        """
        page_token = None
        while True:
            calendar_list = self.service.calendarList().list(pageToken=page_token).execute()
            for calendar_list_entry in calendar_list['items']:
                print(calendar_list_entry['summary'], calendar_list_entry['id'])
            page_token = calendar_list.get('nextPageToken')
            if not page_token:
                break

    def create_event(
            self,
            summary,
            location,
            dateTime_time_start,
            dateTime_date_start,
            dateTime_time_end,
            dateTime_date_end,
            email,
            freq='DAILY',
            freq_count=1,  # test
            description='Какое то событие',
            visibility='default',
            calendarId='o685qbmudvk6su1j6di61utk8k@group.calendar.google.com'):
        """
        :param summary: Название
        :param location: Место
        :param dateTime_time_start: Время начало (21:00:00)
        :param dateTime_date_start: Дата начало (2020-03-02)
        :param dateTime_time_end: Время начало (21:00:00)
        :param dateTime_date_end: Дата начало (2020-03-02)
        :param email: Участник (его эл почта)
        :param freq: Частота повторения (DAILY ежедневно, WEEKLY - еженедельно, MONTHLY - ежемесячно) (необязательно)
        :param freq_count: Частота повторения количество (необязательно)
        :param description: Описание события (необязательно)
        :param visibility: Настройка приватности
        :param calendarId: ID календаря (необязательно)
        :return: Ссылка на событие
        """
        event = {
            'summary': summary,
            'location': location,
            'description': description,
            'start': {
                'dateTime': dateTime_date_start + 'T' + dateTime_time_start + '+05:00',
                'timeZone': 'Asia/Yekaterinburg',
            },
            'end': {
                'dateTime': dateTime_date_end + 'T' + dateTime_time_end + '+05:00',
                'timeZone': 'Asia/Yekaterinburg',
            },
            'recurrence': [
                f'RRULE:FREQ={freq};COUNT={freq_count}'
            ],
            'attendees': [  # Участники
                {'email': email},
            ],
            'reminders': {
                'useDefault': False,
                'overrides': [
                    {'method': 'email', 'minutes': 60},
                    {'method': 'popup', 'minutes': 10},
                ],
            },
            "visibility": visibility
        }

        event = self.service.events().insert(calendarId=calendarId, body=event).execute()
        print('Event created: %s' % (event.get('htmlLink')))
        return event.get('htmlLink')


if __name__ == '__main__':
    main()
