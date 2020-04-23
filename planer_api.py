from __future__ import print_function
import datetime, json
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import calendar
from datetime import date, timedelta

# Указание прав
SCOPES = ['https://www.googleapis.com/auth/calendar']


def main():
    method = Methods()
    method.get_calendar_list()
    method.get_events_list(level='13', room='3', tmin='2020-04-19T07:33:24.149205Z', tmax='2020-05-10T07:33:24.149205Z')
    method.create_event(
        summary="Нефть все",
        location="РАБота",
        dateTime_time_start='22:00:00',
        dateTime_date_start='2020-04-20',
        dateTime_time_end='23:00:00',
        dateTime_date_end='2020-04-20',
        visibility='public',
        description='Цена на майский фьючерс WTI достиг -40$ за баррель (спойлер: это пиз**ц).',
        level='13',
        room='2'
    )


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
    def get_events_list(self, level, room, tmin='now', tmax='through_mounth'):
        """
        :param level: Этаж
        :param room: Номер переговорки
        :param tmin: Время начала в формате 2020-03-03T07:33:24.149205Z(необязательно) - по умолчанию сегодня
        :param tmax: Время окончанния в формате 2020-03-03T07:33:24.149205Z(необязательно) - по умолчанию через месяц
        :param quantity: Количество событий (необязательно) - по умолчанию 100 - удалено!
        :return: Предстоящие quantity событий в указанном календаре
        """
        with open('bot_log', 'a') as f:
            f.write(datetime.datetime.today().strftime("%Y.%m.%d-%H:%M:%S") + " " +
                    "get_events_list; " + f"level: {level}; room: {room}; tmin: {tmin}; tmax: {tmax};\n")

        # Сегодня
        now = datetime.datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time

        # Дата через месяц
        today = date.today()
        days = calendar.monthrange(today.year, today.month)[1]
        through_mounth = (today + timedelta(days=days)).isoformat() + 'T23:59:59.000000Z'

        if tmin == 'now': tmin = now
        if tmax == 'through_mounth': tmax = through_mounth

        # Получение calendarId
        calendarId = self.get_calendar_id(level, room)

        events_result = self.service.events().list(calendarId=calendarId,
                                                   timeMin=tmin,
                                                   timeMax=tmax,
                                                   singleEvents=True,
                                                   orderBy='startTime').execute()
        events = events_result.get('items', [])

        # нужен объект реквест, но не нужен список человеческий
        # if not events:
        #     print('No upcoming events found.')
        # for event in events:
        #     start = event['start'].get('dateTime', event['start'].get('date'))
        #     print(start, event['summary'])
        return events

    def get_calendar_list(self):
        """
        :return: Список календарей
        """
        with open('bot_log', 'a') as f:
            f.write(datetime.datetime.today().strftime("%Y.%m.%d-%H:%M:%S") + " " +
                    "get_calendar_list; \n")
        page_token = None
        while True:
            calendar_list = self.service.calendarList().list(pageToken=page_token).execute()
            for calendar_list_entry in calendar_list['items']:
                print(calendar_list_entry['summary'], calendar_list_entry['id'], calendar_list_entry['colorId'])
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
            level,
            room,
            freq='DAILY',
            freq_count=1,  # test
            description='Какое то событие',
            visibility='default'):
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

        with open('bot_log', 'a') as f:
            f.write(datetime.datetime.today().strftime("%Y.%m.%d-%H:%M:%S") + " " +
                    "create_event; " +
                    f"summary: {summary} ; location: {location} ; level: {level}; room: {room}; dateTime_time_start: {dateTime_time_start};  dateTime_time_end: {dateTime_time_end};\n")

        calendarId = self.get_calendar_id(level, room)
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
            "visibility": visibility
        }

        event = self.service.events().insert(calendarId=calendarId, body=event).execute()
        print('Event created: %s' % (event.get('htmlLink')))
        return event.get('htmlLink')

    def get_calendar_id(self, level, room):
        """
        Возвращеает индентификатор календаря для укзанной переговорки и указывает нужный токен для нужного этажа
        :param level: Этаж
        :param room: Номер переговорки
        :return: Идентификатор
        """
        # Справочник календарей для переговорок
        with open("alexe.json", "r") as read_file:
            get_room = json.load(read_file)

        if level == '13':
            if room == '1': return get_room['LEVEL13_ROOM1']
            if room == '2': return get_room['LEVEL13_ROOM2']
            if room == '3': return get_room['LEVEL13_ROOM3']
            if room == '4': return get_room['LEVEL13_ROOM4']
            if room == '5': return get_room['LEVEL13_ROOM5']
            if room == 'vece': return get_room['LEVEL13_VECE']

        elif level == '14':
            pass

        else:
            print("ХЗ какой календарь")
if __name__ == '__main__':
    main()
