# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

# isAvailable: https://learn.microsoft.com/en-us/graph/api/calendar-getschedule?view=graph-rest-1.0&tabs=http
# createEvent: https://learn.microsoft.com/en-us/graph/api/calendar-post-events?view=graph-rest-1.0&tabs=http
# updateEvent: https://learn.microsoft.com/en-us/graph/api/event-update?view=graph-rest-1.0&tabs=http
# previewSchedule: https://learn.microsoft.com/en-us/graph/api/calendar-list-calendarview?view=graph-rest-1.0&tabs=http
# suggestAlternateTime: https://learn.microsoft.com/en-us/graph/api/user-findmeetingtimes?view=graph-rest-1.0&tabs=http

import time
from datetime import datetime, timedelta
from configparser import SectionProxy
from azure.identity import DeviceCodeCredential
from msgraph import GraphServiceClient
from msgraph.generated.users.item.user_item_request_builder import UserItemRequestBuilder
from msgraph.generated.users.item.calendar.get_schedule.get_schedule_request_builder import GetScheduleRequestBuilder
from msgraph.generated.users.item.calendar.get_schedule.get_schedule_post_request_body import GetSchedulePostRequestBody
from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
from msgraph.generated.users.item.events.events_request_builder import EventsRequestBuilder
from msgraph.generated.models.event import Event
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.location import Location
from msgraph.generated.models.attendee import Attendee
from msgraph.generated.models.attendee_type import AttendeeType
from msgraph.generated.models.attendee_base import AttendeeBase
from msgraph.generated.models.email_address import EmailAddress
from msgraph.generated.models.response_status import ResponseStatus
from msgraph.generated.users.item.calendar.calendar_view.calendar_view_request_builder import CalendarViewRequestBuilder
from msgraph.generated.users.item.find_meeting_times.find_meeting_times_request_builder import FindMeetingTimesRequestBuilder
from msgraph.generated.users.item.find_meeting_times.find_meeting_times_post_request_body import FindMeetingTimesPostRequestBody
from msgraph.generated.models.time_constraint import TimeConstraint
from msgraph.generated.models.time_slot import TimeSlot



class Graph:
    settings: SectionProxy
    device_code_credential: DeviceCodeCredential
    user_client: GraphServiceClient

    def __init__(self, config: SectionProxy):
        self.settings = config
        client_id = self.settings['clientId']
        tenant_id = self.settings['tenantId']
        graph_scopes = self.settings['graphUserScopes'].split(' ')

        self.device_code_credential = DeviceCodeCredential(client_id, tenant_id = tenant_id)
        self.user_client = GraphServiceClient(self.device_code_credential, graph_scopes)

    async def get_user_token(self):
        graph_scopes = self.settings['graphUserScopes']
        access_token = self.device_code_credential.get_token(graph_scopes)
        return access_token.token

    async def get_user(self):
        # Only request specific properties using $select
        query_params = UserItemRequestBuilder.UserItemRequestBuilderGetQueryParameters(
            select=['displayName', 'mail', 'userPrincipalName']
        )

        request_config = UserItemRequestBuilder.UserItemRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

        user = await self.user_client.me.get(request_configuration=request_config)
        return user

    async def isAvailable(self, user_email: str, start: datetime, end: datetime):
        request_body = GetSchedulePostRequestBody(
            schedules = [
                user_email,
            ],
            start_time = DateTimeTimeZone(
                date_time = start.isoformat(),
                time_zone = "Pacific Standard Time",
            ),
            end_time = DateTimeTimeZone(
                date_time = end.isoformat(),
                time_zone = "Pacific Standard Time",
            ),
            availability_view_interval = 30,
        )

        request_configuration = GetScheduleRequestBuilder.GetScheduleRequestBuilderPostRequestConfiguration()
        request_configuration.headers.add("Prefer", "outlook.timezone=\"Pacific Standard Time\"")
        result = await self.user_client.me.calendar.get_schedule.post(request_body, request_configuration = request_configuration)
    
        if len(result.value) == 0:
            return False
        else:
            for event in result.value[0].schedule_items:
                e_start = datetime.fromisoformat(event.start.date_time.split(".")[0])
                e_end = datetime.fromisoformat(event.end.date_time.split(".")[0])

                if (start >= e_start and start <= e_end) or (end >= e_start and end <= e_end) or (start < e_start and end > e_end):
                    return False

            return True
    
    async def createEvent(self, user_email: str, start: datetime, end: datetime, title: str):
        if (not (await self.isAvailable(user_email, start, end))):
            return ""

        request_body = Event(
            subject = title,
            start = DateTimeTimeZone(
                    date_time = start.isoformat(),
                    time_zone = "Pacific Standard Time",
            ),
            end = DateTimeTimeZone(
                    date_time = end.isoformat(),
                    time_zone = "Pacific Standard Time",
            ),
            allow_new_time_proposals = True,
            attendees = [
		Attendee(
                    email_address = EmailAddress(
                        address = user_email
		    ),
                    type = AttendeeType.Required,
		),
            ],
        )

        request_configuration = EventsRequestBuilder.EventsRequestBuilderPostRequestConfiguration()
        request_configuration.headers.add("Prefer", "outlook.timezone=\"Pacific Standard Time\"")
        result = await self.user_client.me.events.post(request_body, request_configuration = request_configuration)
        return result.id

    async def updateEvent(self, user_email: str, new_start: datetime, new_end: datetime, title: str, event_id: str):
        if (not (await self.isAvailable(user_email, new_start, new_end))):
            return False

        request_body = Event(
            original_start_time_zone = "originalStartTimeZone-value",
            original_end_time_zone = "originalEndTimeZone-value",
            subject = title,
            start = DateTimeTimeZone(
                    date_time = new_start.isoformat(),
                    time_zone = "Pacific Standard Time",
            ),
            end = DateTimeTimeZone(
                    date_time = new_end.isoformat(),
                    time_zone = "Pacific Standard Time",
            ),
        )

        result = await self.user_client.me.events.by_event_id(event_id).patch(request_body)
        result_start = datetime.fromisoformat(result.start.date_time.split(".")[0])
        result_end = datetime.fromisoformat(result.end.date_time.split(".")[0])
        return  result_start == new_start and result_end == new_end 	# might be sketchy

    async def previewSchedule(self, toggle_week_view: bool):
        today = datetime.now()
        tomorrow = today + timedelta(1)
        endofday = datetime(tomorrow.year, tomorrow.month, tomorrow.day, 0, 0, 0)
        endofweek = today + timedelta(6 - today.weekday())
        altzone = time.altzone if time.daylight and time.localtime().tm_isdst > 0 else time.timezone
        tz = '{}{:0>2}:{:0>2}'.format('-' if altzone > 0 else '+', abs(altzone) // 3600, abs(altzone // 60) % 60)

        print(today.strftime("%Y-%m-%dT%H:%M:%S") + tz)
        print((endofweek if toggle_week_view else end_of_day).strftime("%Y-%m-%dT%H:%M:%S") + tz)
        query_params = CalendarViewRequestBuilder.CalendarViewRequestBuilderGetQueryParameters(
	    start_date_time = today.strftime("%Y-%m-%dT%H:%M:%S") + tz,
	    end_date_time = (endofweek if toggle_week_view else end_of_day).strftime("%Y-%m-%dT%H:%M:%S") + tz,
        )

        request_configuration = CalendarViewRequestBuilder.CalendarViewRequestBuilderGetRequestConfiguration(
            query_parameters = query_params,
        )

        result = await self.user_client.me.calendar.calendar_view.get(request_configuration = request_configuration)
        return result

    async def suggestAlternativeTimes(self, user_email: str, start_window: datetime, end_window: datetime, mins: int):
        if (mins == 0):
            return None
        if (mins // 60 == 0):
            iso_duration = "PT" + str(mins) + "M"
        elif (mins % 60 == 0):
            iso_duration = "PT" + str(mins / 60) + "H"
        else:
            iso_duration = "PT" + str(mins // 60) + "H" + str(mins % 60) + "M"
    
        request_body = FindMeetingTimesPostRequestBody(
            attendees = [
                AttendeeBase(
                    type = AttendeeType.Required,
                    email_address = EmailAddress(
                        address = user_email,
                    ),
                ),
            ],
            time_constraint = TimeConstraint(
                time_slots = [
                    TimeSlot(
                        start = DateTimeTimeZone(
                            date_time = start_window.isoformat(),
                            time_zone = "Pacific Standard Time",
                        ),
                        end = DateTimeTimeZone(
                            date_time = end_window.isoformat(),
                            time_zone = "Pacific Standard Time",
                        ),
                    ),
                ],
            ),
            meeting_duration = iso_duration,
        )

        request_configuration = FindMeetingTimesRequestBuilder.FindMeetingTimesRequestBuilderPostRequestConfiguration()
        request_configuration.headers.add("Prefer", "outlook.timezone=\"Pacific Standard Time\"")
        result = await self.user_client.me.find_meeting_times.post(request_body, request_configuration = request_configuration)
        return result
