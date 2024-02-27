# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

import asyncio
import configparser
from datetime import datetime, timedelta
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from graph import Graph

async def main():
    print('Python Graph Tutorial\n')

    # Load settings
    config = configparser.ConfigParser()
    config.read(['config.cfg', 'config.dev.cfg'])
    azure_settings = config['azure']

    graph: Graph = Graph(azure_settings)

    await greet_user(graph)

    choice = -1

    while choice != 0:
        print('Please choose one of the following options:')
        print('0. Exit')
        print('1. Display access token')
        print('2. Check availability')
        print('3. Create and update event')
        print('4. Preview week')
        print('5. Suggest free times')

        try:
            choice = int(input())
        except ValueError:
            choice = -1

        try:
            if choice == 0:
                print('Goodbye...')
            elif choice == 1:
                await display_access_token(graph)
            elif choice == 2:
                await test_is_available(graph)
            elif choice == 3:
                await create_update_event(graph)
            elif choice == 4:
                await preview_week(graph)
            elif choice == 5:
                await suggest_free_times(graph)
            else:
                print('Invalid choice!\n')
        except ODataError as odata_error:
            print('Error:')
            if odata_error.error:
                print(odata_error.error.code, odata_error.error.message)

async def greet_user(graph: Graph):
    user = await graph.get_user()
    if user:
        print('Hello,', user.display_name)
        # For Work/school accounts, email is in mail property
        # Personal accounts, email is in userPrincipalName
        print('Email:', user.mail or user.user_principal_name, '\n')

async def display_access_token(graph: Graph):
    token = await graph.get_user_token()
    print('User token:', token, '\n')

async def test_is_available(graph: Graph):
    user = await graph.get_user()
    if user:
        print("Start datetime:")
        start = promptDateTime()
        print("End datetime:")
        end = promptDateTime()
        print("You chose:", start, "to", end, "\n")

        res = await graph.isAvailable(user.mail or user.user_principal_name, start, end)
        print("Is time available? " + str(res) + "\n")

async def create_update_event(graph: Graph):
    user = await graph.get_user()
    if user:
        print("Start datetime:")
        start = promptDateTime()
        print("End datetime:")
        end = promptDateTime()
        print("Title:")
        title = input()
        print("You chose:", start, "to", end, "with title", title, "\n")

        res = await graph.createEvent(user.mail or user.user_principal_name, start, end, title)
        if (res is None):                       # < This check doesnt catch nulls, fix it in another universe
            print("Failed to create event")
            return

        print("Event id: " + res)
        res = await graph.updateEvent(user.mail or user.user_principal_name, start + timedelta(1), end + timedelta(1), title, res)
        print("Event moved 1 day ahead? " + str(res) + "\n")

async def preview_week(graph: Graph):
    res = await graph.previewSchedule(True)
    print(str(res) + "\n")

async def suggest_free_times(graph: Graph):
    user = await graph.get_user()
    if user:
        print("Start of window datetime:")
        start = promptDateTime()
        print("End of window datetime:")
        end = promptDateTime()
        print("You chose:", start, "to", end, "\n")

        res = await graph.suggestAlternativeTimes(user.mail or user.user_principal_name, start, end, 60)
        print(str(res) + "\n")

def promptDateTime():
    print("Year: ")
    year = int(input())
    print("Month: ")
    month = int(input())
    print("Day: ")
    day = int(input())
    print("Hour: ")
    hour = int(input())
    print("Min: ")
    minute = int(input())
    print("\n")
    return datetime(year, month, day, hour=hour, minute=minute)
    
# Run main
asyncio.run(main())
