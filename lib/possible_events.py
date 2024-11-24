'''Class containing the logic to check which event is (in)valid for a certain swimmer'''

from lib.meet_management import SwimMeet, SwimMeetEvent
from lib.club_management import Club, Swimmer

class PossibleEvents:
    '''Contains the logic to check which event is (in)valid for a certain swimmer'''

    def __init__(self, meet: SwimMeet, club: Club):
        self.meet = meet
        self.club = club
        self.swimmer_possible_event_dict: dict = {}
        self.swimmer_invalid_event_dict: dict = {}

    def __check_possible_event(self, swimmer: Swimmer, event: SwimMeetEvent) -> bool:
        # Check gender
        if event.gender == "F" and swimmer.gender == '1':
            return False
        if event.gender == "M" and swimmer.gender == '2':
            return False

        if event.min_age > swimmer.get_age_at(self.meet.age_date) or \
                event.max_age < swimmer.get_age_at(self.meet.age_date):
            return False

        # check if we can register with NT
        # limit times

        return True

    def generate_possible_events_dict(self, groups_to_use: list[str]):
        '''Check for all the swimmers, which event in the meet they
           can compete in'''
        for group in groups_to_use:
            for swimmer in self.club.get_swimmers_from_group(group):
                self.swimmer_possible_event_dict[swimmer.name] = []
                self.swimmer_invalid_event_dict[swimmer.name] = []
                for event in self.meet.get_all_events():
                    if self.__check_possible_event(swimmer, event):
                        self.swimmer_possible_event_dict[swimmer.name].append(event)
                    else:
                        self.swimmer_invalid_event_dict[swimmer.name].append(event)

    def get_valid_events_for_swimmer(self, swimmer_name: str) -> list[SwimMeetEvent]:
        '''Get a list of all the valid events for a given swimmer'''
        return self.swimmer_possible_event_dict[swimmer_name]

    def get_invalid_events_for_swimmer(self, swimmer_name: str) -> list[SwimMeetEvent]:
        '''Get a list of all the invalid events for a given swimmer'''
        return self.swimmer_invalid_event_dict[swimmer_name]
