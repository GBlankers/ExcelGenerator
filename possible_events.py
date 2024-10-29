from meet_management import SwimMeet, SwimMeetEvent
from club_management import Club, Swimmer

class PossibleEvents:

    def __init__(self, meet: SwimMeet, club: Club):
        self.meet = meet
        self.club = club
        self.swimmer_possible_event_dict = dict()
        self.swimmer_invalid_event_dict = dict()

    def __check_possible_event(self, swimmer: Swimmer, event: SwimMeetEvent) -> bool:
        # Check gender
        if event.gender == "F" and swimmer.gender == '1':
            return False
        elif event.gender == "M" and swimmer.gender == '2':
            return False
        
        if event.min_age > swimmer.get_age_at(self.meet.age_date) or event.max_age < swimmer.get_age_at(self.meet.age_date):
            return False

        # TODO: check if we can register with NT
        # TODO: limit times
    
        return True

    def generate_possible_events_dict(self, groups_to_use: list[str]):
        for group in groups_to_use:
            for swimmer in self.club.get_swimmers_from_group(group):
                self.swimmer_possible_event_dict[swimmer.name] = []
                self.swimmer_invalid_event_dict[swimmer.name] = []
                for event in self.meet.get_all_events():
                    if self.__check_possible_event(swimmer, event):
                        self.swimmer_possible_event_dict[swimmer.name].append(event)
                    else:
                        self.swimmer_invalid_event_dict[swimmer.name].append(event)

    def get_invalid_events_for_swimmer(self, swimmer_name: str) -> list[SwimMeetEvent]:
        return self.swimmer_invalid_event_dict[swimmer_name]
