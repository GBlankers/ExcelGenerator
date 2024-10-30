'''
Contains classes to facilitate lenex handling, extracting 
the lenex and reading all the meet information from it
'''

import os
import logging

from dataclasses import dataclass
from zipfile import ZipFile
from datetime import date

import xml.etree.ElementTree as ET

from easygui import fileopenbox

class LenexHelper:
    '''Helper class with methods to read and extract the xml
       from the lenex'''
    def __init__(self, log: logging.Logger, start_dir: str):
        self.start_dir = start_dir
        self.log = log
        self.full_path: str = None
        self.basename: str = None
        self.dirname: str = None
        self.extracted_filename: str = None
        self.xml_root: ET.Element = None

        # Create tmp folder to unzip/rename
        if not os.path.isdir("tmp"):
            os.mkdir("tmp")
            self.log.debug("Created tmp folder")

    def load_lenex(self):
        '''Select the lenex using a gui file explorer'''
        # Select the lenex
        self.full_path = fileopenbox(default=self.start_dir,
                                title="Select competition lenex file")

        if self.full_path is None:
            raise ValueError("Invalid file selected")

        self.basename = os.path.basename(self.full_path)
        self.dirname = os.path.dirname(self.full_path)

        self.log.info(f"Selected lenex: {self.basename}")

    def extract_lef_from_lenex(self):
        '''Extract the lef from the lenex. This boils down to unzipping
           the lenex file and placing it into tmp'''
        # Check if the lef is already extracted
        if not os.path.exists(f"tmp/{os.path.splitext(self.basename)[0]}.lef"):
            # Extract the lef from the lxf
            with ZipFile(self.full_path, 'r') as zipped_file:
                # There will only be 1 file, get the name
                self.extracted_filename = zipped_file.namelist()[0]
                # Extract the file
                zipped_file.extractall(path="tmp")

        self.log.info(f"Lef extracted from lenex (tmp/{self.extracted_filename})")

    def load_xml_from_lef(self):
        '''Get the xml root from the previously extracted lef'''
        self.xml_root = ET.parse(f"tmp/{self.extracted_filename}").getroot()

        if self.xml_root.tag != "LENEX":
            raise ValueError("Extracted xml is not a lenex!")

@dataclass
class SwimMeetEvent:
    '''Dataclass containing information about a swim meet event'''
    number: int
    gender: str
    style: str
    min_age: int
    max_age: int
    simplified_age: str
    round: str

    def __str__(self):
        return f"{self.round} #{self.number} {self.gender} {self.style} {self.simplified_age}"

    def __hash__(self):
        return hash(self.__str__())

class SwimMeet:
    """Class to group the information of a meet"""

    def __init__(self, log: logging.Logger) -> None:
        # General meet information
        self.meet_name = None
        self.course = None
        self.qualify_date_range = None
        self.deadline = None
        self.age_date = None
        self.program: dict = {}
        self.city: str = None

        self.log = log

    def __extract_general_information(self, meet_root: ET.Element):
        today = date.today().strftime("%d%m%Y")

        self.meet_name = meet_root.attrib.get('name', '?').replace('/', '-')
        self.city = meet_root.attrib.get('city', '?')
        self.course = meet_root.attrib.get('course', 'LCM')
        self.deadline = meet_root.attrib.get('deadline', today)

        for meet_info in meet_root:
            if meet_info.tag == "QUALIFY":
                qualify_from = meet_info.attrib.get("from", today)
                qualify_until = meet_info.attrib.get("until", self.deadline)
                self.qualify_date_range = f"{qualify_from} -> {qualify_until}"
            elif meet_info.tag == "AGEDATE":
                self.age_date = meet_info.attrib.get("value", today)

    def __parse_swimstyle_node(self, node: ET.Element) -> str:
        if int(node.attrib["relaycount"]) > 1:
            return f"{node.attrib['relaycount']} x {node.attrib['distance']}\
                {node.attrib['stroke']}"

        return f"{node.attrib['distance']} {node.attrib['stroke']}"

    def __parse_agegroups_node(self, node: ET.Element) -> str:
        """Construct a structured string with all ages"""
        age = ""
        for age_group in node:
            min_age = age_group.attrib["agemin"]
            max_age = age_group.attrib["agemax"]
            if min_age == "-1" and max_age == "-1":
                age = f"{age}/open"
            elif max_age == "-1":
                age = f"{age}/{min_age}+"
            else:
                if min_age == max_age:
                    age = f"{age}/{min_age}"
                else:
                    age = f"{age}/{min_age}-{max_age}"
        return age

    def __simplify_age(self, age_str: str) -> tuple[int, int, str]:
        if not age_str or age_str == "":
            return 0, 99, "open"

        temp_str = ("/".join(age_str.split("-"))).replace("+", "").replace("open", "")
        age_array = [int(x) for x in temp_str.split("/") if x != ""]
        age_max = max(age_array)
        age_min = min(age_array)

        if "+" in age_str or "open" in age_str:
            simplified_str = f"{age_min} -> open"
            age_max = 99
        else:
            simplified_str = f"{age_min} -> {age_max}"

        return age_min, age_max, simplified_str

    def __extract_event_information(self, event: ET.Element) -> SwimMeetEvent:
        # Event number
        number = int(event.attrib.get("number", "?"))

        # Gender
        gender = event.attrib.get("gender", "X")
        if gender == "X":
            gender = "Mixed"

        event_round = event.attrib.get("round", "PRE")

        style = ""
        age_string = ""
        for event_info in event:
            if event_info.tag == "SWIMSTYLE":
                style = self.__parse_swimstyle_node(event_info)
            elif event_info.tag == "AGEGROUPS":
                age_string += self.__parse_agegroups_node(event_info)

        min_age, max_age, simplified_age = self.__simplify_age(age_string)


        return SwimMeetEvent(number, gender, style, min_age, max_age, simplified_age,
                             event_round)

    def __parse_events(self, events_root: ET.Element, s: dict):
        for event in events_root:
            event_info_node = self.__extract_event_information(event)

            if event_info_node.round == "FIN":
                self.log.debug(f"Skipped final number {event_info_node.number}:\t\
                               {event_info_node.gender}\t{event_info_node.simplified_age}\
                                \t{event_info_node.style}")

            self.log.debug(f"Extracted number {event_info_node.number}:\t{event_info_node.gender}\
                           \t{event_info_node.simplified_age}\t{event_info_node.style}")
            s["events"].append(event_info_node)

    def __parse_session(self, session_root: ET.Element):
        # Extract general session information
        session_attributes = session_root.attrib
        s = {'session_number': session_attributes.get("number", "1"),
             'session_date': session_attributes.get("number", "1"),
             'session_start': session_attributes.get("daytime", "?"),
             'session_end': session_attributes.get("endtime", "?"),
             'session_warmup_start': session_attributes.get("warmupfrom", "?"),
             'session_warmup_until': session_attributes.get("warmupuntil", "?"),
             'session_official_meeting': session_attributes.get("officialmeeting", "?"),
             'events': []}

        events_root = None
        for event_node in session_root:
            if event_node.tag == "EVENTS":
                events_root = event_node

        if events_root is None:
            raise ValueError("Session does not have events")

        # Parse events and add to temp session dict s
        self.__parse_events(events_root, s)

        # Append this session to the program
        self.program[session_attributes.get("name", f"Session {s['session_number']}")] = s

    def __parse_sessions(self, meet_root: ET.Element):
        session_root = None
        for meet_info in meet_root:
            if meet_info.tag == "SESSIONS":
                session_root = meet_info
                break

        if session_root is None:
            raise ValueError("No sessions found in lenex")

        for session in session_root:
            self.__parse_session(session)

    def load_from_xml(self, lef_root_node: ET.Element):
        '''Extract all the meet information from the given lef root node'''
        # Extract all information from the xml
        meet_root = None
        for meets in lef_root_node:
            if meets.tag == "MEETS":
                # Meets will mostly be only 1 meet, multi meet lenex not supported
                meet_root = meets[0]
                break

        if meet_root is None:
            raise ValueError("Swimmeet could not be created from the given lenex")

        self.__extract_general_information(meet_root)
        self.__parse_sessions(meet_root)

    def get_all_events(self) -> list[SwimMeetEvent]:
        '''Get a list with all the events in this swim meet'''
        return_list: list[SwimMeetEvent] = []
        for _, events in self.program.items():
            for event in events["events"]:
                return_list.append(event)
        return return_list

    def get_events_in_session(self, session_name: str) -> list[SwimMeetEvent]:
        '''Get all the events in a session'''
        return self.program[session_name]["events"]

    def __str__(self):
        return_str = f"MEET:\n{self.meet_name} in {self.city} ({self.course})\n"
        for session, events in self.program.items():
            return_str += f"{session}\n"
            for event in events["events"]:
                return_str += f"\t{str(event)}\n"
        return return_str

@dataclass
class RankingsEntry:
    '''Container for information about rankings (name, place, time,...)'''
    placing: int
    swim_time: str
    swimmer_name: str
    nationality: str
    club: str

class MeetResults:
    '''Class to extract results from the results lenex and construct rankings
       where filters can be applies (filter out nationalities/clubs). '''

    @dataclass
    class ResultIdEntry:
        '''Hashmap value for result_id -> athlete_id and swimtime'''
        athlete_id: str
        swim_time: str

        def __str__(self) -> str:
            return f"{self.athlete_id}: {self.swim_time}"

    @dataclass
    class RelayIdEntry:
        '''Hashmap value for result_id of relays'''
        swimmer_ids: list[str]
        swim_time: str
        club: str

        def __str__(self) -> str:
            return f"{self.swimmer_ids} - {self.club}"

    @dataclass
    class AthleteIdEntry:
        '''Hashmap value for athlete_id -> name, club, nationality'''
        swimmer_name: str
        nationality: str
        club: str

        def __str__(self) -> str:
            return f"{self.swimmer_name} - {self.club}"

    @dataclass
    class EventResult:
        '''Container with the results of a certain event'''
        event_round: str
        gender: str
        event_name: str
        age_group: str
        rankings: list[str]
        relay: bool

        def __str__(self) -> str:
            return f"{self.event_round} {self.gender} {self.event_name} " + \
                   f"{self.age_group}: {str(self.rankings)}"

    def __init__(self, log: logging.Logger, results_root: ET.Element) -> None:
        self.log = log

        # Meet and lenex information
        self.meet_name = None
        self.meet_root = None
        self.__extract_meet_root(results_root)

        # Bookkeeping the id's and results
        self.result_ids: dict[str, self.ResultIdEntry] = {}
        self.relay_ids: dict[str, self.RelayIdEntry] = {}
        self.athlete_ids: dict[str, self.AthleteIdEntry] = {}
        self.meet_results: list[self.EventResult] = []
        self.results: dict[str, list[RankingsEntry]] = {}
        self.results_relays: dict[str, list[RankingsEntry]] = {}
        self.results_filters: list[str] = []

    def __extract_meet_root(self, results_root: ET.Element) -> ET.Element:
        if self.meet_root is None:
            for node in results_root:
                if node.tag == "MEETS":
                    self.meet_root = node[0]

        if self.meet_root is None:
            raise ValueError("Results could not be extracted from given lenex")

        return self.meet_root

    def __extract_general_information(self) -> None:
        '''Extract some generic information from the lenex 
           For now this will be limited to meet_name'''
        self.meet_name = self.meet_root.attrib.get('name', '?').replace('/', '-')

    def __extract_personal_results(self, athlete_node: ET.Element, club: str):
        # Get general athelete information
        last_name = athlete_node.attrib.get("lastname", "?")
        first_name = athlete_node.attrib.get("firstname", "?")
        athlete_entry = self.AthleteIdEntry(f"{first_name} {last_name}",
                                            athlete_node.attrib.get("nation", "BEL"),
                                            club)

        athlete_id = athlete_node.attrib.get("athleteid", "?")
        self.athlete_ids[athlete_id] = athlete_entry

        # Get the results node
        results_node = None
        for node in athlete_node:
            if node.tag == "RESULTS":
                results_node = node
                break

        # Can be relay only swimmer
        if results_node is None:
            self.log.debug(f"No results found for {athlete_entry.swimmer_name}")
            return

        for result_node in results_node:
            # Remove unneccesary hour prefix
            swim_time = result_node.attrib.get("swimtime", "?")
            if swim_time.startswith("00:"):
                swim_time = swim_time[3:]

            result_entry = self.ResultIdEntry(athlete_id, swim_time)
            self.result_ids[result_node.attrib.get("resultid", "?")] = result_entry

    def __extract_relay_result(self, result_node: ET.Element, club_name: str):
        result_id = result_node.attrib.get("resultid", "?")
        swim_time = result_node.attrib.get("swimtime", "?")
        if swim_time.startswith("00:"):
            swim_time = swim_time[3:]

        relaypositions_nodes = None
        for node in result_node:
            if node.tag == "RELAYPOSITIONS":
                relaypositions_nodes = node
                break

        if relaypositions_nodes is None:
            self.log.debug("No athletes in relay")
            return

        relay_athlete_ids: list[str] = []
        for relay_position in relaypositions_nodes:
            relay_athlete_ids.append(relay_position.attrib.get("athleteid", "?"))

        self.relay_ids[result_id] = self.RelayIdEntry(relay_athlete_ids, swim_time, club_name)

    def __extract_relays_results(self, relays_node: ET.Element, club_name: str):
        for relay_node in relays_node:
            results_node = None
            for node in relay_node:
                if node.tag == "RESULTS":
                    results_node = node
                    break
            if results_node is None:
                continue

            for result_node in results_node:
                self.__extract_relay_result(result_node, club_name)

    def __extract_club_results(self, club_name: str, club_node: ET.Element):
        athletes_node = None
        relays_node = None
        for node in club_node:
            if node.tag == "ATHLETES":
                athletes_node = node
            elif node.tag == "RELAYS":
                relays_node = node

        if athletes_node is None:
            return

        if relays_node is not None:
            self.__extract_relays_results(relays_node, club_name)

        for athlete in athletes_node:
            self.__extract_personal_results(athlete, club_name)

    def __get_clubs_node(self, meet_root: ET.Element) -> ET.Element:
        '''Iterate over the different clubs in the lenex and return the club
           corresponding to the given club name'''

        clubs_node = None
        for node in meet_root:
            if node.tag == "CLUBS":
                clubs_node = node
                break

        if clubs_node is None:
            raise ValueError("Cannot find clubs node in lenex")

        return clubs_node

    def __parse_individual_results(self):
        clubs_node = self.__get_clubs_node(self.meet_root)

        for club_node in clubs_node:
            self.__extract_club_results(club_node.attrib.get("code", "?"), club_node)

    def __parse_age(self, agemin: str, agemax: str) -> str:
        if agemin == "-1" and agemax == "-1":
            return "open"
        if agemax == "-1":
            return f"{agemin}+"
        if agemin == "-1":
            return f"-{agemax}"
        if agemin == agemax:
            return agemax

        return f"{agemin}-{agemax}"

    def __parse_agegroups(self, agegroups_node: ET.Element, gender: str,
                          event_round: str, event_name: str, relay: bool) -> None:
        for agegroup_node in agegroups_node:
            age = self.__parse_age(agegroup_node.attrib.get("agemin", "-1"),
                                   agegroup_node.attrib.get("agemax", "-1"))

            rankings_node = None
            for node in agegroup_node:
                if node.tag == "RANKINGS":
                    rankings_node = node
                    break

            if rankings_node is None:
                self.log.debug(f"No rankings found for {event_name}")
                return

            order = []
            for ranking in rankings_node:
                order.append(ranking.attrib.get("resultid", "?"))

            self.meet_results.append(self.EventResult(event_round, gender, event_name,
                                                      age, order, relay))

    def __extract_results_from_event(self, event_node: ET.Element):
        event_gender = event_node.attrib.get("gender", "?")
        event_round = event_node.attrib.get("round", "PRE")

        agegroups_node = None
        event_name = ""
        relay = False
        for node in event_node:
            if node.tag == "SWIMSTYLE":
                relaycount = node.attrib.get("relaycount", "1")
                relay = False
                relay_prepend = ""
                if relaycount != "1":
                    relay_prepend = f"{relaycount} x "
                    relay = True

                event_name = f"{relay_prepend}{node.attrib.get('distance')}" + \
                             f"{node.attrib.get('stroke')}"
            elif node.tag == "AGEGROUPS":
                agegroups_node = node

        if agegroups_node is None:
            self.log.debug(f"No agegroups for event {event_name}")
            return

        self.__parse_agegroups(agegroups_node, event_gender, event_round, event_name, relay)

    def __extract_results_from_session(self, session_node: ET.Element):
        for node in session_node:
            if node.tag == "EVENTS":
                events_node = node
                break

        if events_node is None:
            raise ValueError("No events found in session")

        for event in events_node:
            self.__extract_results_from_event(event)

    def __extract_meet_results(self):
        for node in self.meet_root:
            if node.tag == "SESSIONS":
                sessions_node = node
                break

        if sessions_node is None:
            raise ValueError(f"No sessions found for {self.meet_name}")

        for session_node in sessions_node:
            self.__extract_results_from_session(session_node)

    def __get_ranking_entry_for_result_id(self, result_id: str, placing: int) -> RankingsEntry:
        result_id_entry = self.result_ids.get(result_id)
        if result_id_entry is None:
            raise ValueError(f"Unkown result id {result_id}")

        athlete_id_entry = self.athlete_ids.get(result_id_entry.athlete_id)
        if athlete_id_entry is None:
            raise ValueError(f"Unknown athlete id {result_id_entry.athlete_id}")

        return RankingsEntry(placing, result_id_entry.swim_time, athlete_id_entry.swimmer_name,
                             athlete_id_entry.nationality, athlete_id_entry.club)

    def __get_ranking_entry_for_result_id_relay(self, result_id: str,
                                                placing: int) -> RankingsEntry:
        relay_id_entry = self.relay_ids.get(result_id)
        if relay_id_entry is None:
            raise ValueError(f"Unknown relay id {result_id}")

        swimmer_names = ""
        for swimmer_id in relay_id_entry.swimmer_ids:
            athlete_entry = self.athlete_ids.get(swimmer_id)
            swimmer_names = f"{swimmer_names} {athlete_entry.swimmer_name} - "

        swimmer_names = swimmer_names[:-3]

        return RankingsEntry(placing, relay_id_entry.swim_time, swimmer_names,
                             "", relay_id_entry.club)

    def __parse_filters(self, filters: list[str]) -> tuple[list, str, str]:
        return_filters = []
        return_nat = ""
        return_club = ""

        for f in filters:
            if "ONLY_NAT" in f:
                return_nat = f.split("=")[1]
            elif "ONLY_CLUB" in f:
                return_club = f.split("=")[1]
            else:
                return_filters.append(f)

        return return_filters, return_nat, return_club

    def construct_rankings(self, filters: list[str]):
        '''Create rankings from the lenex results and use filters to
           get only club/nationalities that we are interested in'''
        # Preparation to be able to construct the rankings
        self.__extract_general_information()
        self.__parse_individual_results()
        self.__extract_meet_results()

        self.results_filters = filters
        rest_filters, nationality, club = self.__parse_filters(filters)
        # Construct the total rankings
        for res in self.meet_results:
            event_name = f"{res.event_round} {res.gender} {res.event_name} {res.age_group}"

            if "ONLY_FINALS" in rest_filters and res.event_round == "PRE":
                continue

            # DNS, DQ, DNF -> placed at the end so doesn't matter that they are included
            placing = 1
            for result_id in res.rankings:
                if "ONLY_PODIUM" in filters and placing >= 4:
                    break

                if not res.relay:
                    ranking_entry = self.__get_ranking_entry_for_result_id(result_id, placing)

                    if nationality != "" and ranking_entry.nationality != nationality:
                        continue

                    if (club != "" and ranking_entry.club == club) or club == "":
                        # Only create results entry is there is an swimmer
                        if event_name not in self.results:
                            self.results[event_name] = []
                        self.results[event_name].append(ranking_entry)
                elif "NO_RELAYS" not in rest_filters: # Relay
                    ranking_entry = self.__get_ranking_entry_for_result_id_relay(result_id, placing)

                    if (club != "" and ranking_entry.club == club) or club == "":
                        if event_name not in self.results_relays:
                            self.results_relays[event_name] = []
                        self.results_relays[event_name].append(ranking_entry)

                placing += 1

    def print_rankings(self):
        '''Print the constructed rankings to the console'''
        self.log.info(f"Printing rankings with following filters applied: {self.results_filters}")
        for event, ranking in self.results.items():
            if ranking == []:
                continue

            self.log.info(f"Rankings for {event}:")
            for e in ranking:
                self.log.info(f"{e.placing}: {e.swimmer_name} - {e.club} - {e.swim_time}")
            self.log.info(" ")

        for event, ranking in self.results_relays.items():
            if ranking == []:
                continue

            self.log.info(f"Results for {event}")
            for e in ranking:
                self.log.info(f"{e.placing}: {e.swimmer_name} - {e.club} - {e.swim_time}")
            self.log.info(" ")

    def __str__(self) -> str:
        return str(self.results)
