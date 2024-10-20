from easygui import fileopenbox
from zipfile import ZipFile
from datetime import date

import xml.etree.ElementTree as ET
import os

class LenexHelper:
    def __init__(self, start_dir: str):
        self.start_dir = start_dir

        # Create tmp folder to unzip/rename
        if not os.path.isdir("tmp"):
            os.mkdir("tmp")
    
    def load_lenex(self):
        # Select the lenex
        self.full_path = fileopenbox(default=self.start_dir,
                                title="Select competition lenex file")

        if self.full_path == None:
            raise ValueError("Invalid file selected")
        
        self.basename = os.path.basename(self.full_path)
        self.dirname = os.path.dirname(self.full_path)

        print(f"Selected lenex: {self.basename}")
        
    def extract_lef_from_lenex(self):
        # Check if the lef is already extracted
        if not os.path.exists(f"tmp/{os.path.splitext(self.basename)[0]}.lef"):
            # Extract the lef from the lxf
            with ZipFile(self.full_path, 'r') as zipped_file:
                # There will only be 1 file, get the name
                self.extracted_filename = zipped_file.namelist()[0]
                # Extract the file
                zipped_file.extractall(path="tmp")

        print(f"Lef extracted from lenex (tmp/{self.extracted_filename})")

    def load_xml_from_lef(self):
        self.xml_root = ET.parse(f"tmp/{self.extracted_filename}").getroot()

        if self.xml_root.tag != "LENEX":
            raise ValueError("Extracted xml is not a lenex!")


class SwimMeet:
    """Class to group the information of a meet"""

    def __init__(self) -> None:
        # General meet information
        self.meet_name = None
        self.course = None
        self.qualify_date_range = None
        self.deadline = None
        self.age_date = None
        self.program = dict()
    
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
            return f"{node.attrib['relaycount']} x {node.attrib['distance']} {node.attrib['stroke']}"

        return f"{node.attrib['distance']} {node.attrib['stroke']}"

    def __parse_agegroups_node(self, node: ET.Element) -> str:
        """Construct a structured string with all ages"""
        age = ""
        for age_group in node:
            min_age = age_group.attrib["agemin"]
            max_age = age_group.attrib["agemax"]
            if min_age == "-1" and max_age == "-1":
                age = age + "/open"
            elif max_age == "-1":
                age = age + "/" + min_age + "+"
            else:
                if min_age == max_age:
                    age = age + "/" + min_age
                else:
                    age = age + "/" + min_age + "-" + max_age
    
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
    
    def __extract_event_information(self, event: ET.Element) -> dict:
        # Event number
        number = event.attrib.get("number", "?")

        # Gender
        gender = event.attrib.get("gender", "X")
        if gender == "X":
            gender = "Mixed"
        
        style = ""
        for event_info in event:
            if event_info.tag == "SWIMSTYLE":
                style = self.__parse_swimstyle_node(event_info)
            elif event_info.tag == "AGEGROUPS":
                age_string = self.__parse_agegroups_node(event_info)

        min_age, max_age, simplified_age = self.__simplify_age(age_string)

        print(f"Extracted number {number}:\t{gender}\t{simplified_age}\t{style}")
        
        return dict(number = number,
                    gender = gender,
                    style = style,
                    min_age = min_age,
                    max_age = max_age,
                    simplified_age = simplified_age
                    )
    
    def __parse_events(self, events_root: ET.Element, s: dict):
        for event in events_root:
            s["events"].append(self.__extract_event_information(event))
    
    def __parse_session(self, session_root: ET.Element):
        # Extract general session information
        session_attributes = session_root.attrib
        s = dict(session_number = session_attributes.get("number", "1"),
                 session_date = session_attributes.get("date", "?"),
                 sessuib_start = session_attributes.get("daytime", "?"),
                 session_end = session_attributes.get("endtime", "?"),
                 session_warmup_start = session_attributes.get("warmupfrom", "?"),
                 session_warmup_until = session_attributes.get("warmupuntil", "?"),
                 session_official_meeting = session_attributes.get("officialmeeting", "?"),
                 events = []
                )
        
        for event_node in session_root:
            if event_node.tag == "EVENTS":
                events_root = event_node
        
        if events_root == None:
            raise ValueError("Session does not have events")
        
        # Parse events and add to temp session dict s
        self.__parse_events(events_root, s)
        
        # Append this session to the program
        self.program[session_attributes.get("name", f"Session {s['session_number']}")] = s
    
    def __parse_sessions(self, meet_root: ET.Element):
        for meet_info in meet_root:
            if meet_info.tag == "SESSIONS":
                session_root = meet_info
                break
        
        if session_root == None:
            raise ValueError("No sessions found in lenex")
        
        for session in session_root:
            self.__parse_session(session)

    def load_from_xml(self, lef_root_node: ET.Element):
        # Extract all information from the xml
        for meets in lef_root_node:
            if meets.tag == "MEETS":
                # Meets will mostly be only 1 meet, multi meet lenex not supported
                meet_root = meets[0]
        
        if meet_root == None:
            raise ValueError("Swimmeet could not be created from the given lenex")

        self.__extract_general_information(meet_root)
        self.__parse_sessions(meet_root)

    def __str__(self):
        return f"MEET:\n{self.meet_name} in {self.city} ({self.course})\n{self.program}"
