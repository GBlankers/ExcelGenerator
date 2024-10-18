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

    def load_from_xml(self, lef_root_node: ET.Element):
        # Extract all information from the xml
        for meets in lef_root_node:
            if meets.tag == "MEETS":
                meet_root = meets[0]
        
        if meet_root == None:
            raise ValueError("Swimmeet could not be created from the given lenex")

        self.__extract_general_information(meet_root)

    def __str__(self):
        return f"MEET:\n{self.meet_name} in {self.city} ({self.course})"
