from dataclasses import dataclass
from enum import Enum

import os
import csv
import subprocess

class Gender(Enum):
    MALE = 1
    FEMALE = 2

@dataclass
class Swimmer:
    name: str
    age: int
    gender: Gender
    group: str

class Club:
    ACTIVE = "T"

    def __init__(self, name: str):
        self.members = dict()
        self.club_name = name

    def __fill_club_from_members_list(self, members_list: list):
        # Get the indices of the different fields of the headers
        index_last_name = members_list[0].index("LASTNAME")
        index_first_name = members_list[0].index("FIRSTNAME")
        index_gender = members_list[0].index("GENDER")
        index_active = members_list[0].index("ACTIVE")
        index_groups = members_list[0].index("GROUPS")
        index_birth_data = members_list[0].index("BIRTHDATE")

        # Remove the first entry as this contains the headers
        members_list.pop(0)

        for athlete in members_list:
            # Possibility that there are empty elements
            if not athlete:
                continue

            # Do not use inactive members
            if athlete[index_active] != self.ACTIVE:
                continue

            # TODO: clean up the mdb, remove inactive members, correct group names
            # Iterate over all the group of the athlete
            for group in athlete[index_groups].split(", "):
                # Create empty list if unkown group
                if group not in self.members:
                    self.members[group] = []
            
                # Append the athlete to the correct group
                self.members[group].append(athlete[index_first_name])

    def __read_members_from_mdb(self, mdbPath: str) -> str:
        if os.path.splitext(mdbPath)[1] != '.mdb':
            raise ValueError("Given path to database is not an mdb")

        if os.path.exists(os.path.splitext(mdbPath)[0] + '.ldb'):
            raise RuntimeError("Database is locked")

        try:
            result = subprocess.run(['mdb-export', mdbPath, 'MEMBERS'], 
                                    capture_output=True, 
                                    text=True, 
                                    check=True)
        except subprocess.CalledProcessError as e:
            raise RuntimeError(f"Error extracting data from database: {e}")
        
        return result.stdout

    def fill_using_team_manager_mdb(self, mdb_path: str):
        # Extract members from the team manager database
        members_csv = self.__read_members_from_mdb(mdb_path)
        if members_csv == "":
            raise RuntimeError("Empty members csv")

        members_list = list(csv.reader(members_csv.split('\n')))
        if len(members_list) <= 1:
            raise ValueError("Members list is empty")

        self.__fill_club_from_members_list(members_list)

    def __str__(self):
        return f"{self.club_name} contains the following members:\n" + str(self.members)

