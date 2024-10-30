'''
Contains classes to facilitate mdb handling and reading
the groups, athletes
'''
from dataclasses import dataclass
from datetime import datetime

import re
import os
import csv
import subprocess
import logging

@dataclass
class Swimmer:
    '''Dataclass with information about a swimmer'''
    name: str
    birth_data: str
    gender: str
    group: str

    def get_age_at(self, date_str: str) -> int:
        '''Get the age of a swimmer at a certain point in
           time. Expect date_str as dd/mm/yy'''

        birth_year_short = int(self.birth_data[-2:])
        meet_year = re.search(r"(\d\d\d\d)", date_str)

        if meet_year is None:
            meet_year_short = str(datetime.now().year)[-2]
        else:
            meet_year_short = int(meet_year.group()[-2:])

        if meet_year_short - birth_year_short < 0:
            return meet_year_short + (100-birth_year_short)

        return meet_year_short - birth_year_short

    def __str__(self):
        return self.name


class Club:
    '''Container with all the data of the club and its members'''
    ACTIVE = "T"

    def __init__(self, log: logging.Logger, name: str):
        self.members = dict()
        self.club_name = name
        self.log = log

    def __fill_club_from_members_list(self, members_list: list):
        # Get the indices of the different fields of the headers
        index_last_name = members_list[0].index("LASTNAME")
        index_first_name = members_list[0].index("FIRSTNAME")
        index_gender = members_list[0].index("GENDER")
        index_active = members_list[0].index("ACTIVE")
        index_groups = members_list[0].index("GROUPS")
        index_birth_date = members_list[0].index("BIRTHDATE")

        # Remove the first entry as this contains the headers
        members_list.pop(0)

        for athlete in members_list:
            # Possibility that there are empty elements
            if not athlete:
                continue

            # Do not use inactive members
            if athlete[index_active] != self.ACTIVE:
                continue

            # Iterate over all the group of the athlete
            for group in athlete[index_groups].split(", "):
                # Create empty list if unkown group
                if group not in self.members:
                    self.members[group] = []

                # Append the athlete to the correct group
                athlete_name = f"{athlete[index_first_name]} {athlete[index_last_name]}"
                self.members[group].append(Swimmer(athlete_name,
                                                   athlete[index_birth_date].split(' ')[0],
                                                   athlete[index_gender],
                                                   group))

    def __read_members_from_mdb(self, mdb_path: str) -> str:
        if os.path.splitext(mdb_path)[1] != '.mdb':
            raise ValueError("Given path to database is not an mdb")

        if os.path.exists(os.path.splitext(mdb_path)[0] + '.ldb'):
            raise RuntimeError("Database is locked")

        try:
            result = subprocess.run(['mdb-export', mdb_path, 'MEMBERS'],
                                    capture_output=True,
                                    text=True,
                                    check=True)
        except subprocess.CalledProcessError as e:
            raise RuntimeError(f"Error extracting data from database: {e}") from e

        return result.stdout

    def fill_using_team_manager_mdb(self, mdb_path: str):
        '''Fill the club class using the given mdb'''
        # Extract members from the team manager database
        members_csv = self.__read_members_from_mdb(mdb_path)
        if members_csv == "":
            raise RuntimeError("Empty members csv")

        members_list = list(csv.reader(members_csv.split('\n')))
        if len(members_list) <= 1:
            raise ValueError("Members list is empty")

        self.__fill_club_from_members_list(members_list)

    def get_groups(self) -> list[str]:
        '''Get all the group names in the club'''
        return sorted(list(self.members.keys()))

    def get_swimmers_from_group(self, group_name: str) -> list[Swimmer]:
        '''Get all the swimmers from a given group'''
        return self.members[group_name]

    def get_swimmer_names_from_group(self, group_name: str) -> list[str]:
        '''Sort the list on last name for more easy entry in teammanager'''
        sorted_2d_list = sorted([str(swimmer_name).split(' ', 1) for \
                                 swimmer_name in self.members[group_name]], key=lambda x: x[-1])
        return [' '.join(swimmer) for swimmer in sorted_2d_list]

    def __str__(self):
        return_str = f"{self.club_name}:\n"
        for group, athletes in self.members.items():
            return_str += f"{group}: "
            for athlete in athletes:
                return_str += f"{str(athlete)};"
            return_str += "\n"

        return return_str
