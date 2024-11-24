'''Classes for creating registration excels'''

import os
import logging

import xlsxwriter
import xlsxwriter.utility

from easygui import multchoicebox
from lib.meet_management import SwimMeet
from lib.club_management import Club
from lib.possible_events import PossibleEvents

class _Sheet:
    '''Base class for excel sheets with common methods and variables'''
    def __init__(self, workbook: xlsxwriter.Workbook, name: str, groups_to_use: list[str],
                 log: logging.Logger, club_logo_path: str):
        self.log = log
        self.name = name
        self.workbook = workbook
        self.sheet = self.workbook.add_worksheet(name=self.name)
        self.groups_to_use = groups_to_use
        self.club_logo_path = club_logo_path

        self.__create_styles()

    def __create_styles(self):
        '''Some generic styles. Not having to reinitialize every time we need to use these'''
        self.styles = {}
        self.styles["bold"] = self.workbook.add_format({"bold": True})
        self.styles["group_name"] = self.workbook.add_format({"bold": True, 'bg_color': 'gray',
                                                              'align': 'center', 'right': True})

    def add_general_information(self, meet: SwimMeet) -> int:
        '''Add general meet information at the top of the sheet
           return the row index at which the sheet specific part can start'''

        row_number = 0
        col_number = 0

        # Set column sizes
        self.sheet.set_column(col_number, col_number, 30)
        self.sheet.set_column(col_number+1, col_number+150, 15)

        # Add club image
        self.sheet.set_row(row_number, 51)
        self.sheet.insert_image(row_number, col_number, self.club_logo_path,
                           {'x_scale': 0.60, 'y_scale': 0.485})

        # First row is taken up by image and start info in second column
        row_number += 1
        col_number += 1

        # Meet name
        style = self.workbook.add_format({'bold': True})
        style.set_top()
        style.set_left()
        self.sheet.write(row_number, col_number, "Wedstrijd:", style)
        style = self.workbook.add_format()
        style.set_top()
        self.sheet.merge_range(row_number, col_number+1, row_number, col_number+3, meet.meet_name,
                               style)
        col_number += 4

        # Deadline
        style = self.workbook.add_format({'bold': True})
        style.set_top()
        self.sheet.write(row_number, col_number, "Deadline:", style)
        style = self.workbook.add_format()
        style.set_top()
        style.set_right()
        self.sheet.merge_range(row_number, col_number+1, row_number, col_number+3, meet.deadline,
                               style)

        # Next row
        row_number += 1
        col_number = 1

        # City
        style = self.workbook.add_format({'bold': True})
        style.set_bottom()
        style.set_left()
        self.sheet.write(row_number, col_number, "Zwembad:", style)
        style = self.workbook.add_format()
        style.set_bottom()
        self.sheet.merge_range(row_number, col_number+1, row_number, col_number+3, meet.city, style)
        col_number += 4

        # Qualify
        style = self.workbook.add_format({'bold': True})
        style.set_bottom()
        self.sheet.write(row_number, col_number, "Qualify:", style)
        style = self.workbook.add_format()
        style.set_bottom()
        style.set_right()
        self.sheet.merge_range(row_number, col_number+1, row_number, col_number+3,
                               meet.qualify_date_range, style)

        return row_number+2

class OverviewRegistrationSheet(_Sheet):
    '''Sheet containing an overview of the shedule with all the swimmers'''

    def __init__(self, workbook: xlsxwriter.Workbook, name: str, groups_to_use: list[str],
                 log: logging.Logger, club_logo_path: str):
        super().__init__(workbook, name, groups_to_use, log, club_logo_path)

        # Keep track at which location certain elements are placed
        self.event_row_nr: int = -1
        self.event_name_row_nr: int = -1
        self.final_column: int = -1
        self.swimmer_to_row_number: dict = {}
        self.event_to_column_number: dict = {}

    def __add_structure(self, start_row: int) -> int:
        row_counter = start_row
        # Add headers
        self.sheet.write(row_counter, 0, "Event NR:", self.styles["bold"])
        self.event_row_nr = row_counter
        self.sheet.write(row_counter+1, 0, "Gender:", self.styles["bold"])
        self.sheet.write(row_counter+2, 0, "Event:", self.styles["bold"])
        self.event_name_row_nr = row_counter + 2
        self.sheet.write(row_counter+3, 0, "Age:", self.styles["bold"])
        row_counter += 4

        self.sheet.set_row(row_counter, 2)
        self.sheet.freeze_panes(row_counter, 1)

        return row_counter + 1

    def __add_swimmers(self, club: Club, start_row: int):
        '''Add all the groups and swimmers to the first column of the sheet'''
        col_number = 0
        row_number = start_row

        for group in self.groups_to_use:
            self.sheet.write(row_number, col_number, group, self.styles["group_name"])
            row_number += 1

            for swimmer_name in club.get_swimmer_names_from_group(group):
                self.sheet.write(row_number, col_number, swimmer_name)
                self.swimmer_to_row_number[swimmer_name] = row_number
                row_number += 1

        self.log.info(','.join(self.groups_to_use) + " added to the register overview sheet")

        return row_number-1

    def __add_events(self, meet: SwimMeet, start_row_events: int, start_row_swimmers: int,
                     end_row: int):
        col_number = 1

        session_cell_format: dict = {'rotation': -90, 'bold': True, 'bg_color': "gray",
                                     'align': "center", 'valign': "vcenter", 'text_wrap': True}
        session_cell_style = self.workbook.add_format(session_cell_format)

        for session in meet.program:
            session_info = meet.program[session]

            self.sheet.merge_range(start_row_swimmers, col_number, end_row, col_number,
                                   f"{session} - Start: {session_info['session_start']} - \
                                   Warmup: {session_info['session_warmup_start']}",
                                   session_cell_style)

            col_number += 1
            for event in meet.get_events_in_session(session_name=session):
                if event.round == "FIN":
                    continue
                self.event_to_column_number[event] = col_number
                self.sheet.write(start_row_events, col_number, f"{event.round} # {event.number}")
                self.sheet.write(start_row_events+1, col_number, event.gender)
                self.sheet.write(start_row_events+2, col_number, event.style)
                self.sheet.write(start_row_events+3, col_number, event.simplified_age)
                col_number += 1

        self.final_column = col_number

    def fill_sheet(self, meet: SwimMeet, club: Club):
        '''Fill in the overview registration sheet. I.e. add the
           general information, swimmers, events and cross out
           the events that are invalid'''
        # Put the general information at the top of the excel
        start_row_events = super().add_general_information(meet)

        # Create the general sheet structure
        start_row_swimmers = self.__add_structure(start_row_events)

        # Add Swimmers as first column
        final_row = self.__add_swimmers(club, start_row_swimmers)

        # Add the events
        self.__add_events(meet, start_row_events, start_row_swimmers, final_row)

    def cross_invalid_events(self, possible_events: PossibleEvents, club: Club):
        '''In the event/swimmer matrix, cross out the events that the swimmer may 
           not participate in'''
        cross_cell_style = self.workbook.add_format({'diag_type': 3, 'diag_border': 1,
                                                     'diag_color': 'black', 'bg_color': 'gray'})
        for group in self.groups_to_use:
            for swimmer_name in club.get_swimmer_names_from_group(group):
                for invalid_event in possible_events.get_invalid_events_for_swimmer(swimmer_name):
                    # Finals will not be included in the overview
                    if invalid_event.round == "FIN":
                        continue
                    self.sheet.write(self.swimmer_to_row_number[swimmer_name],
                                     self.event_to_column_number[invalid_event],
                                     "", cross_cell_style)

class SummarySheet(_Sheet):
    '''Registration excel sheet to give an overview of the different
       events that are selected for every swimmer'''

    def __init__(self, workbook: xlsxwriter.Workbook, name: str, groups_to_use: list[str],
                 log: logging.Logger, club_logo_path: str):
        super().__init__(workbook, name, groups_to_use, log, club_logo_path)

        # Keep track of different row/column numbers
        self.swimmer_to_row_number: dict = {}

    def __add_swimmers(self, club: Club, start_row: int):
        '''Add all the groups and swimmers to the first column of the sheet'''
        col_number = 0
        row_number = start_row

        header_style = self.workbook.add_format({'bold': True, 'align': 'center'})
        header_style.set_bottom()

        self.sheet.write(row_number, col_number, "Swimmers", header_style)
        self.sheet.write(row_number, col_number+1, "Selected events", header_style)
        self.sheet.write(row_number, col_number+2, "Selected event numbers", header_style)
        self.sheet.set_column(col_number+1, col_number+2, 45)

        row_number += 1

        for group in self.groups_to_use:
            self.sheet.write(row_number, col_number, group, self.styles["group_name"])
            row_number += 1

            for swimmer_name in club.get_swimmer_names_from_group(group):
                self.sheet.write(row_number, col_number, swimmer_name)
                self.swimmer_to_row_number[swimmer_name] = row_number
                row_number += 1

        self.log.info(','.join(self.groups_to_use) + " added to the summary sheet")

    def __add_summary(self, club: Club, register_sheet: OverviewRegistrationSheet):
        final_column_letter = xlsxwriter.utility.xl_col_to_name(register_sheet.final_column-1)

        for group in self.groups_to_use:
            for swimmer_name in club.get_swimmer_names_from_group(group):
                local_row = self.swimmer_to_row_number[swimmer_name]
                register_row = register_sheet.swimmer_to_row_number[swimmer_name] + 1
                form = f'=_xlfn.TEXTJOIN(", ", TRUE, IF(ISBLANK(Inschrijving!C{register_row}:{final_column_letter + str(register_row)}), "", {register_sheet.name}!C{register_sheet.event_name_row_nr+1}:{final_column_letter}{register_sheet.event_name_row_nr+1}))'
                self.log.debug(form)
                self.sheet.write_array_formula(local_row, 1, local_row, 1, form)

    def fill_sheet(self, meet: SwimMeet, club: Club, register_sheet: OverviewRegistrationSheet):
        '''Fill in the summary sheet. I.e. add the general information, swimmers
           and summary'''
        # Add the standard information at the top of the sheet
        start_row = super().add_general_information(meet)

        # Add all the swimmers in the first column
        self.__add_swimmers(club, start_row)

        # Add the excel formula to create a summary
        self.__add_summary(club, register_sheet)

class ValidEventsSheet(_Sheet):
    '''Registration excel sheet to give an overview of the different
       events that are possible for every swimmer'''

    def __add_swimmers_and_events(self, club: Club, possible_events: PossibleEvents, s_row: int):
        col_number = 0
        row_number = s_row

        header_style = self.workbook.add_format({'bold': True, 'align': 'center'})
        header_style.set_bottom()

        self.sheet.write(row_number, col_number, "Swimmers", header_style)
        self.sheet.write(row_number, col_number+1, "Possible events", header_style)
        self.sheet.write(row_number, col_number+2, "Selected events", header_style)

        self.sheet.set_column(col_number+1, col_number+2, 30)

        row_number += 1

        for group in self.groups_to_use:
            self.sheet.merge_range(row_number, col_number, row_number, col_number+2,
                                   group, self.styles["group_name"])
            row_number += 1

            for swimmer_name in club.get_swimmer_names_from_group(group):
                self.sheet.write(row_number, col_number, swimmer_name)
                events = possible_events.get_valid_events_for_swimmer(swimmer_name)

                if len(events) == 0:
                    self.sheet.write(row_number, col_number+1,
                                     "No events possible for this swimmer")
                    continue

                for event in events:
                    self.sheet.write(row_number, col_number+1,
                                     f"#{event.number} {event.style} \
                                     {event.simplified_age}")
                    row_number += 1

                row_number += 1


    def fill_sheet(self, meet: SwimMeet, club: Club, possible_events: PossibleEvents):
        '''Fill in the valid events sheet. I.e. For every swimmer,
           give an overview of all the different events he/she can
           pariticipate in'''

        # Add the general information
        start_row = super().add_general_information(meet)

        # Start adding the swimmers and events
        self.__add_swimmers_and_events(club, possible_events, start_row)


class RegistrationExcel:
    '''Class to group all the data concering the registration excel'''
    def __init__(self, log: logging.Logger, meet_name: str, club_logo_path: str):
        self.groups_to_use: list[str] = None
        self.log = log
        self.sheets: list[_Sheet] = []

        self.possible_events: PossibleEvents = None
        self.groups_to_use: list[str] = None
        self.club_logo_path: str = club_logo_path

        self.__check_tmp_dir()
        self.__create_empty_excel(meet_name)

    def __check_tmp_dir(self):
        # check tmp folder present
        if not os.path.isdir("tmp"):
            os.mkdir("tmp")
            self.log.debug("Created tmp folder")

    def __create_empty_excel(self, meet_name: str):
        '''Create empty base excel'''
        # String sanitizing
        name = meet_name.replace(' ', '-')
        self.file_path = f"tmp/inschrijving_{name}.xlsx"
        self.workbook = xlsxwriter.Workbook(self.file_path)
        self.log.debug("Excel initialized")

    def __get_groups_to_use(self, club: Club):
        if self.groups_to_use is None:
            self.groups_to_use = multchoicebox("Select the groups to use",
                                               "Group selection", club.get_groups())
            self.groups_to_use.sort()
            self.log.info(f"Selected groups: {self.groups_to_use}")

        return self.groups_to_use

    def __get_possible_events(self, meet: SwimMeet, club: Club) -> PossibleEvents:
        if self.possible_events is not None:
            return self.possible_events

        self.possible_events = PossibleEvents(meet, club)
        self.possible_events.generate_possible_events_dict(self.__get_groups_to_use(club))

        return self.possible_events

    def add_overview_registration_sheet(self, meet: SwimMeet, club: Club):
        '''Add a sheet with an overview of all the events, sessions and swimmers on
           which a selection of the different event for the swimmer
           can be made'''
        # Select which groups we want to include in the excel
        groups = self.__get_groups_to_use(club)

        # Get all the possible and invalid events for each swimmer
        possible_events = self.__get_possible_events(meet, club)

        ors = OverviewRegistrationSheet(self.workbook, "Inschrijving", groups, self.log,
                                        self.club_logo_path)
        ors.fill_sheet(meet, club)
        ors.cross_invalid_events(possible_events, club)
        self.sheets.append(ors)

    def add_summary_sheet(self, meet: SwimMeet, club: Club):
        '''Add a sheet with a summary of the selected events for every
           swimmer'''
        # Get the groups we want to include
        groups = self.__get_groups_to_use(club)

        # Check if there is a overview registration sheet
        ors: OverviewRegistrationSheet = None
        for sheet in self.sheets:
            if isinstance(sheet, OverviewRegistrationSheet):
                ors = sheet

        if ors is None:
            raise ValueError("Cannot create a summary if there is not \
                             an overview registration sheet")

        sum_s = SummarySheet(self.workbook, "Summary", groups, self.log, self.club_logo_path)
        sum_s.fill_sheet(meet, club, ors)
        self.sheets.append(sum_s)

    def add_valid_events_sheet(self, meet: SwimMeet, club: Club):
        '''Add a sheet with all the valid events per swimmers'''

        # Get the groups we want to include
        groups = self.__get_groups_to_use(club)

        # Get the possible events for every swimmer
        pos_events = self.__get_possible_events(meet, club)

        valid_events_sheet = ValidEventsSheet(self.workbook, "Individueel", groups, self.log,
                                              self.club_logo_path)
        valid_events_sheet.fill_sheet(meet, club, pos_events)
        self.sheets.append(valid_events_sheet)

    def close(self):
        '''Close and save the registration excel'''
        self.workbook.close()
        print(f"Excel saved at {self.file_path}")
