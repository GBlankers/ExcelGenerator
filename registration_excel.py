import xlsxwriter
import os
import logging

from easygui import multchoicebox
import xlsxwriter.utility
from meet_management import SwimMeet, SwimMeetEvent
from club_management import Club, Swimmer
from possible_events import PossibleEvents

class _Sheet:
    def __init__(self, workbook: xlsxwriter.Workbook, name: str, groups_to_use: list[str], log: logging.Logger):
        self.log = log
        self.name = name
        self.workbook = workbook
        self.sheet = self.workbook.add_worksheet(name=self.name)
        self.groups_to_use = groups_to_use

        self.__create_styles()

    def __create_styles(self):
        '''Some generic styles. Not having to reinitialize every time we need to use these'''
        self.styles = dict()
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
        self.sheet.insert_image(row_number, col_number, "kazsc.png",
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
        self.sheet.merge_range(row_number, col_number+1, row_number, col_number+3, meet.meet_name, style)
        col_number += 4

        # Deadline
        style = self.workbook.add_format({'bold': True})
        style.set_top()
        self.sheet.write(row_number, col_number, "Deadline:", style)
        style = self.workbook.add_format()
        style.set_top()
        style.set_right()
        self.sheet.merge_range(row_number, col_number+1, row_number, col_number+3, meet.deadline, style)

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
        self.sheet.merge_range(row_number, col_number+1, row_number, col_number+3, meet.qualify_date_range, style)

        return row_number+2

    def fill_sheet(self, meet: SwimMeet, club: Club):
        raise NotImplementedError("Fill sheet called on base class")

class OverviewRegistrationSheet(_Sheet):
    def __init__(self, workbook: xlsxwriter.Workbook, name: str, groups_to_use: list[str], log: logging.Logger):
        super().__init__(workbook, name, groups_to_use, log)

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

        # Keep track at which row number this swimmer is located
        self.swimmer_to_row_number = dict()

        for group in self.groups_to_use:
            self.sheet.write(row_number, col_number, group, self.styles["group_name"])
            row_number += 1

            for swimmer_name in club.get_swimmer_names_from_group(group):
                self.sheet.write(row_number, col_number, swimmer_name)
                self.swimmer_to_row_number[swimmer_name] = row_number
                row_number += 1

        self.log.info(','.join(self.groups_to_use) + " added to the register overview sheet")

        return row_number-1

    def __add_events(self, meet: SwimMeet, start_row_events: int, start_row_swimmers: int, end_row: int):
        col_number = 1

        session_cell_format = dict(rotation = -90, bold = True, bg_color = "gray",
                                   align = "center", valign = "vcenter", text_wrap = True)
        session_cell_style = self.workbook.add_format(session_cell_format)

        self.event_to_column_number = dict()
        for session in meet.program:
            session_info = meet.program[session]

            self.sheet.merge_range(start_row_swimmers, col_number, end_row, col_number,
                             f"{session} - Start: {session_info['session_start']} - Warmup: {session_info['session_warmup_start']}",
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
        # Put the general information at the top of the excel
        start_row_events = super().add_general_information(meet)

        # Create the general sheet structure
        start_row_swimmers = self.__add_structure(start_row_events)

        # Add Swimmers as first column
        final_row = self.__add_swimmers(club, start_row_swimmers)

        # Add the events
        self.__add_events(meet, start_row_events, start_row_swimmers, final_row)

    def cross_invalid_events(self, possible_events: PossibleEvents, club: Club):
        cross_cell_style = self.workbook.add_format({'diag_type': 3, 'diag_border': 1, 'diag_color': 'black', 'bg_color': 'gray'})
        for group in self.groups_to_use:
            for swimmer_name in club.get_swimmer_names_from_group(group):
                for invalid_event in possible_events.get_invalid_events_for_swimmer(swimmer_name):
                    # Finals will not be included in the overview
                    if invalid_event.round == "FIN":
                        continue
                    self.sheet.write(self.swimmer_to_row_number[swimmer_name], self.event_to_column_number[invalid_event],
                                     "", cross_cell_style)

class SummarySheet(_Sheet):
    def __init__(self, workbook: xlsxwriter.Workbook, name: str, groups_to_use: list[str], log: logging.Logger):
        super().__init__(workbook, name, groups_to_use, log)

    def __add_swimmers(self, club: Club, start_row: int):
        '''Add all the groups and swimmers to the first column of the sheet'''
        col_number = 0
        row_number = start_row

        # Keep track at which row number this swimmer is located
        self.swimmer_to_row_number = dict()

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
                form = (f'=_xlfn.TEXTJOIN(", ", TRUE, IF(ISBLANK(Inschrijving!C{register_row}:'
                        f'{final_column_letter + str(register_row)}), "", {register_sheet.name}'
                        f'!C{register_sheet.event_name_row_nr+1}:{final_column_letter}{register_sheet.event_name_row_nr+1}))')
                self.sheet.write_array_formula(local_row, 1, local_row, 1, form)

    def fill_sheet(self, meet: SwimMeet, club: Club, register_sheet: OverviewRegistrationSheet):
        # Add the standard information at the top of the sheet
        start_row = super().add_general_information(meet)

        # Add all the swimmers in the first column
        self.__add_swimmers(club, start_row)

        # Add the excel formula to create a summary
        self.__add_summary(club, register_sheet)

class RegistrationExcel:
    def __init__(self, log: logging.Logger, meet_name: str):
        self.groups_to_use: list[str] = None
        self.log = log
        self.sheets: list[_Sheet] = [] 

        self.possible_events: PossibleEvents = None

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
        if self.groups_to_use == None:
            self.groups_to_use = multchoicebox("Select the groups to use", "Group selection", club.get_groups())
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
        # Select which groups we want to include in the excel
        groups = self.__get_groups_to_use(club)

        # Get all the possible and invalid events for each swimmer
        possible_events = self.__get_possible_events(meet, club)

        ors = OverviewRegistrationSheet(self.workbook, "Inschrijving", groups, self.log)
        ors.fill_sheet(meet, club)
        ors.cross_invalid_events(possible_events, club)
        self.sheets.append(ors)

    def add_summary_sheet(self, meet: SwimMeet, club: Club):
        # Get the groups we want to include
        groups = self.__get_groups_to_use(club)

        # Check if there is a overview registration sheet
        ors: OverviewRegistrationSheet = None
        for sheet in self.sheets:
            if type(sheet) == OverviewRegistrationSheet:
                ors = sheet 

        if ors is None:
            raise ValueError("Cannot create a summary if there is not an overview registration sheet")

        sum_s = SummarySheet(self.workbook, "Summary", groups, self.log)
        sum_s.fill_sheet(meet, club, ors)
        self.sheets.append(sum_s)
    
    def close(self):
        self.workbook.close()
        print(f"Excel saved at {self.file_path}")
