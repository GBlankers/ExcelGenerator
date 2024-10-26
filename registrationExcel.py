import xlsxwriter
import os

from easygui import multchoicebox
from meet_management import SwimMeet, SwimMeetEvent
from club_management import Club, Swimmer, Gender

class RegistrationExcel():

    def __init__(self, meet_name: str):
        self.groups_to_use: list[str] = None
        self.__check_tmp_dir()
        self.__create_empty_excel(meet_name)
        self.__create_styles()

    def __check_tmp_dir(self):
        # check tmp folder present
        if not os.path.isdir("tmp"):
            os.mkdir("tmp")
    
    def __create_empty_excel(self, meet_name: str):
        '''Create empty base excel'''
        # String sanitizing
        name = meet_name.replace(' ', '-')
        self.file_path = f"tmp/inschrijving_{name}.xlsx"
        self.workbook = xlsxwriter.Workbook(self.file_path)
        print("Excel initialized")
    
    def __create_styles(self):
        self.styles = dict()
        self.styles["bold"] = self.workbook.add_format({"bold": True})
        self.styles["group_name"] = self.workbook.add_format({"bold": True,
                                                              'bg_color': 'gray',
                                                              'align': 'center',
                                                              'right': True})

    def __select_groups_to_use(self, club: Club):
        if self.groups_to_use == None:
            self.groups_to_use = multchoicebox("Select the groups to use", "Group selection", club.get_groups())
            self.groups_to_use.sort()
        print(f"Selected groups: {self.groups_to_use}")

    def __add_general_information(self, sheet, meet: SwimMeet) -> int:
        '''Add general meet information at the top of the sheet
           return the row index at which the sheet specific part can start'''
        
        row_number = 0
        col_number = 0

        # Set column sizes
        sheet.set_column(col_number, col_number, 30)
        sheet.set_column(col_number+1, col_number+60, 15)

        # Add club image
        sheet.set_row(row_number, 51)
        sheet.insert_image(row_number, col_number, "kazsc.png",
                           {'x_scale': 0.60, 'y_scale': 0.485})

        # First row is taken up by image and start info in second column
        row_number += 1
        col_number += 1
        
        # Meet name
        style = self.workbook.add_format({'bold': True})
        style.set_top()
        style.set_left()
        sheet.write(row_number, col_number, "Wedstrijd:", style)
        style = self.workbook.add_format()
        style.set_top()
        sheet.merge_range(row_number, col_number+1, row_number, col_number+3, meet.meet_name, style)
        col_number += 4

        # Deadline
        style = self.workbook.add_format({'bold': True})
        style.set_top()
        sheet.write(row_number, col_number, "Deadline:", style)
        style = self.workbook.add_format()
        style.set_top()
        style.set_right()
        sheet.merge_range(row_number, col_number+1, row_number, col_number+3, meet.deadline, style)

        # Next row
        row_number += 1
        col_number = 1

        # City
        style = self.workbook.add_format({'bold': True})
        style.set_bottom()
        style.set_left()
        sheet.write(row_number, col_number, "Zwembad:", style)
        style = self.workbook.add_format()
        style.set_bottom()
        sheet.merge_range(row_number, col_number+1, row_number, col_number+3, meet.city, style)
        col_number += 4

        # Qualify
        style = self.workbook.add_format({'bold': True})
        style.set_bottom()
        sheet.write(row_number, col_number, "Qualify:", style)
        style = self.workbook.add_format()
        style.set_bottom()
        style.set_right()
        sheet.merge_range(row_number, col_number+1, row_number, col_number+3, meet.qualify_date_range, style)

        return row_number+2

    def __add_overview_registration_sheet_structure(self, sheet, meet: SwimMeet, start_row: int) -> int:
        row_counter = start_row
        # Add headers
        sheet.write(row_counter, 0, "Event NR:", self.styles["bold"])
        sheet.write(row_counter+1, 0, "Gender:", self.styles["bold"])
        sheet.write(row_counter+2, 0, "Event:", self.styles["bold"])
        sheet.write(row_counter+3, 0, "Age:", self.styles["bold"])
        row_counter += 4

        sheet.set_row(row_counter, 2)
        sheet.freeze_panes(row_counter, 1)

        return row_counter + 1

    def __add_swimmers_overview_registration_sheet(self, sheet, club: Club, start_row: int):
        '''Add all the groups and swimmers to the first column of the sheet'''
        col_number = 0
        row_number = start_row

        # Keep track at which row number this swimmer is located
        self.swimmer_to_row_number = dict()

        for group in self.groups_to_use:
            sheet.write(row_number, col_number, group, self.styles["group_name"])
            row_number += 1

            for swimmer_name in club.get_swimmer_names_from_group(group):
                sheet.write(row_number, col_number, swimmer_name)
                self.swimmer_to_row_number[swimmer_name] = row_number
                row_number += 1

        print(','.join(self.groups_to_use) + " added to the register overview sheet")

        return row_number-1

    def __add_events_overview_registration_sheet(self, sheet, meet: SwimMeet, start_row_events: int, start_row_swimmers: int, end_row: int):
        col_number = 1

        session_cell_style = self.workbook.add_format({"rotation": -90, 
                                                       "bold": True, 
                                                       "bg_color": "gray", 
                                                       "align": "center",
                                                       "valign": "vcenter", 
                                                       "text_wrap": True})

        self.event_to_column_number = dict()
        for session in meet.program:
            session_info = meet.program[session]

            sheet.merge_range(start_row_swimmers, col_number, end_row, col_number,
                              f"{session} - Start: {session_info['session_start']} - Warmup: {session_info['session_warmup_start']}",
                              session_cell_style)
            
            col_number += 1
            for event in session_info["events"]:
                self.event_to_column_number[event] = col_number
                sheet.write(start_row_events, col_number, f"# {event.number}")
                sheet.write(start_row_events+1, col_number, event.gender)
                sheet.write(start_row_events+2, col_number, event.style)
                sheet.write(start_row_events+3, col_number, event.simplified_age)
                col_number += 1

    def __check_possible_event(self, swimmer: Swimmer, event: SwimMeetEvent, meet_age_date: str) -> bool:
        # Check gender
        if event.gender == "F" and swimmer.gender == '1':
            return False
        elif event.gender == "M" and swimmer.gender == '2':
            return False
        
        if event.min_age > swimmer.get_age_at(meet_age_date) or event.max_age < swimmer.get_age_at(meet_age_date):
            return False
    
        return True

    def __add_applicability_matrix_registration_sheet(self, sheet, meet: SwimMeet, club: Club):
        cross_cell_style = self.workbook.add_format({'diag_type': 3, 'diag_border': 1, 'diag_color': 'black', 'bg_color': 'gray'})

        for group in self.groups_to_use:
            for swimmer in club.get_swimmers_from_group(group):
                for event in meet.get_all_events():
                    if not self.__check_possible_event(swimmer, event, meet.age_date):
                        sheet.write(self.swimmer_to_row_number[swimmer.name], self.event_to_column_number[event],
                                    "", cross_cell_style)



    def create_overview_registration_sheet(self, meet: SwimMeet, club: Club):
        sheet = self.workbook.add_worksheet(name="Inschrijving")

        # Select which groups we want to include in the excel
        self.__select_groups_to_use(club)

        # Put the general information at the top of the excel
        start_row_events = self.__add_general_information(sheet, meet)

        # Create the general sheet structure
        start_row_swimmers = self.__add_overview_registration_sheet_structure(sheet, meet, start_row_events)

        # Add Swimmers as first column
        final_row = self.__add_swimmers_overview_registration_sheet(sheet, club, start_row_swimmers)

        # Add the events
        self.__add_events_overview_registration_sheet(sheet, meet, start_row_events, start_row_swimmers, final_row)

        # Check possible events -> blackout/cross cells if event cannot be done by swimmer
        self.__add_applicability_matrix_registration_sheet(sheet, meet, club)

    
    def close(self):
        self.workbook.close()
        print(f"Excel saved at {self.file_path}")
