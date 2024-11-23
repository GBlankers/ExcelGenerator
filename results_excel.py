import xlsxwriter
import os
import logging

import xlsxwriter.exceptions

from meet_management import RankingsEntry


class ResultsExcel:
    IND_RESULTS_COL=1
    RELAY_RESULTS_COL=5
    
    def __init__(self, log: logging.Logger, excel_name: str) -> None:
        self.log = log

        self.file_path: str = ""
        self.workbook: xlsxwriter.Workbook = None

        self.__create_empty_excel(excel_name)

    def __check_tmp_dir(self):
        # check tmp folder present
        if not os.path.isdir("tmp"):
            os.mkdir("tmp")
            self.log.debug("Created tmp folder")
    
    def __create_empty_excel(self, excel_name: str):
        '''Create empty base excel'''
        # Check if the tmp folder is present
        self.__check_tmp_dir()
        # String sanitizing
        self.file_path = f"tmp/{excel_name}.xlsx"
        # Create the excel
        self.workbook = xlsxwriter.Workbook(self.file_path)
        self.log.debug("Excel initialized")

    def __structure_sheet(self, sheet) -> None:
        # 0th column and row very small for cleanness
        sheet.set_column(0, 0, 3)
        sheet.set_row(0, 3)
        
        # First column for place
        sheet.set_column(self.IND_RESULTS_COL, self.IND_RESULTS_COL, 3)
        # Second column for name
        sheet.set_column(self.IND_RESULTS_COL+1, self.IND_RESULTS_COL+1, 35)
        # Third column for time
        sheet.set_column(self.IND_RESULTS_COL+2, self.IND_RESULTS_COL+2, 13)
        # Fourth column is open
        # Fifth column for place relay
        sheet.set_column(self.RELAY_RESULTS_COL, self.RELAY_RESULTS_COL, 3)
        # Sixth column for relay names
        sheet.set_column(self.RELAY_RESULTS_COL+1, self.RELAY_RESULTS_COL+1, 75)
        # Seventh column for time
        sheet.set_column(self.RELAY_RESULTS_COL+2, self.RELAY_RESULTS_COL+2, 13)

        # First row for meet name
        sheet.set_row(1, 30)

    def __add_header_to_sheet(self, sheet, meet_name: str) -> int:
        bold_underline_s = self.workbook.add_format({"bold": True,
                                                     "valign": "vcenter",
                                                     "align": "center"})
        bold_s = bold_underline_s
        bold_underline_s.set_bottom()
        
        sheet.merge_range(1, 1, 1, 7, f"Results for {meet_name}", bold_underline_s)
        sheet.merge_range(2, 1, 2, 3, "Individual", bold_s)
        sheet.merge_range(2, 5, 2, 7, "Relays", bold_s)
        
        # Leave line open for first ranking
        return 4
        
    def __add_result_to_sheet(self, sheet, event_name: str, 
                              event_ranking: list[RankingsEntry],
                              row_number: int, col_number: int) -> int:
        # Write the event name
        bold_s = self.workbook.add_format({'bold': True})
        sheet.merge_range(row_number, col_number, row_number, col_number+2,
                          event_name, bold_s) 

        gold_bg_s = self.workbook.add_format({"bg_color": "#FDDC5C"})
        silver_bg_s = self.workbook.add_format({"bg_color": "#D7D7D7"})
        broze_bg_s = self.workbook.add_format({"bg_color": "#a77044"})
        podium_s = [gold_bg_s, silver_bg_s, broze_bg_s]
        normal_s = self.workbook.add_format()

        for ranking_entry in event_ranking:
            if ranking_entry.placing < 4:
                style = podium_s[int(ranking_entry.placing)-1]
            else:
                style = normal_s
            row_number += 1
            sheet.write(row_number, col_number, f"{ranking_entry.placing}.", style)
            sheet.write(row_number, col_number+1, f"{ranking_entry.swimmer_name}", style)
            sheet.write(row_number, col_number+2, f"{ranking_entry.swim_time}", style)
        
        return row_number + 2
            
    def add_results_to_excel(self, rankings_individual: dict[str, list[RankingsEntry]],
                             rankings_relay: dict[str, list[RankingsEntry]],
                             meet_name: str) -> None:
        # Create empty sheet
        try:
            sheet = self.workbook.add_worksheet(name=meet_name)
        except xlsxwriter.exceptions.DuplicateWorksheetName:
            sheet = self.workbook.add_worksheet(name=f"{meet_name}_V2")
        except xlsxwriter.exceptions.InvalidWorksheetName:
            sheet = self.workbook.add_worksheet()

        # Set column/row sizes
        self.__structure_sheet(sheet)

        # Keep track of the row to place the entries below each other
        row_number_begin = self.__add_header_to_sheet(sheet, meet_name)

        # Iterate over all the individual results and write to excel
        row_number = row_number_begin
        for event_name, event_ranking in rankings_individual.items():
            row_number = self.__add_result_to_sheet(sheet, event_name,
                                                    event_ranking, row_number,
                                                    self.IND_RESULTS_COL)
        
        # Iterate over the relay results and write to excel
        row_number = row_number_begin
        for event_name, event_ranking in rankings_relay.items():
            row_number = self.__add_result_to_sheet(sheet, event_name,
                                                    event_ranking, row_number,
                                                    self.RELAY_RESULTS_COL)

    def close(self) -> None:
        self.workbook.close()
        print(f"Excel saved at {self.file_path}")        
