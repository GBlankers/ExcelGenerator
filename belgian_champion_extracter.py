'''
Import the results lenex, extract all the rankings. Apply filters and put the result
into an excel
'''


from settings import Settings
from lib.meet_management import LenexHelper, MeetResults
from lib.results_excel import ResultsExcel

def add_results_to_workbook(log, excel, filters):
    '''Load lenex, create results and add to excel'''
    lenex = LenexHelper(log, "C:/Users/brabo/Lenex_register-Excel-Generator/")
    lenex.load_lenex()
    lenex.extract_lef_from_lenex()
    lenex.load_xml_from_lef()

    meet_results = MeetResults(log, lenex.xml_root)
    meet_results.construct_rankings(filters)
    meet_results.print_rankings()

    excel.add_results_to_excel(meet_results.results,
                               meet_results.results_relays,
                               meet_results.meet_name)

def main():
    '''Main'''
    # Create the settings for logging
    log = Settings.get_logger()

    # Create the filters
    basic_filters = ["ONLY_NAT=BEL", "ONLY_CLUB=BRABO", "ONLY_PODIUM"]
    basic_filters_finals = basic_filters.copy()
    basic_filters_finals.append("ONLY_FINALS")

    results_excel = ResultsExcel(log, "results")
    log.info("BK Open")
    add_results_to_workbook(log, results_excel, basic_filters_finals)
    # log.info("BK 25M")
    # add_results_to_workbook(log, results_excel, basic_filters_finals)
    # log.info("BK Cat 1")
    # add_results_to_workbook(log, results_excel, basic_filters)
    # log.info("BK Cat 2")
    # add_results_to_workbook(log, results_excel, basic_filters_finals)
    results_excel.close()

if __name__ == "__main__":
    main()
