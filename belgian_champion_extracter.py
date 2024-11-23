from settings import Settings
from meet_management import LenexHelper, MeetResults
from results_excel import ResultsExcel

def add_results_to_workbook(log, excel, filters):
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
    # Create the settings for logging
    log = Settings.get_logger()

    results_excel = ResultsExcel(log, "results")
    log.info("BK Open")
    add_results_to_workbook(log, results_excel, ["ONLY_NAT=BEL", "ONLY_CLUB=BRABO", "ONLY_FINALS", "ONLY_PODIUM"])
    # log.info("BK 25M")
    # add_results_to_workbook(log, results_excel, ["ONLY_NAT=BEL", "ONLY_CLUB=BRABO", "ONLY_FINALS", "ONLY_PODIUM"])
    # log.info("BK Cat 1")
    # add_results_to_workbook(log, results_excel, ["ONLY_NAT=BEL", "ONLY_CLUB=BRABO", "ONLY_PODIUM"])
    # log.info("BK Cat 2")
    # add_results_to_workbook(log, results_excel, ["ONLY_NAT=BEL", "ONLY_CLUB=BRABO", "ONLY_FINALS", "ONLY_PODIUM"])
    results_excel.close()
    

if __name__ == "__main__":
    main()

