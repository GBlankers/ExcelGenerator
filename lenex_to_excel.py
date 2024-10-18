import csv

from club_management import *
from settings import *
from meet_management import *

def main() -> None:
    # Load or create the settings
    settings = Settings.init_settings()

    # Create a club using the provided mdb
    club = Club(settings.club_name)
    club.fill_using_team_manager_mdb(settings.mdb_path)
    
    lenex = LenexHelper(settings.default_competition_path)
    lenex.load_lenex()
    lenex.extract_lef_from_lenex()
    lenex.load_xml_from_lef()

    meet = SwimMeet()
    meet.load_from_xml(lenex.xml_root)

    print(meet)


if __name__ == "__main__":
    main()