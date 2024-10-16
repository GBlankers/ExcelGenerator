import csv

from club_management import *
from settings import *

def main() -> None:
    # Load or create the settings
    settings = Settings.init_settings()

    # Create a club using the provided mdb
    club = Club(settings.club_name)
    club.fill_using_team_manager_mdb(settings.mdb_path)
    
    print(club)


if __name__ == "__main__":
    main()