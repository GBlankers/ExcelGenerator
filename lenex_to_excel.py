import csv

from club_management import *

def main() -> None:
    club = Club("KAZSC")
    club.fill_using_team_manager_mdb("KAZSC.mdb")
    
    print(club)


if __name__ == "__main__":
    main()