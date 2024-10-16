import csv

from club_management import *



def main() -> None:
    # Extract members from the team manager database
    members_csv = read_members_from_mdb("KAZSC.mdb")
    if members_csv == "":
        raise RuntimeError("Empty members csv")

    members_list = list(csv.reader(members_csv.split('\n')))
    if len(members_list) <= 1:
        raise ValueError("Members list is empty")

    club = Club(members_list)
    
    print(club)


if __name__ == "__main__":
    main()