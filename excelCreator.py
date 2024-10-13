# TODO: Black out the cells if a time was needed and the swimmer didn't swim it in the required period

import os
import easygui
import shutil
import xml.etree.ElementTree as ET
import xlsxwriter
import json
import datetime
import re

from zipfile import ZipFile
from dataclasses import dataclass
from variables import *

today = datetime.date.today().strftime("%d/%m/%Y")

@dataclass
class Swimmer:
    name: str
    age: int
    gender: str
    blockedNumbers = []

class SwimMeet:
    """Container to save the information of a meet"""

    def __init__(self) -> None:
        # Root Nodes of Lef
        self.sessionRoot = None
        self.meetRoot = None

        # Meet Information
        self.name = None
        self.ageDate = None
        self.city = None
        self.course = None
        self.qualifyFrom = None
        self.qualify = ""
        self.qualifyUntil = None
        self.deadline = None
        self.program = {}

    def __simplifyAge(self, age: str) -> str:
        if not age:
            return 0, 99, "open"

        ageString = ("/".join(age.split("-"))).replace("+", "").replace("open", "")
        ageArray = [int(x) for x in ageString.split("/") if x != ""]
        ageMax = max(ageArray)
        ageMin = min(ageArray)

        if "+" in age or "open" in age:
            simplifiedAgeString = str(ageMin) + " -> open"
            ageMax = 99
        else:
            simplifiedAgeString = str(ageMin) + " -> " + str(ageMax)
        
        return ageMin, ageMax, simplifiedAgeString

    def __extractEventInformation(self, eventNode) -> dict:
        # Gender
        gender = eventNode.attrib.get("gender", "X")
        if gender == "X":
            gender = "Mixed"
        
        # Event number
        number = eventNode.attrib.get("number", "?")

        style = ""
        age = ""
        for eventInfo in eventNode:
            if eventInfo.tag == "SWIMSTYLE":
                if int(eventInfo.attrib["relaycount"]) > 1:
                    style = eventInfo.attrib["relaycount"] + "x" + eventInfo.attrib["distance"] + " " + eventInfo.attrib["stroke"]
                else:
                    style = eventInfo.attrib["distance"] + " " + eventInfo.attrib["stroke"]
            elif eventInfo.tag == "AGEGROUPS":
                for ageGroup in eventInfo:
                    minAge = ageGroup.attrib["agemin"]
                    maxAge = ageGroup.attrib["agemax"]
                    if minAge == "-1" and maxAge == "-1":
                        age = age + "/open"
                    elif maxAge == "-1":
                        age = age + "/" + minAge + "+"
                    else:
                        if minAge == maxAge:
                            age = age + "/" + minAge
                        else:
                            age = age + "/" + minAge + "-" + maxAge

        minAge, maxAge, simplifiedAge = self.__simplifyAge(age)

        print(f"{simplifiedAge}\t{number}\t{gender}\t{style} extracted")

        return dict(simplifiedAge = simplifiedAge,
                    minAge = minAge,
                    maxAge = maxAge,
                    number = number,
                    age = age,
                    gender = gender,
                    style = style)

    def __parseProgram(self) -> int:        
        for session in self.sessionRoot:
            sessionAttr = session.attrib

            s = dict(sessionNumber = sessionAttr.get("number", "1"), 
                     sessionDate = sessionAttr.get("date", "?"),
                     sessionStart = sessionAttr.get("daytime", "?"),
                     sessionEnd = sessionAttr.get("endtime", "?"),
                     sessionWarmupStart = sessionAttr.get("warmupfrom", "?"),
                     sessionWarmupUntil = sessionAttr.get("warmupuntil", "?"),
                     sessionOfficialMeeting = sessionAttr.get("officialmeeting", "?"),
                     events = [])
            
            # Find the events in the session node
            eventId = 0
            for i, j in enumerate(session):
                if j.tag == "EVENTS":
                    eventId = i
                    break
            
            # Get all the events in the session
            for event in session[eventId]:
                s["events"].append(self.__extractEventInformation(event))

            # Add this session to the program
            self.program["Session " + s["sessionNumber"]] = s

        return 0

    def createFromLef(self, root) -> int:
        for meets in root:
            if meets.tag == "MEETS":
                self.meetRoot = meets[0]

        if self.meetRoot == None:
            print("Swimmeet could not be created from the given lenex")
            return -1

        self.name = self.meetRoot.attrib.get('name', "?").replace("/", "-")
        self.city = self.meetRoot.attrib.get('city', "?")
        self.course = self.meetRoot.attrib.get('course', "LCM")

        for meetInfo in self.meetRoot:
            if meetInfo.tag == "QUALIFY":
                self.qualifyFrom = meetInfo.attrib.get("from", today)
                self.qualifyUntil = meetInfo.attrib.get("until", self.deadline)
                self.qualify = self.qualifyFrom + " -> " + self.qualifyUntil
            elif meetInfo.tag == "SESSIONS":
                self.sessionRoot = meetInfo
            elif meetInfo.tag == "AGEDATE":
                self.ageDate = meetInfo.attrib.get("value", today)

        return self.__parseProgram()
    
class Club:
    groups = ["LS", "S1", "S2", "S3"]

    def __init__(self) -> None:
        self.swimmers = {}
        self.amountOfSwimmers = 0
        self.groupsToUse = self.groups
    
    def loadFromFile(self, fileName: str, groupsToUse: list[int], ageDate: str) -> None:
        with open(fileName, 'r') as f:
            club = json.load(f)

        extractYear = re.search("(\d\d\d\d)", ageDate)
        if extractYear is None:
            currentYear = 2022
        else:
            currentYear = int(extractYear.group())

        self.groupsToUse = groupsToUse

        for group in groupsToUse:
            tempGroup = {}
            for swimmer in club[group]:
                tempGroup[swimmer] = Swimmer(swimmer, currentYear - int(club[group][swimmer]["yearOfBirth"]), club[group][swimmer]["gender"])
                self.amountOfSwimmers += 1
            self.amountOfSwimmers += 1
            self.swimmers[group] = tempGroup
    
    def checkPossibleEvents(self, swimMeet: SwimMeet) -> None:
        for group in self.swimmers:
            for swimmer in self.swimmers[group]:
                age = self.swimmers[group][swimmer].age
                gender = self.swimmers[group][swimmer].gender
                blockedNumbers = []

                excelIndex = 2
                for session in swimMeet.program:
                    for event in swimMeet.program[session]['events']:
                        excelIndex += 1
                        if age < event["minAge"] or age > event["maxAge"]:
                            blockedNumbers.append(excelIndex)
                        if event["gender"] != "Mixed" and event["gender"] != gender:
                            blockedNumbers.append(excelIndex)
                    excelIndex += 1
                self.swimmers[group][swimmer].blockedNumbers = blockedNumbers

class Excel:
    workbook = None
    registerSheet = None
    summarySheet = None

    # Make the cells of events where a swimmer cannot compete in grey
    blackout = True

    def __init__(self) -> None:
        self.__checkExcelFolder()

    def __checkExcelFolder(self) -> None:
        if not os.path.isdir("excel"):
            os.mkdir("excel")

    def __createStyles(self) -> None:
        self.s_bold = self.workbook.add_format({"bold": True})
        self.s_summaryHeader = self.workbook.add_format({"bold": True, 
                                                         "bottom": True})
        self.s_groupName = self.workbook.add_format({"bold": True, 
                                                     "bg_color": "gray", 
                                                     "align": "center", 
                                                     "right": True})
        self.s_swimmer = self.workbook.add_format({"shrink": True, 
                                                   "right": True})
        self.s_blackedOut = self.workbook.add_format({"bg_color": "gray"})
        self.s_sessionCell = self.workbook.add_format({"rotation": -90, 
                                                       "bold": True, 
                                                       "bg_color": "gray", 
                                                       "align": "center",
                                                       "valign": "vcenter", 
                                                       "text_wrap": True})
        self.s_ageCell = self.workbook.add_format({"text_wrap": True})

    def __createRegisterSheet(self) -> None:
        self.registerSheet = self.workbook.add_worksheet(name="Inschrijving")
        
        self.registerSheet.write("A1", "Wedstrijd: ", self.s_bold)
        self.registerSheet.write("A2", "Zwembad: ", self.s_bold)
        self.registerSheet.write("F1", "Deadline: ", self.s_bold)
        self.registerSheet.write("F2", "Qualify: ", self.s_bold)
        self.registerSheet.write("A4", "Event nr.", self.s_bold)
        self.registerSheet.write("A5", "Geslacht", self.s_bold)
        self.registerSheet.write("A6", "Event", self.s_bold)
        self.registerSheet.write("A7", "Leeftijd", self.s_bold)

        self.registerSheet.set_column("A:A", 30)
        self.registerSheet.freeze_panes(7, 0)

        print("Created the register sheet")

    def __createSummarySheet(self) -> None:
        self.summarySheet = self.workbook.add_worksheet(name="Samenvatting")

        self.summarySheet.set_column("A:A", 30)
        self.summarySheet.set_column("B:C", 45)

        self.summarySheet.write("B1", "Wedstrijd", self.s_summaryHeader)
        self.summarySheet.write("C1", "Wedstrijd Nummer", self.s_summaryHeader)

    def createBaseExcel(self, meetName: str) -> None:
        '''
        Create the base template of the excel
        '''
        self.filename = "excels/inschrijving_" + meetName + ".xlsx"
        self.workbook = xlsxwriter.Workbook(self.filename)
        print("Created Excel")
        
        self.__createStyles()

        self.__createRegisterSheet()
        self.__createSummarySheet()

    def __fillRegisterSheet(self, swimMeet: SwimMeet, club: Club) -> None:
        self.registerSheet.merge_range("B1:E1", swimMeet.name)
        self.registerSheet.merge_range("B2:E2", swimMeet.city)
        self.registerSheet.merge_range("H1:J1", swimMeet.deadline)
        self.registerSheet.merge_range("H2:J2", swimMeet.qualify)

        # Start adding the groups/swimmers after the general meet information
        rowIndex = 8
        for group in club.groupsToUse:
            self.registerSheet.write("A"+str(rowIndex), group, self.s_groupName)
            for nameOfSwimmer in club.swimmers[group]:
                rowIndex += 1
                self.registerSheet.write("A"+str(rowIndex), nameOfSwimmer, self.s_swimmer)
                if self.blackout:
                    for blockedNumber in club.swimmers[group][nameOfSwimmer].blockedNumbers:
                        self.registerSheet.write(xlsxwriter.utility.xl_col_to_name(blockedNumber-1)+str(rowIndex), "", self.s_blackedOut)
            rowIndex += 1
        print(", ".join(club.groupsToUse) + " added to the register sheet")

    def __addEventsToRegisterSheet(self, swimMeet: SwimMeet, club: Club) -> None:
        self.registerColumnCounter = 1

        for session in swimMeet.program:
            sessionInfo = swimMeet.program[session]
            # Split the rows in the different sessions
            self.registerSheet.merge_range(7, self.registerColumnCounter, 6 + club.amountOfSwimmers, self.registerColumnCounter, 
                                        "Sessie " + sessionInfo["sessionNumber"] +
                                        " - Start : " + sessionInfo["sessionStart"] +
                                        " - Warm-up: " + sessionInfo["sessionWarmupStart"] +
                                        " -> " + sessionInfo["sessionWarmupUntil"], self.s_sessionCell)
            
            self.registerColumnCounter += 1
            eventCounter = 0

            # Fill up the excel
            for event in sessionInfo["events"]:
                self.registerSheet.write(3, self.registerColumnCounter, "#" + event["number"])
                self.registerSheet.write(4, self.registerColumnCounter, event["gender"])
                self.registerSheet.write(5, self.registerColumnCounter, event["style"])
                self.registerSheet.write(6, self.registerColumnCounter, event["simplifiedAge"], self.s_ageCell)
                self.registerColumnCounter += 1
                eventCounter += 1

            self.registerSheet.set_column(self.registerColumnCounter-eventCounter, self.registerColumnCounter-1, 15)
        
        print("Sessions and events added to the register sheet")

    def __fillSummarySheet(self, club: Club) -> None:
        # Add the swimmers
        rowIndex = 2
        for group in club.groupsToUse:
            self.summarySheet.write("A"+str(rowIndex), group, self.s_groupName)
            self.summarySheet.write("A"+str(rowIndex), "", self.s_blackedOut)
            for nameOfSwimmer in club.swimmers[group]:
                rowIndex += 1
                self.summarySheet.write("A"+str(rowIndex), nameOfSwimmer, self.s_swimmer)
            rowIndex += 1

        # Add the excel formula's
        columnLetter = xlsxwriter.utility.xl_col_to_name(self.registerColumnCounter-1)
        for i in range(club.amountOfSwimmers):
            self.summarySheet.write_array_formula('B'+str(2+i), '=_xlfn.TEXTJOIN(", ", TRUE, IF(ISBLANK(Inschrijving!C'+str(8+i) + ':' + 
                                            columnLetter + str(8+i) +'), "", Inschrijving!$C$6:$' + columnLetter + '$6))')
            self.summarySheet.write_array_formula('C'+str(2+i), '=_xlfn.TEXTJOIN(", ", TRUE, IF(ISBLANK(Inschrijving!C'+str(8+i) + ':' + 
                                            columnLetter + str(8+i) +'), "", Inschrijving!$C$4:$' + columnLetter + '$4))')

    def fillAndSaveExcel(self, swimMeet: SwimMeet, club: Club) -> None:
        self.__fillRegisterSheet(swimMeet, club)
        self.__addEventsToRegisterSheet(swimMeet, club)
        self.__fillSummarySheet(club)
        self.workbook.close()
        print(f"Excel saved at {self.filename}")

class TimeStandards:
    standards = []
    timeStandardsRoot = None
    timeStandardsNames = set()

    def __init__(self, root) -> None:
        for node in root:
            if node.tag == "TIMESTANDARDLISTS":
                self.timeStandardsRoot = node
                break
        
        if self.timeStandardsRoot == None:
            print("No time standards present in the lenex")
            return
        
        for i in self.timeStandardsRoot:
            self.timeStandardsNames.add(i.attrib.get("name", "Invalid"))
        print(f"Time standards present: {self.timeStandardsNames}")


    def loadFromLef(self, name: str):
        self.standards[name] = {}
        for timestandardList in self.timeStandardsRoot:
            if timestandardList.attrib.get("name", "?") != name:
                continue

            self.standards[name]["type"] = timestandardList.attrib.get("type", "?")
            
            for node in timestandardList:
                if node.tag == "AGEGROUP":
                    temp["age"] = node.attrib.get("agemin") + " -> " + node.attrib.get("agemax")
                if node.tag == "TIMESTANDARDS":
                    for timestandard in node:
                        print(timestandard.attrib.get("swimtime", "?"))


def extractAndLoadLenex():
    print("Select the lxf file.")
    filepath = easygui.fileopenbox(default= competitionFolder, title="Select Lenex File")
    filename = filepath.split("\\")[-1]
    print(f"{filename} is selected.")

    if not os.path.isdir("zips"):
        os.mkdir("zips")
    if not os.path.isdir("lefs"):
        os.mkdir("lefs")

    # If the zip is present then assume that the lef is also extracted
    if not os.path.exists("zips/" + filename + ".zip"):
        shutil.copy(filepath, "zips")
        os.rename("zips/"+filename, "zips/"+filename+".zip")

        with ZipFile("zips/"+filename+".zip", 'r') as zipped:
            zipped.extractall(path="lefs")

        print("Extracted the lef from the lenex file. Available in the lefs folder.")
    else:
        print("Lef file is already extracted")

    root = ET.parse("lefs/" + filename.split(".")[0] + ".lef").getroot()

    if root.tag != "LENEX":
        raise Exception("File is not a lenex file")
    
    print("Xml correctly extracted from the lenex file.")

    return root

def main():
    # Open GUI to select a lenex, extract the xml and save in root
    root = extractAndLoadLenex()
    
    # Create a club and select which groups need to be added to the excel
    club = Club()
    groupsToUse = easygui.multchoicebox("Select the groups to use in the excel file", "Group selection", club.groups, preselect="2")

    # Create a swimmeet
    swimMeet = SwimMeet()
    if swimMeet.createFromLef(root) != 0:
        return
    
    club.loadFromFile(jsonFileName, groupsToUse, swimMeet.ageDate)

    time = TimeStandards(root)
    #time.loadFromLef("Loodsvisje 2023")

    # Check which swimmers can compete in which events
    club.checkPossibleEvents(swimMeet)

    # Create, fill and save the excel
    excel = Excel()
    excel.createBaseExcel(swimMeet.name)
    excel.fillAndSaveExcel(swimMeet, club)

if __name__ == "__main__":
    main()
