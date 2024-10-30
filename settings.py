'''Class to save the settings in between runs'''
import pickle
import os
import logging
import logging.config

from easygui import diropenbox, fileopenbox

class Settings:
    '''Containers to save different global setting
       and save them to use it in next runs'''
    SAVE_FILE = "data/.settings.pk"
    log = None

    def __init__(self):
        self.club_name = ""
        self.mdb_path = ""
        self.default_competition_path = ""
        self.club_logo_path = ""

        if not os.path.isdir("data"):
            os.mkdir("data")

    @staticmethod
    def init_settings():
        '''Check if there is a save file present, else
           get user input for all the different global settings'''
        # Try loading from saved file
        if os.path.exists(Settings.SAVE_FILE):
            with open(Settings.SAVE_FILE, "rb") as fi:
                return pickle.load(fi)

        # No save file present, ask user input
        setting = Settings()
        print("No settings save file found")
        setting.club_name = input("Club name: ")
        setting.mdb_path = fileopenbox(title="Select the team manager mdb")
        setting.default_competition_path = diropenbox(title="Select the directory \
                                                      with the competition folders")
        setting.club_logo_path = fileopenbox(title="Select the club logo to add\
                                             to all the excel sheets")

        # Save current settings
        with open(Settings.SAVE_FILE, "wb") as fi:
            pickle.dump(setting, fi)

        return setting

    @staticmethod
    def get_logger():
        '''Get the global logger'''
        if Settings.log is not None:
            return Settings.log

        # Create the logging facilities
        logging.config.fileConfig('data/logging.conf')
        Settings.log = logging.getLogger('lenexToExcelLog')

        Settings.log.debug("Logging initialized")

        return Settings.log
