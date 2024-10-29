from easygui import diropenbox, fileopenbox

import pickle
import os
import logging
import logging.config

class Settings:
    SAVE_FILE = ".settings.pk"
    log = None

    def __init__(self):
        self.club_name = ""
        self.mdb_path = ""
        self.default_competition_path = ""
    
    @staticmethod
    def init_settings():
        # Try loading from saved file
        if os.path.exists(Settings.SAVE_FILE):
            with open(Settings.SAVE_FILE, "rb") as fi:
                return pickle.load(fi)
        
        # No save file present, ask user input
        setting = Settings()
        print("No settings save file found")
        setting.club_name = input("Club name: ")
        setting.mdb_path = fileopenbox(title="Select the team manager mdb")
        setting.default_competition_path = diropenbox(title="Select the directory with the competition folders")

        # Save current settings
        with open(Settings.SAVE_FILE, "wb") as fi:
            pickle.dump(setting, fi)
        

        return setting
    
    @staticmethod
    def get_logger():
        if Settings.log is not None:
            return Settings.log
        
        # Create the logging facilities
        logging.config.fileConfig('logging.conf')
        Settings.log = logging.getLogger('lenexToExcelLog')

        Settings.log.debug("Logging initialized")

        return Settings.log


        
