import pickle
import os

class Settings:
    SAVE_FILE = ".settings.pk"

    def __init__(self):
        self.club_name = ""
        self.mdb_path = ""
    
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
        setting.mdb_path = input("Team manager mdb path: ") # TODO using gui to select path

        # Save current settings
        with open(Settings.SAVE_FILE, "wb") as fi:
            pickle.dump(setting, fi)

        return setting
        
