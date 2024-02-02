# -*- coding: utf-8 -*-

# Imports
# =======================================================
import json
import os

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *


# Settings Functions
# ====================================================================
def ext_dir():
    """Methode to locate the extention directory

    Returns:
        _type_: directory (os.path)
    """
    # split directorie
    directories = __file__.split(os.sep)
    # find extention level
    ext_index = directories.index("CMRP.extension")
    levels = len(directories) - ext_index - 1
    # return directorie
    result_dir = __file__
    for l in range(levels):
        result_dir = os.path.dirname(result_dir)
    return result_dir


def list_settings_dir():
    return os.path.join(ext_dir(), "bin", "list_settings.json")


class MissingSettingsError(Exception):
    def __init__(self, message):
        self.message = message


# Settings Class


class Settings:
    def __init__(
        self,
        project_name,
        project_address,
        architect_name,
        model_levels_grids,
        model_rooms_windows,
        model_shafts,
        param_window_height,
        param_window_width,
        selected_window_families,
        mapping_file,
    ):
        self.project_name = project_name
        self.project_address = project_address
        self.architect_name = architect_name
        self.model_levels_grids = model_levels_grids
        self.model_rooms_windows = model_rooms_windows
        self.model_shafts = model_shafts
        self.param_window_height = param_window_height
        self.param_window_width = param_window_width
        self.selected_window_families = selected_window_families
        self.mapping_file = mapping_file

    def __str__(self):
        return (
            "Settings:\n"
            "Project Name: {}\n"
            "Project Address: {}\n"
            "Architect Name: {}\n"
            "Model Levels and Grids: {}\n"
            "Model Rooms and Windows: {}\n"
            "Model Shafts: {}\n"
            "Parameter Window Height: {}\n"
            "Parameter Window Width: {}\n"
            "Selected Window Families: {}\n"
            "Mapping File: {}\n".format(
                self.project_name,
                self.project_address,
                self.architect_name,
                self.model_levels_grids,
                self.model_rooms_windows,
                self.model_shafts,
                self.param_window_height,
                self.param_window_width,
                self.selected_window_families,
                self.mapping_file,
            )
        )


class SettingsList:
    def __init__(self):
        self.list = []

    def add_settings(self, new_settings):
        # Check if the project_name already exists in the list and replace it if it does
        for index, settings in enumerate(self.list):
            if settings.project_name == new_settings.project_name:
                self.list.pop(index)
                break
        self.list.append(new_settings)

    def get_settings(self, project_name):
        for settings in self.list:
            if settings.project_name == project_name:
                return settings

    def get_project_names(self):
        return [settings.project_name for settings in self.list]

    def to_json(self, filename):
        with open(filename, "w") as json_file:
            settings_data = [settings.__dict__ for settings in self.list]
            json.dump(settings_data, json_file, indent=4, sort_keys=False)

    @classmethod
    def from_json(_, file_name):
        settings_list = SettingsList()
        with open(file_name, "r") as json_file:
            settings_data = json.load(json_file)
            for data in settings_data:
                settings = Settings(**data)
                settings.selected_window_families = {}
                family_dict = data["selected_window_families"]
                for family_name in family_dict:
                    settings.selected_window_families[family_name] = family_dict[
                        family_name
                    ]
                settings_list.list.append(settings)
        return settings_list


def load_current_settings(document):
    project_name = (
        FilteredElementCollector(document).OfClass(ProjectInfo).FirstElement().Name
    )
    file_name = list_settings_dir()
    settings_list = SettingsList.from_json(file_name)
    return settings_list.get_settings(project_name)
