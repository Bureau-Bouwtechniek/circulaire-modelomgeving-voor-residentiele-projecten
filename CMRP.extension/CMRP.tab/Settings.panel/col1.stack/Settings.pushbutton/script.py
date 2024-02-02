# -*- coding: utf-8 -*-

# Imports
# =======================================================
import clr, os, sys

clr.AddReference("System.Windows.Forms")
clr.AddReference("IronPython.Wpf")

from pyrevit import forms, revit, script
from Autodesk.Revit.DB import *
from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Controls import CheckBox, StackPanel, TreeViewItem

from CMRPSettings import *
from CMRP import get_linked_document_from_name

# RevitDocument
# =======================================================
doc = revit.doc
uidoc = revit.uidoc

# Definitions and global variables
# =======================================================
logger = script.get_logger()

# Read the settings file
SETTINGS_FILE = list_settings_dir()
SETTINGS_LIST = SettingsList.from_json(SETTINGS_FILE)

# Main
# =======================================================


class ErrorPopup(forms.WPFWindow):
    def __init__(self, message):
        forms.WPFWindow.__init__(self, "error_popup.xaml")
        self.error_message.Text = message

    def close_button_Click(self, sender, e):
        self.Close()


def show_error_popup(message):
    error_popup = ErrorPopup(message)
    error_popup.ShowDialog()


def get_window_types_and_families(model_rooms_windows):
    """
    Create a dictionary:
        key:    window family
        value:  list of window types
    """
    windows = (
        FilteredElementCollector(model_rooms_windows)
        .OfCategory(BuiltInCategory.OST_Windows)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    family_type_dict = {}
    for window in windows:
        family_name = window.Symbol.Family.Name
        type_name = window.Symbol.LookupParameter("Type Name").AsString()
        if family_name in family_type_dict.keys():
            if type_name not in family_type_dict[family_name]:
                family_type_dict[family_name].append(type_name)
        else:
            family_type_dict[family_name] = [type_name]
    return family_type_dict


class SettingsWindow(forms.WPFWindow):
    # TODO add bool for initializing
    def __init__(self, xaml_file_name):
        self.initializing = True
        forms.WPFWindow.__init__(self, xaml_file_name)
        # Check if settings are loaded from the file, load if that is not the case
        global SETTINGS_LIST
        if SETTINGS_LIST is None or SETTINGS_LIST.list is None:
            SETTINGS_LIST = SettingsList.from_json(SETTINGS_FILE)
        self.init_project_selection_combobox()
        # Get the project name from the project information
        project_name = (
            FilteredElementCollector(doc).OfClass(ProjectInfo).FirstElement().Name
        )
        if project_name is None or project_name == "":
            logger.critical(
                "CMRP Settings: The ProjectName is missing in the Project Information."
            )
            sys.exit(1)
        if project_name in SETTINGS_LIST.get_project_names():
            # Settings for this project already exist in the json file -> show settings
            self.project_selection_cb.SelectedItem = project_name
            self.populate_ui_fields(project_name)
        else:
            # Settings this project don't yet exist in the json file -> show empty ui
            self.populate_model_combo_boxes(None)
        self.initializing = False

    def init_project_selection_combobox(self):
        self.project_selection_cb.ItemsSource = ObservableCollection[str]()
        for project_name in SETTINGS_LIST.get_project_names():
            self.project_selection_cb.ItemsSource.Add(project_name)

    def populate_ui_fields(self, project_name):
        settings = SETTINGS_LIST.get_settings(project_name)
        family_dict = settings.selected_window_families
        self.project_name_tb.Text = settings.project_name
        self.project_address_tb.Text = settings.project_address
        self.architect_name_tb.Text = settings.architect_name
        self.populate_model_combo_boxes(settings)
        self.param_window_height_tb.Text = settings.param_window_height
        self.param_window_width_tb.Text = settings.param_window_width
        self.populate_window_families_treeview(family_dict)
        self.mapping_file_t.Text = settings.mapping_file

    def populate_model_combo_boxes(self, settings):
        linked_models = [
            link.Name
            for link in (
                FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
            )
        ]
        self.model_levels_grids_cb.ItemsSource = ObservableCollection[str]()
        self.model_rooms_windows_cb.ItemsSource = ObservableCollection[str]()
        self.model_shafts_cb.ItemsSource = ObservableCollection[str]()
        for name in linked_models:
            self.model_levels_grids_cb.ItemsSource.Add(name)
            self.model_rooms_windows_cb.ItemsSource.Add(name)
            self.model_shafts_cb.ItemsSource.Add(name)
        if settings is not None:
            self.model_levels_grids_cb.SelectedItem = settings.model_levels_grids
            self.model_rooms_windows_cb.SelectedItem = settings.model_rooms_windows
            self.model_shafts_cb.SelectedItem = settings.model_shafts

    def populate_window_families_treeview(self, selected_window_families):
        self.window_families_treeview.Items.Clear()
        model_rooms_windows = get_linked_document_from_name(
            doc, self.model_rooms_windows_cb.SelectedItem
        )
        if model_rooms_windows is None:
            return
        family_type_dict = get_window_types_and_families(model_rooms_windows)
        for family, window_types in family_type_dict.items():
            family_treeview = TreeViewItem()
            family_checkbox = CheckBox()
            family_checkbox.Content = family
            family_checkbox.IsChecked = (
                selected_window_families is not None
                and family in selected_window_families.keys()
            )
            family_checkbox.Checked += self.update_check_boxes_of_family
            family_checkbox.Unchecked += self.update_check_boxes_of_family
            family_treeview.Header = family_checkbox
            for window_type in window_types:
                window_type_treeview = TreeViewItem()
                window_type_checkbox = CheckBox()
                window_type_checkbox.Content = window_type
                window_type_checkbox.IsChecked = (
                    selected_window_families is not None
                    and family in selected_window_families.keys()
                    and window_type in selected_window_families[family]
                )
                window_type_treeview.Header = window_type_checkbox
                family_treeview.Items.Add(window_type_treeview)
            self.window_families_treeview.Items.Add(family_treeview)

    # Functions called by Buttons
    def project_selection_changed(self, sender, args):
        if self.project_selection_cb.SelectedItem:
            self.initializing = True
            self.populate_ui_fields(self.project_selection_cb.SelectedItem)
            self.initializing = False

    def model_rooms_windows_changed(self, sender, args):
        if self.initializing:
            return
        self.populate_window_families_treeview(None)

    def update_check_boxes_of_family(self, sender, args):
        if self.initializing:
            return
        treeview = sender.Parent
        for item in treeview.Items:
            item.Header.IsChecked = sender.IsChecked

    def select_mapping_file(self, sender, args):
        new_path = forms.pick_excel_file()
        if new_path:
            new_path = os.path.normpath(new_path)
            self.mapping_file_t.Text = str(new_path)

    def save_settings(self, sender, args):
        if self.project_name_tb.Text is None or self.project_name_tb.Text == "":
            show_error_popup("The project name is required.")
            return

        # Read input data from UI window
        selected_window_families = {}
        for treeview in self.window_families_treeview.items:
            family_name = treeview.Header.Content
            types = []
            if not treeview.Header.IsChecked:
                continue
            for item in treeview.Items:
                if item.Header.IsChecked:
                    types.append(item.Header.Content)
            selected_window_families[family_name] = types
        settings = Settings(
            self.project_name_tb.Text,
            self.project_address_tb.Text,
            self.architect_name_tb.Text,
            self.model_levels_grids_cb.SelectedItem,
            self.model_rooms_windows_cb.SelectedItem,
            self.model_shafts_cb.SelectedItem,
            self.param_window_height_tb.Text,
            self.param_window_width_tb.Text,
            selected_window_families,
            self.mapping_file_t.Text,
        )
        # Add the new settings to the list
        SETTINGS_LIST.add_settings(settings)
        # Save the updated list to the JSON file
        SETTINGS_LIST.to_json(SETTINGS_FILE)
        # Close window
        self.Close()


# Main function: starts GUI
if __name__ == "__main__":
    SettingsWindow("settings_window.xaml").show_dialog()
