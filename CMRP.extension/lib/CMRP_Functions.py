# -*- coding: utf-8 -*-

# Imports
# =======================================================
import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
from Microsoft.Office.Interop import Excel

from pyrevit import script
from pyrevit.forms import ProgressBar
from Autodesk.Revit.DB import *

from CMRPParameters import (
    FUNCTIONS_TEMPLATE_NAME,
    Comments,
    CMRP_CP_TXT_Function,
    MissingTemplateError,
)
from CMRPSettings import MissingSettingsError
from CMRP import *


# Definitions and global variables
# =======================================================
logger = script.get_logger()


def get_mapping_data(filename):
    """
    Reads an excel file and stores the mapping data in a dictionary

    Parameters:
    ----------
    filename: str
        The path to the excel file

    Returns:
    -------
    dict{Room Name: [Room Type, Category]}: a dictionary with the Room name as key
    """
    # Create Excel application object
    excel_app = Excel.ApplicationClass()
    excel_app.Visible = False

    if filename:
        # Open the selected Excel workbook
        workbook = excel_app.Workbooks.Open(filename)
    else:
        raise MissingSettingsError(
            "{}: No excel file with the name '{}' was found.".format(
                FUNCTIONS_TEMPLATE_NAME, filename
            )
        )

    # check each worksheet for the name "Project_Informatie"
    found_sheet = False
    for worksheet in workbook.Worksheets:
        if worksheet.Name == "Room_Information":
            sheet = worksheet
            found_sheet = True
            break

    if not found_sheet:
        workbook.Close(False)  # Close without saving changes
        excel_app.Quit()
        raise MissingSettingsError(
            "{}: No sheet with name '{}' was found in the mapping file '{}'.".format(
                FUNCTIONS_TEMPLATE_NAME, "Room_Information", filename
            )
        )

    # Dictionary to store the data
    data = {}

    # Assuming the first row is header information
    # Starting from the second row
    for row in range(2, sheet.UsedRange.Rows.Count + 1):
        # Get Room Type and Category from the first and second column
        room_type = sheet.Cells(row, 1).Value2
        category = sheet.Cells(row, 2).Value2

        # Room names
        for col in range(3, sheet.UsedRange.Columns.Count + 1):
            if sheet.Cells(row, col).Value2 == None:
                continue
            data[sheet.Cells(row, col).Value2] = [room_type, category]

    # close Excel
    workbook.Close(False)  # Close without saving changes
    excel_app.Quit()

    return data


def retrieve_data_from_document(document, model_levels_grids, model_rooms_windows):
    """
    Retrieve data from the Revit (linked) document

    Parameters:
    ----------
    document: RevitDocument
        The document from which to retrieve data
    model_rooms_windows: RevitDocument
        The document that is linked to 'document' from which to retrieve data

    Returns:
    -------
    Level[], Room[]: arrays of Revit Classes Level and Room
    """
    levels = FilteredElementCollector(document).OfClass(Level).ToElements()
    imported_level_ids = get_imported_level_ids(levels, model_levels_grids)
    linked_rooms = get_rooms_of_imported_levels(imported_level_ids, model_rooms_windows)
    return levels, linked_rooms


def check_necessary_settings(settings):
    if settings.model_levels_grids is None or settings.model_levels_grids == "":
        raise MissingSettingsError(
            "{}: The model for Levels & Grids is missing in the settings.".format(
                FUNCTIONS_TEMPLATE_NAME
            )
        )
    if settings.model_rooms_windows is None or settings.model_rooms_windows == "":
        raise MissingSettingsError(
            "{}: The model for Rooms & Windows is missing in the settings.".format(
                FUNCTIONS_TEMPLATE_NAME
            )
        )
    if settings.mapping_file is None or settings.mapping_file == "":
        raise MissingSettingsError(
            "{}: The mapping excel file is missing in the settings.".format(
                FUNCTIONS_TEMPLATE_NAME
            )
        )


# Main
# ==================================================


def run_functions_script(document, uidocument, settings):
    """
    This script creates areas in Autodesk Revit and asigns them a category based on their name (use).

    It performs the following tasks:
        0. Delete previously created areas and area plans.
        1. Retrieve data from the Revit document, including room information.
        2. Create area plans for viewing the area's.
        3. Iterate through rooms to create area's representing the boundary of the room.
        4. Retrieve the category mapping data from an Excel file.
        5. Map the rooms to a room_type and category.
        6. Fill in the room_type and category parameters in the area corresponding to the room.
    """
    try:
        check_necessary_settings(settings)
        # 0. Delete previously created areas and area plans
        delete_old_areas_lines_and_areaplans(
            document, uidocument, FUNCTIONS_TEMPLATE_NAME
        )
        # Retrieve data from settings
        model_levels_grids = get_linked_document_from_name(
            document, settings.model_levels_grids
        )
        model_rooms_windows = get_linked_document_from_name(
            document, settings.model_rooms_windows
        )
        mapping_file = settings.mapping_file
        # 1. Retrieve data from the Revit document, including room information.
        levels, linked_rooms = retrieve_data_from_document(
            document, model_levels_grids, model_rooms_windows
        )
        # 2. Create area plans for viewing the area's.
        area_plans = create_area_plans(document, FUNCTIONS_TEMPLATE_NAME)

        # 3. Iterate through rooms to create area's representing the boundary of the room.
        rooms_dict = create_sketchplans_and_areas(
            document, linked_rooms, levels, area_plans, FUNCTIONS_TEMPLATE_NAME
        )

        # 4. Retrieve the category mapping data from an Excel file
        mapping_data = get_mapping_data(mapping_file)
        # 5. Map the rooms to a room_type and category
        with Transaction(
            document, "CMRP {}: Map rooms to categories".format(FUNCTIONS_TEMPLATE_NAME)
        ) as transaction:
            transaction.Start()
            value = 0
            max_value = len(rooms_dict.keys())
            with ProgressBar(
                title="CMRP "
                + FUNCTIONS_TEMPLATE_NAME
                + ": Mapping rooms to categories: ({value} of {max_value})"
            ) as progress_bar:
                for values in rooms_dict.values():
                    area = values[1]
                    category_found = False
                    for room_name in mapping_data.keys():
                        if Element.Name.GetValue(area).startswith(room_name):
                            data_row = mapping_data[room_name]
                            room_type = data_row[0]
                            category = data_row[1]
                            # 6. Fill in the room_type and category parameters in the area corresponding to the room
                            area.LookupParameter(Comments).Set(room_type)
                            area.LookupParameter(CMRP_CP_TXT_Function).Set(category)
                            category_found = True
                            break
                    if not category_found:
                        logger.warning(
                            "{}: No mapping was found for the room name: {}".format(
                                FUNCTIONS_TEMPLATE_NAME, Element.Name.GetValue(area)
                            )
                        )
                    value += 1
                    progress_bar.update_progress(value, max_value)
            transaction.Commit()
        logger.success(
            "{}: Script completed without critical errors.".format(
                FUNCTIONS_TEMPLATE_NAME
            )
        )
    except (MissingSettingsError, MissingTemplateError) as error:
        logger.critical(error.message)
