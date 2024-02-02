# -*- coding: utf-8 -*-

# Imports
# =======================================================

from pyrevit import script

from CMRP import *
from CMRPParameters import SURFACES_TEMPLATE_NAME, MissingTemplateError
from CMRPSettings import MissingSettingsError

# Definitions and global variables
# =======================================================
logger = script.get_logger()


def retrieve_data_from_document(document, model_levels_grids, model_rooms_windows):
    """
    Retrieve data from the Revit (linked) document

    Parameters:
    ----------
    document: RevitDocument
        The document from which to retrieve data
    model_levels_grids: RevitDocument
        The document that is linked to 'document' from which to retrieve data
    model_rooms_windows: RevitDocument
        The document that is linked to 'document' from which to retrieve data
    """
    levels = FilteredElementCollector(document).OfClass(Level).ToElements()
    imported_level_ids = get_imported_level_ids(levels, model_levels_grids)
    linked_rooms = get_rooms_of_imported_levels(imported_level_ids, model_rooms_windows)
    return levels, linked_rooms


def check_necessary_settings(settings):
    if settings.model_levels_grids is None or settings.model_levels_grids == "":
        raise MissingSettingsError(
            "{}: The model for Levels & Grids is missing in the settings.".format(
                SURFACES_TEMPLATE_NAME
            )
        )
    if settings.model_rooms_windows is None or settings.model_rooms_windows == "":
        raise MissingSettingsError(
            "{}: The model for Rooms & Windows is missing in the settings.".format(
                SURFACES_TEMPLATE_NAME
            )
        )


# Main
# ==================================================


def run_surfaces_script(document, uidocument, settings):
    """
    This script creates areaplans for each level and areas for each room
    
    It performs the following tasks:
        0. Delete previously created areas and area plans.
        1. Retrieve data from the Revit document, including room information.
        2. Create area plans for viewing the area's.
        3. Iterate through rooms to create area's representing the boundary of the room.
    """
    try:
        check_necessary_settings(settings)
        # 0. Delete previously created areas and area plans
        delete_old_areas_lines_and_areaplans(
            document, uidocument, SURFACES_TEMPLATE_NAME
        )
        # Retrieve data from settings
        model_levels_grids = get_linked_document_from_name(
            document, settings.model_levels_grids
        )
        model_rooms_windows = get_linked_document_from_name(
            document, settings.model_rooms_windows
        )
        # 1. Retrieve data from the Revit document, including room information.
        levels, linked_rooms = retrieve_data_from_document(
            document, model_levels_grids, model_rooms_windows
        )
        # 2. Create area plans for viewing the area's.
        area_plans = create_area_plans(document, SURFACES_TEMPLATE_NAME)
        # 3. Iterate through rooms to create area's representing the boundary of the room.
        create_sketchplans_and_areas(
            document, linked_rooms, levels, area_plans, SURFACES_TEMPLATE_NAME
        )
        logger.success(
            "{}: Script completed without critical errors.".format(
                SURFACES_TEMPLATE_NAME
            )
        )
    except (MissingSettingsError, MissingTemplateError) as error:
        logger.critical(error.message)
