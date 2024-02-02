# -*- coding: utf-8 -*-

# Imports
# =======================================================
import traceback
from pyrevit import revit, script
from pyrevit.forms import ProgressBar

from Autodesk.Revit.DB import *

from CMRPGUI import linked_element_selection
from CMRPSettings import load_current_settings
from CMRP import get_linked_document_from_name


# Definitions and global variables
# =======================================================
logger = script.get_logger()


def update_grids(document, model_levels_grids):
    """
    Update the grids from model_levels_grids to document

    Parameters
    ----------
    document: RevitDocument
        The current opened document
    model_levels_grids: RevitDocument
        The linked document selected by the user
    """
    # Show popup for the user to select the grids that need to be updated
    selected_grids = linked_element_selection(
        "Select the grids u want to import or update",
        FilteredElementCollector(model_levels_grids).OfClass(Grid).ToElements(),
        True,
    )
    if selected_grids is None:
        return

    # Retrieve Current Grids
    current_grids = FilteredElementCollector(document).OfClass(Grid).ToElements()

    grids_to_be_deleted = []
    grids_to_be_updated = {}
    selected_grid_names = [selected_grid.Name for selected_grid in selected_grids]
    for grid in current_grids:
        if grid.Name in selected_grid_names:
            # Get the grids that need to be updated:
            # grids that are present in the document and are selected by the user
            grids_to_be_updated[grid.Name] = grid
        else:
            # Get the grids that need to be deleted:
            # grids are present in the document but aren't selected by the user
            grids_to_be_deleted.append(grid.Id)

    if len(grids_to_be_deleted) > 0:
        # Delete unnecessary grids
        with Transaction(document, "CMRP: Deleting old grids") as transaction:
            transaction.Start()
            value = 0
            max_value = len(grids_to_be_deleted)
            with ProgressBar(
                title="CMRP: Deleting old grids: ({value} of {max_value})"
            ) as progress_bar:
                for id in grids_to_be_deleted:
                    document.Delete(id)
                    value += 1
                    progress_bar.update_progress(value, max_value)
            transaction.Commit()

    if len(selected_grids) > 0:
        # Import / Create new selected grids
        with Transaction(document, "CMRP: Creating/Updating grids") as transaction:
            transaction.Start()
            value = 0
            max_value = len(selected_grids)
            with ProgressBar(
                title="CMRP: Creating/Updating grids: ({value} of {max_value})"
            ) as progress_bar:
                for grid in selected_grids:
                    if grid.Name in grids_to_be_updated.keys():
                        # grid already exists in the current document -> update
                        try:
                            new_grid = grids_to_be_updated[grid.Name]
                            new_grid.Curve = grid.Curve
                        except Exception:
                            logger.error(traceback.format_exc())
                            continue
                    else:
                        # grid doesn't exist in current document -> create
                        new_grid = Grid.Create(document, grid.Curve)
                        new_grid.Name = grid.Name

                    value += 1
                    progress_bar.update_progress(value, max_value)
            transaction.Commit()


def update_levels(document, model_levels_grids):
    """
    Update the levels from model_levels_grids to document

    Parameters
    ----------
    document: RevitDocument
        The current opened document
    model_levels_grids: RevitDocument
        The linked document selected by the user
    """
    # Show popup for the user to select the levels that need to be updated
    selected_levels = linked_element_selection(
        "Select the levels u want to import or update",
        FilteredElementCollector(model_levels_grids).OfClass(Level).ToElements(),
        True,
    )
    if selected_levels is None:
        return

    # Retrieve Current Grids
    current_levels = FilteredElementCollector(document).OfClass(Level).ToElements()

    levels_to_be_deleted = []
    levels_to_be_updated = {}
    selected_level_names = [selected_level.Name for selected_level in selected_levels]
    for level in current_levels:
        if level.Name in selected_level_names:
            # Get the levels that need to be updated:
            # levels that are present in the document and are selected by the user
            levels_to_be_updated[level.Name] = level
        else:
            # Get the levels that need to be deleted:
            # levels are present in the document but aren't selected by the user
            levels_to_be_deleted.append(level.Id)

    if len(levels_to_be_deleted) > 0:
        # Delete unnecessary levels
        with Transaction(document, "CMRP: Deleting old levels") as transaction:
            transaction.Start()
            value = 0
            max_value = len(levels_to_be_deleted)
            with ProgressBar(
                title="CMRP: Deleting old levels: ({value} of {max_value})"
            ) as progress_bar:
                for id in levels_to_be_deleted:
                    element_ids = (
                        FilteredElementCollector(document)
                        .WhereElementIsNotElementType()
                        .WherePasses(ElementLevelFilter(id))
                        .ToElementIds()
                    )
                    document.Delete(element_ids)
                    try:
                        document.Delete(id)
                    except Exception:
                        logger.error("{} {}".format(traceback.format_exc(), id))
                    value += 1
                    progress_bar.update_progress(value, max_value)
            transaction.Commit()

    if len(selected_levels) > 0:
        # Import / Create new selected levels
        with Transaction(document, "CMRP: Creating/Updating levels") as transaction:
            transaction.Start()
            value = 0
            max_value = len(selected_levels)
            with ProgressBar(
                title="CMRP: Creating/Updating levels: ({value} of {max_value})"
            ) as progress_bar:
                for level in selected_levels:
                    if level.Name in levels_to_be_updated.keys():
                        # linked_level already exists in the current document -> update
                        try:
                            new_level = levels_to_be_updated[level.Name]
                            new_level.Elevation = level.Elevation
                        except Exception:
                            logger.error(traceback.format_exc())
                            continue
                    else:
                        # linked_level doesn't exist in current document -> create
                        new_level = Level.Create(document, level.Elevation)
                        new_level.Name = level.Name
                    value += 1
                    progress_bar.update_progress(value, max_value)
            transaction.Commit()


# Main function: runs script
if __name__ == "__main__":
    # RevitDocument
    doc = revit.doc
    uidoc = revit.uidoc
    # Load settings from json file
    SETTINGS = load_current_settings(doc)
    model_levels_grids = get_linked_document_from_name(doc, SETTINGS.model_levels_grids)
    update_grids(doc, model_levels_grids)
    update_levels(doc, model_levels_grids)
