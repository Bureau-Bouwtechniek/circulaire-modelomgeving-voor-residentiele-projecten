# -*- coding: utf-8 -*-

# Imports
# =======================================================
import math
import traceback

from pyrevit import script
from pyrevit.forms import ProgressBar

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import Structure, Material

from CMRP import (
    delete_elements,
    get_linked_document_from_name,
    get_size_from_window,
    delete_area_plans,
    create_area_plans,
)
from CMRPUnit import units_to_internal
from CMRPColors import CMRP_colors_hard
from CMRPParameters import *
from CMRPSettings import MissingSettingsError

# Definitions and global variables
# =======================================================
logger = script.get_logger()


def delete_old_window_blocks(document):
    """
    Delete the window blocks created by a previous execution of this script

    Parameters:
    ----------
    document: RevitDocument
        The document from which to delete the shafts

    Returns:
    ----------
    FamilySymbol
        The FamilySymbol with the name 'CMRP_31_GM_UN_WindowBlock'
    """
    window_block_family = (
        FilteredElementCollector(document)
        .OfClass(FamilySymbol)
        .WhereElementIsElementType()
        .WherePasses(
            ElementParameterFilter(
                FilterStringRule(
                    ParameterValueProvider(
                        ElementId(BuiltInParameter.SYMBOL_NAME_PARAM)
                    ),
                    FilterStringEquals(),
                    CMRP_31_GM_UN_WindowBlock,
                )
            )
        )
        .FirstElement()
    )
    if window_block_family is None:
        raise MissingTemplateError(
            "{}: Can't find the FamilySymbol '{}'.".format(
                WINDOWTYPES_TEMPLATE_NAME, CMRP_31_GM_UN_WindowBlock
            )
        )
    old_window_block_ids = (
        FilteredElementCollector(document)
        .OfCategoryId(window_block_family.Category.Id)
        .OfClass(FamilyInstance)
        .WherePasses(
            ElementParameterFilter(
                FilterStringRule(
                    ParameterValueProvider(
                        ElementId(BuiltInParameter.SYMBOL_NAME_PARAM)
                    ),
                    FilterStringEquals(),
                    CMRP_31_GM_UN_WindowBlock,
                )
            )
        )
        .ToElementIds()
    )
    delete_elements("Deleting old window blocks", document, old_window_block_ids)
    return window_block_family


def create_window_type_material_dict(document, window_types):
    """
    Create dictionary window_type names -> material ids

    The dictionary exist of:
    - the key: name of the window_type
    - the value: an array of:
        - the id of the corresponding material
        - the color of the material in hexadecimal
        - a counter to count the amount of windows with this window_type

    Parameters:
    ----------
    ocument: RevitDocument
        The current document
    window_types: FamilyType[]
        A list of unique window types

    Returns:
    -------
    dictionary
    """
    cmrp_materials = (
        FilteredElementCollector(document)
        .OfClass(Material)
        .WhereElementIsNotElementType()
        .WherePasses(
            ElementParameterFilter(
                FilterStringRule(
                    ParameterValueProvider(ElementId(BuiltInParameter.MATERIAL_NAME)),
                    FilterStringBeginsWith(),
                    "CMRP",
                )
            )
        )
        .ToElements()
    )
    if cmrp_materials is None or len(cmrp_materials) == 0:
        raise MissingTemplateError(
            "{}: No materials that start with 'CMRP' found.".format(
                WINDOWTYPES_TEMPLATE_NAME
            )
        )

    window_type_material_dict = {}
    colors = CMRP_colors_hard()
    index = 0
    for window_type in window_types:
        color = colors[index]
        index += 1
        if index >= len(colors):
            logger.warning(
                "{}: There aren't enough different colors. Recycling used colors.".format(
                    WINDOWTYPES_TEMPLATE_NAME
                )
            )
            index = 0

        material_name = "CMRP_Hard_{},{},{}".format(color.Red, color.Green, color.Blue)
        for material in cmrp_materials:
            # Find the correct material by name
            if material.Name == material_name:
                hex_color = "#{:02X}{:02X}{:02X}".format(
                    int(color.Red), int(color.Green), int(color.Blue)
                )
                window_type_material_dict[
                    window_type.LookupParameter("Type Name").AsString()
                ] = [material.Id, hex_color, 0, window_type.Family.Name]
                break
    return window_type_material_dict


def retrieve_data_from_document(model_rooms_windows, selected_window_families):
    """
    Retrieve data from the Revit (linked) document

    Logs errors for values the script absolutly needs -> throws exit code -1

    Parameters:
    ----------
    model_rooms_windows: RevitDocument
        The document that is linked to 'document' from which to retrieve data
    selected_window_families: Family[]
        A list of Revit Families selected by the user to use as filter for windows

    Returns:
    -------
    Window[], WindowType[]:
        A list of windows filtered by 'selected_window_families'
        A list of unique window types of the windows
    """
    windows = (
        FilteredElementCollector(model_rooms_windows)
        .OfCategory(BuiltInCategory.OST_Windows)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    if windows is None or len(windows) == 0:
        raise MissingTemplateError(
            "{}: No windows found in the document.".format(WINDOWTYPES_TEMPLATE_NAME)
        )

    filtered_windows = []
    window_type_names = []
    window_types = set()
    for window in windows:
        family_name = window.Symbol.Family.Name
        type_name = window.Symbol.LookupParameter("Type Name").AsString()
        if (
            family_name in selected_window_families.keys()
            and type_name in selected_window_families[family_name]
        ):
            filtered_windows.append(window)
            if type_name not in window_type_names:
                window_type_names.append(type_name)
                window_types.add(window.Symbol)

    sorted_window_types = sorted(
        window_types, key=lambda wt: wt.LookupParameter("Type Name").AsString()
    )
    return filtered_windows, sorted_window_types


def draw_type_window_function_result(window_type_material_dict):
    """
    Draw a Pie chart to show an overview of the window types and a count
    """
    chart = script.get_output().make_pie_chart()
    chart.data.labels = [
        "{}: {}".format(window_type_material_dict[type][3], type)
        for type in window_type_material_dict.keys()
    ]
    chart.options.title = {
        "display": True,
        "position": "top",
        "fullWidth": True,
        "fontSize": 30,
        "fontFamily": "Arial",
        "fontColor": "#666",
        "fontStyle": "bold",
        "padding": 10,
        "text": "Window Types",
    }

    data_set = chart.data.new_dataset("Window Types")
    data_set.data = [item[2] for item in window_type_material_dict.values()]
    data_set.backgroundColor = [item[1] for item in window_type_material_dict.values()]
    chart.draw()


def check_necessary_settings(settings):
    if settings.model_rooms_windows is None or settings.model_rooms_windows == "":
        raise MissingSettingsError(
            "{}: The model for Rooms & Windows is missing in the settings.".format(
                WINDOWTYPES_TEMPLATE_NAME
            )
        )
    if settings.param_window_height is None or settings.param_window_height == "":
        raise MissingSettingsError(
            "{}: The parameter for the height of windows is missing in the settings.".format(
                WINDOWTYPES_TEMPLATE_NAME
            )
        )
    if settings.param_window_width is None or settings.param_window_height == "":
        raise MissingSettingsError(
            "{}: The parameter for the height of windows is missing in the settings.".format(
                WINDOWTYPES_TEMPLATE_NAME
            )
        )
    if (
        settings.selected_window_families is None
        or len(settings.selected_window_families.keys()) == 0
    ):
        raise MissingSettingsError(
            "{}: No window families or type were selected, check in the settings.".format(
                WINDOWTYPES_TEMPLATE_NAME
            )
        )


# Main
# =======================================================


def run_window_types_script(document, uidocument, settings):
    """
    This script creates a volume for each window and adds a color for each window type
    therefore creating a 3D visual overview of where window types are placed.

    It performs the following tasks:
        0. Delete previously created window blocks
        1. Retrieve data from the Revit document, including window information and relevant elements
        and filter the windows by the user selected families
        2. Create area plans to view window blocks
        3. Create a dictionary to map window types to a color, family and counter
        4. Iterate through the windows and calculates the boundary for the window block
        5. Place the windowblocks and assign the mapped color
        6. Draw a pie chart to show a overview of window types and the amount for each
    """
    try:
        check_necessary_settings(settings)
        # 0. Delete previously created window blocks
        window_block_family = delete_old_window_blocks(document)
        delete_area_plans(document, uidocument, WINDOWTYPES_TEMPLATE_NAME[:3])
        # Retrieve data from settings
        model_rooms_windows = get_linked_document_from_name(
            document, settings.model_rooms_windows
        )
        selected_window_families = settings.selected_window_families
        # 1. Retrieve data from the Revit document
        (
            filtered_windows,
            window_types,
        ) = retrieve_data_from_document(model_rooms_windows, selected_window_families)
        # 2. Create area plans to view window blocks
        views = create_area_plans(document, WINDOWTYPES_TEMPLATE_NAME)

        # 3. Create a dictionary to map window types to a color, family and counter
        window_type_material_dict = create_window_type_material_dict(
            document, window_types
        )
        # 4. Iterate through the windows and calculates the boundary for the window block
        with Transaction(
            document, "CMRP {}: Create Window Blocks".format(WINDOWTYPES_TEMPLATE_NAME)
        ) as transaction:
            transaction.Start()
            value = 0
            max_value = len(filtered_windows)
            with ProgressBar(
                title="CMRP "
                + WINDOWTYPES_TEMPLATE_NAME
                + ": Creating New Window Blocks: ({value} of {max_value})"
            ) as progress_bar:
                for window in filtered_windows:
                    window_type = window.Symbol.LookupParameter("Type Name").AsString()
                    location = window.Location.Point
                    height, width = get_size_from_window(
                        window,
                        settings.param_window_height,
                        settings.param_window_width,
                    )
                    if height is None or width is None:
                        max_value -= 1
                        continue
                    facing_orientation = window.FacingOrientation

                    # Get the correct orientation of the window and then rotate the window_block
                    angle = (
                        -math.copysign(1.0, facing_orientation.X)
                        * math.acos(facing_orientation.Y)
                        + math.pi
                    )
                    axis = Line.CreateBound(
                        location, location + XYZ(0, 0, 1) * units_to_internal(500)
                    )
                    try:
                        # 5. Place the windowblocks and assign the mapped color
                        # Create a window_block in a certain color at the correct location
                        window_block = document.Create.NewFamilyInstance(
                            location,
                            window_block_family,
                            Structure.StructuralType.NonStructural,
                        )
                        ElementTransformUtils.RotateElement(
                            document, window_block.Id, axis, angle
                        )

                        # window_block.LookupParameter(CMRP_CF_TX_WindowType).set(window_type)

                        # Update the size of the window_block to match the size of the window
                        # window_block.LookupParameter("CMRP_CF_LE_Depth").Set()
                        window_block.LookupParameter(CMRP_CF_LE_Height).Set(height)
                        window_block.LookupParameter(CMRP_CF_LE_Width).Set(width)

                        window_block.LookupParameter(CMRP_CF_MA_Material).Set(
                            window_type_material_dict[window_type][0]
                        )
                        window_type_material_dict[window_type][2] += 1
                    except:
                        logger.error(
                            "{}: Couldn't create window block for window ({}).\n".format(
                                WINDOWTYPES_TEMPLATE_NAME, window.Id
                            )
                            + traceback.format_exc()
                        )
                    value += 1
                    progress_bar.update_progress(value, max_value)
            transaction.Commit()
        # 6. Draw a pie chart to show a overview of window types and the amount for each
        draw_type_window_function_result(window_type_material_dict)

        logger.success(
            "{}: Script completed without critical errors.".format(
                WINDOWTYPES_TEMPLATE_NAME
            )
        )
    except (MissingSettingsError, MissingTemplateError) as error:
        logger.critical(error.message)
