# -*- coding: utf-8 -*-

# Imports
# =======================================================
from pyrevit import script
from pyrevit.forms import ProgressBar
from Autodesk.Revit.DB import *

from CMRP import *
from CMRPUnit import units_to_internal
from CMRPParameters import (
    CMRP_UniqueId,
    CMRP_CP_LEN_Height,
    FREE_HEIGHT_TEMPLATE_NAME,
    ray3D_View,
    MissingTemplateError,
)
from CMRPSettings import MissingSettingsError

# Definitions and global variables
# =======================================================
logger = script.get_logger()


def execute_ray_trace(reference_intersector, location_point, room_id):
    """
    Raytraces from 'location_point' in the up and down z-axis direction in order to calculate the free height in a room.

    Parameters:
    ----------
    reference_intersector: ReferenceIntersector
        a reference intersector: a Revit Object that holds the relevant information from the Revit Document
        it will be used to tell when we're colliding with an object during raytracing
    location_point: XYZ
        The location point of a room
    room_id: ElementId
        Used for error logging

    Returns:
    -------
    double: the free height expressed as double in the internal Revit units
    """
    # Upwards Raytrace
    raytrace_distance_up = 0.0
    nearest_reference = reference_intersector.FindNearest(location_point, XYZ(0, 0, 1))
    if nearest_reference is not None:
        try:
            nearest_reference_z = nearest_reference.GetReference().GlobalPoint.Z
            raytrace_distance_up = abs(nearest_reference_z - location_point.Z)
        except:
            pass
    else:
        logger.warning(
            "{}: No reference found in the up direction for room ({}).".format(
                FREE_HEIGHT_TEMPLATE_NAME, room_id
            )
        )
        # TODO welke waarde nemen we hier en wat doen we er later mee
        raytrace_distance_up = units_to_internal(20000)

    # Downwards Raytrace
    raytrace_distance_down = 0.0
    nearest_reference = reference_intersector.FindNearest(location_point, XYZ(0, 0, -1))
    if nearest_reference is not None:
        try:
            nearest_reference_z = nearest_reference.GetReference().GlobalPoint.Z
            raytrace_distance_down = abs(location_point.Z - nearest_reference_z)
        except:
            pass
    else:
        logger.warning(
            "{}: No reference found in the down direction for room ({}).".format(
                FREE_HEIGHT_TEMPLATE_NAME, room_id
            )
        )
        # TODO welke waarde nemen we hier en wat doen we er later mee
        raytrace_distance_up = units_to_internal(1000)

    # Calculate total free height
    return raytrace_distance_up + raytrace_distance_down


def calculate_free_heights(document, room_dict, ray_3d_view):
    """
    Calculates the free height in rooms by iterating through all the rooms in the document and using ray tracing to calculate the free height.

    Parameters:
    ----------
    document: RevitDocument
        The document from which to retrieve data
    room_dict: dict{Id: [XYZ, Area]}
        a dictionary that maps the unique id of a room to its location point and corresponding area
    ray_3d_view: ViewPlan
        a 3D view plan used for ray tracing
    """

    # Tracing settings
    reference_intersector = ReferenceIntersector(ray_3d_view)
    reference_intersector.FindReferencesInRevitLinks = True

    # Start raytracing 1 meter above the floor
    Z_DIRECTION_OFFSET = 3.28084
    with Transaction(
        document,
        "CMRP {}: Calculate and write free height".format(FREE_HEIGHT_TEMPLATE_NAME),
    ) as transaction:
        transaction.Start()
        value = 0
        max_value = len(room_dict.keys())
        with ProgressBar(
            title="CMRP "
            + FREE_HEIGHT_TEMPLATE_NAME
            + ": Calculating free height in rooms: ({value} of {max_value})"
        ) as progress_bar:
            for unique_id, values in room_dict.items():
                location = values[0]
                area = values[1]
                # Translate each point in room_points so that it is above the floor
                free_height = execute_ray_trace(
                    reference_intersector,
                    location.Add(XYZ(0, 0, Z_DIRECTION_OFFSET)),
                    values[2],
                )
                if get_parameter_value(area, CMRP_UniqueId) == unique_id:
                    # Set the height parameter for the matching area
                    set_parameter_value(area, CMRP_CP_LEN_Height, free_height)
                value += 1
                progress_bar.update_progress(value, max_value)
        transaction.Commit()


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
    selected_window_families: str[]
        A list of window families that are selected by the user
    """
    levels = FilteredElementCollector(document).OfClass(Level).ToElements()
    imported_level_ids = get_imported_level_ids(levels, model_levels_grids)
    linked_rooms = get_rooms_of_imported_levels(imported_level_ids, model_rooms_windows)
    ray_3d_views = (
        FilteredElementCollector(document)
        .OfClass(View)
        .WhereElementIsNotElementType()
        .WherePasses(
            ElementParameterFilter(
                FilterStringRule(
                    ParameterValueProvider(ElementId(BuiltInParameter.VIEW_NAME)),
                    FilterStringEquals(),
                    ray3D_View,
                )
            )
        )
        .ToElements()
    )
    if ray_3d_views is None or len(ray_3d_views) == 0:
        raise MissingTemplateError(
            "{}: Couldn't find 3D raytracing view.".format(FREE_HEIGHT_TEMPLATE_NAME)
        )
    return levels, linked_rooms, ray_3d_views[0]


def check_necessary_settings(settings):
    if settings.model_levels_grids is None or settings.model_levels_grids == "":
        raise MissingSettingsError(
            "{}: The model for Levels & Grids is missing in the settings.".format(
                FREE_HEIGHT_TEMPLATE_NAME
            )
        )
    if settings.model_rooms_windows is None or settings.model_rooms_windows == "":
        raise MissingSettingsError(
            "{}: The model for Rooms & Windows is missing in the settings.".format(
                FREE_HEIGHT_TEMPLATE_NAME
            )
        )


# Main
# ==================================================


def run_freeheight_script(document, uidocument, settings):
    """
    This script calculates the free height in rooms in Autodesk Revit based on ray tracing.

    It performs the following tasks:
        0. Delete previously created areas and area plans.
        1. Retrieve data from the Revit document, including room information.
        2. Create area plans for viewing the area's.
        3. Iterate through rooms to create area's representing the boundary of the room.
        4. Raytrace up and down in the 'middle' of the room and calculates the free height.
        5. Fill in the height in a paramater in each area.
    """
    try:
        check_necessary_settings(settings)
        # 0. Delete previously created areas and area plans
        delete_old_areas_lines_and_areaplans(
            document, uidocument, FREE_HEIGHT_TEMPLATE_NAME
        )
        # Retrieve data from settings
        model_levels_grids = get_linked_document_from_name(
            document, settings.model_levels_grids
        )
        model_rooms_windows = get_linked_document_from_name(
            document, settings.model_rooms_windows
        )
        # 1. Retrieve data from the Revit document, including room information.
        levels, linked_rooms, ray_3d_view = retrieve_data_from_document(
            document, model_levels_grids, model_rooms_windows
        )
        # 2. Create area plans for viewing the area's.
        area_plans = create_area_plans(document, FREE_HEIGHT_TEMPLATE_NAME)
        # 3. Iterate through rooms to create area's representing the boundary of the room.
        rooms_dict = create_sketchplans_and_areas(
            document, linked_rooms, levels, area_plans, FREE_HEIGHT_TEMPLATE_NAME
        )
        # 4. Raytrace up and down in the 'middle' of the room and calculates the free height
        # 5. Fill in the height in a paramater in each area
        calculate_free_heights(document, rooms_dict, ray_3d_view)

        logger.success(
            "{}: Script completed without critical errors.".format(
                FREE_HEIGHT_TEMPLATE_NAME
            )
        )
    except (MissingSettingsError, MissingTemplateError) as error:
        logger.critical(error.message)
