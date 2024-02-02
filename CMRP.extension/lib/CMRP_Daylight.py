# -*- coding: utf-8 -*-

# Imports
# =======================================================
import math
import traceback

from pyrevit import revit, script
from pyrevit.forms import ProgressBar
from Autodesk.Revit.DB import *
import clr

clr.AddReference("RevitAPI")

from CMRPCalculations import *
from CMRPParameters import (
    CMRP_Daylight,
    CMRP_DaylightBoundary,
    DAYLIGHT_TEMPLATE_NAME,
    MissingTemplateError,
)
from CMRPSettings import MissingSettingsError
from CMRPUnit import units_to_internal
from CMRP import *

# Definitions and global variables
# =======================================================
logger = script.get_logger()


def get_largest_curveloop_from_room_boundary(room_boundary_segments):
    """
    Remove holes in the room boundary by selecting the loop with the largest surface area
    The room boundary can exist out of multiple curve loops
    for example when columns in the middle of the room create a 'donut' shaped room boundary
    This code can't handle boundaries shaped like a 'donut' and needs one continuous loop

    In order to keep it it simple the function searches for the curve loop with the largest surface area

    Parameters
    ----------
    room_boundary_segments: BoundarySegment[][]
        The room boundary as a nested array of boundary segments (which hold curves)

    Returns
    -------
    Curve[]
        The adapted room boundary as an array of curves
    """
    if len(room_boundary_segments) == 1:
        return [segment.GetCurve() for segment in room_boundary_segments[0]]
    result = None
    largest_area = 0.0
    for segments in room_boundary_segments:
        temp = []
        for segment in segments:
            temp.append(segment.GetCurve())
        surface_area = calculate_surface_area_of_boundary(temp)
        if surface_area > largest_area:
            result = temp
    return result


def adapt_room_boundary_to_useful_curveloop(room_boundary_segments):
    """
    Adapts the boundary of a room to CurveLoop that can be used later on
    The boundary is structured as an array within an array with Curves
    In order to make the boundary usefull we need to remove certain imperfections such as
     - holes created by columns
     - imperfections in corner due to misaligned walls or finishes
    The function takes follwing steps in order:
     - remove holes in the room boundary by selected the loop with the biggest surface area
     - remove imperfections in corners by searching for intersections within the boundary

    Parameters
    ----------
    room_boundary_segments: BoundarySegment[][]
        The room boundary as a nested array of boundary segments (which hold curves)

    Returns
    -------
    CurveLoop()
        The adapted room boundary
    """
    room_boundary = get_largest_curveloop_from_room_boundary(room_boundary_segments)
    if room_boundary is None:
        return None
    room_boundary = remove_imperfections_from_boundary(room_boundary)
    curve_loop = CurveLoop()
    for curve in room_boundary:
        curve_loop.Append(curve)
    return curve_loop


def get_correct_facing_orientation(window, room_boundary):
    """
    Finds the correct facing orientation by following process
     - get the 'FacingOrientation' vector from the window
     - create an opposite or reverse vector
     - now use both vectors to 'shoot' a line straight from the window with the length of 0,50 m
     - project both lines and the roomboundary to the XY plane through (0,0,0) to remove the Z coordinates
      -> this is to make sure all lines are in the same plane and can intersect
     - check which of the lines intersects with a curve from the room boundary
      -> return the correct orientation

    Parameters
    ----------
    window: Revit OST_Window
        The window from which to determine the correct orientation

    room_boundary: Curve[]
        The room boundary from the room adjacent to the window, found by the 'FromRoom' parameter

    Returns
    -------
    XYZ vector
        The correct orientation of the window
    """
    # Get data from window
    location_point = window.Location.Point
    facing_orientation = window.FacingOrientation
    reverse_orientation = XYZ(
        -facing_orientation.X, -facing_orientation.Y, -facing_orientation.Z
    )
    # Shoot lines with length 0,50 m using the facing and reverse orientation vector
    facing_line = set_z_coordinate_of_line(
        Line.CreateBound(
            location_point,
            location_point + facing_orientation * units_to_internal(1000),
        )
    )
    reverse_line = set_z_coordinate_of_line(
        Line.CreateBound(
            location_point,
            location_point + reverse_orientation * units_to_internal(1000),
        )
    )
    results = clr.Reference[IntersectionResultArray]()
    for curve in room_boundary:
        if (
            set_z_coordinate_of_line(curve).Intersect(reverse_line, results)
            == SetComparisonResult.Overlap
        ):
            return reverse_orientation, XYZ(
                results.Item[0].XYZPoint.X,
                results.Item[0].XYZPoint.Y,
                curve.GetEndPoint(0).Z,
            )
        elif (
            set_z_coordinate_of_line(curve).Intersect(facing_line, results)
            == SetComparisonResult.Overlap
        ):
            return facing_orientation, XYZ(
                results.Item[0].XYZPoint.X,
                results.Item[0].XYZPoint.Y,
                curve.GetEndPoint(0).Z,
            )
    logger.warning(
        "{}: Couldn't find a correct facing orientation for window: {}".format(
            DAYLIGHT_TEMPLATE_NAME, window.Id
        )
    )
    return None, None


def find_from_room(window, room_boundaries):
    for room_boundary in room_boundaries:
        (
            facing_orientation,
            intersection_point,
        ) = get_correct_facing_orientation(window, room_boundary)
        if facing_orientation is not None:
            return facing_orientation, intersection_point
    return None, None


def calculate_full_daylight_zone(
    window,
    facing_orientation,
    intersection_point,
    param_window_height,
    param_window_width,
):
    """
    Calculates the full daylight zone, create a trapezium:
     - use 45° angles from the edges from the window
     - use the height of the window as height(/depth) of the trapezium

    Parameters
    ----------
    window: Element
        The window of which to calculate the daylight zone
    room_boundary: Curve[]
        The boundary of the room in which the window is placed
    param_window_height: str
        The name of the Parameter where the height of the window is stored
    param_window_width: str
        The name of the Parameter where the width of the window is stored

    Returns
    -------
    CurveLoop
        The boundary of the daylightzone as CurveLoop
    """
    # Get information from window
    height, width = get_size_from_window(
        window, param_window_height, param_window_width
    )
    if height is None or width is None:
        return None
    location = window.Location.Point

    # Get the angle between the facing orientation and the south-pointing vector
    angle = (
        -math.copysign(1.0, facing_orientation.X) * math.acos(facing_orientation.Y)
        + math.pi
    )

    # Create the daylight zone at (0, 0, 0)
    p1 = XYZ(-width / 2.0, 0, 0)
    p2 = XYZ(width / 2.0, 0, 0)
    p3 = XYZ(height + width / 2.0, -height, 0)
    p4 = XYZ(-height - width / 2.0, -height, 0)

    # Move to window location point to the edge of the room boundary
    new_location = XYZ(intersection_point.X, intersection_point.Y, location.Z)
    # Translate/Move the daylight zone to the location point of the window and rotate it to the correct orientation
    p1 = rotate_z(p1, angle) + new_location
    p2 = rotate_z(p2, angle) + new_location
    p3 = rotate_z(p3, angle) + new_location
    p4 = rotate_z(p4, angle) + new_location

    # Create a new CurveLoop for the boundary
    daylight_zone = CurveLoop()

    # Create boundary lines based on window geometry
    daylight_zone.Append(Line.CreateBound(p1, p2))
    daylight_zone.Append(Line.CreateBound(p2, p3))
    daylight_zone.Append(Line.CreateBound(p3, p4))
    daylight_zone.Append(Line.CreateBound(p4, p1))

    return daylight_zone


def calculate_common_area(daylight_zone, room_boundary):
    """
    Find the common area of 2 boundaries in following steps:
    1. Create solids of both boundaries by extruding them in the z direction
    2. Calculate the intersection between both solids
    3. Convert the intersection solid back to a boundary

    Parameters
    ----------
    daylight_zone: CurveLoop
        The boundary of the daylight zone
    room_boundary: CurveLoop
        The boundary of a room

    Returns
    -------
    CurveLoop
        The boundary of the intersection
    """
    try:
        # Create Solid object from the daylight zone
        daylight_solid = GeometryCreationUtilities.CreateExtrusionGeometry(
            [daylight_zone], XYZ.BasisZ, 1
        )

        # Create Solid object from the room boundary
        room_solid = None
        if room_boundary:
            room_solid = GeometryCreationUtilities.CreateExtrusionGeometry(
                [room_boundary], XYZ.BasisZ, 100
            )

        # Calculate the common area or intersection of the daylight zone and the room surface area
        # This will result in the daylight zone not going through walls outside of the room boundary
        if room_solid:
            intersection = BooleanOperationsUtils.ExecuteBooleanOperation(
                daylight_solid, room_solid, BooleanOperationsType.Intersect
            )
            if intersection:
                intersection_boundary = get_boundary_from_solid(intersection)
                return intersection_boundary
    except Exception:
        logger.error("{}: {}".format(DAYLIGHT_TEMPLATE_NAME, traceback.format_exc()))
    return None


def get_room_boundaries_of_level(model_rooms_windows, level_id, boundary_options):
    """
    Creates a dictionary of room id's and the room boundaries for a certain level

    Parameters:
    ----------
    model_rooms_windows: RevitDocument
        The document frow which the function retrieves its data
    level_id: ElementId
        The id of the level in question
    boundary_options: SpatialElementBoundaryOptions
        Determines which boundary we retrieve from the room

    Returns:
    -------
    dict: {room id: room boundary}: a dictionary with the room id as key and room boundary as value
    """
    rooms = (
        FilteredElementCollector(model_rooms_windows)
        .OfCategory(BuiltInCategory.OST_Rooms)
        .WhereElementIsNotElementType()
        .WherePasses(
            LogicalAndFilter(
                ElementParameterFilter(
                    FilterElementIdRule(
                        ParameterValueProvider(
                            ElementId(BuiltInParameter.ROOM_LEVEL_ID)
                        ),
                        FilterNumericEquals(),
                        level_id,
                    )
                ),
                ElementParameterFilter(
                    FilterInverseRule(
                        FilterStringRule(
                            ParameterValueProvider(
                                ElementId(BuiltInParameter.ROOM_NAME)
                            ),
                            FilterStringBeginsWith(),
                            "Terrasse",
                        )
                    )
                ),
            )
        )
        .ToElements()
    )
    room_boundary_dict = {}
    for from_room in rooms:
        # Get Room Boundary
        room_boundary_segments = from_room.GetBoundarySegments(boundary_options)
        if room_boundary_segments is None:
            logger.warning(
                "{}: Room ({}) has no BoundarySegments.".format(
                    DAYLIGHT_TEMPLATE_NAME, from_room.Id
                )
            )
            continue
        room_boundary = adapt_room_boundary_to_useful_curveloop(room_boundary_segments)
        if room_boundary is None:
            logger.warning(
                "{}: Room ({}) has no useful boundary.".format(
                    DAYLIGHT_TEMPLATE_NAME, from_room.Id
                )
            )
            continue
        room_boundary_dict[from_room.Id] = room_boundary
    return room_boundary_dict


def map_windows_to_rooms(
    imported_level_ids, model_rooms_windows, selected_window_families
):
    boundary_options = SpatialElementBoundaryOptions()
    boundary_options.SpatialElementBoundaryLocation = (
        SpatialElementBoundaryLocation.Finish
    )
    level_room_dict = {}
    window_room_dict = {}
    value = 0
    max_value = len(imported_level_ids)
    with ProgressBar(
        title="CMRP "
        + DAYLIGHT_TEMPLATE_NAME
        + ": Mapping windows to their room: ({value} of {max_value})"
    ) as progress_bar:
        for level_id in imported_level_ids:
            windows = (
                FilteredElementCollector(model_rooms_windows)
                .OfCategory(BuiltInCategory.OST_Windows)
                .WhereElementIsNotElementType()
                .WherePasses(ElementLevelFilter(level_id))
                .ToElements()
            )
            room_boundary_dict = get_room_boundaries_of_level(
                model_rooms_windows, level_id, boundary_options
            )

            for window in windows:
                family_name = window.Symbol.Family.Name
                type_name = window.Symbol.LookupParameter("Type Name").AsString()
                from_room_boundary = None
                if (
                    family_name not in selected_window_families.keys()
                    or type_name not in selected_window_families[family_name]
                ):
                    continue
                from_room = get_room_of_window(model_rooms_windows, window)
                if from_room is None:
                    # Get data from window
                    location_point = window.Location.Point
                    facing_orientation = window.FacingOrientation
                    reverse_orientation = XYZ(
                        -facing_orientation.X,
                        -facing_orientation.Y,
                        -facing_orientation.Z,
                    )
                    rooms = (
                        FilteredElementCollector(model_rooms_windows)
                        .OfCategory(BuiltInCategory.OST_Rooms)
                        .WhereElementIsNotElementType()
                        .WherePasses(
                            LogicalAndFilter(
                                ElementLevelFilter(level_id),
                                ElementParameterFilter(
                                    FilterInverseRule(
                                        FilterStringRule(
                                            ParameterValueProvider(
                                                ElementId(BuiltInParameter.ROOM_NAME)
                                            ),
                                            FilterStringBeginsWith(),
                                            "Terrasse",
                                        )
                                    )
                                ),
                            )
                        )
                        .WherePasses(
                            BoundingBoxContainsPointFilter(
                                location_point
                                + facing_orientation * units_to_internal(500),
                                units_to_internal(200),
                            )
                        )
                        .ToElements()
                    )
                    rooms = [room for room in rooms if room.Location is not None]
                    # TODO map "terasse"
                    if rooms is None or len(rooms) == 0:
                        rooms = (
                            FilteredElementCollector(model_rooms_windows)
                            .OfCategory(BuiltInCategory.OST_Rooms)
                            .WhereElementIsNotElementType()
                            .WherePasses(
                                LogicalAndFilter(
                                    ElementLevelFilter(level_id),
                                    ElementParameterFilter(
                                        FilterInverseRule(
                                            FilterStringRule(
                                                ParameterValueProvider(
                                                    ElementId(
                                                        BuiltInParameter.ROOM_NAME
                                                    )
                                                ),
                                                FilterStringBeginsWith(),
                                                "Terrasse",
                                            )
                                        )
                                    ),
                                )
                            )
                            .WherePasses(
                                BoundingBoxContainsPointFilter(
                                    location_point
                                    + reverse_orientation * units_to_internal(500),
                                    units_to_internal(200),
                                )
                            )
                            .ToElements()
                        )
                        rooms = [room for room in rooms if room.Location is not None]
                    if rooms is None or len(rooms) == 0:
                        continue
                    else:
                        from_room = rooms[0]
                try:
                    from_room_boundary = room_boundary_dict[from_room.Id]
                except:
                    for dict in level_room_dict.values():
                        try:
                            from_room_boundary = dict[from_room.Id]
                            break
                        except:
                            pass
                if from_room_boundary is not None:
                    window_room_dict[window] = from_room_boundary
                else:
                    logger.warning(
                        "{}: Window ({}) has no 'FromRoom'.".format(
                            DAYLIGHT_TEMPLATE_NAME, window.Id
                        )
                    )
            level_room_dict[level_id] = room_boundary_dict

            value += 1
            progress_bar.update_progress(value, max_value)
    return level_room_dict, window_room_dict


def retrieve_data_from_document(
    document, model_levels_grids, model_rooms_windows, selected_window_families
):
    """
    Retrieve data from the Revit (linked) document

    Logs errors for values the script absolutly needs -> throws exit code 1

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

    # Get FilledRegionType Id
    filled_region_types = (
        FilteredElementCollector(document)
        .OfClass(FilledRegionType)
        .WhereElementIsElementType()
        .WherePasses(
            ElementParameterFilter(
                FilterStringRule(
                    ParameterValueProvider(
                        ElementId(BuiltInParameter.SYMBOL_NAME_PARAM)
                    ),
                    FilterStringEquals(),
                    CMRP_Daylight,
                )
            )
        )
        .ToElements()
    )
    if filled_region_types is None or len(filled_region_types) == 0:
        raise MissingTemplateError(
            "{}: Can't find FilledRegionType '{}'".format(
                DAYLIGHT_TEMPLATE_NAME, CMRP_Daylight
            )
        )
    filled_region_type_id = filled_region_types[0].Id

    # Get LineStyle Id
    line_styles = [
        style
        for style in FilteredElementCollector(document)
        .OfClass(GraphicsStyle)
        .ToElements()
        if style.GraphicsStyleCategory.Name == CMRP_DaylightBoundary
    ]
    if line_styles is None or len(line_styles) == 0:
        raise MissingTemplateError(
            "{}: Can't find LineStyle '{}'".format(
                DAYLIGHT_TEMPLATE_NAME, CMRP_DaylightBoundary
            )
        )

    line_style_id = line_styles[0].Id

    # Create dictionary of levels and room boundaries
    levels = FilteredElementCollector(document).OfClass(Level).ToElements()
    imported_level_ids = get_imported_level_ids(levels, model_levels_grids)

    level_room_dict, window_room_dict = map_windows_to_rooms(
        imported_level_ids, model_rooms_windows, selected_window_families
    )

    return filled_region_type_id, line_style_id, level_room_dict, window_room_dict


def check_necessary_settings(settings):
    if settings.model_levels_grids is None or settings.model_levels_grids == "":
        raise MissingSettingsError(
            "{}: The model for Levels & Grids is missing in the settings.".format(
                DAYLIGHT_TEMPLATE_NAME
            )
        )
    if settings.model_rooms_windows is None or settings.model_rooms_windows == "":
        raise MissingSettingsError(
            "{}: The model for Rooms & Windows is missing in the settings.".format(
                DAYLIGHT_TEMPLATE_NAME
            )
        )
    if settings.param_window_height is None or settings.param_window_height == "":
        raise MissingSettingsError(
            "{}: The parameter for the height of windows is missing in the settings.".format(
                DAYLIGHT_TEMPLATE_NAME
            )
        )
    if settings.param_window_width is None or settings.param_window_height == "":
        raise MissingSettingsError(
            "{}: The parameter for the height of windows is missing in the settings.".format(
                DAYLIGHT_TEMPLATE_NAME
            )
        )
    if (
        settings.selected_window_families is None
        or len(settings.selected_window_families.keys()) == 0
    ):
        raise MissingSettingsError(
            "{}: No window families or type were selected, check in the settings.".format(
                DAYLIGHT_TEMPLATE_NAME
            )
        )


# Main
# ==================================================


def run_daylight_script(document, uidocument, settings):
    """
    This script creates daylight zones in Autodesk Revit based on window configurations.

    It performs the following tasks:
        0. Delete previously created daylight zones and area plans.
        1. Retrieve data from the Revit document, including window information and relevant elements.
        2. Create area plans for viewing the daylight zones.
        3. Iterate through windows and their corresponding rooms to calculate and create daylight zones.
        4. Iterate through the rooms on the same level of each window to split the daylight zones per room.
        5. Create filled regions to represent daylight zones.
    """
    try:
        check_necessary_settings(settings)
        # 0. Delete previously created daylight zones and area plans
        delete_area_plans(document, uidocument, DAYLIGHT_TEMPLATE_NAME)
        # Retrieve data from settings
        model_levels_grids = get_linked_document_from_name(
            document, settings.model_levels_grids
        )
        model_rooms_windows = get_linked_document_from_name(
            document, settings.model_rooms_windows
        )
        param_window_height = settings.param_window_height
        param_window_width = settings.param_window_width
        # 1. Retrieve data from the Revit document, including window information and relevant elements.
        (
            filled_region_type_id,
            line_style_id,
            level_room_dict,
            window_room_dict,
        ) = retrieve_data_from_document(
            document,
            model_levels_grids,
            model_rooms_windows,
            settings.selected_window_families,
        )

        # 2. Create area plans for viewing the daylight zones.
        views = create_area_plans(document, DAYLIGHT_TEMPLATE_NAME)
        # 3. Iterate through windows and their corresponding rooms to calculate and create daylight zones.
        with Transaction(
            document,
            "CMRP {}: Create new daylight zones".format(DAYLIGHT_TEMPLATE_NAME),
        ) as transaction:
            transaction.Start()
            value = 0
            max_value = len(window_room_dict)
            with ProgressBar(
                title="CMRP "
                + DAYLIGHT_TEMPLATE_NAME
                + ": Creating new daylight zones: ({value} of {max_value})"
            ) as progress_bar:
                for window in window_room_dict.keys():
                    room_boundaries = level_room_dict[window.LevelId]

                    from_room_boundary = window_room_dict[window]
                    facing_orientation, intersection_point = None, None
                    if from_room_boundary is not None:
                        (
                            facing_orientation,
                            intersection_point,
                        ) = get_correct_facing_orientation(window, from_room_boundary)
                    else:
                        (
                            facing_orientation,
                            intersection_point,
                        ) = find_from_room(window, room_boundaries.values())
                    if facing_orientation is None:
                        value += 1
                        progress_bar.update_progress(value, max_value)
                        continue

                    # Create the full trapezium as daylight zone

                    full_daylight_zone = calculate_full_daylight_zone(
                        window,
                        facing_orientation,
                        intersection_point,
                        param_window_height,
                        param_window_width,
                    )
                    if full_daylight_zone is None:
                        value += 1
                        progress_bar.update_progress(value, max_value)
                        continue
                    current_view = find_view_by_level_elevation(
                        views, model_levels_grids.GetElement(window.LevelId).Elevation
                    )
                    # 4. Iterate through the rooms on the same level of each window to split the daylight zones per room.
                    for room_boundary in room_boundaries.values():
                        # Find the common area between the daylight zone and the boundary of each room
                        daylight_zone = calculate_common_area(
                            full_daylight_zone, room_boundary
                        )
                        if daylight_zone is None:
                            continue

                        # 5. Create filled regions to represent daylight zones.
                        filled_region = FilledRegion.Create(
                            document,
                            filled_region_type_id,
                            current_view.Id,
                            daylight_zone,
                        )
                        filled_region.SetLineStyleId(line_style_id)

                    value += 1
                    progress_bar.update_progress(value, max_value)
            transaction.Commit()

        logger.success(
            "{}: Script completed without critical errors.".format(
                DAYLIGHT_TEMPLATE_NAME
            )
        )
    except (MissingSettingsError, MissingTemplateError) as error:
        logger.critical(error.message)
