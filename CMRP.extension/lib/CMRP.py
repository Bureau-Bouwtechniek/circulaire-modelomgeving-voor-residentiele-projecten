# -*- coding: utf-8 -*-

# Imports
# =======================================================
import sys

from pyrevit import revit, script
from pyrevit.forms import ProgressBar
import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import *
import Autodesk.Revit.Exceptions as RevitExceptions

from CMRPParameters import CMRP_UniqueId

# Definitions and global variables
# =======================================================
logger = script.get_logger()

# Document Functions
# =======================================================


def get_linked_document_from_name(document, link_name):
    linked_instances = (
        FilteredElementCollector(document).OfClass(RevitLinkInstance).ToElements()
    )
    for link in linked_instances:
        if link.Name == link_name:
            return link.GetLinkDocument()
    return None


def get_parameter_value(element, param_name):
    param = element.LookupParameter(param_name)
    if param:
        return param.AsString()
    return None


def set_parameter_value(element, param_name, value):
    param = element.LookupParameter(param_name)
    if param:
        param.Set(value)


# Level Functions
# =======================================================


def get_level_from_linked_level(levels, linked_level):
    """
    Match the level of the linked document with a level of the current document.
    """
    for level in levels:
        if level.Elevation == linked_level.Elevation:
            return level
    return None


def get_imported_level_ids(levels, model_levels_grids):
    """
    Get a list of ids of levels in 'model_levels_grids' matching the levels in the current document.
    """
    level_elevations = [level.Elevation for level in levels]
    linked_levels = (
        FilteredElementCollector(model_levels_grids).OfClass(Level).ToElements()
    )
    imported_level_ids = []
    for linked_level in linked_levels:
        if linked_level.Elevation in level_elevations:
            imported_level_ids.append(linked_level.Id)
    return imported_level_ids


# View / AreaPlan Functions
# =======================================================


def find_view_by_level_elevation(views, elevation):
    for view in views:
        if view.ViewType == ViewType.AreaPlan:
            if view.GenLevel.Elevation == elevation:
                return view
    return None


def find_view_by_level_id(views, level_id):
    for view in views:
        if view.ViewType == ViewType.AreaPlan:
            if view.GenLevel.Id == level_id:
                return view
    return None


def get_area_plans_by_template_code(document, template_code):
    """
    Retrieves the area plans of which the name starts with 'template_code'
    and it filters out those plans that are viewtemplates.

    Parameters:
    ----------
    document: RevitDocument
        The document from which to retrieve data
    template_code: str
        The string to compare the name of the views with

    Returns:
    -------
    View[]: An array of Views/AreaPlans that match the search criteria.
    """
    views = (
        FilteredElementCollector(document)
        .OfClass(View)
        .WhereElementIsNotElementType()
        .WherePasses(
            ElementParameterFilter(
                FilterStringRule(
                    ParameterValueProvider(ElementId(BuiltInParameter.VIEW_NAME)),
                    FilterStringBeginsWith(),
                    template_code,
                )
            )
        )
        .ToElements()
    )
    area_plans = []
    for view in views:
        if not view.IsTemplate and view.ViewType == ViewType.AreaPlan:
            area_plans.append(view)
    return area_plans


def get_area_scheme_id(document, area_scheme_name):
    """
    Retrieves a specific AreaScheme from the model based on its name.

    Parameters:
    ----------
    document: RevitDocument
        The document from which to retrieve data.
    areascheme_name: str
        The name of the AreaScheme to retrieve.

    Returns:
    -------
    ElementId or None: The ElementId of the AreaScheme if found, or None if not found.
    """
    area_scheme = (
        FilteredElementCollector(document)
        .OfClass(AreaScheme)
        .WhereElementIsNotElementType()
        .WherePasses(
            ElementParameterFilter(
                FilterStringRule(
                    ParameterValueProvider(
                        ElementId(BuiltInParameter.AREA_SCHEME_NAME)
                    ),
                    FilterStringEquals(),
                    area_scheme_name,
                )
            )
        )
        .FirstElement()
    )
    if area_scheme:
        return area_scheme.Id
    else:
        return None


def create_area_plans(document, scheme_name, template_name=None):
    """
    Creates Area Plans based on a scheme and an optional template.

    Parameters:
    ----------
    document: RevitDocument
        The document in which to create areaplans.
    scheme_name: str
        The name of the Area Scheme to use.
    template_name: str
        default: None
        The name of the template to be applied to the created views

    Note:
        If template_name is not provided, it defaults to scheme_name.
    """
    if template_name is None:
        template_name = scheme_name

    # Get the first 3 characters of the scheme_name as the template code
    template_code = scheme_name[:3]

    # Get the Area Scheme Id
    area_scheme_id = get_area_scheme_id(document, scheme_name)
    if area_scheme_id == None:
        logger.error(
            "{}: The area scheme '{}' was not found".format(
                template_name[3:], scheme_name
            )
        )

    # Get views and templates from model
    template = (
        FilteredElementCollector(document)
        .OfClass(View)
        .WhereElementIsNotElementType()
        .WherePasses(
            ElementParameterFilter(
                FilterStringRule(
                    ParameterValueProvider(ElementId(BuiltInParameter.VIEW_NAME)),
                    FilterStringEquals(),
                    template_name,
                )
            )
        )
        .FirstElement()
    )
    # Find the template with the specified name
    if template is None:
        logger.critical(
            "{}: The template '{}' was not found".format(
                template_name[3:], template_name
            )
        )
        raise

    area_plans = get_area_plans_by_template_code(document, template_code)
    area_plan_names = [area_plan.Name for area_plan in area_plans]

    # Get levels in model
    levels = FilteredElementCollector(document).OfClass(Level).ToElements()

    # Create Area Plans
    with Transaction(
        document, "CMRP {}: Create Areaplans: ".format(template_name)
    ) as transaction:
        transaction.Start()
        value = 0
        max_value = len(levels)
        with ProgressBar(
            title="CMRP "
            + template_name
            + ": Creating new area plans: ({value} of {max_value})"
        ) as progress_bar:
            for level in levels:
                if template_code + level.Name in area_plan_names:
                    continue
                else:
                    area_plan = ViewPlan.CreateAreaPlan(
                        document, area_scheme_id, level.Id
                    )
                    area_plan.Name = template_code + level.Name
                    area_plan.ViewTemplateId = template.Id
                    area_plans.append(area_plan)

                value += 1
                progress_bar.update_progress(value, max_value)
        transaction.Commit()
    return area_plans


def delete_elements(transaction_name, document, list_of_elements):
    """
    Deletes elements in 'document' and shows the progress using ProgressBar.

    Parameters:
    ----------
    transaction_name: str
        The name for the transaction and title for the progressbar.
    document: RevitDocument
        The current document from which to delete certain elements.
    list_of_elements: Element[] ot ElementId[]
        A list of elements or elementids, the function can handle both.
    """
    with Transaction(document, transaction_name) as transaction:
        transaction.Start()
        value = 0
        max_value = len(list_of_elements)
        with ProgressBar(
            title=transaction_name + ": ({value} of {max_value})"
        ) as progress_bar:
            for item in list_of_elements:
                if isinstance(item, ElementId):
                    document.Delete(item)
                else:
                    document.Delete(item.Id)
                value += 1
                progress_bar.update_progress(value, max_value)
        transaction.Commit()


def delete_area_plans(document, uidocument, template_name):
    """
    Deletes area plans of which the name starts with 'template_code'.
    In order for all functions to create area's, the area plans and elements based on the plans need to be removed and recreated.

    Parameters:
    ----------
    document: RevitDocument
        The current document from which to delete area plans.
    uidocument: UIDocument
        The current uidocument used to check active views.
    template_name:
        The name of the template for the area plans.
    """
    area_plans = get_area_plans_by_template_code(document, template_name[:3])
    with Transaction(
        document, "CMRP {}: Deleting old area plans".format(template_name)
    ) as transaction:
        transaction.Start()
        value = 0
        max_value = len(area_plans)
        with ProgressBar(
            title="CMRP "
            + template_name
            + ": Deleting old area plans: ({value} of {max_value})"
        ) as progress_bar:
            for area_plan in area_plans:
                if area_plan.Id == uidocument.ActiveView.Id:
                    logger.error(
                        "CMRP {}: Please close the view: {}".format(
                            template_name, area_plan.Name
                        )
                    )
                else:
                    document.Delete(area_plan.Id)
                value += 1
                progress_bar.update_progress(value, max_value)
        transaction.Commit()


# Area Functions
# =======================================================


def create_area_for_room(document, room, area_plan, sketch_plane):
    """
    Create area using the room boundary.

    Parameters:
    ----------
    document: RevitDocument
        The current document from which to delete area plans.
    room: Room
        The room for which to create the area.
    area_plan: ViewPlan
        The view plan on which to create the area.
    skecth_plane: SketchPlane
        The plane on which to create lines for the boundary of the area.

    Returns:
    -------
    Area
        The created area.
    """
    options = SpatialElementBoundaryOptions()
    options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish

    room_boundary = []
    room_boundary_segments = room.GetBoundarySegments(options)
    for segment in room_boundary_segments:
        temp = []
        for curve in segment:
            temp.append(curve.GetCurve())
        room_boundary += remove_imperfections_from_boundary(temp)

    # Create area boundary
    for curve in room_boundary:
        document.Create.NewAreaBoundaryLine(sketch_plane, curve, area_plan)
    uv_point = UV(room.Location.Point.X, room.Location.Point.Y)
    # Create area
    area = document.Create.NewArea(area_plan, uv_point)
    area.Name = Element.Name.GetValue(room)
    area.LookupParameter(CMRP_UniqueId).Set(str(room.UniqueId))
    return area


def create_sketchplans_and_areas(
    document, linked_rooms, levels, area_plans, template_name
):
    """
    Create sketchplanes for each level.
    Create areas for each room using its boundary.
    Create a dictionary:
        key: uniqueId of the room
        values: a list of:
            - locationpoint of the room
            - area representing the room
            - id of the room

    Parameters:
    ----------
    document: RevitDocument
        The current document from which to delete area plans.
    linked_rooms: Room[]
        An array of rooms from the linked document.
    levels: Level[]
        An array of levels of the current document.
    area_plans: ViewPlan[]
        An array of viewplans for each level.
    template_name: str
        A string representing the name of the template for the current running function.

    Returns:
    -------
    dictionary
    """
    rooms_dict = {}
    with Transaction(
        document, "CMRP {}: Creating areas for each room".format(template_name)
    ) as transaction:
        transaction.Start()
        sketch_planes = {}
        for level in levels:
            sketch_plane = SketchPlane.Create(document, level.Id)
            sketch_plane.Name = template_name[:3] + level.Name
            sketch_planes[level.Id] = sketch_plane
        value = 0
        max_value = len(linked_rooms)
        with ProgressBar(
            title="CMRP "
            + template_name
            + ": Creating areas for each room: ({value} of {max_value})"
        ) as progress_bar:
            # Iterate over each room
            for linked_room in linked_rooms:
                if not isinstance(linked_room.Location, LocationPoint):
                    logger.error(
                        "{}: The location of room {} is not a LocationPoint object".format(
                            template_name, linked_room.Id
                        )
                    )
                    continue
                # Get Level
                level = get_level_from_linked_level(levels, linked_room.Level)
                if level is None:
                    continue
                # Get ViewPlan²
                area_plan = find_view_by_level_id(area_plans, level.Id)
                if area_plan is None:
                    logger.error(
                        "{}: Couldn't find view for level with id: {}".format(
                            template_name, level.Id
                        )
                    )
                    continue
                # Get SketchPlane
                sketch_plane = sketch_planes[level.Id]
                # Create Area
                area = create_area_for_room(
                    document, linked_room, area_plan, sketch_plane
                )
                if area is None:
                    continue
                rooms_dict[linked_room.UniqueId] = [
                    linked_room.Location.Point,
                    area,
                    linked_room.Id,
                ]
                value += 1
                progress_bar.update_progress(value, max_value)
        transaction.Commit()
    return rooms_dict


def delete_old_areas_lines_and_areaplans(document, uidocument, template_name):
    # Delete old areas
    areas = (
        FilteredElementCollector(document)
        .OfCategory(BuiltInCategory.OST_Areas)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    filtered_areas = []
    for area in areas:
        if area.AreaScheme and area.AreaScheme.Name.startswith(template_name[:3]):
            filtered_areas.append(area)
    delete_elements(
        "CMRP {}: Delete areas".format(template_name), document, filtered_areas
    )
    # Delete old area plans (views), lines and sketchplanes
    area_plans = get_area_plans_by_template_code(document, template_name[:3])
    with Transaction(
        document, "CMRP {}: Delete old areaplans and lines".format(template_name[3:])
    ) as transaction:
        transaction.Start()
        value = 0
        max_value = len(area_plans)
        with ProgressBar(
            title="CMRP "
            + template_name
            + ": Deleting Lines, SketchPlanes and AreaPlans: ({value} of {max_value})"
        ) as progress_bar:
            for area_plan in area_plans:
                line_ids = (
                    FilteredElementCollector(document, area_plan.Id)
                    .OfClass(CurveElement)
                    .WhereElementIsNotElementType()
                    .ToElementIds()
                )
                sketch_plane_ids = (
                    FilteredElementCollector(document, area_plan.Id)
                    .OfClass(SketchPlane)
                    .WhereElementIsNotElementType()
                    .WherePasses(ElementLevelFilter(area_plan.GenLevel.Id))
                    .ToElementIds()
                )
                document.Delete(line_ids)
                document.Delete(sketch_plane_ids)
                if area_plan.Id == uidocument.ActiveView.Id:
                    logger.error(
                        "{}: Please close the view '{}', it can't be deleted/updated while opened.".format(
                            template_name, area_plan.Name
                        )
                    )
                else:
                    document.Delete(area_plan.Id)
                value += 1
                progress_bar.update_progress(value, max_value)
        transaction.Commit()


# Room Functions
# =======================================================


def get_room_of_window(linked_document, window):
    room = None
    for phase in linked_document.Phases:
        try:
            room = window.FromRoom[phase]
        except:
            continue
    return room


def get_rooms_of_imported_levels(imported_level_ids, model_rooms_windows):
    linked_rooms = []
    for level_id in imported_level_ids:
        rooms = (
            FilteredElementCollector(model_rooms_windows)
            .OfCategory(BuiltInCategory.OST_Rooms)
            .WhereElementIsNotElementType()
            .WherePasses(ElementLevelFilter(level_id))
            .ToElements()
        )
        for room in rooms:
            if room.Location is not None:
                linked_rooms.append(room)
    return linked_rooms


# Room Boundary Functions
# =======================================================


def calculate_surface_area_of_boundary(boundary):
    """
    Calculates the surface area of a boundary

    Parameters
    ----------
    boundary: Curve[]
        A continuous loop of curves from which to calculate the surface area

    Returns
    -------
    double
        The surface area
    """
    surface_area = 0.0
    for curve in boundary:
        if curve.IsBound:
            surface_area += (
                curve.Length * curve.GetEndParameter(0) * curve.GetEndParameter(1)
            )
    return abs(surface_area) / 2.0


def remove_single_imperfection_from_boundary(room_boundary, nr_of_recursive=0):
    results = clr.Reference[IntersectionResultArray]()
    for current_index in range(len(room_boundary)):
        for intersect_index in range(current_index + 2, len(room_boundary)):
            if current_index == 0 and intersect_index == len(room_boundary) - 1:
                continue  # Skip the first and last curves

            intersect = room_boundary[current_index].Intersect(
                room_boundary[intersect_index], results
            )

            if intersect != SetComparisonResult.Overlap:
                continue
            if current_index == 0 and nr_of_recursive < 4:
                # The first curve is intersecting another curve in the boundary
                # in order to handle this specific situation we move the curves in the array
                # one space to the left (the first curve becomes the last)
                # and try this function again from the start
                return remove_single_imperfection_from_boundary(
                    room_boundary[1:] + [room_boundary[0]], nr_of_recursive + 1
                )
            intersection_point = results.Item[0].XYZPoint

            # Situation 1: Shorten the intersecting curves to the intersection point
            new_curves_situation_1 = [
                Line.CreateBound(
                    room_boundary[current_index].GetEndPoint(0),
                    intersection_point,
                ),
                Line.CreateBound(
                    intersection_point,
                    room_boundary[intersect_index].GetEndPoint(1),
                ),
            ]

            boundary_situation_1 = (
                room_boundary[0:current_index]
                + new_curves_situation_1
                + room_boundary[intersect_index + 1 :]
            )
            try:
                # Situation 2: Shorten the intersecting curves to the intersection point
                new_curves_situation_2 = [
                    Line.CreateBound(
                        room_boundary[intersect_index].GetEndPoint(0),
                        intersection_point,
                    ),
                    Line.CreateBound(
                        intersection_point,
                        room_boundary[current_index].GetEndPoint(1),
                    ),
                ]

                boundary_situation_2 = (
                    room_boundary[current_index + 1 : intersect_index]
                    + new_curves_situation_2
                )

                # Calculate the surface area for both situations
                # Choose the situation with the larger surface area
                if calculate_surface_area_of_boundary(
                    boundary_situation_1
                ) > calculate_surface_area_of_boundary(boundary_situation_2):
                    return boundary_situation_1
                else:
                    return boundary_situation_2
            except RevitExceptions.ArgumentsInconsistentException:
                return boundary_situation_1
    return room_boundary


def remove_imperfections_from_boundary(room_boundary, max_iterations=10):
    for _ in range(max_iterations):
        new_room_boundary = remove_single_imperfection_from_boundary(room_boundary)
        if len(new_room_boundary) == len(room_boundary):
            return new_room_boundary
        room_boundary = new_room_boundary

    return room_boundary


def get_size_from_window(window, param_window_height, param_window_width):
    param_height = window.LookupParameter(param_window_height)
    if param_height is None:
        param_height = window.Symbol.LookupParameter(param_window_height)
    param_width = window.LookupParameter(param_window_width)
    if param_width is None:
        param_width = window.Symbol.LookupParameter(param_window_width)
    try:
        height = param_height.AsDouble()
        width = param_width.AsDouble()
        if height == 0 or width == 0:
            logger.error(
                "Window Types: Height or width of window ({}) is zero.".format(
                    window.Id
                )
            )
            return None, None
        return height, width
    except AttributeError:
        logger.error(
            "Window Types: Couldn't get height or width from window ({}).".format(
                window.Id
            )
        )
        return None, None
