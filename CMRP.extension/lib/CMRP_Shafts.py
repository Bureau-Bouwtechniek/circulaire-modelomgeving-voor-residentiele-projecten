# -*- coding: utf-8 -*-

# Imports
# =======================================================
import traceback

from pyrevit import script
from pyrevit.forms import ProgressBar

from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import CurveLoop

from CMRPCalculations import set_z_coordinate_of_line
from CMRP import (
    get_linked_document_from_name,
    delete_elements,
    delete_area_plans,
    create_area_plans,
)
from CMRPParameters import Comments, SHAFTS_TEMPLATE_NAME, MissingTemplateError
from CMRPSettings import MissingSettingsError
from CMRPUnit import units_to_internal

# Definitions and global variables
# =======================================================
logger = script.get_logger()


def delete_old_shafts(document):
    """
    Delete the shafts created by a previous execution of this script

    Parameters:
    ----------
    document: RevitDocument
        The document from which to delete the shafts
    """
    old_shaft_ids = (
        FilteredElementCollector(document).OfClass(DirectShape).ToElementIds()
    )
    delete_elements("Deleting old shafts", document, old_shaft_ids)


def expand_boundary(curve_loop, offset_distance, shaft_height):
    """
    Expand the boundary from the shaft by a given distance
    by moving each curve in the boundary in the outwards direction by the given distance,
    taking into account that the endpoints of touching curves still match.
    This function takes following steps in order:
        0. Calculate the direction of expansion of each curve in the boundary:
            the function calculates both directions and determines the correct one later on
        1. Expand the border by translating each endpoints by the combined vector of both curves that start/end at/in that point
        2. Extrude the 'expanded' boundaries to the shaft_height
        3. Compare the volume of the extrusion of the 'expanded' boundary of each direction to determine the correct one

    """
    vectors = []
    reverse_vectors = []
    # 0. Calculate the direction of expansion of each curve in the boundary.
    for curve in curve_loop:
        start_point = curve.GetEndPoint(0)
        end_point = curve.GetEndPoint(1)
        # Normal vector
        direction = (start_point - end_point).Normalize()
        offset_vector = direction.CrossProduct(XYZ.BasisZ) * offset_distance
        vectors.append(offset_vector)
        # Reverse vector
        direction = (end_point - start_point).Normalize()
        offset_vector = direction.CrossProduct(XYZ.BasisZ) * offset_distance
        reverse_vectors.append(offset_vector)

    expanded_solid = None
    try:
        expanded_curve_loop = CurveLoop()
        index = 0
        for curve in curve_loop:
            offset_startpoint = (
                curve.GetEndPoint(0)
                + vectors[index]
                + vectors[(index - 1) % len(vectors)]
            )
            offset_endpoint = (
                curve.GetEndPoint(1)
                + vectors[index]
                + vectors[(index + 1) % len(vectors)]
            )
            expanded_curve_loop.Append(
                Line.CreateBound(offset_startpoint, offset_endpoint)
            )
            index += 1
        expanded_solid = GeometryCreationUtilities.CreateExtrusionGeometry(
            [expanded_curve_loop], XYZ(0, 0, 1), shaft_height
        )
    except:
        pass
    reverse_expanded_solid = None
    try:
        reverse_expanded_curve_loop = CurveLoop()
        index = 0
        for curve in curve_loop:
            offset_startpoint = (
                curve.GetEndPoint(0)
                + reverse_vectors[index]
                + reverse_vectors[(index - 1) % len(reverse_vectors)]
            )
            offset_endpoint = (
                curve.GetEndPoint(1)
                + reverse_vectors[index]
                + reverse_vectors[(index + 1) % len(reverse_vectors)]
            )
            reverse_expanded_curve_loop.Append(
                Line.CreateBound(offset_startpoint, offset_endpoint)
            )
            index += 1
        reverse_expanded_solid = GeometryCreationUtilities.CreateExtrusionGeometry(
            [reverse_expanded_curve_loop], XYZ(0, 0, 1), shaft_height
        )
    except:
        pass
    if expanded_solid is None:
        return reverse_expanded_solid
    elif reverse_expanded_solid is None:
        return expanded_solid
    else:
        if expanded_solid.Volume > reverse_expanded_solid.Volume:
            return expanded_solid
        else:
            return reverse_expanded_solid


# Main
# =======================================================


def check_necessary_settings(settings):
    if settings.model_levels_grids is None or settings.model_levels_grids == "":
        raise MissingSettingsError(
            "{}: The model for Levels & Grids is missing in the settings.".format(
                SHAFTS_TEMPLATE_NAME
            )
        )
    if settings.model_shafts is None or settings.model_shafts == "":
        raise MissingSettingsError(
            "{}: The model for Shafts is missing in the settings.".format(
                SHAFTS_TEMPLATE_NAME
            )
        )


def run_shafts_script(document, uidocument, settings, from_shafts_button=True):
    """
    This script creates volumes the visually show where shafts are placed in the model.

    It performs the following tasks:
        0. Delete previously created shafts (DirectShapes).
        1. Retrieve shaft data from the Revit document.
        2. Create area plans for viewing the shafts.
        3. Create a volume for each shaft using it's 3D shape, base level and height.
    """
    try:
        check_necessary_settings(settings)
        # 0. Delete previously created shafts and area plans (DirectShapes)
        delete_old_shafts(document)
        if from_shafts_button:
            delete_area_plans(document, uidocument, SHAFTS_TEMPLATE_NAME[:3])
        # Retrieve data from settings
        model_levels_grids = get_linked_document_from_name(
            document, settings.model_levels_grids
        )
        model_shafts = get_linked_document_from_name(document, settings.model_shafts)
        # 1. Retrieve shaft data from the Revit document.
        shafts = (
            FilteredElementCollector(model_shafts)
            .OfCategory(BuiltInCategory.OST_ShaftOpening)
            .ToElements()
        )
        # 2. Create area plans for viewing the shafts.
        if from_shafts_button:
            area_plans = create_area_plans(document, SHAFTS_TEMPLATE_NAME)
        # Start a transaction to create DirectShape elements
        with Transaction(
            document, "CMRP Shafts: Create DirectShape Shaft Elements"
        ) as transaction:
            transaction.Start()
            value = 0
            max_value = len(shafts)
            with ProgressBar(
                title="CMRP Shafts: Creating new daylight zones: ({value} of {max_value})"
            ) as progress_bar:
                for shaft in shafts:
                    # Get the Sketch object using SketchId property
                    sketch = model_levels_grids.GetElement(shaft.SketchId)
                    if sketch is None:
                        logger.error(
                            "{}: Sketch ({}) for shaft ({}) not found.".format(
                                SHAFTS_TEMPLATE_NAME, shaft.SketchId, shaft.Id
                            )
                        )
                        continue
                    # Access the Profile property of the Sketch object
                    profile = sketch.Profile
                    level_id = shaft.LookupParameter("Base Constraint").AsElementId()
                    shaft_bottom_z = (
                        model_levels_grids.GetElement(level_id).Elevation
                        + shaft.LookupParameter("Base Offset").AsDouble()
                    )
                    shaft_height = shaft.LookupParameter(
                        "Unconnected Height"
                    ).AsDouble()
                    boundaries = []
                    result_solid = None
                    for curve_array in profile:
                        curve_loop = CurveLoop()
                        for curve in curve_array:
                            curve_loop.Append(
                                set_z_coordinate_of_line(curve, shaft_bottom_z)
                            )
                        boundaries.append(curve_loop)
                        if from_shafts_button:
                            expanded_extrusion_solid = expand_boundary(
                                curve_loop, units_to_internal(1000), shaft_height
                            )
                            if result_solid is None:
                                result_solid = expanded_extrusion_solid
                            else:
                                result_solid = (
                                    BooleanOperationsUtils.ExecuteBooleanOperation(
                                        result_solid,
                                        expanded_extrusion_solid,
                                        BooleanOperationsType.Union,
                                    )
                                )

                    # 3. Create a volume for each shaft using it's 3D shape, base level and height
                    # Create a solid by extruding the profile curves
                    extrusion_solid = GeometryCreationUtilities.CreateExtrusionGeometry(
                        boundaries, XYZ(0, 0, 1), shaft_height
                    )
                    if from_shafts_button:
                        intersection = BooleanOperationsUtils.ExecuteBooleanOperation(
                            result_solid,
                            extrusion_solid,
                            BooleanOperationsType.Difference,
                        )
                        # Create a DirectShape element from the extrusion solid
                        ds_category = (
                            BuiltInCategory.OST_GenericModel
                        )  # You can change this to another category if needed
                        direct_shape = DirectShape.CreateElement(
                            document, ElementId(ds_category)
                        )
                        direct_shape.SetShape([intersection])
                        direct_shape.Name = "Extruded Shaft"
                        direct_shape.LookupParameter(Comments).Set("CMRP_ExpandedShaft")
                    else:
                        # Create a DirectShape element from the extrusion solid
                        ds_category = (
                            BuiltInCategory.OST_GenericModel
                        )  # You can change this to another category if needed
                        direct_shape = DirectShape.CreateElement(
                            document, ElementId(ds_category)
                        )
                        direct_shape.SetShape([extrusion_solid])
                        direct_shape.Name = "Extruded Shaft"
                        direct_shape.LookupParameter(Comments).Set("CMRP_Shaft")
                    value += 1
                    progress_bar.update_progress(value, max_value)
            transaction.Commit()

        logger.success(
            "{}: Script completed without critical errors.".format(SHAFTS_TEMPLATE_NAME)
        )
    except (MissingSettingsError, MissingTemplateError) as error:
        logger.critical(error.message)
