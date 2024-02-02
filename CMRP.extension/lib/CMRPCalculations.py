# -*- coding: utf-8 -*-

# Imports
# =======================================================
from pyrevit import revit
from Autodesk.Revit.DB import *


# Definitions and global variables
# =======================================================
def set_z_coordinate_of_line(line, new_z_coordinate=0):
    """
    This function projects a line/curve onto the XY plane with new_z_coordinate as elevation
    by replacing the z-coordinates from the endpoints of the line/curve

    Parameters
    ----------
    line: Line or Curve
        The line/curve to adapt
    new_z_coordinate: double
        default: 0
        The new z coordinate of both endpoints of the line

    Returns
    -------
    Line or Curve
        The adapted line/curve
    """
    start = line.GetEndPoint(0)
    end = line.GetEndPoint(1)
    return Line.CreateBound(
        XYZ(start.X, start.Y, new_z_coordinate), XYZ(end.X, end.Y, new_z_coordinate)
    )


def set_z_coordinate_of_boundary(boundary, new_z_coordinate=0):
    if isinstance(boundary, CurveLoop):
        new_boundary = CurveLoop()
        for curve in boundary:
            new_boundary.Append(set_z_coordinate_of_line(curve, new_z_coordinate))
        return new_boundary
    else:
        new_boundary = []
        for curve_loop in boundary:
            new_boundary.append(set_z_coordinate_of_boundary(curve_loop))
        return new_boundary


def rotate_z(point, angle):
    """
    Rotates the point around the z-axis

    Parameters
    ----------
    point: XYZ
        The original point
    angle: double
        The angle of rotation in radians

    Returns
    -------
    XYZ
        The rotated point
    """
    rotation_axis = XYZ(0, 0, 1)
    rotation_angle_radians = angle
    rotation = Transform.CreateRotation(rotation_axis, rotation_angle_radians)
    return rotation.OfPoint(point)


def get_boundary_from_solid(solid):
    """
    Calculates the XY plane boundary from a solid object or volume
    by iterating through the faces and finding the ones that are parallel to the XY plane

    Parameters
    ----------
    solid: Solid
        A 3D solid from which to get the boundary

    Returns
    -------
    List[CurveLoop]()
        The boundary as a list of CurveLoops
    """
    # Define a tolerance to identify XY plane faces
    tolerance = 0.001  # Adjust as needed

    # Iterate through the faces of the solid
    for face in solid.Faces:
        # Use the face normal to check if it's approximately parallel to the XY plane
        normal = face.FaceNormal
        if (
            abs(normal.X) < tolerance
            and abs(normal.Y) < tolerance
            and abs(normal.Z - 1.0) < tolerance
        ):
            # Return the first found face in the XY plane as CurveLoop
            return face.GetEdgesAsCurveLoops()
