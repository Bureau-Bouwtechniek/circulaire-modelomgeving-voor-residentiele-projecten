# -*- coding: utf-8 -*-

# Imports
# =======================================================
import datetime

starttime = datetime.datetime.today()

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import SpecTypeId
from pyrevit import revit

# RevitDocument
# =======================================================
doc = revit.doc
uidoc = revit.uidoc


# Definitions and global variables
# =======================================================
def get_units():
    try:
        UIunit = (
            Document.GetUnits(doc).GetFormatOptions(UnitType.UT_Length).DisplayUnits
        )
    except:
        UIunit = (
            Document.GetUnits(doc).GetFormatOptions(SpecTypeId.Length).GetUnitTypeId()
        )
    return UIunit


def units_to_internal(items):
    UIunit = get_units()
    if isinstance(items, list):
        return [UnitUtils.ConvertToInternalUnits(i, UIunit) for i in items]
    else:
        return UnitUtils.ConvertToInternalUnits(items, UIunit)


def units_from_internal(items):
    UIunit = get_units()
    if isinstance(items, list):
        return [round(UnitUtils.ConvertFromInternalUnits(i, UIunit), 2) for i in items]
    if isinstance(items, XYZ):
        return XYZ(
            round(UnitUtils.ConvertFromInternalUnits(items.X, UIunit), 2),
            round(UnitUtils.ConvertFromInternalUnits(items.Y, UIunit), 2),
            round(UnitUtils.ConvertFromInternalUnits(items.Z, UIunit), 2),
        )
    else:
        return round(
            UnitUtils.ConvertFromInternalUnits(items, UIunit),
        )
