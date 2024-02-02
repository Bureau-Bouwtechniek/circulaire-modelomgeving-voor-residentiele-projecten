# -*- coding: utf-8 -*-

# Imports
# =======================================================
from pyrevit import revit

from Autodesk.Revit.DB import *
from CMRPSettings import load_current_settings
from CMRP import delete_area_plans, create_area_plans
from CMRPParameters import STRUCTURE_TEMPLATE_NAME
from CMRP_Shafts import run_shafts_script


# Definitions and global variables
# =======================================================


# Main
# ===================================================
if __name__ == "__main__":
    # RevitDocument
    doc = revit.doc
    uidoc = revit.uidoc
    # Load settings from json file
    SETTINGS = load_current_settings(doc)
    # Structure
    delete_area_plans(doc, uidoc, STRUCTURE_TEMPLATE_NAME[:3])
    create_area_plans(doc, STRUCTURE_TEMPLATE_NAME)
    # Shafts
    run_shafts_script(doc, uidoc, SETTINGS, False)
