# -*- coding: utf-8 -*-

# Imports
# =======================================================
from pyrevit import revit
from CMRPSettings import load_current_settings
from CMRP_FreeHeight import run_freeheight_script
from CMRP_Functions import run_functions_script
from CMRP_Surfaces import run_surfaces_script

# Main
# =======================================================
if __name__ == "__main__":
    # RevitDocument
    doc = revit.doc
    uidoc = revit.uidoc
    # Load settings from json file
    SETTINGS = load_current_settings(doc)

    # Group:    20_Program
    # Function:     21_Free Height
    run_freeheight_script(doc, uidoc, SETTINGS)

    # Function:     22_Functions
    run_functions_script(doc, uidoc, SETTINGS)

    # Function:     23_Surfaces
    run_surfaces_script(doc, uidoc, SETTINGS)
