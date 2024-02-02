# -*- coding: utf-8 -*-

# Imports
# =======================================================
from pyrevit import revit
from CMRPSettings import load_current_settings

from CMRP_Structure import run_structure_script
from CMRP_FreeHeight import run_freeheight_script
from CMRP_Functions import run_functions_script
from CMRP_Surfaces import run_surfaces_script
from CMRP_Daylight import run_daylight_script
from CMRP_Shafts import run_shafts_script
from CMRP_WindowTypes import run_window_types_script

# Main
# ===================================================
if __name__ == "__main__":
    # RevitDocument
    doc = revit.doc
    uidoc = revit.uidoc
    # Load settings from json file
    SETTINGS = load_current_settings(doc)

    # Group:    10_Structure
    # Function:     11_Structure
    run_structure_script(doc, uidoc, SETTINGS)

    # Group:    20_Program
    # Function:     21_Free Height
    run_freeheight_script(doc, uidoc, SETTINGS)

    # Function:     22_Functions
    run_functions_script(doc, uidoc, SETTINGS)

    # Function:     23_Surfaces
    run_surfaces_script(doc, uidoc, SETTINGS)

    # Group:    30_Daylight
    # Function:     31_Daylight
    run_daylight_script(doc, uidoc, SETTINGS)

    # Group:    40_Shafts
    # Function:     41_Shafts
    run_shafts_script(doc, uidoc, SETTINGS)

    # Group:    50_Window Types
    # Function:     51_Window Types
    run_window_types_script(doc, uidoc, SETTINGS)
