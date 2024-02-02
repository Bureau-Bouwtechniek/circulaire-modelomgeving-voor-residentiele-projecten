# -*- coding: utf-8 -*-

# Imports
# =======================================================
from pyrevit import revit
from CMRPSettings import load_current_settings
from CMRP_WindowTypes import run_window_types_script

# Main
# =======================================================
if __name__ == "__main__":
    # RevitDocument
    doc = revit.doc
    uidoc = revit.uidoc
    # Load settings from json file
    SETTINGS = load_current_settings(doc)
    run_window_types_script(doc, uidoc, SETTINGS)
