# -*- coding: utf-8 -*-

# Imports
# =======================================================
from pyrevit import script

from CMRP import delete_area_plans, create_area_plans
from CMRPParameters import STRUCTURE_TEMPLATE_NAME, MissingTemplateError
from CMRPSettings import MissingSettingsError

# Definitions and global variables
# =======================================================
logger = script.get_logger()

# Main
# ==================================================


def run_structure_script(document, uidocument, settings):
    try:
        delete_area_plans(document, uidocument, STRUCTURE_TEMPLATE_NAME)
        create_area_plans(document, STRUCTURE_TEMPLATE_NAME)
        logger.success(
            "{}: Script completed without critical errors.".format(
                STRUCTURE_TEMPLATE_NAME
            )
        )
    except (MissingSettingsError, MissingTemplateError) as error:
        logger.critical(error.message)
