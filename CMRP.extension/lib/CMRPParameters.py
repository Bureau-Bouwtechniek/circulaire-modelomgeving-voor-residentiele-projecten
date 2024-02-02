# -*- coding: utf-8 -*-

# CMRP Parameters definitions
# =======================================================

# Common Parameters
CMRP_UniqueId = "CMRP_UniqueId"

# Structure Script
STRUCTURE_TEMPLATE_NAME = "11_Structure"

# FreeHeight Script
FREE_HEIGHT_TEMPLATE_NAME = "21_Free Height"
CMRP_CP_LEN_Height = "CMRP_CP_LEN_Height"
ray3D_View = "FreeHeight_Raytrace_3D"

# Functions Script
FUNCTIONS_TEMPLATE_NAME = "22_Functions"
Comments = "Comments"
CMRP_CP_TXT_Function = "CMRP_CP_TXT_Function"

# Surfaces Script
SURFACES_TEMPLATE_NAME = "23_Surfaces"

# Daylight Script
DAYLIGHT_TEMPLATE_NAME = "31_Daylight"
CMRP_Daylight = "CMRP_Daylight"
CMRP_DaylightBoundary = "CMRP_DaylightBoundary"

# Window Type Script
CMRP_31_GM_UN_WindowBlock = "CMRP_31_GM_UN_WindowBlock"
CMRP_CF_TX_WindowType = "CMRP_CF_TX_WindowType"
CMRP_CF_LE_Height = "CMRP_CF_LE_Height"
CMRP_CF_LE_Width = "CMRP_CF_LE_Width"
CMRP_CF_MA_Material = "CMRP_CF_MA_Material"

# Shafts Script
SHAFTS_TEMPLATE_NAME = "41_Shafts"
mm_of_expansion = 1000

# WindowTypes Script
WINDOWTYPES_TEMPLATE_NAME = "51_Window Types"


class MissingTemplateError(Exception):
    def __init__(self, message):
        self.message = message
