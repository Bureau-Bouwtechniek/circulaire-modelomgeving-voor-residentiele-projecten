# imports
# ===================================================
import clr

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *


# definitions
# ===================================================
def CMRP_colors_gray():
    colors = iter([])  # to be updated
    return colors


def CMRP_colors_zacht():
    colors = [
        Color(253, 215, 215),
        Color(215, 241, 226),
        Color(235, 251, 245),
        Color(226, 243, 249),
        Color(193, 228, 241),
        Color(255, 224, 195),
        Color(255, 241, 205),
        Color(255, 253, 205),
        Color(246, 254, 208),
        Color(217, 234, 194),
        Color(247, 203, 226),
        Color(252, 239, 246),
        Color(240, 225, 254),
        Color(217, 207, 243),
        Color(237, 236, 218),
        Color(249, 247, 225),
    ]
    return colors


def CMRP_colors_hard():
    colors = [
        Color(251, 123, 124),
        Color(126, 209, 160),
        Color(192, 242, 223),
        Color(158, 219, 240),
        Color(51, 165, 207),
        Color(254, 151, 58),
        Color(255, 208, 91),
        Color(255, 248, 91),
        Color(226, 251, 99),
        Color(131, 185, 54),
        Color(231, 81, 157),
        Color(242, 147, 195),
        Color(206, 155, 254),
        Color(131, 95, 215),
        Color(200, 195, 132),
        Color(237, 233, 157),
    ]
    return colors
