# -*- coding: utf-8 -*-

# Imports
# =======================================================
from pyrevit import forms
from Autodesk.Revit.DB import *


# Definitions and global variables
# =======================================================
def linked_element_selection(title, linked_instances, multiselect=False):
    """Creates a popup window to select links
    Parameters
    ----------
    title: str
        The title for the popup

    linked_instances: array of linked instances
        The data for the selection menu
    """
    link_names = [link.Name for link in linked_instances]
    selected_links = forms.SelectFromList.show(
        link_names, title=title, width=400, height=500, multiselect=multiselect
    )
    if not selected_links or selected_links is None:
        return None
    if isinstance(selected_links, str):
        return [linked_instances[link_names.index(selected_links)]]
    return [linked_instances[link_names.index(name)] for name in selected_links]
