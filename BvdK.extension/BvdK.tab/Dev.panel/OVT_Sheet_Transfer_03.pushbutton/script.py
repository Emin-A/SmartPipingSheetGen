# -*- coding: utf-8 -*-
__title__ = "OVT Sheet\nTransfer"
__doc__ = """Version = 1.0
Date    = 03.08.2025
________________________________________________________________
Description:
Automates transfer of sheets, views, and scope boxes from OVT1 to OVT2.
________________________________________________________________
How-To:

1. [Hold ALT + CLICK] on the button to open its source folder.
2. Create your boundary (with detail lines) in the view.
3. Click the button and follow the prompts.
________________________________________________________________
Author: Emin Avdovic"""

# ==================================================
# Imports
# ==================================================
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FamilySymbol,
    FamilyInstance,
    FilteredElementCollector,
    FormatOptions,
    FilterStringRule,
    FilterStringRuleEvaluator,
    FilterStringBeginsWith,
    FilterStringContains,
    FilterStringEquals,
    XYZ,
    Transaction,
    TextNote,
    TextNoteType,
    TextNoteOptions,
    IndependentTag,
    UV,
    UnitTypeId,
    Reference,
    TagMode,
    TagOrientation,
    ViewSchedule,
    ViewSheet,
    ViewDuplicateOption,
    ViewDiscipline,
    Viewport,
    ViewPlan,
    ParameterValueProvider,
    ParameterFilterElement,
    ScheduleSheetInstance,
    ScheduleFilter,
    ScheduleFilterType,
    ScheduleSortGroupField,
    ScheduleSortOrder,
    StorageType,
    SectionType,
    Category,
)
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI import UIDocument
from Autodesk.Revit.DB.Structure import *
from Autodesk.Revit.Exceptions import *
from Autodesk.Revit.Attributes import *
from Autodesk.Revit.Exceptions import ArgumentException
from System.Collections.Generic import List

import clr
import System
import System.IO

clr.AddReference("System")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
clr.AddReference("RevitServices")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("WindowsBase")
from RevitServices.Persistence import DocumentManager
from System.Windows.Forms import (
    FormBorderStyle,
    AnchorStyles,
    AutoScaleMode,
    Form,
    ComboBox,
    ListBox,
    PictureBox,
    PictureBoxSizeMode,
    DataGridView,
    DataGridViewTextBoxColumn,
    DataGridViewButtonColumn,
    DataGridViewAutoSizeColumnsMode,
    DataGridViewSelectionMode,
    DockStyle,
    TextBox,
    Button,
    MessageBox,
    DialogResult,
    Label,
    ScrollBars,
    Application,
)
from System.Drawing import Image, Point, Color, Rectangle, Size
from System.IO import MemoryStream
from System.Windows.Forms import DataGridViewButtonColumn

from System import Array
from pyrevit import forms
import math, re, sys

# ==================================================
# Revit Document Setup
# ==================================================
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

VERBOSE = False

# === TO HIDE THE DEBUG CONSOLE ===
# def debug(*args):
#     if VERBOSE:
#         print(" ".join([str(a) for a in args]))


# Find the target document by name or prompt user to pick from open docs
def find_target_doc(target_title_substring):
    from Autodesk.Revit.ApplicationServices import Application

    app = doc.Application
    for open_doc in app.Documents:
        if target_title_substring in open_doc.Title and open_doc != doc:
            return open_doc
    raise Exception("Target doc with '{}' not found.".format(target_title_substring))


# Set this to part of your target doc filename (e.g., "OVT2")
target_doc = find_target_doc("BNZ_OMO_OV2_N_IT_MPO_Sanitaire Installatie")


# === STEP 1 & 2: COLLECT VIEWS, SHEETS, SCOPE BOXES
def collect_views_and_sheets(source_doc, main_view_name):
    # Find main view (floor plan) by name
    main_view = None
    for v in FilteredElementCollector(source_doc).OfClass(ViewPlan):
        if v.Name == main_view_name and not v.IsTemplate:
            main_view = v
            break
    if not main_view:
        raise Exception("Main view '{}' not found.".format(main_view_name))

    # Find dependent views
    dependent_views = []
    for v in FilteredElementCollector(source_doc).OfClass(ViewPlan):
        if (
            not v.IsTemplate
            and v.GetPrimaryViewId().IntegerValue == main_view.Id.IntegerValue
        ):
            dependent_views.append(v)
    all_views = [main_view] + dependent_views
    print("DEBUG: Views found: ", [v.Name for v in all_views])
    # Find sheets containing any of these views
    view_ids = set(v.Id.IntegerValue for v in all_views)
    sheets = []
    for s in FilteredElementCollector(source_doc).OfClass(ViewSheet):
        for vp_id in s.GetAllViewports():
            vp = source_doc.GetElement(vp_id)
            if hasattr(vp, "ViewId") and vp.ViewId.IntegerValue in view_ids:
                sheets.append(s)
                break
    print("DEBUG: Sheets found: ", [s.SheetNumber for s in sheets])
    # Find all unique scope boxes
    scope_box_ids = set()
    for s in sheets:
        for vp_id in s.GetAllViewports():
            vp = source_doc.GetElement(vp_id)
            if hasattr(vp, "ViewId") and vp.ViewId.IntegerValue in view_ids:
                # Try to get scope box from the viewport (works for placed dependent views)
                scope_box_param = vp.LookupParameter("Scope Box")
                if (
                    scope_box_param
                    and scope_box_param.StorageType == StorageType.ElementId
                ):
                    sb_id = scope_box_param.AsElementId()
                    if sb_id != ElementId.InvalidElementId:
                        scope_box_ids.add(sb_id)
    print(
        "DEBUG: Scope boxes found: ",
        [source_doc.GetElement(sb).Name for sb in scope_box_ids],
    )
    return all_views, sheets, scope_box_ids


# === STEP 3: COPY SCOPE BOXES
def copy_scope_boxes(source_doc, target_doc, scope_box_ids):
    from Autodesk.Revit.DB import CopyPasteOptions, ElementTransformUtils

    if not scope_box_ids:
        return {}
    copy_options = CopyPasteOptions()
    # Correct: convert to .NET List[ElementId]
    scope_box_ids_list = List[ElementId]()
    for eid in scope_box_ids:
        scope_box_ids_list.Add(eid)

    with Transaction(target_doc, "Copy Scope Boxes") as t:
        t.Start()
        # CopyElements returns new element ids in the target doc
        id_map = ElementTransformUtils.CopyElements(
            source_doc, scope_box_ids_list, target_doc, None, copy_options
        )
        # id_map: Dict[ElementId, ElementId]
        t.Commit()
    # ElementIdMap -> Python dict:
    id_dict = {}
    it = id_map.ForwardIterator()
    it.Reset()
    while it.MoveNext():
        id_dict[it.Key] = it.Value
    return id_dict


# === STEP 4: COPY VIEWS
def copy_views(source_doc, target_doc, views, scopebox_id_map):
    from Autodesk.Revit.DB import CopyPasteOptions, ElementTransformUtils

    view_id_map = {}
    copy_options = CopyPasteOptions()
    views_list = List[ElementId]([v.Id for v in views])
    with Transaction(target_doc, "Copy Views") as t:
        t.Start()
        id_map = ElementTransformUtils.CopyElements(
            source_doc, views_list, target_doc, None, copy_options
        )
        for i, v in enumerate(views):
            new_view_id = id_map[v.Id]
            new_view = target_doc.GetElement(new_view_id)
            view_id_map[v.Id] = new_view_id
            # Assign scope box if applicable
            if hasattr(v, "ScopeBox") and v.ScopeBox and v.ScopeBox.IntegerValue != -1:
                if v.ScopeBox in scopebox_id_map:
                    new_view.ScopeBox = scopebox_id_map[v.ScopeBox]
            # Copy view template if exists
            new_view.ViewTemplateId = v.ViewTemplateId
        t.Commit()
    return view_id_map


# === STEP 5: COPY SHEETS AND PLACE VIEWS
def copy_sheets(source_doc, target_doc, sheets, view_id_map):
    from Autodesk.Revit.DB import Viewport

    existing_sheet_numbers = set(
        s.SheetNumber for s in FilteredElementCollector(target_doc).OfClass(ViewSheet)
    )
    copied_sheet_count = 0
    with Transaction(target_doc, "Copy Sheets and Place Views") as t:
        t.Start()
        for s in sheets:
            new_sheet_number = s.SheetNumber
            # If conflict, auto-rename with _copy, _copy2, etc.
            count = 1
            orig_sheet_number = new_sheet_number
            while new_sheet_number in existing_sheet_numbers:
                new_sheet_number = "{}_copy{}".format(
                    orig_sheet_number, count if count > 1 else ""
                )
                count += 1
            # Copy titleblock family
            titleblock = (
                source_doc.GetElement(s.GetTitleBlockIds()[0])
                if s.GetTitleBlockIds()
                else None
            )
            if titleblock:
                new_sheet = ViewSheet.Create(target_doc, titleblock.GetTypeId())
            else:
                new_sheet = ViewSheet.Create(target_doc, ElementId.InvalidElementId)
            new_sheet.SheetNumber = new_sheet_number
            new_sheet.Name = s.Name
            # Place new views at same viewport locations
            for vp_id in s.GetAllViewports():
                vp = source_doc.GetElement(vp_id)
                old_view_id = vp.ViewId
                if old_view_id in view_id_map:
                    new_view_id = view_id_map[old_view_id]
                    box_center = vp.GetBoxCenter()
                    Viewport.Create(target_doc, new_sheet.Id, new_view_id, box_center)
            existing_sheet_numbers.add(new_sheet_number)
            copied_sheet_count += 1
        t.Commit()
    return copied_sheet_count


# === MAIN FUNCTION
def main():
    # Set this to your main floor plan or sheet view name to copy
    def pick_main_view_name(doc):
        view_names = sorted(
            [
                v.Name
                for v in FilteredElementCollector(doc).OfClass(ViewPlan)
                if not v.IsTemplate
            ]
        )
        result = forms.SelectFromList.show(
            view_names, title="Select Main Floor Plan", multiselect=False
        )
        return result

    main_view_name = pick_main_view_name(doc)
    if not main_view_name:
        TaskDialog.Show("Cancelled", "No main view selected.")
        import sys

        sys.exit()
    # 1. Collect all views, sheets, and scope boxes needed
    all_views, sheets, scope_box_ids = collect_views_and_sheets(doc, main_view_name)
    if not all_views or not sheets or not scope_box_ids:
        TaskDialog.Show("Nothing Found", "No matching views, sheets, or scope boxes.")
        return
    # 2. Copy scope boxes to target project
    scopebox_id_map = copy_scope_boxes(doc, target_doc, scope_box_ids)
    # 3. Copy views to target, reassigning new scope box IDs
    view_id_map = copy_views(doc, target_doc, all_views, scopebox_id_map)
    # 4. Copy sheets and place new views
    sheet_count = copy_sheets(doc, target_doc, sheets, view_id_map)
    # 5. Report
    msg = "Copied {} views, {} sheets, and {} scope boxes.\n".format(
        len(all_views), sheet_count, len(scope_box_ids)
    )
    TaskDialog.Show("Transfer Complete", msg)


main()
