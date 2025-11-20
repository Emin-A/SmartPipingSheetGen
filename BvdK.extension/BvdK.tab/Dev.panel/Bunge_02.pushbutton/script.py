# -*- coding: utf-8 -*-
__title__ = "Bunge Sheet Gen"
__doc__ = """Version = 1.1  (safe-mode)
Date    = 21.07.2025
________________________________________________________________
Description:
Creates a cropped prefab plan from a user-drawn boundary, numbers
Pipes/Fittings 'Comments', auto-tags pipes (no leaders), and places
the view on a new sheet (safe defaults to avoid regen thrash).
________________________________________________________________
How-To:
1. Draw a closed loop with Detail Lines around the region.
2. Run the tool, edit/confirm codes, press OK.
3. Pick a title block, done.
________________________________________________________________
Author: Emin Avdovic (safe-mode refactor)
"""

# =========================
# SAFE MODE SWITCHES
# =========================
SAFE_CFG = {
    "duplicate_with_detailing": False,  # False => minimal duplication (safer)
    "place_text_note": True,  # place tiny text note inside crop
    "autotag_pipes": True,  # create tags for untagged pipes
    "tag_use_leader": False,  # leaderless tags = lighter regen
    "create_3d_view": False,  # disabled for stability
    "place_schedule": False,  # disabled for stability
    "hide_crop_region_controls": True,  # keep the crop visible OFF
    "annotation_crop_on": False,  # annotation crop OFF for stability
    "view_scale": 50,  # plan scale
    "view_discipline": "Coordination",  # Coordination is stable
}

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
    XYZ,
    Transaction,
    TextNote,
    TextNoteType,
    TextNoteOptions,
    IndependentTag,
    UnitTypeId,
    Reference,
    TagMode,
    TagOrientation,
    ViewSheet,
    ViewDuplicateOption,
    ViewDiscipline,
    Viewport,
    ViewPlan,
    StorageType,
    Category,
)
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.UI import *
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
    Form,
    ListBox,
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
    Application,
)
from System.Drawing import Point, Color, Rectangle, Size
from System import Array
import math, re, sys

# ==================================================
# Revit Document Setup
# ==================================================
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

VERBOSE = False


def debug(*args):
    if VERBOSE:
        print(" ".join([str(a) for a in args]))


# ==================================================
# Utilities
# ==================================================
def safe_int(value, default=0):
    try:
        return int(str(value))
    except:
        return default


def mm_to_ft(mm):
    return UnitUtils.ConvertToInternalUnits(mm, UnitTypeId.Millimeters)


Z_BAND = mm_to_ft(1500.0)
XY_PAD = mm_to_ft(50.0)
INSET = mm_to_ft(15.0)  # tiny inset to avoid grazing zeros


def points_are_close(pt1, pt2, tol=1e-6):
    return (
        abs(pt1.X - pt2.X) < tol
        and abs(pt1.Y - pt2.Y) < tol
        and abs(pt1.Z - pt2.Z) < tol
    )


def order_segments_to_polygon(segments):
    if not segments:
        return None
    polygon = [segments[0][0], segments[0][1]]
    segments.pop(0)
    changed = True
    while segments and changed:
        changed = False
        last_pt = polygon[-1]
        for idx, seg in enumerate(segments):
            ptA, ptB = seg
            if points_are_close(last_pt, ptA):
                polygon.append(ptB)
                segments.pop(idx)
                changed = True
                break
            elif points_are_close(last_pt, ptB):
                polygon.append(ptA)
                segments.pop(idx)
                changed = True
                break
    if polygon and points_are_close(polygon[0], polygon[-1]):
        polygon.pop()
        return polygon
    return None


def is_point_inside_polygon(point, polygon):
    x, y = point.X, point.Y
    inside = False
    n = len(polygon)
    j = n - 1
    for i in range(n):
        xi, yi = polygon[i].X, polygon[i].Y
        xj, yj = polygon[j].X, polygon[j].Y
        if ((yi > y) != (yj > y)) and (
            x < (xj - xi) * (y - yi) / ((yj - yi) or 1e-12) + xi
        ):
            inside = not inside
        j = i
    return inside


def polygon_bounds_xy(poly):
    minx = min(p.X for p in poly)
    maxx = max(p.X for p in poly)
    miny = min(p.Y for p in poly)
    maxy = max(p.Y for p in poly)
    # safety inset/outset
    return (minx + INSET, miny + INSET, maxx - INSET, maxy - INSET)


class DetailLineSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return bool(
            elem.Category
            and elem.Category.Id.IntegerValue == int(BuiltInCategory.OST_Lines)
        )

    def AllowReference(self, ref, point):
        return False


def convert_param_to_string(param_obj):
    if not param_obj:
        return ""
    try:
        s = param_obj.AsValueString()
        if s and s.strip():
            return s
    except:
        pass
    try:
        val_mm = param_obj.AsDouble() * 304.8
        return str(int(round(val_mm))) + " mm"
    except:
        return ""


def select_boundary_and_gather():
    try:
        selection_refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            DetailLineSelectionFilter(),
            "Select boundary detail lines (closed loop)",
        )
    except Exception:
        return None, None
    if not selection_refs:
        return None, None

    segments = []
    for r in selection_refs:
        el = doc.GetElement(r)
        try:
            crv = el.GeometryCurve
            segments.append((crv.GetEndPoint(0), crv.GetEndPoint(1)))
        except:
            pass

    polygon = order_segments_to_polygon(segments[:])
    if polygon is None:
        MessageBox.Show("The selected lines do not form a closed loop.", "Error")
        return None, None

    active_plan = uidoc.ActiveView
    level_z = 0.0
    if isinstance(active_plan, ViewPlan) and active_plan.GenLevel:
        level_z = doc.GetElement(active_plan.GenLevel.Id).Elevation
    z_min = level_z - Z_BAND
    z_max = level_z + Z_BAND

    def rects_overlap(ax1, ay1, ax2, ay2, bx1, by1, bx2, by2):
        return not (ax2 < bx1 or bx2 < ax1 or ay2 < by1 or by2 < ay1)

    pminx, pminy, pmaxx, pmaxy = polygon_bounds_xy(polygon)
    collector = (
        FilteredElementCollector(doc, uidoc.ActiveView.Id)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    elements_inside = []

    for elem in collector:
        bb = elem.get_BoundingBox(uidoc.ActiveView)
        if not bb:
            continue
        if bb.Max.Z < z_min or bb.Min.Z > z_max:
            continue
        if not rects_overlap(
            bb.Min.X, bb.Min.Y, bb.Max.X, bb.Max.Y, pminx, pminy, pmaxx, pmaxy
        ):
            continue

        center = XYZ((bb.Min.X + bb.Max.X) * 0.5, (bb.Min.Y + bb.Max.Y) * 0.5, 0)
        inside = is_point_inside_polygon(center, polygon)
        if inside:
            elements_inside.append(elem)

    MessageBox.Show(
        "Found {0} element(s) in region.".format(len(elements_inside)), "Boundary"
    )
    return elements_inside, polygon


# ==================================================
# Data grid UI (unchanged core behavior; cleaned a bit)
# ==================================================
from System.Windows.Forms import (
    Panel,
    DataGridViewTextBoxColumn,
    DataGridViewButtonColumn,
    DataGridViewAutoSizeColumnsMode,
)


class ElementEditorForm(Form):
    def __init__(self, elements_data, region_elements=None):
        self.Text = "Edit Element Codes"
        self.Width = 1050
        self.Height = 500
        self.MinimumSize = Size(700, 400)
        self.SuspendLayout()
        self.regionElements = region_elements

        self.buttonPanel = Panel()
        self.buttonPanel.Height = 48
        self.buttonPanel.Dock = DockStyle.Bottom
        self.Controls.Add(self.buttonPanel)
        self.gridPanel = Panel()
        self.gridPanel.Dock = DockStyle.Fill
        self.Controls.Add(self.gridPanel)

        self.dataGrid = DataGridView()
        self.dataGrid.Dock = DockStyle.Fill
        self.dataGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        self.dataGrid.MultiSelect = True
        self.dataGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        self.dataGrid.SelectionChanged += self.on_row_selected
        self.gridPanel.Controls.Add(self.dataGrid)

        # columns
        cols = []

        def add_col(name, header, ro=True):
            c = DataGridViewTextBoxColumn()
            c.Name = name
            c.HeaderText = header
            c.ReadOnly = ro
            cols.append(c)

        add_col("Id", "Element Id")
        add_col("Category", "Category")
        add_col("Name", "Name")
        add_col("Warning", "Warning")
        add_col("Bend45", "2x45°")
        add_col("DefaultCode", "Default Code")
        add_col("NewCode", "New Code", ro=False)
        add_col("OutsideDiameter", "Outside Diameter")
        add_col("Length", "Length")
        add_col("Size", "Size")
        add_col("GEB_Article_Number", "GEB Article No.")

        self.colTagStatus = DataGridViewButtonColumn()
        self.colTagStatus.Name = "TagStatus"
        self.colTagStatus.HeaderText = "Tags"
        self.colTagStatus.UseColumnTextForButtonValue = False

        self.dataGrid.Columns.AddRange(Array[DataGridViewTextBoxColumn](cols))
        self.dataGrid.Columns.Add(self.colTagStatus)
        self.dataGrid.CellContentClick += self.dataGrid_CellContentClick

        # text note input
        self.txtTextNoteCode = TextBox()
        self.txtTextNoteCode.Width = 160
        self.txtTextNoteCode.ForeColor = Color.Gray
        self.txtTextNoteCode.Text = "prefab 5.5.5"
        self.txtTextNoteCode.GotFocus += self._clear_placeholder
        self.txtTextNoteCode.LostFocus += self._restore_placeholder
        self.buttonPanel.Controls.Add(self.txtTextNoteCode)

        # buttons
        def add_btn(text, handler, w=150):
            b = Button()
            b.Text = text
            b.Width = w
            b.Click += handler
            self.buttonPanel.Controls.Add(b)
            return b

        self.btnAutoFill = add_btn("Auto-Fill Tag Codes", self.autoFillPipeTagCodes)
        self.btnOK = add_btn("OK", self.okButton_Click, 80)
        self.btnCancel = add_btn("Cancel", self._cancel, 80)
        self.buttonPanel.Resize += self._layout_buttons

        self.textNotePlaced = False
        self.Result = None

        # populate
        for ed in elements_data:
            ridx = self.dataGrid.Rows.Add()
            row = self.dataGrid.Rows[ridx]
            row.Cells["Id"].Value = ed["Id"]
            row.Cells["Category"].Value = ed["Category"]
            row.Cells["Name"].Value = ed["Name"]
            row.Cells["Warning"].Value = ed.get("Warning", "")
            row.Cells["Bend45"].Value = ed.get("Bend45", "")
            row.Cells["DefaultCode"].Value = ed["DefaultCode"]
            row.Cells["NewCode"].Value = ed["NewCode"]
            row.Cells["OutsideDiameter"].Value = ed["OutsideDiameter"]
            row.Cells["Length"].Value = ed["Length"]
            row.Cells["Size"].Value = ed.get("Size", "")
            row.Cells["GEB_Article_Number"].Value = ed.get("GEB_Article_Number", "")

            cat = ed["Category"]
            if cat == "Pipes":
                row.Cells["TagStatus"].Value = (
                    "Remove Tag" if ed["TagStatus"] == "Yes" else "Add/Place Tag"
                )
                row.DefaultCellStyle.BackColor = Color.LightBlue
            elif cat == "Pipe Tags":
                row.Cells["TagStatus"].Value = "Remove Tag"
                row.DefaultCellStyle.BackColor = Color.LightGreen
            elif cat == "Pipe Fittings":
                row.Cells["TagStatus"].Value = ""
                row.Cells["TagStatus"].ReadOnly = True
                row.DefaultCellStyle.BackColor = Color.LightGoldenrodYellow
            else:
                row.Cells["TagStatus"].Value = ""
                row.DefaultCellStyle.BackColor = Color.LightGray

        self.ResumeLayout(False)

    def _layout_buttons(self, sender, e):
        ctrls = list(self.buttonPanel.Controls)
        total_w = sum(c.Width for c in ctrls)
        avail = self.buttonPanel.Width - total_w
        spacing = max(10, avail // (len(ctrls) + 1))
        x = spacing
        for c in ctrls:
            c.Location = Point(x, (self.buttonPanel.Height - c.Height) // 2)
            x += c.Width + spacing

    def _clear_placeholder(self, sender, e):
        if self.txtTextNoteCode.Text == "prefab 5.5.5":
            self.txtTextNoteCode.Text = ""
            self.txtTextNoteCode.ForeColor = Color.Black

    def _restore_placeholder(self, sender, e):
        if not self.txtTextNoteCode.Text.strip():
            self.txtTextNoteCode.Text = "prefab 5.5.5"
            self.txtTextNoteCode.ForeColor = Color.Gray

    def autoFillPipeTagCodes(self, sender, e):
        raw = self.txtTextNoteCode.Text.strip()
        m = re.search(r"([\d\.]+)", raw)
        if not m:
            MessageBox.Show("Could not parse base code from text note.", "Error")
            return
        base = m.group(1)

        pipe_rows, fit_rows, tag_rows = [], [], []
        for i in range(self.dataGrid.Rows.Count):
            cat = self.dataGrid.Rows[i].Cells["Category"].Value
            if cat == "Pipes":
                pipe_rows.append(i)
            elif cat == "Pipe Fittings":
                fit_rows.append(i)
            elif cat == "Pipe Tags":
                tag_rows.append(i)

        # fittings = base
        for i in fit_rows:
            self.dataGrid.Rows[i].Cells["NewCode"].Value = base

        # pipes numbered by X,Y
        entries = []
        for i in pipe_rows:
            rid = safe_int(self.dataGrid.Rows[i].Cells["Id"].Value)
            elem = doc.GetElement(ElementId(rid))
            bb = elem.get_BoundingBox(uidoc.ActiveView) if elem else None
            ctr = (
                XYZ(
                    (bb.Min.X + bb.Max.X) * 0.5,
                    (bb.Min.Y + bb.Max.Y) * 0.5,
                    (bb.Min.Z + bb.Max.Z) * 0.5,
                )
                if bb
                else XYZ(0, 0, 0)
            )
            entries.append((i, ctr))
        entries.sort(key=lambda x: (x[1].X, x[1].Y))
        for n, (i, _) in enumerate(entries, 1):
            self.dataGrid.Rows[i].Cells["NewCode"].Value = "{}.{}".format(base, n)

        # mirror to tag rows (count-limited)
        for n in range(1, len(entries) + 1):
            if n - 1 < len(tag_rows):
                tr = tag_rows[n - 1]
                self.dataGrid.Rows[tr].Cells["NewCode"].Value = "{}.{}".format(base, n)

    def dataGrid_CellContentClick(self, sender, e):
        # Only handle basic Pipe Tag add/remove here (leaderless); no fitting flips in safe-mode
        if self.dataGrid.Columns[e.ColumnIndex].Name != "TagStatus":
            return
        row = self.dataGrid.Rows[e.RowIndex]
        cat = row.Cells["Category"].Value
        val = row.Cells["TagStatus"].Value
        if cat != "Pipes":
            return

        host_id = safe_int(row.Cells["Id"].Value)
        host = doc.GetElement(ElementId(host_id))
        if not host:
            return

        if val == "Add/Place Tag":
            with Transaction(doc, "Add Tag (safe)") as tr:
                tr.Start()
                bb = host.get_BoundingBox(uidoc.ActiveView)
                if bb:
                    ctr = XYZ(
                        (bb.Min.X + bb.Max.X) * 0.5,
                        (bb.Min.Y + bb.Max.Y) * 0.5,
                        (bb.Min.Z + bb.Max.Z) * 0.5,
                    )
                    IndependentTag.Create(
                        doc,
                        doc.ActiveView.Id,
                        Reference(host),
                        SAFE_CFG["tag_use_leader"],
                        TagMode.TM_ADDBY_CATEGORY,
                        TagOrientation.Horizontal,
                        ctr,
                    )
                tr.Commit()
            row.Cells["TagStatus"].Value = "Remove Tag"
        elif val == "Remove Tag":
            # remove any tag in active view pointing to this host
            tag_to_delete = None
            for t in (
                FilteredElementCollector(doc, uidoc.ActiveView.Id)
                .OfCategory(BuiltInCategory.OST_PipeTags)
                .WhereElementIsNotElementType()
                .ToElements()
            ):
                try:
                    if hasattr(t, "GetTaggedElementIds"):
                        ids = t.GetTaggedElementIds()
                        for rid in ids:
                            hid = (
                                rid.HostElementId
                                if hasattr(rid, "HostElementId")
                                else rid
                            )
                            if hid.IntegerValue == host.Id.IntegerValue:
                                tag_to_delete = t.Id
                                break
                    elif hasattr(t, "TaggedElementId") and t.TaggedElementId:
                        if t.TaggedElementId.IntegerValue == host.Id.IntegerValue:
                            tag_to_delete = t.Id
                except:
                    pass
                if tag_to_delete:
                    break

            if tag_to_delete:
                with Transaction(doc, "Remove Tag (safe)") as tr:
                    tr.Start()
                    doc.Delete(tag_to_delete)
                    tr.Commit()
                row.Cells["TagStatus"].Value = "Add/Place Tag"

    def okButton_Click(self, sender, e):
        updated = []
        for r in self.dataGrid.Rows:
            updated.append(
                {
                    "Id": r.Cells["Id"].Value,
                    "Category": r.Cells["Category"].Value,
                    "Name": r.Cells["Name"].Value,
                    "DefaultCode": r.Cells["DefaultCode"].Value,
                    "NewCode": r.Cells["NewCode"].Value,
                }
            )
        self.Result = {
            "Elements": updated,
            "TextNotePlaced": self.textNotePlaced,
            "TextNote": self.txtTextNoteCode.Text.strip(),
        }
        self.DialogResult = DialogResult.OK
        self.Close()

    def _cancel(self, sender, e):
        self.Result = None
        self.DialogResult = DialogResult.Cancel
        self.Close()

    def on_row_selected(self, sender, e):
        row = self.dataGrid.CurrentRow
        if not row:
            return
        id_val = row.Cells["Id"].Value
        if not id_val:
            return
        try:
            eid = ElementId(int(str(id_val)))
            elem = doc.GetElement(eid)
            if elem and elem.IsValidObject:
                uidoc.Selection.SetElementIds(List[ElementId]([eid]))
        except:
            return


def show_element_editor(elements_data, region_elements=None):
    f = ElementEditorForm(elements_data, region_elements)
    return f.ShowDialog() == DialogResult.OK and f.Result or None


# ==================================================
# Filtering for relevant elements
# ==================================================
def filter_relevant_elements(gathered_elements):
    relevant = []
    all_pipe_tags = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_PipeTags)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    pipe_ids = {
        e.Id.IntegerValue
        for e in gathered_elements
        if e.Category and e.Category.Name == "Pipes"
    }

    # include any pipe tags whose hosts are in region
    for tag in all_pipe_tags:
        try:
            host = None
            if hasattr(tag, "GetTaggedElementIds"):
                ids = tag.GetTaggedElementIds()
                if ids and ids.Count > 0:
                    rid = ids[0]
                    host_eid = (
                        rid.HostElementId if hasattr(rid, "HostElementId") else rid
                    )
                    host = doc.GetElement(host_eid)
            elif hasattr(tag, "TaggedElementId"):
                host = doc.GetElement(tag.TaggedElementId)
            if host and host.Id.IntegerValue in pipe_ids:
                relevant.append(
                    {
                        "Id": str(tag.Id),
                        "Category": "Pipe Tags",
                        "Name": tag.Name or "",
                        "Warning": "",
                        "Bend45": "",
                        "DefaultCode": (
                            (host.LookupParameter("Comments").AsString() or "")
                            if host
                            else ""
                        ),
                        "NewCode": (
                            (host.LookupParameter("Comments").AsString() or "")
                            if host
                            else ""
                        ),
                        "OutsideDiameter": (
                            convert_param_to_string(
                                host.LookupParameter("Outside Diameter")
                            )
                            if host
                            else ""
                        ),
                        "Length": (
                            convert_param_to_string(host.LookupParameter("Length"))
                            if host
                            else ""
                        ),
                        "Size": "",
                        "GEB_Article_Number": "",
                        "TagStatus": "Yes",
                    }
                )
        except:
            pass

    for e in gathered_elements:
        if not e.Category:
            continue
        cat = e.Category.Name
        if cat not in ("Pipes", "Pipe Fittings", "Pipe Tags", "Text Notes"):
            continue

        com = e.LookupParameter("Comments")
        default_code = com.AsString() if com and com.AsString() else ""

        warning_val = ""
        bend45_val = ""
        outside_diam = ""
        length_val = ""
        art_num = ""
        tag_status = ""

        if cat == "Pipes":
            odp = e.LookupParameter("Outside Diameter")
            lp = e.LookupParameter("Length")
            outside_diam = convert_param_to_string(odp)
            length_val = convert_param_to_string(lp)

            tag_status = "No"
            for tag in all_pipe_tags:
                try:
                    if hasattr(tag, "GetTaggedElementIds"):
                        for rid in tag.GetTaggedElementIds():
                            eid = (
                                rid.HostElementId
                                if hasattr(rid, "HostElementId")
                                else rid
                            )
                            if eid.IntegerValue == e.Id.IntegerValue:
                                tag_status = "Yes"
                                break
                    elif hasattr(tag, "TaggedElementId") and tag.TaggedElementId:
                        if tag.TaggedElementId.IntegerValue == e.Id.IntegerValue:
                            tag_status = "Yes"
                except:
                    pass
                if tag_status == "Yes":
                    break

        elif cat == "Pipe Fittings":
            p_warn = e.LookupParameter("waarschuwing")
            warning_val = p_warn.AsString() if p_warn else ""
            p_bend = e.LookupParameter("2x45°")
            if p_bend and p_bend.StorageType == StorageType.Integer:
                bend45_val = "Yes" if p_bend.AsInteger() == 1 else "No"
            for pname in ("Outside Diameter", "Diameter", "Nominal Diameter"):
                p = e.LookupParameter(pname)
                if p:
                    outside_diam = convert_param_to_string(p)
                    break
            lp = e.LookupParameter("Length")
            length_val = convert_param_to_string(lp)
            ap = e.LookupParameter("GEB_Article_Number")
            art_num = ap.AsString() if (ap and ap.AsString()) else ""
            tag_status = ""  # fittings not tagged here in safe-mode

        elif cat == "Pipe Tags":
            tag_status = "Yes"
            host = None
            try:
                if hasattr(e, "GetTaggedElementIds"):
                    ids = e.GetTaggedElementIds()
                    if ids and ids.Count > 0:
                        rid = ids[0]
                        host_eid = (
                            rid.HostElementId if hasattr(rid, "HostElementId") else rid
                        )
                        host = doc.GetElement(host_eid)
                if not host and hasattr(e, "TaggedElementId"):
                    host = doc.GetElement(e.TaggedElementId)
            except:
                host = None
            if host:
                outside_diam = convert_param_to_string(
                    host.LookupParameter("Outside Diameter")
                )
                length_val = convert_param_to_string(host.LookupParameter("Length"))

        relevant.append(
            {
                "Id": str(e.Id),
                "Category": cat,
                "Name": e.Name if hasattr(e, "Name") else "",
                "Warning": warning_val,
                "Bend45": bend45_val,
                "DefaultCode": default_code,
                "NewCode": default_code,
                "OutsideDiameter": outside_diam,
                "Length": length_val,
                "Size": "",
                "GEB_Article_Number": art_num,
                "TagStatus": tag_status,
            }
        )
    return relevant


# ==================================================
# Tagging helpers (leaderless, single transaction)
# ==================================================
def _midpoint_or_center(elem, view):
    try:
        loc = getattr(elem, "Location", None)
        if loc and hasattr(loc, "Curve") and loc.Curve:
            return loc.Curve.Evaluate(0.5, True)
    except:
        pass
    bb = elem.get_BoundingBox(view)
    if not bb:
        return XYZ(0, 0, 0)
    return XYZ(
        (bb.Min.X + bb.Max.X) / 2.0,
        (bb.Min.Y + bb.Max.Y) / 2.0,
        (bb.Min.Z + bb.Max.Z) / 2.0,
    )


def get_untagged_pipes_in_view(doc, view):
    pipes = list(
        FilteredElementCollector(doc, view.Id)
        .OfCategory(BuiltInCategory.OST_PipeCurves)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    tagged = set()
    for t in (
        FilteredElementCollector(doc, view.Id)
        .OfCategory(BuiltInCategory.OST_PipeTags)
        .WhereElementIsNotElementType()
        .ToElements()
    ):
        try:
            if hasattr(t, "GetTaggedElementIds"):
                for rid in t.GetTaggedElementIds():
                    hid = rid.HostElementId if hasattr(rid, "HostElementId") else rid
                    tagged.add(hid.IntegerValue)
            elif hasattr(t, "TaggedElementId") and t.TaggedElementId:
                tagged.add(t.TaggedElementId.IntegerValue)
        except:
            pass
    return [p for p in pipes if p.Id.IntegerValue not in tagged]


def auto_tag_pipes_in_view(doc, view, use_leader=False):
    untagged = get_untagged_pipes_in_view(doc, view)
    if not untagged:
        return 0
    created = 0
    with Transaction(doc, "Auto-tag pipes (safe)") as t:
        t.Start()
        for p in untagged:
            try:
                pt = _midpoint_or_center(p, view)
                IndependentTag.Create(
                    doc,
                    view.Id,
                    Reference(p),
                    use_leader,
                    TagMode.TM_ADDBY_CATEGORY,
                    TagOrientation.Horizontal,
                    pt,
                )
                created += 1
            except:
                continue
        t.Commit()
    return created


# ==================================================
# MAIN
# ==================================================
gathered_elements, polygon = select_boundary_and_gather()
if not gathered_elements:
    MessageBox.Show("No elements were gathered. Cancelled.", "Error")
    sys.exit("Cancelled")

filtered = filter_relevant_elements(gathered_elements)
if not filtered:
    MessageBox.Show("No relevant elements found in region.", "Error")
    sys.exit("Cancelled")

result = show_element_editor(filtered, region_elements=gathered_elements)
if result is None:
    sys.exit("Cancelled")

uidoc.Selection.SetElementIds(List[ElementId]())

# Force fittings to base
baseCode = result["TextNote"]
for eData in result["Elements"]:
    if eData["Category"] == "Pipe Fittings":
        eData["NewCode"] = baseCode

# Renumber pipes if no placed text note
if not result.get("TextNotePlaced", False):
    base_raw = (result.get("TextNote", "") or "").strip()
    m = re.search(r"([\d\.]+)", base_raw)
    base = m.group(1) if m else "0"
    pipe_entries = []
    for idx, eData in enumerate(result["Elements"]):
        if eData["Category"] == "Pipes":
            elem = doc.GetElement(ElementId(int(str(eData["Id"]))))
            if elem:
                bb = elem.get_BoundingBox(uidoc.ActiveView)
                if bb:
                    center = XYZ(
                        (bb.Min.X + bb.Max.X) / 2.0,
                        (bb.Min.Y + bb.Max.Y) / 2.0,
                        (bb.Min.Z + bb.Max.Z) / 2.0,
                    )
                    pipe_entries.append((idx, center))
    pipe_entries.sort(key=lambda x: (x[1].X, x[1].Y))
    ctr = 1
    for i, _ in pipe_entries:
        result["Elements"][i]["NewCode"] = base + "." + str(ctr)
        ctr += 1
    for eData in result["Elements"]:
        if eData["Category"] == "Pipe Fittings":
            eData["NewCode"] = base

# Update Comments
with Transaction(doc, "Update Comments (safe)") as t:
    t.Start()
    for eData in result["Elements"]:
        id_val = eData.get("Id")
        try:
            eid = int(str(id_val))
        except:
            continue
        elem = doc.GetElement(ElementId(eid))
        if not elem:
            continue
        p = elem.LookupParameter("Comments")
        if not p or p.IsReadOnly:
            continue
        if eData["Category"] == "Pipe Fittings":
            p.Set(result["TextNote"])
        else:
            p.Set(str(eData["NewCode"]))
    t.Commit()

# -------- Create cropped duplicate view (no detailing) --------
orig = uidoc.ActiveView
if orig.ViewType != ViewType.FloorPlan:
    MessageBox.Show("Active view is not a Floor Plan!", "Error")
    sys.exit()

# level Z for 3D box if ever needed
if isinstance(orig, ViewPlan) and orig.GenLevel:
    level_z = doc.GetElement(orig.GenLevel.Id).Elevation
else:
    level_z = 0.0

m = re.search(r"([\d\.]+)", result["TextNote"])
base = m.group(1) if m else (result["TextNote"] or "").strip()

pminx, pminy, pmaxx, pmaxy = polygon_bounds_xy(polygon)
orig_bb = orig.CropBox

region_min = XYZ(min(pminx, pmaxx), min(pminy, pmaxy), orig_bb.Min.Z)
region_max = XYZ(max(pminx, pmaxx), max(pminy, pmaxy), orig_bb.Max.Z)

dup_mode = (
    ViewDuplicateOption.WithDetailing
    if SAFE_CFG["duplicate_with_detailing"]
    else ViewDuplicateOption.Duplicate
)

with Transaction(doc, "Create Cropped Plan View (safe)") as tx:
    tx.Start()
    new_id = orig.Duplicate(dup_mode)
    new_view = doc.GetElement(new_id)

    # Detach from template and set light discipline/scale
    new_view.ViewTemplateId = ElementId.InvalidElementId
    new_view.Scale = SAFE_CFG["view_scale"]
    try:
        new_view.Discipline = getattr(ViewDiscipline, SAFE_CFG["view_discipline"])
    except:
        pass

    try:
        new_view.Name = base
    except ArgumentException:
        new_view.Name = "{} ({})".format(base, new_view.Id.IntegerValue)

    # Apply identity crop (no transform reuse)
    bb = BoundingBoxXYZ()
    bb.Min = region_min
    bb.Max = region_max
    new_view.CropBoxActive = True
    new_view.CropBoxVisible = False  # always off in safe-mode
    # Annotation crop OFF in safe-mode
    annoParam = new_view.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE)
    if annoParam and not annoParam.IsReadOnly:
        annoParam.Set(1 if SAFE_CFG["annotation_crop_on"] else 0)

    new_view.CropBox = bb

    # Tiny text note (optional)
    if SAFE_CFG["place_text_note"]:
        try:
            nt = FilteredElementCollector(doc).OfClass(TextNoteType).FirstElement()
            if nt:
                note_pt = XYZ(pminx + XY_PAD, pminy + XY_PAD, 0)
                TextNote.Create(
                    doc,
                    new_view.Id,
                    note_pt,
                    result.get("TextNote", ""),
                    TextNoteOptions(nt.Id),
                )
        except:
            pass

    # Hide crop region controls if wanted
    if SAFE_CFG["hide_crop_region_controls"]:
        visParam = new_view.get_Parameter(BuiltInParameter.VIEWER_CROP_REGION_VISIBLE)
        if visParam and not visParam.IsReadOnly:
            visParam.Set(0)

    tx.Commit()

# Tag untagged pipes (leaderless, one transaction)
if SAFE_CFG["autotag_pipes"]:
    try:
        auto_tag_pipes_in_view(doc, new_view, use_leader=SAFE_CFG["tag_use_leader"])
    except:
        pass

# -------- Title-block picker & sheet creation (simple) --------
# collect title-blocks
all_tbs = list(
    FilteredElementCollector(doc)
    .OfCategory(BuiltInCategory.OST_TitleBlocks)
    .OfClass(FamilySymbol)
    .ToElements()
)

if not all_tbs:
    MessageBox.Show("No title-block types found.", "Error")
    sys.exit()


class TBPicker(Form):
    def __init__(self, tbs):
        self.tbs = tbs
        self.Text = "Choose a Title-Block"
        self.ClientSize = Size(340, 360)
        self.lb = ListBox()
        self.lb.Bounds = Rectangle(10, 10, 320, 300)
        for sym in tbs:
            fam = sym.Family.Name if sym.Family else ""
            type_name = (
                sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() or ""
            )
            self.lb.Items.Add(fam + " - " + type_name)
        self.Controls.Add(self.lb)
        ok = Button(Text="OK", DialogResult=DialogResult.OK, Location=Point(10, 320))
        ca = Button(
            Text="Cancel", DialogResult=DialogResult.Cancel, Location=Point(100, 320)
        )
        self.Controls.Add(ok)
        self.Controls.Add(ca)
        self.AcceptButton = ok
        self.CancelButton = ca


existing_sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
existing_numbers = {s.SheetNumber for s in existing_sheets}

picker = TBPicker(all_tbs)
if picker.ShowDialog() != DialogResult.OK or picker.lb.SelectedIndex < 0:
    MessageBox.Show("Sheet creation cancelled.", "Info")
    sys.exit()

title_block = all_tbs[picker.lb.SelectedIndex]
sheet_code = base
if sheet_code in existing_numbers:
    sheet_code = "{} ({})".format(base, new_view.Id.IntegerValue)

# Ensure TB active
try:
    if hasattr(title_block, "IsActive") and (not title_block.IsActive):
        with Transaction(doc, "Activate title block") as t_act:
            t_act.Start()
            title_block.Activate()
            t_act.Commit()
except:
    pass

with Transaction(doc, "Create sheet & place plan (safe)") as t3:
    t3.Start()
    sheet = ViewSheet.Create(doc, title_block.Id)
    sheet.SheetNumber = sheet_code
    sheet.Name = "Prefab " + sheet_code

    # center-ish placement
    if not Viewport.CanAddViewToSheet(doc, sheet.Id, new_view.Id):
        raise Exception("Plan view cannot be placed on this sheet.")
    center = XYZ(
        (sheet.Outline.Min.U + sheet.Outline.Max.U) / 2.0,
        (sheet.Outline.Min.V + sheet.Outline.Max.V) / 2.0,
        0,
    )
    Viewport.Create(doc, sheet.Id, new_view.Id, center)

    # schedules & 3D intentionally skipped in safe-mode
    t3.Commit()

MessageBox.Show(
    "Prefab sheet '{}' created and plan placed.".format(sheet.SheetNumber), "Done"
)
