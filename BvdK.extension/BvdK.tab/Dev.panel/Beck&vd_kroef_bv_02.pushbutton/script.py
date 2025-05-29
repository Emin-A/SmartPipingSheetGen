# -*- coding: utf-8 -*-
__title__ = "AutoFix\nReducers"
__doc__ = """Version = 1.0
Date    = 23.05.2025
________________________________________________________________
Description:

This integrated script lets you:
  1. Fix reducer orientations and parameters automatically.
  2. Enforce family/type rulse based on pipe diameter threshold.
  3. Replace long fittings with short variants.
  4. Control eccentricity, length options, and elbow configurations dynamically.
________________________________________________________________
How-To:
Button Behavior: When clicked, the script:
1. Scans all visible elements (in current 2D/3D view);
2. Identifies:
  - Main pipes Family: System Family: Pipe Types Type: 'NLRS_52_PI_PE buis (OD)_geb' (>= 160 mm)
  - Side pipes -||- (>= 125 mm)
  - Connected T-fittings Family:'NLRS_52_PIF_UN_PE multi T-stuk_geb' Type: 'Liggend - Var. DN/OD' (>= 125 mm)
  - Vertical pipes 'NLRS_52_PI_PE buis (OD)_geb' (>= 110 mm)
  - Vertical elbows Family: 'NLRS_52_PID_UN_PE multibocht_geb' Type: 'Var. DN/OD'
  - Reducers Family: 'NLRS_52_PIF_UN_PE multireducer_geb' Type: Var. DN/OD
3. Applies parameter toggles based on:
  - Pipe diameter
  - Direction in plan (relative angle vector between pipes)
  - Elevation (middle elevation = 0.0)
________________________________________________________________
Author: Emin Avdovic"""
# Imports

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Plumbing import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.UI import *
from Autodesk.Revit.Exceptions import *
from Autodesk.Revit.DB.Structure import *
from Autodesk.Revit.Attributes import *
from Autodesk.Revit.Exceptions import ArgumentException
from System.Collections.Generic import List
from pyrevit import revit, DB, script

import math
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
from System.Windows.Forms import MessageBox

# Revit Document Setup

app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView


def is_pipe_of_type(pipe, name, min_diam_mm):
    return name in pipe.Name and pipe.Diameter * 304.8 >= min_diam_mm


def get_direction_vector(pipe):
    try:
        curve = pipe.Location.Curve
        return (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize()
    except:
        return None


def get_branch_pipe(fitting, main_pipe_id):
    try:
        for c in fitting.MEPModel.ConnectorManager.Connectors:
            for r in c.AllRefs:
                other = r.Owner
                if isinstance(other, Pipe) and other.Id != main_pipe_id:
                    return other
    except:
        return None


def set_yesno_param(elem, param_name, on=True):
    if param_name == "reducer_eccentric" and not on:
        return
    p = elem.LookupParameter(param_name)
    if p and p.StorageType == StorageType.Integer:
        try:
            p.Set(1 if on else 0)
        except:
            pass


def vectors_are_aligned(vec1, vec2):
    if vec1 is None or vec2 is None:
        return False
    dot = vec1.Normalize().DotProduct(vec2.Normalize())
    return dot >= 0.5  # ~60 degrees or less difference


def try_update_fitting(fitting, param_map, flip=False):
    t = Transaction(doc, "Update Fitting")
    t.Start()
    try:
        for name in param_map:
            set_yesno_param(fitting, name, param_map[name])
        try:
            if flip and hasattr(fitting, "CanFlipHand") and fitting.CanFlipHand:
                fitting.FlipHand()
        except:
            pass
        t.Commit()
        return True
    except:
        t.RollBack()
        return False


def is_reducer_fully_connected(fitting):
    try:
        connectors = fitting.MEPModel.ConnectorManager.Connectors
        connected = [
            r for c in connectors for r in c.AllRefs if r.Owner.Id != fitting.Id
        ]
        return len(connected) >= 2
    except:
        return False


# === CORE FUNCTION ===


def auto_fix():
    pipes = FilteredElementCollector(doc).OfClass(Pipe).ToElements()
    fittings = FilteredElementCollector(doc).OfClass(FamilyInstance).ToElements()
    visited = set()
    updated = 0
    skipped = 0

    tg = TransactionGroup(doc, "Safe Reducer Update")
    tg.Start()

    for pipe in pipes:
        if not is_pipe_of_type(pipe, "NLRS_52_PI_PE buis", 160):
            continue

        main_dir = get_direction_vector(pipe)

        for conn in pipe.ConnectorManager.Connectors:
            for ref in conn.AllRefs:
                other_id = ref.Owner.Id
                if other_id.IntegerValue in visited:
                    continue
                visited.add(other_id.IntegerValue)
                try:
                    other = doc.GetElement(other_id)
                    if other is None or not other.IsValidObject:
                        continue
                    if not isinstance(other, FamilyInstance):
                        continue
                    # Access symbol only if still valid
                    symbol = other.Symbol
                    if symbol is None or not symbol.IsValidObject:
                        continue
                    family = symbol.Family
                    if family is None or not family.IsValidObject:
                        continue
                    family_name = family.Name
                    if not family_name:
                        continue
                except:
                    continue
                if not family_name.startswith("NLRS_52_PIF_UN_PE multi T-stuk"):
                    continue

                branch_pipe = get_branch_pipe(other, pipe.Id)
                branch_dir = get_direction_vector(branch_pipe)

                aligned = vectors_are_aligned(main_dir, branch_dir)

                param_map = {
                    "kort_verloop (kleinste)": True,
                    "kort_verloop (grootste)": True,
                    "reducer_eccentric": True,
                    "switch_excentriciteit": not aligned,  # ON if opposite to main flow
                }

                if try_update_fitting(other, param_map, flip=True):
                    updated += 1
                else:
                    skipped += 1

    # Elbows and reducers
    for f in fittings:
        try:
            if f is None or not f.IsValidObject:
                continue
            symbol = f.Symbol
            if symbol is None or not symbol.IsValidObject:
                continue
            family = symbol.Family
            if family is None or not family.IsValidObject:
                continue
            name = family.Name.lower()
        except:
            continue

        if "multibocht" in name:
            param_map = {"2x45°": False, "buis_invogen": False}
            if try_update_fitting(f, param_map):
                updated += 1
            else:
                skipped += 1

        elif "multireducer" in name:
            try:
                if f is None or not f.IsValidObject or is_reducer_fully_connected(f):
                    skipped += 1
                    continue
            except:
                skipped += 1
                continue
            param_map = {
                "kort_verloop (kleinste)": False,
                "kort_verloop (grootste)": False,
                "switch_excentriciteit": False,
                "reducer_eccentric": False,
            }
            if try_update_fitting(f, param_map):
                updated += 1
            else:
                skipped += 1

    tg.Assimilate()
    MessageBox.Show(
        "✅ Finished\nUpdated: " + str(updated) + "\nSkipped: " + str(skipped),
        "AutoFix Reducers",
    )


auto_fix()
