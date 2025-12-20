# Geberit

# Smart Piping Sheet Generator for Revit (pyRevit)

Select a prefab region → auto-number pipes/tags → generate a cropped plan + sheet + 3D view + filtered schedules — in one run.

Smart Piping Sheet Generator is a pyRevit automation tool for Revit MEP that helps you create consistent “prefab / piping package” sheets quickly and repeatably. It reduces manual renaming, tagging, view duplication/cropping, and schedule setup.

---

## What it does

1. **Boundary-driven selection**

   - You select detail lines that form a **closed boundary** in the active view.
   - The tool gathers all elements whose bounding-box center falls inside that region.

2. **Collects relevant piping annotation/model elements**

   - Pipes
   - Pipe Fittings
   - Pipe Tags
   - Text Notes

3. **Editable grid (review + adjust)**

   - Shows key info per element (diameter, length, warnings, etc.)
   - Lets you edit the target code (“New Code”) before applying.

4. **Auto-fill “prefab” codes**

   - Uses your **Text Note code** as the base (example: `prefab 5.5.5` → base `5.5.5`)
   - Pipe fittings get the base code: `5.5.5`
   - Pipes + pipe tags get sequential codes: `5.5.5.1`, `5.5.5.2`, …

5. **Writes results into the `Comments` parameter**

   - Pipes / tags: `base.index`
   - Fittings: `base`

6. **Generates deliverables**

   - Duplicates the current plan view **with detailing**
   - Applies a crop matching the selected region
   - Sets discipline/scale and names the new view using the base code
   - Creates a sheet (you pick the title block)
   - Places:
     - the cropped plan view
     - an isometric 3D view (with a section box matching the region)
     - schedules duplicated from master schedules and filtered by `Comments contains <sheet_code>`

7. **Utility actions (in the grid UI)**
   - Add/remove pipe tags
   - Place a text note at the region corner (if missing)
   - Fix reducers / toggle fitting parameters (project/family dependent)

---

## Requirements

- Revit (Floor Plan view is required)
- pyRevit installed
- A project template that uses:
  - the **Comments** parameter for prefab codes (or adapted by you)
  - schedule templates (optional; see “Configuration”)

---

## Installation (pyRevit)

1. Download the ZIP.
2. Copy the `.extension` folder into your pyRevit extensions directory.
3. Reload pyRevit / restart Revit.
4. Run **Smart Piping Sheet Generator** from the toolbar.

Tip: Hold **ALT + click** the button to open the source folder (pyRevit feature).

---

## Typical workflow

1. Go to a **Floor Plan** view.
2. Draw/prepare a closed boundary with **detail lines** around a prefab region.
3. Run the tool and select the boundary lines.
4. In the grid:
   - Enter your base code (example: `prefab 5.5.5`)
   - Click **Auto-Fill Tag Codes**
   - Optionally add/remove tags or place a missing text note
5. Click **OK** to apply codes and generate:
   - a cropped plan view named with the base code
   - a sheet numbered with the base code
   - a 3D view and filtered schedules placed on the sheet

---

## Configuration (important)

This tool can be adapted to your standards:

- Schedule names (master schedules to duplicate)
- View template names
- Parameter names used by specific family libraries
- Sheet layout placement

If your template differs, you can:

- adjust the configuration (recommended)
- or request a tailored version for your company

---

## Compatibility

**Works with Revit MEP piping workflows.**  
Supports vendor-specific workflows (e.g., "Company X" family/schedule naming) via configurable mappings.

---

## Limitations / notes

- The active view must be a **Floor Plan**.
- Boundary must be a **closed loop** (detail lines).
- The script currently expects certain schedule/view template names (can be customized).
- Writes to the **Comments** parameter — ensure your BIM standards allow this.

---

## Support / customization

If you want:

- company-specific schedule templates
- different numbering logic
- different parameters (e.g., shared parameters)
- a cleaner UI / additional QA checks
- additional sheet layouts

Contact: **BIMCode Solutions** or DM me on LinkedIn
Custom work and feature extensions are available as a paid add-on.

---

## License

Single-user license by default (unless a company license is purchased).  
Do not redistribute without permission.

---

## Changelog

- v1.0 (2025-04-12): Initial release

## Roadmap

- v1.1: Config file (schedule names, view template name, text-note defaults)
- v1.2: Dry-run mode + export report (CSV)
- v1.3: Better boundary handling + element inclusion options (categories on/off)
- v1.4: Presets per company template (one-click setup)
