"""
SolidWorks Export & Reference Geometry Operations
-------------------------------------------------
Document export (STL/STEP/IGES/PDF) and reference planes.

Version: 4.1.0 (SolidWorks 2027 API groundwork)
Author: Samsaam Ali Baig

API grounding (verified against the official SolidWorks 2027 API help):
- Export: IModelDocExtension.SaveAs / SaveAs3 (current in 2027;
  IModelDoc2.SaveAs3 is obsolete). "The file extension indicates any
  conversion that should be performed (for example, Part1.igs to save
  in IGES format)." Options/Version 0 follows the repo's proven
  save_document pattern. Trap (third-party bridge, SW 2025): SaveAs
  exports only the selected bodies if a selection is active — the
  selection is cleared first.
- Reference plane: IFeatureManager.InsertRefPlane (current in 2027;
  ICreatePlaneAtSurface3 is obsolete). Constraints are
  swRefPlaneReferenceConstraints_e; Distance = 8 is corroborated by
  the repo's own working code. Other constraint values are per
  secondary sources — see the docstring.
"""

import logging
import os
import traceback
from typing import Dict

import win32com.client
import pythoncom

from ..constants import SwErrors

logger = logging.getLogger(__name__)


class ExportOperations:
    """
    Mixin class for export and reference-geometry operations.

    Requires parent class to have:
    - get_active_doc(): Document access method
    - _result(): Result factory method
    - _units: UnitConverter instance
    """

    # File extension -> format label (the extension drives the conversion)
    EXPORT_FORMATS = {
        "stl": "STL",
        "step": "STEP",
        "stp": "STEP",
        "iges": "IGES",
        "igs": "IGES",
        "pdf": "PDF",
        "x_t": "Parasolid",
        "x_b": "Parasolid",
    }

    def export_document(self, filepath: str, format: str = "") -> Dict:
        """
        Export the active document to another format.
        API: IModelDocExtension.SaveAs (current in SolidWorks 2027).

        The file extension drives the format: .stl, .step/.stp, .iges/.igs,
        .pdf, .x_t/.x_b. If `format` is given it is appended as the
        extension when the filepath has none.

        Args:
            filepath: Full output path (extension drives the format)
            format: Optional format override ("stl", "step", "iges",
                "pdf", "x_t") used only when filepath has no extension

        Returns:
            Result dictionary
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err

            # Resolve the extension
            _, ext = os.path.splitext(filepath)
            ext = ext.lstrip(".").lower()
            if not ext and format:
                ext = format.lower().lstrip(".")
                filepath = f"{filepath}.{ext}"
            if ext not in self.EXPORT_FORMATS:
                return self._result(False,
                    f"Unsupported export format '.{ext}'. Use one of: "
                    f"{sorted(self.EXPORT_FORMATS)}",
                    SwErrors.swInvalidInput)

            # SaveAs exports only the selected bodies if a selection is
            # active — clear it first.
            try:
                doc.ClearSelection2(True)
            except Exception:
                pass

            dir_path = os.path.dirname(filepath)
            if dir_path and not os.path.exists(dir_path):
                os.makedirs(dir_path)

            empty_export = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
            errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
            warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

            result = doc.Extension.SaveAs(filepath, 0, 0, empty_export,
                                          errors, warnings)
            if not result:
                return self._result(False,
                    f"Export to {filepath} failed (error code "
                    f"{errors.value})",
                    SwErrors.swExportError)

            return self._result(True,
                f"Exported: {filepath} [{self.EXPORT_FORMATS[ext]}]",
                SwErrors.swSuccess,
                {"path": filepath, "format": self.EXPORT_FORMATS[ext],
                 "api_method": "Extension.SaveAs"})

        except Exception as e:
            logger.error(f"Export error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swExportError)

    def create_reference_plane(self, offset: float,
                               plane_name: str = "Front Plane",
                               unit: str = None) -> Dict:
        """
        Create a reference plane offset from an existing plane.
        API: IFeatureManager.InsertRefPlane (current in SolidWorks 2027).

        Args:
            offset: Offset distance from the reference plane
            plane_name: Name of the plane to offset from
                (e.g. "Front Plane", "Top Plane", "Right Plane")
            unit: Unit for the offset

        Returns:
            Result dictionary

        Note: the Distance constraint value 8 is corroborated by the
        repo's own working code; other swRefPlaneReferenceConstraints_e
        values (Coincident=4, Angle=16, MidPlane=128) are per secondary
        sources and are not used by this tool.
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err

            offset_m = self._units.to_meters(offset, unit)
            unit_str = unit or self._units.default_unit.value

            doc.ClearSelection2(True)
            empty_callout = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
            try:
                selected = doc.Extension.SelectByID2(
                    plane_name, "PLANE", 0, 0, 0, False, 0, empty_callout, 0)
            except Exception:
                selected = False
            if not selected:
                try:
                    selected = doc.Extension.SelectByID2(
                        plane_name, "PLANE", 0, 0, 0, False, 0,
                        pythoncom.Nothing, 0)
                except Exception:
                    selected = False
            if not selected:
                return self._result(False,
                    f"Could not select plane '{plane_name}'",
                    SwErrors.swSelectionError)

            # Constraint 8 = Distance; unused constraints are 0.
            ref_plane = doc.FeatureManager.InsertRefPlane(
                8, offset_m,
                0, 0.0,
                0, 0.0
            )

            if ref_plane is None:
                return self._result(False,
                    f"Reference plane creation failed for '{plane_name}'",
                    SwErrors.swFeatureError)

            try:
                plane_name_out = ref_plane.Name
            except Exception:
                plane_name_out = ""

            return self._result(True,
                f"Reference plane offset {offset}{unit_str} from "
                f"'{plane_name}' [InsertRefPlane]",
                SwErrors.swSuccess,
                {"offset": offset, "unit": unit_str,
                 "from_plane": plane_name, "plane_name": plane_name_out,
                 "api_method": "InsertRefPlane"})

        except Exception as e:
            logger.error(f"Reference plane error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swFeatureError)
